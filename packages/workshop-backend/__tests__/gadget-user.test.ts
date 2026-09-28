// Gadget code learns who made a call through currentUser(), and only the Workshop can say.
// Facet tests run real gadgets in a real OverseerDurableObject (see gadget-restore.test.ts).

import { describe, expect, it } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { OverseerDurableObject } from "../src/overseer.js";
import type { GadgetUser } from "../src/gadget-user.js";
import { openFakeOverseer } from "./fixtures.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const SERVER_JS = `
import { DurableObject, RpcTarget, WorkerEntrypoint } from "cloudflare:workers";
import { currentUser } from "gadgets:user";
class Counter extends RpcTarget { n = 0; inc() { return ++this.n; } }
export class Gadget extends DurableObject {
  bound = this.who.bind(this);
  shadowed = this.shadowed.bind(this);
  shadowed() { return "shadowed ran"; }
  who() { return currentUser(); }
  async slowWho(ms) {
    await new Promise(resolve => setTimeout(resolve, ms));
    return currentUser()?.id ?? null;
  }
  counter() { return new Counter(); }
  fails() { throw new Error("boom"); }
  alarm() { return "alarm ran"; }
}
export class ExportHandler extends WorkerEntrypoint {
  getExportFormats() {
    return [{ id: "who", label: "Who", mode: "server", contentType: "text/plain", fileExtension: ".txt" }];
  }
  async export(gadget) { return new Response((await gadget.who())?.id ?? "nobody").body; }
}
`;

const ALICE: GadgetUser = { id: "u-alice", displayName: "Alice", role: "use" };
const BOB: GadgetUser = { id: "u-bob", displayName: "Bob", role: "use" };

let doCounter = 0;

async function withGadget(files: Record<string, string>,
    fn: (impl: any, logged: string[]) => Promise<void>): Promise<void> {
  await runInDurableObject(env.TEST_OVERSEER.getByName(`gadget-user-${++doCounter}`),
      async (instance: OverseerDurableObject) => {
    let impl = (instance as unknown as { impl: any }).impl;
    let logged: string[] = [];
    impl.deliverGadgetLogs = async (_chatId: number | null, events: { message: unknown[] }[]) => {
      logged.push(...events.map(event => String(event.message[0])));
    };
    let commitId = await impl.gitStore.writeFilesAsCommit(new Map(Object.entries(files)), {
      parents: [], author: { name: "A", email: "a@example.com" }, message: "test",
      timestamp: new Date(1700000000_000),
    });
    impl.storage.gadgets.put({
      type: "gadget", id: 1, title: "G", created: new Date(0), bindingName: "G", bindings: {},
      commitId,
    });
    await fn(impl, logged);
  });
}

// Handling the RpcPromise's rejection directly keeps it from also being reported as unhandled.
function rejection(call: Promise<unknown>): Promise<string | null> {
  return call.then(() => null, (error: unknown) => String(error));
}

describe("currentUser()", () => {
  it("is the user who made each call, and null for other callers",
      () => withGadget({ "server.js": SERVER_JS }, async (impl, logged) => {
    using alice = await impl.getGadgetFacet(1, undefined, undefined, ALICE);
    using bob = await impl.getGadgetFacet(1, undefined, undefined, BOB);
    using facet = await impl.getGadgetFacet(1);  // what the agent, hooks and other gadgets get
    expect(await alice.who()).toEqual(ALICE);
    expect(await facet.who()).toBe(null);
    expect(await Promise.all([alice.slowWho(30), bob.slowWho(5), facet.slowWho(15)]))
        .toEqual(["u-alice", "u-bob", null]);

    // Otherwise calls behave as they do over plain RPC.
    using counter = await alice.counter();
    expect(await counter.inc()).toBe(1);
    expect(await rejection(alice.fails())).toContain("boom");
    expect(logged).toEqual([expect.stringMatching(/boom[\s\S]*at fails\(\)/)]);
    for (let name of ["alarm", "bound", "shadowed"]) {  // as plain RPC refuses them
      expect(await rejection(alice[name]())).toContain(`does not implement the method "${name}"`);
    }
    expect(await rejection(facet.shadowed())).toContain('does not implement the method "shadowed"');
  }));

  it("is the exporting user during an export",
      () => withGadget({ "server.js": SERVER_JS }, async impl => {
    let stream = await impl.exportGadget(1, "who", undefined, ALICE);
    expect(await new Response(stream).text()).toBe("u-alice");
  }));

  it("can only be set by the Workshop", () => withGadget({ "server.js": SERVER_JS },
      async impl => {
    using facet = await impl.getGadgetFacet(1);
    using alice = await impl.getGadgetFacet(1, undefined, undefined, ALICE);
    for (let stub of [facet, alice]) {
      expect(await rejection(stub.callAsUser(BOB, "who", [])))
          .toContain("can only be called by the Workshop");
    }
  }));

  it("leaves gadgets with no, broken or incomplete server code working as before", async () => {
    await withGadget({ "client.js": "0" }, async (impl, logged) => {
      using _alice = await impl.getGadgetFacet(1, undefined, undefined, ALICE);
      expect(logged).toEqual([]);
    });
    await withGadget({ "server.js": "export class Gadget {" }, async impl => {
      using alice = await impl.getGadgetFacet(1, undefined, undefined, ALICE);
      expect(await rejection(alice.who())).toContain("Failed to start Worker");
    });
    // No Gadget class: the rest of the worker (here, its exports) still loads.
    await withGadget({ "server.js": SERVER_JS.replace("export class Gadget", "class Gadget") },
        async impl => {
      using alice = await impl.getGadgetFacet(1, undefined, undefined, ALICE);
      expect(await rejection(alice.who())).toContain('must export a class named "Gadget"');
      expect(await impl.getGadgetExportFormats(1)).toEqual([expect.objectContaining({ id: "who" })]);
    });
  });
});

describe("user ids", () => {
  it("are stable per user within a workspace, and differ across workspaces", async () => {
    let ids: string[] = [];
    for (let i = 0; i < 2; i++) {
      await withGadget({}, async impl => {
        let alice = await impl.gadgetUser("alice@example.com", "use", "Alice");
        expect(alice).toEqual({ id: alice.id, displayName: "Alice", role: "use" });
        expect((await impl.gadgetUser("alice@example.com", "build", "A")).id).toBe(alice.id);
        expect((await impl.gadgetUser("bob@example.com", "use", "Bob")).id).not.toBe(alice.id);
        ids.push(alice.id);
      });
    }
    expect(ids[0]).not.toBe(ids[1]);
  });
});

describe("connectToGadget() presents its session's user", () => {
  it.each([["owner", "build"], ["build", "build"], ["use", "use"]] as const)(
      "as a(n) %s", async (role, expected) => {
    let presented: unknown[] = [];
    let client = await openFakeOverseer({}, {
      role: role === "use" ? "use" : "build",
      impl: {
        // The fixture opens "build" as the owner; any other owner makes it a collaborator.
        ...(role === "build" ? { ownerId: "someone-else" } : {}),
        getGadgetRecord: () => ({ type: "gadget", id: 1 }),
        recordGadgetAnalytics: () => {},
        gadgetUser: async (profileId: string, userRole: string) => ({ profileId, userRole }),
        getGadgetFacet: async (_id: number, _chat?: number, _join?: string, user?: unknown) => {
          presented.push(user);
          return {};
        },
        users: {
          idFromString: (id: string) => id,
          get: () => ({
            id: { toString: () => "user-do" },
            whoami: async () => ({ type: "user", id: "profile-id", name: "Test User" }),
            recordSharedGadgetOpen: async () => {},
          }),
        },
      },
    });
    await (await client.getGadget(1) as any).connectToGadget();
    let profileId = role === "use" ? "viewer-id-profile" : "owner-id-profile";
    expect(presented).toEqual([{ profileId, userRole: expected }]);
  });
});
