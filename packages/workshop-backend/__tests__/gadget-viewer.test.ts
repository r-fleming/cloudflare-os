// connectToGadget() tells a gadget that implements connectViewer() who is connecting, and nothing
// else can call connectViewer().
//
// The facet tests run a real gadget, loaded from a real commit, inside a real
// OverseerDurableObject (see gadget-restore.test.ts); the wiring tests forge sessions via open().

import { describe, expect, it } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { OverseerDurableObject } from "../src/overseer.js";
import type { GadgetViewer } from "../src/gadget-viewer.js";
import { openFakeOverseer } from "./fixtures.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const GADGET_ID = 1;

const VIEWER_AWARE_JS = `
import { DurableObject, RpcTarget } from "cloudflare:workers";

export class Gadget extends DurableObject {
  connected = new Set();
  hello() { return "hi"; }
  connectedViewers() { return [...this.connected]; }
  connectViewer(viewer) {
    if (viewer.displayName === "Refused") throw new Error("not welcome");
    return new Session(this, viewer);
  }
}

class Session extends RpcTarget {
  #gadget; #viewer;
  constructor(gadget, viewer) {
    super();
    this.#gadget = gadget;
    this.#viewer = viewer;
    gadget.connected.add(viewer.id);
  }
  whoami() { return this.#viewer; }
  [Symbol.dispose]() { this.#gadget.connected.delete(this.#viewer.id); }
}
`;

const PLAIN_JS = `
import { DurableObject } from "cloudflare:workers";
export class Gadget extends DurableObject { hello() { return "hi"; } }
`;

const ALICE: GadgetViewer = { id: "v-alice", displayName: "Alice", role: "use" };

let doCounter = 0;

async function withGadget(serverJs: string, fn: (impl: any) => Promise<void>,
    name = `gadget-viewer-${++doCounter}`): Promise<void> {
  await runInDurableObject(env.TEST_OVERSEER.getByName(name),
      async (instance: OverseerDurableObject) => {
    let impl = (instance as unknown as { impl: any }).impl;
    let commitId = await impl.gitStore.writeFilesAsCommit(new Map([["server.js", serverJs]]), {
      parents: [],
      author: { name: "Alice", email: "alice@example.com" },
      message: "test commit",
      timestamp: new Date(1700000000_000),
    });
    impl.storage.gadgets.put({
      type: "gadget", id: GADGET_ID, title: "G", created: new Date(0), bindingName: "G",
      bindings: {}, commitId,
    });
    await fn(impl);
  });
}

describe("the viewer handshake", () => {
  it("hands a viewer-aware gadget's session to the connection",
      () => withGadget(VIEWER_AWARE_JS, async impl => {
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    expect(await session.whoami()).toEqual(ALICE);
  }));

  it("tells the gadget when the connection's session ends",
      () => withGadget(VIEWER_AWARE_JS, async impl => {
    using facet = await impl.getGadgetFacet(GADGET_ID);
    let session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    expect(await facet.connectedViewers()).toEqual(["v-alice"]);

    session[Symbol.dispose]();
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(await facet.connectedViewers()).toEqual([]);
  }));

  it("is refused through every stub it hands out", () => withGadget(VIEWER_AWARE_JS, async impl => {
    // What the agent, other gadgets, hooks and exports receive: no viewer.
    using facet = await impl.getGadgetFacet(GADGET_ID);
    expect(await facet.hello()).toBe("hi");
    await expect(facet.connectViewer({ ...ALICE, id: "v-mallory" }))
        .rejects.toThrow("can only be called by the Workshop");

    // Nor can a connected viewer open a second session as someone else.
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    await expect(session.connectViewer({ ...ALICE, id: "v-mallory" }))
        .rejects.toThrow("can only be called by the Workshop");
  }));

  it("connects a gadget without the handshake to the gadget itself",
      () => withGadget(PLAIN_JS, async impl => {
    using facet = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    expect(await facet.hello()).toBe("hi");
  }));

  it("does not fall back to the gadget itself when the handshake turns a viewer away",
      () => withGadget(VIEWER_AWARE_JS, async impl => {
    let logged: unknown[] = [];
    impl.deliverGadgetLogs = async (_chatId: number | null, events: { message: unknown[] }[]) => {
      logged.push(...events.map(event => String(event.message[0])));
    };
    await expect(impl.getGadgetFacet(GADGET_ID, undefined, undefined,
        { ...ALICE, displayName: "Refused" })).rejects.toThrow("not welcome");
    // ...and the gadget's author sees why in the console, as for any other gadget method error.
    expect(logged).toEqual([expect.stringMatching(/not welcome[\s\S]*at connectViewer\(\)/)]);
  }));
});

describe("viewer ids", () => {
  it("are stable per person within a workspace and opaque", () => withGadget(PLAIN_JS, async impl => {
    let a1 = await impl.gadgetViewer("alice@example.com", "use", "Alice");
    let a2 = await impl.gadgetViewer("alice@example.com", "build", "Alice B.");
    let b = await impl.gadgetViewer("bob@example.com", "use", "Bob");
    expect(a1).toEqual({ id: a1.id, displayName: "Alice", role: "use" });
    expect(a2.id).toBe(a1.id);
    expect(b.id).not.toBe(a1.id);
    expect(a1.id).not.toContain("alice");
  }));

  it("differ across workspaces", async () => {
    let ids: string[] = [];
    for (let i = 0; i < 2; i++) {
      await withGadget(PLAIN_JS, async impl => {
        ids.push((await impl.gadgetViewer("alice@example.com", "use", "Alice")).id);
      });
    }
    expect(ids[0]).not.toBe(ids[1]);
  });
});

// Each session presents its own profile, with the role it was admitted with ("build" for the owner).
describe("connectToGadget() presents the session's viewer", () => {
  async function connectAs(role: "owner" | "build" | "use") {
    let presented: unknown[] = [];
    let client = await openFakeOverseer({}, {
      role: role === "use" ? "use" : "build",
      impl: {
        // The fixture opens "build" as the owner; any other owner makes it a collaborator.
        ...(role === "build" ? { ownerId: "someone-else" } : {}),
        getGadgetRecord: () => ({ type: "gadget", id: GADGET_ID }),
        recordGadgetAnalytics: () => {},
        gadgetViewer: async (profileId: string, viewerRole: string, displayName: string) =>
            ({ profileId, role: viewerRole, displayName }),
        getGadgetFacet: async (_id: number, _chatId?: number, _joinAs?: string,
            viewer?: unknown) => {
          presented.push(viewer);
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
    let gadget = await client.getGadget(GADGET_ID);
    await (gadget as any).connectToGadget();
    return presented;
  }

  it("as the owner", async () => {
    expect(await connectAs("owner")).toEqual(
        [{ profileId: "owner-id-profile", role: "build", displayName: "Test User" }]);
  });

  it("as a build collaborator", async () => {
    expect(await connectAs("build")).toEqual(
        [{ profileId: "owner-id-profile", role: "build", displayName: "Test User" }]);
  });

  it("as a use collaborator", async () => {
    expect(await connectAs("use")).toEqual(
        [{ profileId: "viewer-id-profile", role: "use", displayName: "Test User" }]);
  });
});
