// connectToGadget() tells a gadget's code who is calling, through viewer() or the gadget's own
// connectViewer(), and nothing else can present a viewer.
//
// The facet tests run a real gadget, loaded from a real commit, inside a real
// OverseerDurableObject (see gadget-restore.test.ts); the wiring tests forge sessions via open().

import { describe, expect, it, vi } from "vitest";
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

// A gadget written before viewers existed.
const PLAIN_JS = `
import { DurableObject, RpcTarget } from "cloudflare:workers";
class Counter extends RpcTarget { n = 0; inc() { return ++this.n; } }
export class Gadget extends DurableObject {
  count = 0;
  hello() { return "hi"; }
  bump() { return ++this.count; }
  counter() { return new Counter(); }
  fails() { throw new Error("boom"); }
}
`;

const VIEWER_JS = `
import { DurableObject } from "cloudflare:workers";
import { viewer } from "gadgets:viewer";
export class Gadget extends DurableObject {
  who() { return viewer(); }
  async slowWho(ms) {
    await new Promise(resolve => setTimeout(resolve, ms));
    await this.ctx.storage.get("x");
    return viewer()?.id ?? null;
  }
}
`;

const SESSION_JS = `
import { DurableObject, RpcTarget } from "cloudflare:workers";
export class Gadget extends DurableObject {
  connected = new Set();
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

const ALICE: GadgetViewer = { id: "v-alice", displayName: "Alice", role: "use" };
const BOB: GadgetViewer = { id: "v-bob", displayName: "Bob", role: "use" };

let doCounter = 0;

async function withGadget(files: Record<string, string>,
    fn: (impl: any, logged: string[]) => Promise<void>): Promise<void> {
  await runInDurableObject(env.TEST_OVERSEER.getByName(`gadget-viewer-${++doCounter}`),
      async (instance: OverseerDurableObject) => {
    let impl = (instance as unknown as { impl: any }).impl;
    let logged: string[] = [];
    impl.deliverGadgetLogs = async (_chatId: number | null, events: { message: unknown[] }[]) => {
      logged.push(...events.map(event => String(event.message[0])));
    };
    let commitId = await impl.gitStore.writeFilesAsCommit(new Map(Object.entries(files)), {
      parents: [],
      author: { name: "Alice", email: "alice@example.com" },
      message: "test commit",
      timestamp: new Date(1700000000_000),
    });
    impl.storage.gadgets.put({
      type: "gadget", id: GADGET_ID, title: "G", created: new Date(0), bindingName: "G",
      bindings: {}, commitId,
    });
    await fn(impl, logged);
  });
}

// The error a native RPC call fails with. (Unlike `expect().rejects`, handling the RpcPromise's
// rejection directly keeps it from also being reported as unhandled.)
function rejection(call: Promise<unknown>): Promise<string | null> {
  return call.then(() => null, (error: unknown) => String(error));
}

describe("viewer connections", () => {
  it("reach a gadget that knows nothing of viewers unchanged",
      () => withGadget({ "server.js": PLAIN_JS }, async (impl, logged) => {
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    using facet = await impl.getGadgetFacet(GADGET_ID);
    expect(await session.hello()).toBe("hi");
    expect(await session.bump()).toBe(1);
    expect(await facet.bump()).toBe(2);
    using counter = await session.counter();
    expect(await counter.inc()).toBe(1);
    expect(await rejection(session.fails())).toContain("boom");
    expect(logged).toEqual([expect.stringMatching(/boom[\s\S]*at fails\(\)/)]);
  }));

  it("tell gadget code who made each call, and nobody for other callers",
      () => withGadget({ "server.js": VIEWER_JS }, async impl => {
    using alice = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    using bob = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, BOB);
    using facet = await impl.getGadgetFacet(GADGET_ID);  // what the agent, hooks, etc. get
    expect(await alice.who()).toEqual(ALICE);
    expect(await facet.who()).toBe(null);
    expect(await Promise.all([alice.slowWho(40), bob.slowWho(5), facet.slowWho(20)]))
        .toEqual(["v-alice", "v-bob", null]);
  }));

  it("refuse connectViewer through every stub", () => withGadget({ "server.js": VIEWER_JS },
      async impl => {
    using facet = await impl.getGadgetFacet(GADGET_ID);
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    for (let stub of [facet, session]) {
      expect(await rejection(stub.connectViewer(BOB)))
          .toContain("can only be called by the Workshop");
    }
  }));

  it("don't load a gadget that is never called", () => withGadget({ "client.js": "0" },
      async (impl, logged) => {
    using _session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    await new Promise(resolve => setTimeout(resolve, 20));
    expect(logged).toEqual([]);
  }));

  it("fail on use, not on connect, when server.js is broken",
      () => withGadget({ "server.js": "export class Gadget {" }, async impl => {
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    expect(await rejection(session.hello())).toContain("Failed to start Worker");
  }));
});

describe("a gadget's own connectViewer()", () => {
  it("opens the connection's session", () => withGadget({ "server.js": SESSION_JS }, async impl => {
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    expect(await session.whoami()).toEqual(ALICE);
  }));

  it("learns when the session ends", () => withGadget({ "server.js": SESSION_JS }, async impl => {
    using facet = await impl.getGadgetFacet(GADGET_ID);
    let session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined, ALICE);
    await session.whoami();
    expect(await facet.connectedViewers()).toEqual(["v-alice"]);
    session[Symbol.dispose]();
    await vi.waitFor(async () => expect(await facet.connectedViewers()).toEqual([]));
  }));

  it("can turn a viewer away", () => withGadget({ "server.js": SESSION_JS }, async impl => {
    using session = await impl.getGadgetFacet(GADGET_ID, undefined, undefined,
        { ...ALICE, displayName: "Refused" });
    expect(await rejection(session.whoami())).toContain("not welcome");
  }));
});

describe("viewer ids", () => {
  it("are stable per person within a workspace and opaque", () => withGadget({ "server.js": PLAIN_JS }, async impl => {
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
      await withGadget({ "server.js": PLAIN_JS }, async impl => {
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
