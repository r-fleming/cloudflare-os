// A gadget that defines `connectCaller(caller)` is reached only through the session it returns
// for each caller: every stub the overseer mints to a gadget facet names who will hold it (see
// GadgetCaller and facetRpcStub in overseer.ts), and none can reach the gadget's own methods or
// call connectCaller itself. A gadget without the method is called directly, as before.
//
// Runs real gadgets, loaded from real commits, inside a real OverseerDurableObject.

import { describe, expect, it } from "vitest";
import { env } from "cloudflare:workers";
import { runInDurableObject } from "cloudflare:test";
import type { OverseerDurableObject } from "../src/overseer.js";

declare module "cloudflare:workers" {
  interface ProvidedEnv {
    TEST_OVERSEER: DurableObjectNamespace<OverseerDurableObject>;
  }
}

const GATED = 1;
const PLAIN = 2;

const GATED_JS = `
import { DurableObject, RpcTarget } from "cloudflare:workers";

export class Gadget extends DurableObject {
  disposed = 0;

  // Reachable only if the overseer let a caller past its session.
  rawOnly() {
    return "raw";
  }

  connectCaller(caller) {
    if (caller.kind === "user" && caller.name === "Mallory") {
      throw new Error("Refused: " + caller.name);
    }
    return new Session(this, caller);
  }
}

class Session extends RpcTarget {
  constructor(gadget, caller) {
    super();
    this.gadget = gadget;
    this.caller = caller;
  }
  whoami() {
    return this.caller;
  }
  disposedCount() {
    return this.gadget.disposed;
  }
  [Symbol.dispose]() {
    this.gadget.disposed++;
  }
}
`;

const PLAIN_JS = `
import { DurableObject } from "cloudflare:workers";

export class Gadget extends DurableObject {
  hello() {
    return "hi";
  }
}
`;

const alice = { kind: "user", id: "alice@example.com", name: "Alice", role: "use" } as const;
const mallory = { kind: "user", id: "mallory@example.com", name: "Mallory", role: "use" } as const;
const system = { kind: "system", reason: "test" } as const;

let doCounter = 0;

async function withGadgets(fn: (impl: any) => Promise<void>): Promise<void> {
  let stub = env.TEST_OVERSEER.getByName(`gadget-caller-${++doCounter}`);
  await runInDurableObject(stub, async (instance: OverseerDurableObject) => {
    let impl = (instance as unknown as { impl: any }).impl;
    for (let [id, code] of [[GATED, GATED_JS], [PLAIN, PLAIN_JS]] as const) {
      let commitId = await impl.gitCache.writeFilesAsCommit(new Map([["server.js", code]]), {
        parents: [],
        author: { name: "Alice", email: "alice@example.com" },
        message: "test commit",
        timestamp: new Date(1700000000_000),
      });
      impl.storage.gadgets.put({
        type: "gadget", id, title: `G${id}`, created: new Date(0), bindingName: `G${id}`,
        bindings: {}, commitId,
      });
    }
    await fn(impl);
  });
}

// Await a rejection with a single handler, returning its message: expect(...).rejects forks the
// underlying JsRpcPromise, and the leftover copy is reported as an unhandled rejection.
async function rejection(call: Promise<unknown>): Promise<string> {
  try {
    await call;
  } catch (err) {
    return String(err instanceof Error ? err.message : err);
  }
  throw new Error("expected a rejection");
}

describe("connectCaller", () => {
  it("routes a person's calls to their session, never to the gadget", () =>
      withGadgets(async impl => {
    using gadget = await impl.getGadgetFacet(GATED, alice);
    expect(await gadget.whoami()).toEqual(alice);
    // The session has no rawOnly(), and the gadget's is out of reach.
    expect(await rejection(gadget.rawOnly())).toMatch(/does not implement/);
  }));

  it("can't be called through any stub", () => withGadgets(async impl => {
    using gated = await impl.getGadgetFacet(GATED, alice);
    expect(await rejection(gated.connectCaller({ ...alice, role: "owner" })))
        .toMatch(/reserved for the Workshop/);
    using plain = await impl.getGadgetFacet(PLAIN, alice);
    expect(await rejection(plain.connectCaller(alice))).toMatch(/reserved for the Workshop/);
  }));

  it("opens one session per stub, each for its own caller", () => withGadgets(async impl => {
    let bob = { ...alice, id: "bob@example.com", name: "Bob", role: "owner" } as const;
    using a = await impl.getGadgetFacet(GATED, alice);
    using b = await impl.getGadgetFacet(GATED, bob);
    expect(await a.whoami()).toEqual(alice);
    expect(await b.whoami()).toEqual(bob);
    expect(await a.whoami()).toEqual(alice);
  }));

  it("looks a deferred caller up once, and only when the gadget is first called", () =>
      withGadgets(async impl => {
    let lookups = 0;
    let source = async () => { lookups++; return alice; };
    using gadget = await impl.getGadgetFacet(GATED, source);
    expect(lookups).toBe(0);
    await gadget.whoami();
    await gadget.whoami();
    expect(lookups).toBe(1);
  }));

  it("fails only the refused caller's calls", () => withGadgets(async impl => {
    using refused = await impl.getGadgetFacet(GATED, mallory);
    using allowed = await impl.getGadgetFacet(GATED, alice);
    expect(await rejection(refused.whoami())).toMatch(/Refused: Mallory/);
    expect(await allowed.whoami()).toEqual(alice);
    // A refusal isn't cached as a session: it is asked again, and refused again.
    expect(await rejection(refused.whoami())).toMatch(/Refused: Mallory/);
  }));

  it("names the gadget calling through a binding, and system work as such", () =>
      withGadgets(async impl => {
    using fromGadget = await impl.startGatekeeperSession(
        { type: "gadget", id: GATED }, { from: "gadget", gadgetId: PLAIN });
    expect(await fromGadget.whoami()).toEqual({ kind: "gadget", gadgetId: PLAIN });
    using fromHook = await impl.startGatekeeperSession(
        { type: "gadget", id: GATED }, { from: "hook", gadgetId: GATED });
    expect(await fromHook.whoami()).toEqual({ kind: "system", reason: "hook" });
    expect(await rejection(fromGadget.rawOnly())).toMatch(/does not implement/);
  }));

  it("routes a snapshot's viewer through the snapshot's own connectCaller", () =>
      withGadgets(async impl => {
    let viewer = { ...alice, role: "view" } as const;
    let facetName = `snapshot-${crypto.randomUUID()}`;
    impl.storage.gadgetSnapshots.put({ facetName, gadgetId: GATED });
    using copy = impl.getSnapshotFacet({
      facetName, gadgetId: GATED, commitId: impl.getGadgetHead(GATED)!, refreshedAt: new Date(),
    }, viewer);
    expect(await copy.whoami()).toEqual(viewer);
    expect(await rejection(copy.rawOnly())).toMatch(/does not implement/);
  }));

  it("disposes the session with its stub", () => withGadgets(async impl => {
    using watcher = await impl.getGadgetFacet(GATED, system);
    expect(await watcher.disposedCount()).toBe(0);
    let gadget = await impl.getGadgetFacet(GATED, alice);
    await gadget.whoami();
    gadget[Symbol.dispose]();
    let count = 0;
    for (let i = 0; i < 50 && count === 0; i++) {
      await new Promise(resolve => setTimeout(resolve, 10));
      count = await watcher.disposedCount();
    }
    expect(count).toBe(1);
  }));

  it("leaves a gadget without connectCaller called directly", () => withGadgets(async impl => {
    using asUser = await impl.getGadgetFacet(PLAIN, alice);
    using asSystem = await impl.getGadgetFacet(PLAIN, system);
    expect(await asUser.hello()).toBe("hi");
    expect(await asSystem.hello()).toBe("hi");
  }));
});
