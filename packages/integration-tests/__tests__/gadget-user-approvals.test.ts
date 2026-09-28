// End to end over the real WebSocket API: a gadget learns who made each call from the Workshop
// (currentUser()), not from the browser. Drives the bundled Approvals demo blueprint with an owner
// and a "use" collaborator, each on their own authenticated connection.

import type { RpcStub } from "capnweb";
import { afterAll, beforeAll, expect, it } from "vitest";
import { type Harness, startHarness } from "../src/harness.js";
import { NetworkInterceptor } from "../src/network-interceptor.js";
import { connect, nextUsernames, RpcTarget, signUp, waitFor } from "../src/rpc-client.js";

const BLUEPRINT_ID = "demo.approvals";

let harness: Harness | undefined;
const network = new NetworkInterceptor();

beforeAll(async () => {
  network.install();
  harness = await startHarness({ gatekeepers: [], enableGadgetExecution: true });
});

afterAll(async () => {
  try {
    expect(network.getUnmockedCalls()).toEqual([]);
  } finally {
    network.uninstall();
    await harness?.server.close();
  }
});

function requireHarness(): Harness {
  if (harness === undefined) throw new Error("Workshop harness did not start");
  return harness;
}

type View = {
  me: { id: string; displayName: string; role: string; approver: boolean };
  requests: {
    id: string; title: string; canDecide: boolean;
    requester: { id: string; name: string };
    decision?: { outcome: string; by: { id: string; name: string } };
  }[];
};

type User = { id: string; displayName: string; role: string };

// The Approvals gadget as a connection sees it (see the blueprint's lib/protocol.ts), plus the
// Workshop's reserved dispatch method, which a connection must not be able to call.
type Session = {
  whoami(): User;
  callAsUser(secret: string, user: User, name: string, args: unknown[]): unknown;
  subscribe(listener: Listener): View;
  submit(title: string, detail: string): void;
  decide(requestId: string, outcome: string, note: string): void;
  setApprover(memberId: string, approver: boolean): void;
};

// connectToGadget() is typed for any gadget; narrow it to this one.
async function connectSession(gadget: { connectToGadget(): Promise<unknown> })
    : Promise<RpcStub<Session>> {
  return await gadget.connectToGadget() as RpcStub<Session>;
}

// What the client's listener sees: the latest view the gadget pushed to this connection.
class Listener extends RpcTarget {
  latest: View | undefined;
  update(view: View) { this.latest = view; }
}

it("attributes requests and decisions to the signed-in people, not the browser", async () => {
  const [ownerName, requesterName] = nextUsernames("owner", "requester");
  if (!ownerName || !requesterName) throw new Error("Missing test username");

  using ownerPublic = connect(requireHarness().url);
  using requesterPublic = connect(requireHarness().url);
  using owner = await signUp(ownerPublic, ownerName);
  using requester = await signUp(requesterPublic, requesterName);

  await waitFor("bundled blueprints to install", async () =>
    (await owner.listOutputFormats()).some(offer => offer.blueprintId === BLUEPRINT_ID) || null);
  using workspace = await owner.newGadgetFromBlueprint(BLUEPRINT_ID, {});
  const { id: workspaceId, defaultGadgetId } = await workspace.getMetadata();
  if (defaultGadgetId === undefined) throw new Error("Workspace has no default Gadget");
  await workspace.addCollaborator(requesterName, "use", "requester");

  using ownerGadget = await workspace.getGadget(defaultGadgetId);
  using ownerSession = await connectSession(ownerGadget);
  using requesterWorkspace = await requester.openGadget(workspaceId);
  using requesterGadget = await requesterWorkspace.getGadget(defaultGadgetId);
  using requesterSession = await connectSession(requesterGadget);

  // Each connection is its own person, as the Workshop knows them.
  const ownerMe = await ownerSession.whoami();
  const requesterMe = await requesterSession.whoami();
  expect(ownerMe).toMatchObject({ displayName: ownerName, role: "build" });
  expect(requesterMe).toMatchObject({ displayName: requesterName, role: "use" });
  expect(requesterMe.id).not.toBe(ownerMe.id);
  expect(requesterMe.id).toMatch(/^[0-9a-f]{64}$/);  // opaque, not the username

  // A browser can't act as someone else: only the Workshop may present a user.
  await expect(requesterSession.callAsUser("", ownerMe, "whoami", []))
      .rejects.toThrow("can only be called by the Workshop");

  const ownerView = new Listener();
  const requesterView = new Listener();
  await ownerSession.subscribe(ownerView);
  await requesterSession.subscribe(requesterView);

  // The requester's request is theirs, and they can't approve it.
  await requesterSession.submit("New laptop", "Mine is from 2019.");
  const request = await waitFor("the request to reach the owner", async () =>
    ownerView.latest?.requests.find(each => each.title === "New laptop") ?? null);
  expect(request.requester).toEqual({ id: requesterMe.id, name: requesterName });
  expect(request.canDecide).toBe(true);
  await expect(requesterSession.decide(request.id, "approved", ""))
      .rejects.toThrow("can't decide your own request");
  await expect(requesterSession.setApprover(requesterMe.id, true))
      .rejects.toThrow("Only builders choose approvers");

  // The owner approves it, and the requester sees who did.
  await ownerSession.decide(request.id, "approved", "Go for it.");
  const decided = await waitFor("the decision to reach the requester", async () =>
    requesterView.latest?.requests.find(each => each.id === request.id && each.decision) ?? null);
  expect(decided.decision).toEqual(expect.objectContaining({
    outcome: "approved", by: { id: ownerMe.id, name: ownerName },
  }));

  // Separation of duties cuts both ways: the owner (a builder) can't approve their own request
  // through the gadget either, but an approver they designate can. (A builder could change the
  // gadget's code to allow it; the rules bind "use" users only.)
  await ownerSession.submit("Team offsite", "");
  const ownRequest = await waitFor("the owner's request", async () =>
    ownerView.latest?.requests.find(each => each.title === "Team offsite") ?? null);
  expect(ownRequest.canDecide).toBe(false);
  await expect(ownerSession.decide(ownRequest.id, "approved", ""))
      .rejects.toThrow("can't decide your own request");
  await ownerSession.setApprover(requesterMe.id, true);
  await requesterSession.decide(ownRequest.id, "rejected", "Not this quarter.");
  const rejected = await waitFor("the requester's decision to reach the owner", async () =>
    ownerView.latest?.requests.find(each => each.id === ownRequest.id && each.decision) ?? null);
  expect(rejected.decision).toEqual(expect.objectContaining({
    outcome: "rejected", by: { id: requesterMe.id, name: requesterName },
  }));
});
