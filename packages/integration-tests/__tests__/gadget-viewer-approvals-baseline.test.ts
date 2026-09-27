// The contrast to gadget-viewer-approvals.test.ts on the viewer-identity branch: on a Workshop that
// doesn't introduce viewers, the Approvals demo can only take the browser's word for who it is.

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

type Viewer = { id: string; displayName: string; role: string };

// The Approvals session as a connection sees it (see the blueprint's lib/protocol.ts), plus the
// handshake, which on this Workshop the connection can call itself.
type Session = {
  whoami(): Viewer;
  connectViewer(viewer: Viewer): Session;
  subscribe(listener: Listener): View;
  submit(title: string, detail: string): void;
  decide(requestId: string, outcome: string, note: string): void;
  setApprover(memberId: string, approver: boolean): void;
};

// connectToGadget() is typed for any gadget; narrow it to this one's session.
async function connectSession(gadget: { connectToGadget(): Promise<unknown> })
    : Promise<RpcStub<Session>> {
  return await gadget.connectToGadget() as RpcStub<Session>;
}

// What the client's listener sees: the latest view the gadget pushed to this connection.
class Listener extends RpcTarget {
  latest: View | undefined;
  update(view: View) { this.latest = view; }
}

it("without viewer identity, a use collaborator can claim to be a builder and approve themselves", async () => {
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

  using requesterWorkspace = await requester.openGadget(workspaceId);
  using requesterGadget = await requesterWorkspace.getGadget(defaultGadgetId);
  using gadget = await connectSession(requesterGadget);

  // The Workshop hands the browser the Gadget itself, which doesn't know who is connected.
  await expect(gadget.whoami()).rejects.toThrow('does not implement the method "whoami"');

  // So the browser says who it is -- and nothing checks. The "use" collaborator submits as
  // themselves, then claims to be the owner, with the owner's role, and approves their own request.
  using asThemselves = gadget.connectViewer(
      { id: `claimed:${requesterName}`, displayName: requesterName, role: "use" });
  using asOwner = gadget.connectViewer(
      { id: `claimed:${ownerName}`, displayName: ownerName, role: "build" });
  const view = new Listener();
  await asOwner.subscribe(view);
  await asThemselves.submit("New laptop", "Mine is from 2019.");
  const request = await waitFor("the request", async () =>
    view.latest?.requests.find(each => each.title === "New laptop") ?? null);
  await asOwner.decide(request.id, "approved", "Looks fine to me.");
  const decided = await waitFor("the decision", async () =>
    view.latest?.requests.find(each => each.id === request.id && each.decision) ?? null);
  expect(decided.decision).toEqual(expect.objectContaining(
      { outcome: "approved", by: { id: `claimed:${ownerName}`, name: ownerName } }));
});
