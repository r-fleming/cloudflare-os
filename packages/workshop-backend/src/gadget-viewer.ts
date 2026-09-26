import type { DurableObject, RpcTarget } from "cloudflare:workers";

/**
 * Name of the optional Gadget method through which the Workshop opens a per-viewer session:
 * `connectViewer(viewer: GadgetViewer): RpcTarget`. When a gadget implements it,
 * connectToGadget() hands the client the returned session instead of the gadget itself.
 *
 * Reserved: every gadget stub the Workshop hands out refuses it (see OverseerImpl.getGadgetFacet),
 * so only connectToGadget() can call it. Neither the browser, the agent, another gadget nor a hook
 * can open a session as someone else.
 */
export const GADGET_VIEWER_METHOD = "connectViewer";

/**
 * The person behind one connectToGadget() connection, as established by the Workshop.
 *
 * The Workshop authenticates the viewer; the gadget authorizes. A gadget's rules and stored
 * records are only as trustworthy as everyone who can change its code or data: its "build"
 * collaborators (the owner among them) and any agent they run, including unmerged code an agent
 * tests against the gadget's live storage. So a gadget's own rules bind "use" viewers only.
 */
export type GadgetViewer = {
  /**
   * Opaque id, stable for this person within this workspace. It cannot be computed from their
   * username or email, and is unrelated to their id in any other workspace.
   */
  id: string;

  /** The viewer's display name when the connection opened. For display only. */
  displayName: string;

  /**
   * The access this connection was admitted with, already enforced by the Workshop (the owner
   * connects as "build"). A "build" viewer can change the gadget's code, so the gadget cannot hold
   * them to its own rules.
   */
  role: "build" | "use";
};

/** A gadget facet that implements the GADGET_VIEWER_METHOD handshake. */
export type GadgetWithViewers = DurableObject & {
  connectViewer(viewer: GadgetViewer): RpcTarget;
};

/** Derives GadgetViewer.id: HMAC-SHA-256 of the profile id, keyed by the workspace's key (hex). */
export async function gadgetViewerId(workspaceKey: string, profileId: string): Promise<string> {
  let key = await crypto.subtle.importKey(
      "raw", Uint8Array.fromHex(workspaceKey), { name: "HMAC", hash: "SHA-256" },
      false, ["sign"]);
  let sig = new Uint8Array(await crypto.subtle.sign(
      "HMAC", key, new TextEncoder().encode(profileId)));
  return sig.toHex();
}

/**
 * Whether `error` is the runtime's report that the gadget doesn't implement GADGET_VIEWER_METHOD.
 * Deliberately exact: a gadget whose handshake throws to turn a viewer away must not fall back to
 * handing that viewer the gadget itself.
 */
export function isViewerMethodMissing(error: unknown): boolean {
  return error instanceof Error && error.message ===
      `The RPC receiver does not implement the method "${GADGET_VIEWER_METHOD}".`;
}
