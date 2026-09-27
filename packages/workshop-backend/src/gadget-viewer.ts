import type { DurableObject, RpcTarget } from "cloudflare:workers";

/**
 * Optional Gadget method through which connectToGadget() tells a gadget who is connecting. The
 * client then receives the RpcTarget it returns instead of the gadget. Every gadget stub the
 * Workshop hands out refuses this method (see OverseerImpl.getGadgetFacet).
 */
export const GADGET_VIEWER_METHOD = "connectViewer";

/**
 * The user behind one connectToGadget() call, as authenticated by the Workshop. Gadget rules bind
 * "use" viewers only: "build" viewers (including the owner) can change the gadget's code and data.
 */
export type GadgetViewer = {
  /** Stable for this user within this workspace; unrelated to their username, email or other ids. */
  id: string;

  /** Chosen by the user; not unique. */
  displayName: string;

  role: "build" | "use";
};

/** A gadget that implements GADGET_VIEWER_METHOD. */
export type GadgetWithViewers = DurableObject & {
  connectViewer(viewer: GadgetViewer): RpcTarget;
};

/** HMAC-SHA-256 of `profileId` under the workspace's key (both hex). */
export async function gadgetViewerId(workspaceKey: string, profileId: string): Promise<string> {
  let key = await crypto.subtle.importKey(
      "raw", Uint8Array.fromHex(workspaceKey), { name: "HMAC", hash: "SHA-256" },
      false, ["sign"]);
  let sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(profileId));
  return new Uint8Array(sig).toHex();
}

/**
 * Whether `error` means the gadget doesn't implement GADGET_VIEWER_METHOD. Exact on purpose: a
 * gadget that throws to turn a viewer away must not fall back to handing them the gadget itself.
 */
export function isViewerMethodMissing(error: unknown): boolean {
  return error instanceof Error && error.message ===
      `The RPC receiver does not implement the method "${GADGET_VIEWER_METHOD}".`;
}
