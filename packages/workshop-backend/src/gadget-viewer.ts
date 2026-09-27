import type { CollaboratorRole } from "@gadgets/workshop-shared/api";

/** The user behind a connectToGadget() connection, as gadget code sees them through `viewer()`. */
export type GadgetViewer = {
  /** Stable for this user within this workspace; unrelated to their username, email or other ids. */
  id: string;

  /** Chosen by the user; not unique. */
  displayName: string;

  /** The owner connects as "build". */
  role: CollaboratorRole;
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
 * Method, added to every Gadget class, through which a connectToGadget() stub makes each call on
 * behalf of its viewer. Every gadget stub the Workshop hands out refuses it.
 */
export const GADGET_VIEWER_METHOD = "callAsViewer";

/** Module gadget code imports `viewer()` from. */
export const GADGET_VIEWER_MODULE = "gadgets:viewer";

/** Source of GADGET_VIEWER_MODULE. */
export const GADGET_VIEWER_MODULE_SOURCE = `
import { AsyncLocalStorage } from "node:async_hooks";

const current = new AsyncLocalStorage();

/** The user whose connection made the current call, or null (the agent, a hook, another gadget). */
export function viewer() {
  return current.getStore() ?? null;
}

// Only what RPC could call directly: methods on the class, not fields or Object's own.
function isMethod(object, name) {
  for (let proto = Object.getPrototypeOf(object); proto && proto !== Object.prototype;
      proto = Object.getPrototypeOf(proto)) {
    let desc = Object.getOwnPropertyDescriptor(proto, name);
    if (desc) return typeof desc.value === "function";
  }
  return false;
}

export function withViewers(Base) {
  if (typeof Base !== "function") return Base;
  return class Gadget extends Base {
    ${GADGET_VIEWER_METHOD}(viewer, name, args) {
      if (name === "constructor" || name === "${GADGET_VIEWER_METHOD}" || !isMethod(this, name)) {
        throw new TypeError(\`The RPC receiver does not implement the method "\${name}".\`);
      }
      return current.run(viewer, () => this[name](...args));
    }
  };
}
`;

/** Main module of a gadget worker with a server.js: re-exports it, with its Gadget wrapped. */
export const GADGET_MAIN_MODULE = "gadgets:main";

/** Source of GADGET_MAIN_MODULE. */
export const GADGET_MAIN_MODULE_SOURCE = `
import { withViewers } from "${GADGET_VIEWER_MODULE}";
import * as gadget from "./server.js";
export * from "./server.js";
export const Gadget = withViewers(gadget.Gadget);
`;
