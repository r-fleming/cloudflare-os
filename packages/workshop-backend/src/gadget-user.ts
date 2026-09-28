import type { CollaboratorRole } from "@gadgets/workshop-shared/api";

/** The user behind a connectToGadget() connection, as gadget code sees them via currentUser(). */
export type GadgetUser = {
  /** Stable for this user within this workspace, and not derivable from their profile id. */
  id: string;

  /** Chosen by the user (by default, from their username or email); not unique. */
  displayName: string;

  /** The owner connects as "build". */
  role: CollaboratorRole;
};

/** HMAC-SHA-256 of `profileId` under the workspace's secret (hex). */
export async function gadgetUserId(workspaceSecret: string, profileId: string): Promise<string> {
  let key = await crypto.subtle.importKey(
      "raw", Uint8Array.fromHex(workspaceSecret), { name: "HMAC", hash: "SHA-256" },
      false, ["sign"]);
  let sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(profileId));
  return new Uint8Array(sig).toHex();
}

/**
 * Method, added to every Gadget class, through which a connectToGadget() stub makes each call on
 * behalf of its user. Every gadget stub the Workshop hands out refuses it.
 */
export const GADGET_USER_METHOD = "callAsUser";

/** Module gadget code imports `currentUser()` from. */
export const GADGET_USER_MODULE = "gadgets:user";

/** Source of GADGET_USER_MODULE. */
export const GADGET_USER_MODULE_SOURCE = `
import { AsyncLocalStorage } from "node:async_hooks";

const current = new AsyncLocalStorage();

/**
 * The user whose connection made the current call, or null (the agent, a hook, another gadget).
 * Work a call starts (timers, promises) keeps its user, so read this once and pass it down.
 */
export function currentUser() {
  return current.getStore() ?? null;
}

// Names workerd refuses to call over RPC on a Durable Object.
const RESERVED = new Set(["constructor", "fetch", "connect", "alarm", "webSocketMessage",
    "webSocketClose", "webSocketError", "dup", "${GADGET_USER_METHOD}"]);

// As workerd dispatches RPC: a method of the class, not an own property or one of Object's.
function isRpcMethod(object, name) {
  if (RESERVED.has(name) || Object.hasOwn(object, name)) return false;
  for (let proto = Object.getPrototypeOf(object); proto && proto !== Object.prototype;
      proto = Object.getPrototypeOf(proto)) {
    let desc = Object.getOwnPropertyDescriptor(proto, name);
    if (desc) return typeof desc.value === "function";
  }
  return false;
}

export function withUsers(Base) {
  if (typeof Base !== "function") {
    return class Gadget {
      constructor() { throw new Error('server.js must export a class named "Gadget".'); }
    };
  }
  return class Gadget extends Base {
    ${GADGET_USER_METHOD}(user, name, args) {
      if (!isRpcMethod(this, name)) {
        throw new TypeError(\`The RPC receiver does not implement the method "\${name}".\`);
      }
      return current.run(user, () => this[name](...args));
    }
  };
}
`;

/** Main module of a gadget worker with a server.js: re-exports it, with its Gadget wrapped. */
export const GADGET_MAIN_MODULE = "gadgets:main";

/** Source of GADGET_MAIN_MODULE. */
export const GADGET_MAIN_MODULE_SOURCE = `
import { withUsers } from "${GADGET_USER_MODULE}";
import * as gadget from "./server.js";
export * from "./server.js";
export const Gadget = withUsers(gadget.Gadget);
`;
