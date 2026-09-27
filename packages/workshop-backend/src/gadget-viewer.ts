import type { CollaboratorRole } from "@gadgets/workshop-shared/api";
import type { DurableObject, RpcTarget } from "cloudflare:workers";

/**
 * The user behind one connectToGadget() connection, as authenticated by the Workshop. Gadget rules
 * bind "use" viewers only: "build" viewers (including the owner) can change the gadget's code and
 * data.
 */
export type GadgetViewer = {
  /** Stable for this user within this workspace; unrelated to their username, email or other ids. */
  id: string;

  /** Chosen by the user; not unique. */
  displayName: string;

  /** The access the connection was admitted with; the owner connects as "build". */
  role: CollaboratorRole;
};

/**
 * Method through which connectToGadget() opens a connection's session on the gadget (see
 * GADGET_VIEWER_MODULE_SOURCE). Every gadget stub the Workshop hands out refuses it (see
 * OverseerImpl.getGadgetFacet), so only connectToGadget() can present a viewer.
 */
export const GADGET_VIEWER_METHOD = "connectViewer";

/** What the facet looks like once GADGET_MAIN_MODULE has wrapped it. */
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

/** Module gadget code imports `viewer()` from. */
export const GADGET_VIEWER_MODULE = "gadgets:viewer";

/**
 * Source of GADGET_VIEWER_MODULE. `withViewers()` gives the gadget's class a GADGET_VIEWER_METHOD
 * whose session forwards each call to the gadget, with `viewer()` answering for its duration (and
 * for any timer or callback that call starts). A gadget that defines the method itself keeps it.
 */
export const GADGET_VIEWER_MODULE_SOURCE = `
import { AsyncLocalStorage } from "node:async_hooks";
import { DurableObject, RpcTarget } from "cloudflare:workers";

const current = new AsyncLocalStorage();

/** Who made the current call: a connected user, or null (the agent, a hook, another gadget). */
export function viewer() {
  return current.getStore() ?? null;
}

const sessionTargets = new WeakMap();
const sessionClasses = new WeakMap();

function sessionClassFor(cls) {
  let Session = sessionClasses.get(cls);
  if (Session) return Session;
  Session = class extends RpcTarget {};
  let seen = new Set(["constructor", "${GADGET_VIEWER_METHOD}"]);
  for (let proto = cls.prototype; proto && proto !== DurableObject.prototype &&
      proto !== Object.prototype; proto = Object.getPrototypeOf(proto)) {
    for (let [name, desc] of Object.entries(Object.getOwnPropertyDescriptors(proto))) {
      if (seen.has(name)) continue;
      seen.add(name);
      let forward = (session, use) => {
        let { gadget, viewer } = sessionTargets.get(session);
        return current.run(viewer, () => use(gadget));
      };
      if (typeof desc.value === "function") {
        Session.prototype[name] = function (...args) {
          return forward(this, gadget => gadget[name](...args));
        };
      } else if (desc.get) {
        Object.defineProperty(Session.prototype, name, {
          get() { return forward(this, gadget => gadget[name]); },
        });
      }
    }
  }
  sessionClasses.set(cls, Session);
  return Session;
}

export function withViewers(Base) {
  if (typeof Base !== "function") return Base;
  return class Gadget extends Base {
    ${GADGET_VIEWER_METHOD}(viewer) {
      if (typeof super.${GADGET_VIEWER_METHOD} === "function") {
        return current.run(viewer, () => super.${GADGET_VIEWER_METHOD}(viewer));
      }
      let session = new (sessionClassFor(Base))();
      sessionTargets.set(session, { gadget: this, viewer });
      return session;
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
