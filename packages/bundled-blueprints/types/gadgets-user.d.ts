// The `gadgets:user` module the Workshop gives every gadget's Durable Object (`gadget-user.ts` in
// the backend, which owns the runtime and the `GadgetUser` type this mirrors). Declared here for
// `tsconfig.server.json` and `tsconfig.server-tests.json`; this package can't import the backend,
// which depends on it.

declare module "gadgets:user" {
  /** A signed-in user, as gadget code sees them. */
  export type GadgetUser = {
    /** Stable for this user within this workspace, and not derivable from their profile id. */
    id: string;
    /** Chosen by the user; not unique. */
    displayName: string;
    /** The owner connects as "build". */
    role: "build" | "use";
  };

  /**
   * The user whose client made the current call (or export), or null for calls from the agent,
   * hooks and other gadgets. Work a call starts, such as timers, keeps that call's user.
   */
  export function currentUser(): GadgetUser | null;
}
