// The wire contract between the Approvals client and its Durable Object.

/**
 * A signed-in user, as `currentUser()` from "gadgets:user" returns them. Spelled out rather than
 * imported, because the client imports this module too and only the server has "gadgets:user".
 */
export type User = {
  /** Opaque, stable for this person within this workspace. */
  id: string;
  /** Chosen by the person; not unique. */
  displayName: string;
  role: "build" | "use";
};

/** A person as recorded on a request or decision. */
export type Person = { id: string; name: string };

export type Outcome = "approved" | "rejected";

export type Request = {
  id: string;
  title: string;
  detail: string;
  requester: Person;
  createdAt: string;
  decision?: { outcome: Outcome; by: Person; at: string; note: string };
};

/** Someone who has opened this gadget. */
export type Member = Person & { role: User["role"]; approver: boolean; online: boolean };

/** Everything one user's screen shows, computed for that user. */
export type View = {
  me: User & { approver: boolean };
  /** Newest first. `canDecide` is whether *this user* may decide it. */
  requests: (Request & { canDecide: boolean })[];
  members: Member[];
  canManageApprovers: boolean;
};

/** Implemented by the client: receives a fresh view whenever anything changes. */
export interface ViewListener {
  update(view: View): void;
}

/**
 * What a connected client can call. Every call acts as the user who made it, as the Workshop
 * reports through `currentUser()`; no method takes an identity as a parameter.
 */
export interface ApprovalsApi {
  whoami(): User;
  subscribe(listener: ViewListener): Promise<View>;
  submit(title: string, detail: string): Promise<void>;
  decide(requestId: string, outcome: Outcome, note: string): Promise<void>;
  setApprover(memberId: string, approver: boolean): Promise<void>;
}
