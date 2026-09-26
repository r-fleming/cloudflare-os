// The wire contract between the Approvals client and its Durable Object.

/** Who a connection belongs to, as the Workshop presents it to `Gadget.connectViewer()`. */
export type Viewer = {
  /** Opaque, stable for this person within this workspace. */
  id: string;
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
export type Member = Person & { role: Viewer["role"]; approver: boolean; online: boolean };

/** Everything one viewer's screen shows, computed for that viewer. */
export type View = {
  me: Viewer & { approver: boolean };
  /** Newest first. `canDecide` is whether *this viewer* may decide it. */
  requests: (Request & { canDecide: boolean })[];
  members: Member[];
  canManageApprovers: boolean;
};

/** Implemented by the client: receives a fresh view whenever anything changes. */
export interface ViewListener {
  update(view: View): void;
}

/**
 * What one connected client can do. Every call acts as the connection's viewer; no method takes an
 * identity as a parameter.
 */
export interface ApprovalSessionApi {
  whoami(): Viewer;
  subscribe(listener: ViewListener): View;
  submit(title: string, detail: string): Promise<void>;
  decide(requestId: string, outcome: Outcome, note: string): Promise<void>;
  setApprover(memberId: string, approver: boolean): Promise<void>;
}
