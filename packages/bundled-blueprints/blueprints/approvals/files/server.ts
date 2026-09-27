import { DurableObject, RpcTarget, WorkerEntrypoint } from "cloudflare:workers";
import { SubscriberRegistry } from "@gadgets/bundled-blueprints/libraries/sync/server";
import type {
  ApprovalSessionApi,
  Outcome,
  Person,
  Request,
  View,
  ViewListener,
  Viewer,
} from "./lib/protocol.ts";

/**
 * Approvals: people submit requests, and an approver other than the requester decides them.
 *
 * Identity comes from the Workshop, never from the browser. The Workshop calls `connectViewer()`
 * once per connection with the signed-in person, and the client talks to the `ApprovalSession` it
 * returns. Everything a session does is bound to that viewer when the session is created, so no
 * method -- here or on the session -- accepts an identity as a parameter. Only the Workshop can
 * call `connectViewer()`: the agent, other gadgets and browsers all reach this class through stubs
 * that refuse it.
 *
 * Rules:
 * - "build" viewers (the owner among them) are always approvers, and choose other approvers among
 *   people who have connected.
 * - Nobody decides their own request.
 *
 * These rules bind "use" viewers only. Anyone who can change this gadget's code or data -- "build"
 * viewers and any agent they run -- can bypass them and rewrite the stored requests and history.
 */

type StoredMember = { name: string; role: Viewer["role"]; approver: boolean };
type State = { requests: Request[]; members: Record<string, StoredMember> };

const MAX_TITLE = 200;
const MAX_DETAIL = 2000;
const MAX_NOTE = 500;
const MAX_REQUESTS = 500;

/** What a session may do, already bound to its viewer. Never exposed over RPC. */
type BoundActions = {
  view(): View;
  listen(listener: ViewListener): ViewListener;
  unlisten(handle: ViewListener): void;
  submit(title: string, detail: string): Promise<void>;
  decide(requestId: string, outcome: Outcome, note: string): Promise<void>;
  setApprover(memberId: string, approver: boolean): Promise<void>;
};

function clip(value: unknown, max: number): string {
  return typeof value === "string" ? value.trim().slice(0, max) : "";
}

export class Gadget extends DurableObject<unknown> {
  #state: State = { requests: [], members: {} };
  #viewers = new WeakMap<ViewListener, Viewer>();
  #listeners = new SubscriberRegistry<ViewListener, Viewer>({
    // Joins and leaves change who is online, so everyone gets a fresh view of their own.
    join: (listener) => this.#push(listener),
    leave: (listener) => this.#push(listener),
  });

  constructor(ctx: DurableObjectState, env: unknown) {
    super(ctx, env);
    void ctx.blockConcurrencyWhile(async () => {
      this.#state = (await ctx.storage.get<State>("state")) ?? this.#state;
    });
  }

  /** Read-only list of requests, e.g. for the agent. Acts as nobody. */
  listRequests(): Request[] {
    return this.#state.requests;
  }

  /** Called by the Workshop, once per connection, with the person on the other end. */
  async connectViewer(viewer: Viewer): Promise<ApprovalSession> {
    let known = this.#state.members[viewer.id];
    if (known?.name !== viewer.displayName || known.role !== viewer.role) {
      this.#state.members[viewer.id] =
          { name: viewer.displayName, role: viewer.role, approver: known?.approver ?? false };
      await this.#commit();
    }
    return new ApprovalSession(viewer, {
      view: () => this.#viewFor(viewer),
      listen: (listener) => {
        let handle = this.#listeners.add(listener, viewer);
        this.#viewers.set(handle, viewer);
        return handle;
      },
      unlisten: (handle) => { this.#listeners.remove(handle); },
      submit: (title, detail) => this.#submit(viewer, title, detail),
      decide: (requestId, outcome, note) => this.#decide(viewer, requestId, outcome, note),
      setApprover: (memberId, approver) => this.#setApprover(viewer, memberId, approver),
    });
  }

  #isApprover(viewer: Viewer): boolean {
    return viewer.role === "build" || this.#state.members[viewer.id]?.approver === true;
  }

  #viewFor(viewer: Viewer): View {
    let approver = this.#isApprover(viewer);
    let online = new Set(this.#listeners.members().map(each => each.id));
    return {
      me: { ...viewer, approver },
      requests: this.#state.requests.toReversed().map(request => ({
        ...request,
        canDecide: approver && !request.decision && request.requester.id !== viewer.id,
      })),
      members: Object.entries(this.#state.members).map(([id, member]) => ({
        id,
        name: member.name,
        role: member.role,
        approver: member.role === "build" || member.approver,
        online: online.has(id),
      })),
      canManageApprovers: viewer.role === "build",
    };
  }

  #push(listener: ViewListener): unknown {
    let viewer = this.#viewers.get(listener);
    return viewer && listener.update(this.#viewFor(viewer));
  }

  async #commit(): Promise<void> {
    await this.ctx.storage.put("state", this.#state);
    this.#listeners.broadcast(listener => this.#push(listener));
  }

  async #submit(viewer: Viewer, title: string, detail: string): Promise<void> {
    let cleanTitle = clip(title, MAX_TITLE);
    if (!cleanTitle) throw new Error("A request needs a title.");
    if (this.#state.requests.length >= MAX_REQUESTS) {
      throw new Error("This gadget holds as many requests as it can.");
    }
    this.#state.requests.push({
      id: crypto.randomUUID(),
      title: cleanTitle,
      detail: clip(detail, MAX_DETAIL),
      requester: person(viewer),
      createdAt: new Date().toISOString(),
    });
    await this.#commit();
  }

  async #decide(viewer: Viewer, requestId: string, outcome: Outcome, note: string)
      : Promise<void> {
    if (outcome !== "approved" && outcome !== "rejected") throw new Error("Unknown outcome.");
    let request = this.#state.requests.find(each => each.id === requestId);
    if (!request) throw new Error("No such request.");
    if (request.decision) throw new Error("This request has already been decided.");
    if (request.requester.id === viewer.id) throw new Error("You can't decide your own request.");
    if (!this.#isApprover(viewer)) throw new Error("Only approvers can decide requests.");
    request.decision =
        { outcome, by: person(viewer), at: new Date().toISOString(), note: clip(note, MAX_NOTE) };
    await this.#commit();
  }

  async #setApprover(viewer: Viewer, memberId: string, approver: boolean): Promise<void> {
    if (viewer.role !== "build") throw new Error("Only builders choose approvers.");
    let member = this.#state.members[memberId];
    if (!member) throw new Error("No such person.");
    if (member.role === "build") throw new Error("Builders are always approvers.");
    member.approver = approver === true;
    await this.#commit();
  }
}

function person(viewer: Viewer): Person {
  return { id: viewer.id, name: viewer.displayName };
}

/** One connection's view of the gadget, acting as the viewer the Workshop connected. */
class ApprovalSession extends RpcTarget implements ApprovalSessionApi {
  #viewer: Viewer;
  #actions: BoundActions;
  #listener: ViewListener | undefined;

  constructor(viewer: Viewer, actions: BoundActions) {
    super();
    this.#viewer = viewer;
    this.#actions = actions;
  }

  whoami(): Viewer {
    return this.#viewer;
  }

  subscribe(listener: ViewListener): View {
    if (this.#listener) this.#actions.unlisten(this.#listener);
    this.#listener = this.#actions.listen(listener);
    return this.#actions.view();
  }

  submit(title: string, detail: string): Promise<void> {
    return this.#actions.submit(title, detail);
  }

  decide(requestId: string, outcome: Outcome, note: string): Promise<void> {
    return this.#actions.decide(requestId, outcome, note);
  }

  setApprover(memberId: string, approver: boolean): Promise<void> {
    return this.#actions.setApprover(memberId, approver);
  }

  [Symbol.dispose](): void {
    if (this.#listener) this.#actions.unlisten(this.#listener);
  }
}

export class ExportHandler extends WorkerEntrypoint {
  async getExportFormats() {
    return [
      { id: "html", label: "HTML", mode: "browser", contentType: "text/html", fileExtension: ".html" },
      { id: "pdf", label: "PDF", mode: "browser", contentType: "application/pdf", fileExtension: ".pdf" },
    ];
  }
}
