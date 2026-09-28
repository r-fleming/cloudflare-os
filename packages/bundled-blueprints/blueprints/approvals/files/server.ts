import { DurableObject, WorkerEntrypoint } from "cloudflare:workers";
import { currentUser } from "gadgets:user";
import { SubscriberRegistry } from "@gadgets/bundled-blueprints/libraries/sync/server";
import type {
  ApprovalsApi,
  Outcome,
  Person,
  Request,
  User,
  View,
  ViewListener,
} from "./lib/protocol.ts";

/**
 * Approvals: people submit requests, and an approver other than the requester decides them.
 *
 * Identity comes from the Workshop, never from the browser: every method that acts as someone
 * starts with `this.#me()`, which reads `currentUser()` and refuses callers that aren't signed-in
 * people (the agent, hooks, other gadgets). No method accepts an identity as a parameter.
 *
 * Rules:
 * - "build" users (the owner among them) are always approvers, and choose other approvers among
 *   people who have opened the gadget.
 * - Nobody decides their own request.
 *
 * These rules bind "use" users only. Anyone who can change this gadget's code or data -- "build"
 * users and any agent they run -- can bypass them and rewrite the stored requests and history.
 */

type StoredMember = { name: string; role: User["role"]; approver: boolean };
type State = { requests: Request[]; members: Record<string, StoredMember> };

const MAX_TITLE = 200;
const MAX_DETAIL = 2000;
const MAX_NOTE = 500;
const MAX_REQUESTS = 500;

function clip(value: unknown, max: number): string {
  return typeof value === "string" ? value.trim().slice(0, max) : "";
}

function person(user: User): Person {
  return { id: user.id, name: user.displayName };
}

export class Gadget extends DurableObject<unknown> implements ApprovalsApi {
  #state: State = { requests: [], members: {} };
  #users = new WeakMap<ViewListener, User>();
  #listeners = new SubscriberRegistry<ViewListener, User>({
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

  whoami(): User {
    return this.#me();
  }

  async subscribe(listener: ViewListener): Promise<View> {
    let me = this.#me();
    await this.#remember(me);
    this.#users.set(this.#listeners.add(listener, me), me);
    return this.#viewFor(me);
  }

  async submit(title: string, detail: string): Promise<void> {
    let me = this.#me();
    let cleanTitle = clip(title, MAX_TITLE);
    if (!cleanTitle) throw new Error("A request needs a title.");
    if (this.#state.requests.length >= MAX_REQUESTS) {
      throw new Error("This gadget holds as many requests as it can.");
    }
    this.#state.requests.push({
      id: crypto.randomUUID(),
      title: cleanTitle,
      detail: clip(detail, MAX_DETAIL),
      requester: person(me),
      createdAt: new Date().toISOString(),
    });
    await this.#commit();
  }

  async decide(requestId: string, outcome: Outcome, note: string): Promise<void> {
    let me = this.#me();
    if (outcome !== "approved" && outcome !== "rejected") throw new Error("Unknown outcome.");
    let request = this.#state.requests.find(each => each.id === requestId);
    if (!request) throw new Error("No such request.");
    if (request.decision) throw new Error("This request has already been decided.");
    if (request.requester.id === me.id) throw new Error("You can't decide your own request.");
    if (!this.#isApprover(me)) throw new Error("Only approvers can decide requests.");
    request.decision =
        { outcome, by: person(me), at: new Date().toISOString(), note: clip(note, MAX_NOTE) };
    await this.#commit();
  }

  async setApprover(memberId: string, approver: boolean): Promise<void> {
    let me = this.#me();
    if (me.role !== "build") throw new Error("Only builders choose approvers.");
    let member = this.#state.members[memberId];
    if (!member) throw new Error("No such person.");
    if (member.role === "build") throw new Error("Builders are always approvers.");
    member.approver = approver === true;
    await this.#commit();
  }

  #me(): User {
    let me = currentUser();
    if (!me) throw new Error("Only signed-in people can use Approvals.");
    return me;
  }

  async #remember(user: User): Promise<void> {
    let known = this.#state.members[user.id];
    if (known?.name === user.displayName && known.role === user.role) return;
    this.#state.members[user.id] =
        { name: user.displayName, role: user.role, approver: known?.approver ?? false };
    await this.#commit();
  }

  #isApprover(user: User): boolean {
    return user.role === "build" || this.#state.members[user.id]?.approver === true;
  }

  #viewFor(user: User): View {
    let approver = this.#isApprover(user);
    let online = new Set(this.#listeners.members().map(each => each.id));
    return {
      me: { ...user, approver },
      requests: this.#state.requests.toReversed().map(request => ({
        ...request,
        canDecide: approver && !request.decision && request.requester.id !== user.id,
      })),
      members: Object.entries(this.#state.members).map(([id, member]) => ({
        id,
        name: member.name,
        role: member.role,
        approver: member.role === "build" || member.approver,
        online: online.has(id),
      })),
      canManageApprovers: user.role === "build",
    };
  }

  // Each listener is sent the view of the user who subscribed it, not of whoever caused the change.
  #push(listener: ViewListener): unknown {
    let user = this.#users.get(listener);
    return user && listener.update(this.#viewFor(user));
  }

  async #commit(): Promise<void> {
    await this.ctx.storage.put("state", this.#state);
    this.#listeners.broadcast(listener => this.#push(listener));
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
