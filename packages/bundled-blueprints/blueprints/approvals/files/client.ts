// Approvals client. It never tells the server who it is: the Workshop does, on every call (see
// server.ts), so `gadget` is used directly and no call carries an identity.

import type { ApprovalsApi, Outcome, View, ViewListener } from "./lib/protocol.ts";

type Remote<T> = {
  [K in keyof T]: T[K] extends (...args: infer A) => infer R
    ? (...args: A) => Promise<Awaited<R>> : never;
};

// Defined by the Workshop's iframe bootstrap before this module runs. The stub survives reconnects.
declare const gadget: Remote<ApprovalsApi>;
declare const RpcTarget: { new(): object };

const style = document.createElement("style");
style.textContent = `
:root { color-scheme: light; font: 14px/1.45 system-ui, sans-serif; color: #1d1d1b; }
body { margin: 0; background: #f6f6f4; }
main { max-width: 760px; margin: 0 auto; padding: 24px 20px 48px; display: grid; gap: 20px; }
header { display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap; }
h1 { font-size: 22px; margin: 0; }
h2 { font-size: 13px; text-transform: uppercase; letter-spacing: .04em; color: #6b6b66; margin: 0 0 8px; }
section { background: #fff; border: 1px solid #e3e3df; border-radius: 10px; padding: 16px; }
.me { color: #4a4a46; }
.pill { display: inline-block; padding: 1px 8px; border-radius: 999px; font-size: 12px; background: #efefec; }
.pill.approved { background: #dff3e4; color: #1c6b34; }
.pill.rejected { background: #fbe3e1; color: #9b2c22; }
.pill.pending { background: #fff3d6; color: #7a5a00; }
.request { border-top: 1px solid #efefec; padding: 12px 0; display: grid; gap: 6px; }
.request:first-child { border-top: 0; padding-top: 0; }
.row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.muted { color: #6b6b66; font-size: 12px; }
input[type=text], textarea { font: inherit; width: 100%; box-sizing: border-box; padding: 6px 8px;
  border: 1px solid #d6d6d1; border-radius: 6px; }
textarea { min-height: 56px; resize: vertical; }
button { font: inherit; padding: 5px 12px; border-radius: 6px; border: 1px solid #c9c9c4;
  background: #fff; cursor: pointer; }
button.primary { background: #1d1d1b; color: #fff; border-color: #1d1d1b; }
.error { background: #fbe3e1; color: #9b2c22; border-radius: 8px; padding: 8px 12px; }
.dot { width: 8px; height: 8px; border-radius: 50%; background: #c9c9c4; display: inline-block; }
.dot.online { background: #2f9e55; }
.stack { display: grid; gap: 20px; }
[hidden] { display: none !important; }
ol.history { margin: 0; padding-left: 18px; display: grid; gap: 4px; }
@media print {
  body { background: #fff; }
  .compose, .actions, .error, input[type=checkbox] { display: none !important; }
  section { border: 0; padding: 0; }
}
`;
document.head.append(style);

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K, props: Partial<HTMLElementTagNameMap[K]> = {}, ...children: (Node | string)[]
): HTMLElementTagNameMap[K] {
  let node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

// Display names are chosen by each user and needn't be unique, so show part of the id beside them:
// two people can both be called "alice", but not with the same id.
function who(person: { id: string; name: string }): string {
  return `${person.name} (#${person.id.slice(0, 6)})`;
}

function when(iso: string): string {
  return new Date(iso).toLocaleString();
}

let errorBox = el("div", { className: "error", hidden: true });
function showError(error: unknown) {
  errorBox.textContent = error instanceof Error ? error.message : String(error);
  errorBox.hidden = false;
}
async function attempt(action: () => Promise<void>) {
  errorBox.hidden = true;
  try { await action(); } catch (error) { showError(error); }
}

// The compose form is built once so typing survives re-renders; decision notes are kept per request.
let titleInput = el("input", { type: "text", placeholder: "What do you need approved?" });
let detailInput = el("textarea", { placeholder: "Details (optional)" });
let submitButton = el("button", { className: "primary", textContent: "Submit request" });
submitButton.onclick = () => attempt(async () => {
  await gadget.submit(titleInput.value, detailInput.value);
  titleInput.value = "";
  detailInput.value = "";
});
let notes = new Map<string, string>();

let header = el("header");
let compose = el("section", { className: "compose" },
    el("h2", { textContent: "New request" }), titleInput, detailInput,
    el("div", { className: "row" }, submitButton));
let requestsSection = el("section");
let peopleSection = el("section");
let historySection = el("section");
let app = el("div", { className: "stack", hidden: true },
    compose, requestsSection, peopleSection, historySection);
document.body.append(el("main", {}, header, errorBox, app));

function renderRequests(view: View) {
  let items = view.requests.map(request => {
    let status = request.decision?.outcome ?? "pending";
    let item = el("div", { className: "request" },
        el("div", { className: "row" },
            el("strong", { textContent: request.title }),
            el("span", { className: `pill ${status}`, textContent: status })),
        el("div", { className: "muted",
          textContent: `Requested by ${who(request.requester)} · ${when(request.createdAt)}` }));
    if (request.detail) item.append(el("div", { textContent: request.detail }));
    if (request.decision) {
      let { by, at, note } = request.decision;
      item.append(el("div", { className: "muted",
        textContent: `${status === "approved" ? "Approved" : "Rejected"} by ${who(by)} · ${when(at)}` +
            (note ? ` — “${note}”` : "") }));
    } else if (request.canDecide) {
      let note = el("input", { type: "text", placeholder: "Note (optional)",
        value: notes.get(request.id) ?? "" });
      note.oninput = () => notes.set(request.id, note.value);
      let decide = (outcome: Outcome) => attempt(async () => {
        await gadget.decide(request.id, outcome, note.value);
        notes.delete(request.id);
      });
      let approve = el("button", { className: "primary", textContent: "Approve" });
      approve.onclick = () => decide("approved");
      let reject = el("button", { textContent: "Reject" });
      reject.onclick = () => decide("rejected");
      item.append(el("div", { className: "row actions" }, note, approve, reject));
    } else if (request.requester.id === view.me.id) {
      item.append(el("div", { className: "muted", textContent: "Waiting for another approver." }));
    }
    return item;
  });
  requestsSection.replaceChildren(el("h2", { textContent: "Requests" }),
      ...(items.length ? items : [el("div", { className: "muted", textContent: "No requests yet." })]));
}

function renderPeople(view: View) {
  let rows = view.members.map(member => {
    let row = el("div", { className: "row" },
        el("span", { className: member.online ? "dot online" : "dot" }),
        el("span", { textContent: who(member) }),
        el("span", { className: "pill", textContent: member.role }));
    if (view.canManageApprovers && member.role !== "build") {
      let toggle = el("input", { type: "checkbox", checked: member.approver });
      toggle.onchange = () => attempt(() => gadget.setApprover(member.id, toggle.checked));
      row.append(el("label", { className: "row muted" }, toggle, "approver"));
    } else if (member.approver) {
      row.append(el("span", { className: "muted", textContent: "approver" }));
    }
    return row;
  });
  peopleSection.replaceChildren(el("h2", { textContent: "People" }), ...rows);
}

function renderHistory(view: View) {
  let events = view.requests.flatMap(request => [
    { at: request.createdAt, text: `${who(request.requester)} requested “${request.title}”` },
    ...(request.decision ? [{ at: request.decision.at,
      text: `${who(request.decision.by)} ${request.decision.outcome} “${request.title}”` }] : []),
  ]).sort((a, b) => a.at.localeCompare(b.at));
  historySection.replaceChildren(el("h2", { textContent: "History" }),
      el("ol", { className: "history" },
          ...events.map(event => el("li", {}, `${event.text} `,
              el("span", { className: "muted", textContent: when(event.at) })))));
}

function render(view: View) {
  let me = el("span", { className: "me" },
      "Signed in as ", el("strong", { textContent: who({ id: view.me.id, name: view.me.displayName }) }),
      " ", el("span", { className: "pill", textContent: view.me.role }),
      view.me.approver ? " · approver" : "");
  header.replaceChildren(el("h1", { textContent: "Approvals" }), me);
  renderRequests(view);
  renderPeople(view);
  renderHistory(view);
  app.hidden = false;
}

let listener: Listener | undefined;

class Listener extends RpcTarget implements ViewListener {
  update(view: View) { if (this === listener) render(view); }
  // The connection was lost. The `gadget` stub reconnects, but callbacks passed over it don't
  // survive, so subscribe again.
  [Symbol.dispose]() { if (this === listener) void subscribe(); }
}

async function subscribe() {
  listener = new Listener();
  try {
    render(await gadget.subscribe(listener));
  } catch (error) {
    showError(error);
  }
}

await subscribe();
