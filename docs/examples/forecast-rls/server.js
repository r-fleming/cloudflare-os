// Demo: a renewals forecast where each "use" collaborator sees only their branch of the org
// tree and can save only their own rows. See README.md.
import { DurableObject, RpcTarget } from "cloudflare:workers";

const SAMPLE_ORG = [
  { email: "vp@example.com", name: "Global VP", manager: null },
  { email: "amer-lead@example.com", name: "AMER Lead", manager: "vp@example.com" },
  { email: "emea-lead@example.com", name: "EMEA Lead", manager: "vp@example.com" },
  { email: "amer-rep1@example.com", name: "AMER Rep 1", manager: "amer-lead@example.com" },
  { email: "amer-rep2@example.com", name: "AMER Rep 2", manager: "amer-lead@example.com" },
  { email: "emea-rep1@example.com", name: "EMEA Rep 1", manager: "emea-lead@example.com" },
];

const SAMPLE_ROWS = [
  { id: "a1", account: "Acme Corp", owner: "amer-rep1@example.com", arr: 120000, forecast: 120000 },
  { id: "a2", account: "Globex", owner: "amer-rep1@example.com", arr: 80000, forecast: 60000 },
  { id: "a3", account: "Initech", owner: "amer-rep2@example.com", arr: 200000, forecast: 210000 },
  { id: "a4", account: "Umbrella", owner: "amer-lead@example.com", arr: 500000, forecast: 450000 },
  { id: "e1", account: "Soylent", owner: "emea-rep1@example.com", arr: 150000, forecast: 150000 },
  { id: "e2", account: "Hooli EU", owner: "emea-lead@example.com", arr: 300000, forecast: 320000 },
];

export class Gadget extends DurableObject {
  // Live subscribers: {email (null = sees everything), callback}.
  subscribers = new Set();

  // Called by the Workshop, once per connection. Nothing else can reach this class's methods.
  connectCaller(caller) {
    if (caller.kind === "user" && (caller.role === "use" || caller.role === "view")) {
      if (!this.#org().some(p => p.email === caller.id)) {
        throw new Error(`${caller.id} is not in the forecast's org tree.`);
      }
      return new MemberSession(this, caller);
    }
    // Owners, editors, the agent, bound gadgets and system work are trusted.
    return new AdminSession(this, caller);
  }

  #org() { return this.ctx.storage.kv.get("org") ?? SAMPLE_ORG; }
  #rows() { return this.ctx.storage.kv.get("rows") ?? SAMPLE_ROWS; }

  // Everyone at or below `email` in the org tree.
  branch(email) {
    let org = this.#org();
    let result = new Set([email]);
    let grew = true;
    while (grew) {
      grew = false;
      for (let p of org) {
        if (p.manager && result.has(p.manager) && !result.has(p.email)) {
          result.add(p.email);
          grew = true;
        }
      }
    }
    return result;
  }

  // What `email` may see (null: everything), and which rows they may edit.
  viewFor(email) {
    let org = this.#org();
    let names = Object.fromEntries(org.map(p => [p.email, p.name]));
    let visible = email ? this.branch(email) : null;
    let rows = this.#rows()
      .filter(r => !visible || visible.has(r.owner))
      .map(r => ({ ...r, ownerName: names[r.owner] ?? r.owner,
                   editable: !email || r.owner === email }));
    let total = rows.reduce((sum, r) => sum + r.forecast, 0);
    return { rows, total };
  }

  saveForecast(email, id, forecast) {
    if (typeof forecast !== "number" || !Number.isFinite(forecast) || forecast < 0) {
      throw new Error("Forecast must be a non-negative number.");
    }
    let rows = this.#rows().map(r => ({ ...r }));
    let row = rows.find(r => r.id === id);
    if (!row) throw new Error(`No such row: ${id}`);
    if (email && row.owner !== email) throw new Error("You can only save your own rows.");
    row.forecast = forecast;
    this.ctx.storage.kv.put("rows", rows);
    this.#broadcast();
  }

  setData(org, rows) {
    if (org) this.ctx.storage.kv.put("org", org);
    if (rows) this.ctx.storage.kv.put("rows", rows);
    this.#broadcast();
  }

  subscribe(email, callback) {
    let entry = { email, callback: callback.dup() };
    this.subscribers.add(entry);
    entry.callback.onRpcBroken(() => this.subscribers.delete(entry));
    return this.viewFor(email);
  }

  // Each subscriber gets only its own caller's view.
  #broadcast() {
    for (let { email, callback } of this.subscribers) {
      callback.update(this.viewFor(email)).catch(() => {});
    }
  }
}

// A "use" or "view" collaborator: their branch only; saves to their own rows only.
class MemberSession extends RpcTarget {
  #gadget; #caller;
  constructor(gadget, caller) { super(); this.#gadget = gadget; this.#caller = caller; }
  whoami() { return { ...this.#caller, scope: "branch" }; }
  getView() { return this.#gadget.viewFor(this.#caller.id); }
  subscribe(callback) { return this.#gadget.subscribe(this.#caller.id, callback); }
  saveForecast(id, forecast) {
    if (this.#caller.role === "view") throw new Error("A snapshot can't be edited.");
    return this.#gadget.saveForecast(this.#caller.id, id, forecast);
  }
}

// Trusted callers: everything.
class AdminSession extends RpcTarget {
  #gadget; #caller;
  constructor(gadget, caller) { super(); this.#gadget = gadget; this.#caller = caller; }
  whoami() { return { ...this.#caller, scope: "all" }; }
  getView() { return this.#gadget.viewFor(null); }
  subscribe(callback) { return this.#gadget.subscribe(null, callback); }
  saveForecast(id, forecast) { return this.#gadget.saveForecast(null, id, forecast); }
  // e.g. from executeCode: replace the org tree and rows (from SFDC, Sheets, ...).
  setData(org, rows) { return this.#gadget.setData(org, rows); }
}
