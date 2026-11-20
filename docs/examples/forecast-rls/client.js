import { RpcTarget } from "capnweb";

const money = n => n.toLocaleString("en-US", { style: "currency", currency: "USD",
                                                 maximumFractionDigits: 0 });

function el(tag, props = {}, ...children) {
  let node = Object.assign(document.createElement(tag), props);
  node.append(...children);
  return node;
}

class Updates extends RpcTarget {
  constructor(gadget, render) { super(); this.gadget = gadget; this.render = render; }
  update(view) { this.render(view); }
  [Symbol.dispose]() { this.gadget.subscribe(this).then(this.render); }
}

export default {
  async main(gadget) {
    document.head.append(el("style", { textContent: `
      body { font: 14px system-ui, sans-serif; margin: 16px; }
      table { border-collapse: collapse; width: 100%; max-width: 760px; }
      th, td { text-align: left; padding: 6px 8px; border-bottom: 1px solid #ddd; }
      td.num, th.num { text-align: right; }
      input { width: 110px; text-align: right; }
      .who { padding: 8px 12px; background: #f3f3f3; border-radius: 6px; margin-bottom: 12px; }
      .err { color: #b00020; }
      button { margin-top: 12px; }
    ` }));

    let who = el("div", { className: "who" }, "Connecting…");
    let tableBody = el("tbody");
    let total = el("td", { className: "num" });
    let status = el("div");
    document.body.append(
      el("h2", {}, "Renewals forecast"),
      who,
      el("table", {},
        el("thead", {}, el("tr", {},
          el("th", {}, "Account"), el("th", {}, "Owner"),
          el("th", { className: "num" }, "ARR"), el("th", { className: "num" }, "Forecast"))),
        tableBody,
        el("tfoot", {}, el("tr", {}, el("th", { colSpan: 3 }, "Total"), total))),
      status,
      // Demonstrates that nothing outside the session is reachable from the browser.
      el("button", { textContent: "Try calling the Gadget directly",
                     onclick: async () => {
        try {
          await gadget.connectCaller({ kind: "user", id: "vp@example.com", role: "owner" });
          status.textContent = "Unexpected: connectCaller succeeded.";
        } catch (err) {
          status.className = "err";
          status.textContent = `Refused: ${err.message}`;
        }
      } }));

    let render = view => {
      tableBody.replaceChildren(...view.rows.map(row => {
        let cell = row.editable
          ? el("input", { type: "number", value: row.forecast, onchange: async e => {
              try {
                await gadget.saveForecast(row.id, Number(e.target.value));
                status.className = ""; status.textContent = `Saved ${row.account}.`;
              } catch (err) {
                status.className = "err"; status.textContent = err.message;
              }
            } })
          : money(row.forecast);
        return el("tr", {}, el("td", {}, row.account), el("td", {}, row.ownerName),
                  el("td", { className: "num" }, money(row.arr)),
                  el("td", { className: "num" }, cell));
      }));
      total.textContent = money(view.total);
    };

    try {
      let me = await gadget.whoami();
      who.textContent = `${me.name ?? me.kind} (${me.id ?? me.kind}) — role: ${me.role ?? me.kind}`
          + ` — sees ${me.scope === "all" ? "every row" : "their branch of the org"}`;
      render(await gadget.subscribe(new Updates(gadget, render)));
    } catch (err) {
      who.className = "who err";
      who.textContent = err.message;
    }
  },
};
