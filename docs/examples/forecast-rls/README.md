# Demo: per-caller forecast (`connectCaller`)

A renewals forecast where each "use" collaborator sees only their branch of the org tree, and can save only rows they own. Owners, editors, the agent and bound gadgets see everything. See [Caller sessions](../../sharing.md#caller-sessions-connectcaller).

The org tree is keyed by email: the same address Cloudflare Access gives the Workshop and that SFDC, Workday or Backstage use. A real gadget would load the tree and rows from those systems through its bindings, as the owner. The sample data uses `@example.com` addresses.

## Run it

1. Start a local deployment from this branch and sign in as the owner.
2. Create a gadget with this `server.js` and `client.js`. The owner sees every row, editable, and the header reads `role: owner`.
3. From chat, point the org tree at real test accounts. `executeCode` reaches the agent's (admin) session:
   ```js
   let view = await env.FORECAST.rpc.getView();
   // ...edit an org entry's email to a second local account, e.g. amer-lead -> bob@local
   await env.FORECAST.rpc.setData(org, rows);
   ```
4. Share the gadget with that second account as **Can use**, and open it as them:
   - Only the AMER branch shows: their own row (editable), and their reps' rows (read only).
   - Saving another person's row is refused by the session.
   - Changing the owner's view live updates both windows, each with only its own rows.
   - **Try calling the Gadget directly** shows `connectCaller() is reserved for the Workshop`. In devtools, any method the session lacks (e.g. `setData`) fails too.
5. Share as **Can view** (snapshot) instead: it shows the same branch only, read only, though the snapshot holds every row.
6. Share with an account not in the org tree: the gadget refuses to open for them.
