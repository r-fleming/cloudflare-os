# Approvals

People submit requests; an approver other than the requester approves or rejects them. The
History section shows who did what.

## Identity

The gadget never asks who you are. The Workshop calls `Gadget.connectViewer(viewer)` once per
connection with the signed-in person (`{id, displayName, role}`), and the browser's `gadget` stub is
the `ApprovalSession` returned for that viewer. Keep it that way when editing:

- Never add a method (on `Gadget` or `ApprovalSession`) that takes a user or viewer id as a
  parameter. Session actions are bound to their viewer in `connectViewer()`.
- Public `Gadget` methods are reachable by the agent and other gadgets, which are not viewers, so
  they must not act as anyone (`listRequests()` only reads).

## Rules

- `build` viewers (the owner and anyone who can edit this gadget) are always approvers, and choose
  other approvers from the People list (anyone who has opened the gadget).
- Nobody can decide their own request.

## What these rules are worth

The Workshop vouches for who each viewer is; the rules above are this gadget's own. They bind
`use` collaborators only. Anyone who can change this gadget's code or data -- `build` viewers, and
any agent they run -- can bypass them, and can rewrite the stored requests and History. History is
app data, not a tamper-proof audit log.
