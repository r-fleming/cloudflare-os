# Approvals

People submit requests; an approver other than the requester approves or rejects them. The audit
trail records who did what.

## Identity

The gadget never asks who you are. The Workshop calls `Gadget.connectViewer(viewer)` once per
connection with the signed-in person (`{id, displayName, role}`), and the browser's `gadget` stub is
the `ApprovalSession` returned for that viewer. Keep it that way when editing:

- Never add a method (on `Gadget` or `ApprovalSession`) that takes a user or viewer id as a
  parameter. Session actions are bound to their viewer in `connectViewer()`.
- Public `Gadget` methods are reachable by the agent and other gadgets, which are not viewers, so
  they must not act as anyone (`listRequests()` only reads).

## Rules

- The owner is always an approver and chooses other approvers from the People list (anyone who has
  opened the gadget).
- Nobody can decide their own request.
- These rules bind `use` collaborators. `owner` and `build` viewers can edit this code.
