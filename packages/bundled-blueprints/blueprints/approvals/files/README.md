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

Display names are chosen by each user and needn't be unique, so the UI shows the start of each
person's viewer id beside their name ("alice (#3f9a1c)"). Compare ids, never names.

## Rules

- `build` viewers (the owner and anyone who can edit this gadget) are always approvers, and choose
  other approvers from the People list (anyone who has opened the gadget).
- Nobody can decide their own request.

## What these rules are worth

The Workshop vouches for who each viewer is; the rules above are this gadget's own. They bind
`use` collaborators only. Anyone who can change this gadget's code or data -- `build` viewers, and
any agent they run -- can bypass them, and can rewrite the stored requests and History. History is
app data, not a tamper-proof audit log.

## Without viewer identity

On a Workshop that doesn't call `connectViewer()`, the browser's `gadget` stub is the `Gadget`
itself, which has no `whoami()`. The client then shows a "Who are you?" form and calls
`connectViewer()` itself with whatever name and role the person types -- so anyone can claim to be
anyone, including a `build` approver deciding their own request. That is the point of the
comparison: without the Workshop vouching for viewers, an app can only take the browser's word.

On a Workshop that does introduce viewers, this path never runs, and couldn't work if it did: the
Workshop refuses `connectViewer()` from everyone but itself.
