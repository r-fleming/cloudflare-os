# Approvals

People submit requests; an approver other than the requester approves or rejects them. The
History section shows who did what.

## Identity

The gadget never asks who you are. The Workshop tells it who made each call: every method that
acts as someone starts with `this.#me()`, which reads `currentUser()` from `gadgets:user`
(`{id, displayName, role}`) and refuses callers that aren't signed-in people -- the agent, hooks
and other gadgets. Keep it that way when editing:

- Never add a method that takes a user id as a parameter. Read `currentUser()` instead.
- Read it at the top of the method, before any `await` or timer, and pass it down.
- Methods that don't act as anyone (`listRequests()`) are fine for the agent to call.

Display names are chosen by each user and needn't be unique, so the UI shows the start of each
person's id beside their name ("alice (#3f9a1c)"). Compare ids, never names.

## Rules

- `build` users (the owner and anyone who can edit this gadget) are always approvers, and choose
  other approvers from the People list (anyone who has opened the gadget).
- Nobody can decide their own request.

## What these rules are worth

The Workshop vouches for who each user is; the rules above are this gadget's own. They bind `use`
collaborators only. Anyone who can change this gadget's code or data -- `build` users, and any
agent they run -- can bypass them, and can rewrite the stored requests and History. History is
app data, not a tamper-proof audit log.
