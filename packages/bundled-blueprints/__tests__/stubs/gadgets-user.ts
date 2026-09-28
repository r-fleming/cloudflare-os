// Stand-in for the `gadgets:user` module, which only the Workshop's gadget loader supplies.
//
// No user is ever calling under plain vitest; a test that needs one mocks this module
// (`vi.mock("gadgets:user", ...)`). The real module is exercised by the Workshop backend's workerd
// suite.

export function currentUser(): null {
  return null;
}
