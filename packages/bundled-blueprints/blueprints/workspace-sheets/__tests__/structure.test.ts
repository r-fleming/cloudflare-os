// @vitest-environment node
import { describe, expect, it } from "vitest";
import type { SheetMeta, Structure } from "../files/lib/protocol.ts";
import { replayStructure, structureChanges, structurePatch } from "../files/lib/structure.ts";

function sheet(id: string, extra: Partial<SheetMeta> = {}): SheetMeta {
  return { id, name: id, rows: 100, cols: 26, colWidths: {}, rowHeights: {}, frozenRows: 0, frozenCols: 0, ...extra };
}

function structure(...sheets: SheetMeta[]): Structure {
  return { title: "Book", sheetOrder: sheets.map((each) => each.id), sheets: Object.fromEntries(sheets.map((each) => [each.id, each])) };
}

describe("structure patches", () => {
  it("sends only the fields a tab changed", () => {
    const base = structure(sheet("a"), sheet("b"));
    const local = structure(sheet("a", { name: "Renamed" }), sheet("b"));
    expect(structurePatch(base, local)).toEqual({ sheets: { a: { name: "Renamed" } } });
    expect(structurePatch(base, structure(sheet("a"), sheet("b")))).toBeNull();
  });

  it("lists added and removed sheets explicitly, and the order only when it changed", () => {
    const base = structure(sheet("a"), sheet("b"));
    const local = structure(sheet("c"), sheet("a"));
    expect(structurePatch(base, local)).toEqual({
      sheetOrder: ["c", "a"], addedSheets: ["c"], removedSheets: ["b"], sheets: { c: sheet("c") },
    });
    expect(structurePatch(base, structure(sheet("a")))).toEqual({ removedSheets: ["b"], sheets: {} });
  });
});

describe("replaying pending changes over a collaborator's structure", () => {
  it("keeps both tabs' added sheets and both tabs' field changes", () => {
    const base = structure(sheet("a", { rows: 100 }));
    const local = structure(sheet("a", { name: "Mine" }), sheet("mine"));
    const remote = structure(sheet("a", { rows: 500 }), sheet("theirs"));
    expect(replayStructure(remote, structureChanges(base, local))).toEqual({
      title: "Book",
      sheetOrder: ["a", "mine", "theirs"],
      sheets: { a: sheet("a", { name: "Mine", rows: 500 }), theirs: sheet("theirs"), mine: sheet("mine") },
    });
  });

  it("lets this tab's value win a field both changed, and leaves a sheet the collaborator deleted deleted", () => {
    const base = structure(sheet("a"), sheet("b"));
    const local = structure(sheet("a", { name: "Mine" }), sheet("b", { name: "Edited" }));
    const remote = structure(sheet("a", { name: "Theirs" }));
    expect(replayStructure(remote, structureChanges(base, local))).toEqual(structure(sheet("a", { name: "Mine" })));
  });

  it("keeps the last sheet when both tabs delete the others", () => {
    const base = structure(sheet("a"), sheet("b"));
    const local = structure(sheet("b"));
    const remote = structure(sheet("a"));
    expect(replayStructure(remote, structureChanges(base, local)).sheetOrder).toEqual(["a"]);
  });

  it("removes every old sheet when this tab added one to take their place", () => {
    const base = structure(sheet("a"), sheet("b"));
    const local = structure(sheet("c"));
    const remote = structure(sheet("a", { name: "Renamed" }), sheet("b"));
    expect(replayStructure(remote, structureChanges(base, local))).toEqual(structure(sheet("c")));
  });

  it("applies this tab's order to the sheets that still exist", () => {
    const base = structure(sheet("a"), sheet("b"), sheet("c"));
    const local = structure(sheet("c"), sheet("b"), sheet("a"));
    const remote = structure(sheet("a"), sheet("c"), sheet("new"));
    expect(replayStructure(remote, structureChanges(base, local)).sheetOrder).toEqual(["c", "a", "new"]);
  });

  it("shares no objects with its inputs", () => {
    const base = structure(sheet("a"));
    const local = structure(sheet("a", { colWidths: { 0: 90 } }));
    const remote = structure(sheet("a"));
    const merged = replayStructure(remote, structureChanges(base, local));
    merged.sheets.a.colWidths[1] = 50;
    expect(local.sheets.a.colWidths).toEqual({ 0: 90 });
    expect(remote.sheets.a.colWidths).toEqual({});
  });
});
