// ---------------------------------------------------------------------------
// The client's structure merge: what a tab's pending structure changed relative to what the
// server last held, the update that sends it, and how it is replayed over a collaborator's
// structure. Pure functions over copies, so the client's model is never shared with them.
// ---------------------------------------------------------------------------
import type { SheetMeta, Structure, StructureUpdate } from "./protocol.ts";

/** What a local structure changed relative to its base: `null` for what it left alone. */
export interface StructureChanges {
  title: string | null;
  sheetOrder: string[] | null;
  added: Record<string, SheetMeta>;
  removed: string[];
  changed: Record<string, Partial<SheetMeta>>;
}

export function isEmptyChange(changes: StructureChanges): boolean {
  return changes.title == null && !changes.sheetOrder && !changes.removed.length &&
    !Object.keys(changes.added).length && !Object.keys(changes.changed).length;
}

function copySheetField<K extends keyof SheetMeta>(to: Partial<SheetMeta>, from: SheetMeta, key: K): void { to[key] = from[key]; }

/** What `local` changed relative to `base`, the structure the server last held (none yet: everything is new). */
export function structureChanges(base: Structure | null, local: Structure): StructureChanges {
  const snapshot: Structure = JSON.parse(JSON.stringify(local));
  const from: Structure = base || { title: snapshot.title, sheetOrder: [], sheets: {} };
  const same = (a: unknown, b: unknown) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const changes: StructureChanges = {
    title: snapshot.title !== from.title ? snapshot.title : null,
    sheetOrder: same(snapshot.sheetOrder, from.sheetOrder) ? null : snapshot.sheetOrder,
    added: {},
    removed: from.sheetOrder.filter((id) => !snapshot.sheets[id]),
    changed: {},
  };
  for (const [id, sheet] of Object.entries(snapshot.sheets)) {
    const baseSheet = from.sheets[id];
    if (!baseSheet) { changes.added[id] = sheet; continue; }
    const fields: Partial<SheetMeta> = {};
    const keys = new Set([...Object.keys(sheet), ...Object.keys(baseSheet)] as (keyof SheetMeta)[]);
    for (const key of keys) if (!same(sheet[key], baseSheet[key])) copySheetField(fields, sheet, key);
    if (Object.keys(fields).length) changes.changed[id] = fields;
  }
  return changes;
}

/**
 * The update that sends `local`: only what differs from `base`, since the server keeps every sheet
 * and field an update leaves out. `null` when nothing differs.
 */
export function structurePatch(base: Structure | null, local: Structure): StructureUpdate | null {
  const changes = structureChanges(base, local);
  if (isEmptyChange(changes)) return null;
  const added = Object.keys(changes.added);
  // A removal alone changes the order without reordering anything.
  const kept = (base?.sheetOrder || []).filter((id) => !changes.removed.includes(id));
  const reordered = added.length > 0 || (!!changes.sheetOrder && changes.sheetOrder.join("\n") !== kept.join("\n"));
  return {
    ...(changes.title != null ? { title: changes.title } : {}),
    ...(reordered ? { sheetOrder: local.sheetOrder.slice() } : {}),
    ...(added.length ? { addedSheets: added } : {}),
    ...(changes.removed.length ? { removedSheets: changes.removed } : {}),
    sheets: { ...changes.added, ...changes.changed },
  };
}

/**
 * `remote` with `changes` replayed on top, as the server will hold it once they are saved: the
 * changed fields win, a sheet removed here goes (unless that would leave none), one added here is
 * kept, and one a collaborator removed stays removed.
 */
export function replayStructure(remote: Structure, changes: StructureChanges): Structure {
  const out: Structure = JSON.parse(JSON.stringify(remote));
  const replay: StructureChanges = JSON.parse(JSON.stringify(changes));
  if (replay.title != null) out.title = replay.title;
  for (const [id, sheet] of Object.entries(replay.added)) {
    out.sheets[id] = sheet;
    if (!out.sheetOrder.includes(id)) out.sheetOrder.push(id);
  }
  // Removed after the additions, as the server does, so the last-sheet rule counts them.
  const order = out.sheetOrder.filter((id) => !replay.removed.includes(id));
  if (order.length) {
    for (const id of replay.removed) delete out.sheets[id];
    out.sheetOrder = order;
  }
  for (const [id, fields] of Object.entries(replay.changed)) if (out.sheets[id]) Object.assign(out.sheets[id], fields);
  if (replay.sheetOrder) {
    const listed = replay.sheetOrder.filter((id) => out.sheetOrder.includes(id));
    out.sheetOrder = [...listed, ...out.sheetOrder.filter((id) => !listed.includes(id))];
  }
  return out;
}
