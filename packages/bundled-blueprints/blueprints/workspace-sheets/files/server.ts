import { DurableObject, WorkerEntrypoint } from "cloudflare:workers";
import { MutationQueue, SubscriberRegistry } from "@gadgets/bundled-blueprints/libraries/sync/server";
import { workbookToXlsx } from "./lib/xlsx.ts";
import type {
  Cell,
  CellConflict,
  CellDeletion,
  CellFmt,
  CellMap,
  CellOp,
  CellUpsert,
  CollaboratorInfo,
  Dims,
  DocumentMeta,
  GadgetStub,
  Operation,
  OperationEvent,
  OperationResult,
  PresenceUpdate,
  SheetMeta,
  SheetsDocument,
  SheetsPresenceEvent,
  Structure,
  StructureUpdate,
  SubscriberCallbacks,
} from "./lib/protocol.ts";

// What `applyOperationLocked` hands back: the caller's reply, and the event to broadcast when
// anything changed.
interface LockedOutcome {
  result: OperationResult;
  event?: OperationEvent;
}

// One export option as the platform lists it.
interface ExportFormat {
  id: string;
  label: string;
  mode: "server";
  contentType: string;
  fileExtension: string;
}

const DEFAULT_TITLE = "Untitled spreadsheet";
const DEFAULT_ROWS = 100;
const DEFAULT_COLS = 26;

// ---------------------------------------------------------------------------
// Sheets — authoritative collaboration coordinator.
//
// Storage layout:
//   "meta"          -> { revision, title, sheetOrder:[id], sheets:{id:{...}}, lastModified }
//   "cells:<id>"    -> { "A1": { value, fmt, version }, ... }
//
// Cell content uses per-cell optimistic concurrency (like Docs blocks).
// Structure (sheet list, names, dimensions, column widths, row heights, title)
// is last-writer-wins metadata applied wholesale when a client sends it.
// ---------------------------------------------------------------------------
export class Gadget extends DurableObject<unknown, unknown> {
  declare subscribers: SubscriberRegistry<SubscriberCallbacks, CollaboratorInfo>;
  declare mutations: MutationQueue;

  constructor(ctx: DurableObjectState, env: unknown) {
    super(ctx, env);
    this.ctx = ctx;
    // Presence is announced in this gadget's own callback vocabulary: who arrived, and the bare id
    // of whoever dropped out.
    this.subscribers = new SubscriberRegistry({
      join: (subscriber, who) => subscriber.presence({ type: "join", clientId: who.clientId, name: who.name, color: who.color }),
      leave: (subscriber, who) => subscriber.presence({ type: "leave", clientId: who.clientId }),
    });
    // Overlapping RPC calls are serialized so each observes/commits one
    // authoritative state in strict order. Callbacks to subscribers are never
    // awaited (see the registry), so a callback may itself read or write the document.
    this.mutations = new MutationQueue();
  }

  newId(): string {
    if ((globalThis as typeof globalThis & { crypto?: Crypto }).crypto?.randomUUID) return "s_" + crypto.randomUUID().slice(0, 8);
    return "s_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }

  async loadMeta(): Promise<DocumentMeta> {
    let meta = await this.ctx.storage.get<DocumentMeta>("meta");
    if (!meta) {
      const id = this.newId();
      meta = {
        revision: 0,
        title: DEFAULT_TITLE,
        sheetOrder: [id],
        sheets: { [id]: sheetMeta({ id, name: "Sheet1" }) },
        lastModified: Date.now(),
      };
      await this.ctx.storage.put("meta", meta);
      await this.ctx.storage.put("cells:" + id, {});
      return meta;
    }
    // Every read goes through the same sanitizers as a write: a workbook saved by another version
    // may hold sheet fields this one does not keep, and a structure update is compared against
    // the stored sheets to tell whether it changed anything.
    meta.sheets = normalizeSheets(meta.sheetOrder, meta.sheets);
    return meta;
  }

  async loadCells(id: string): Promise<CellMap> {
    return (await this.ctx.storage.get<CellMap>("cells:" + id)) || {};
  }

  async assembleDocument(meta: DocumentMeta): Promise<SheetsDocument> {
    const cells: Record<string, CellMap> = {};
    for (const id of meta.sheetOrder) cells[id] = await this.loadCells(id);
    return {
      revision: meta.revision,
      title: meta.title,
      sheetOrder: meta.sheetOrder,
      sheets: meta.sheets,
      cells,
      lastModified: meta.lastModified,
    };
  }

  getDocument(): Promise<SheetsDocument> {
    return this.mutations.run(async () => this.assembleDocument(await this.loadMeta()));
  }

  async applyOperation(operation: Operation): Promise<OperationResult> {
    const { result, event } = await this.mutations.run(() => this.applyOperationLocked(operation));
    // Issued after the queue releases, so callbacks may re-enter it, but
    // synchronously here, before the next queued mutation can reach storage,
    // so each subscriber still receives events in revision order.
    if (event) this.broadcast(event);
    return result;
  }

  // Everything is checked before anything is written, so a rejected operation leaves storage as
  // it was.
  async applyOperationLocked(operation: Operation): Promise<LockedOutcome> {
    const meta = await this.loadMeta();
    const rejected = (reason: OperationResult["reason"]): LockedOutcome =>
      ({ result: { status: "rejected", revision: meta.revision, conflicts: [], reason } });

    // --- Structure (last-writer-wins) ------------------------------------
    let nextStructure: Structure | null = null;
    if (operation.structure) {
      const s = operation.structure;
      const order = mergeSheetOrder(meta, s);
      const merged: Record<string, Partial<SheetMeta>> = {};
      for (const id of order) {
        const incoming: Partial<SheetMeta> = s.sheets?.[id] || {};
        const existing: Partial<SheetMeta> = meta.sheets[id] || {};
        // A key the client left out (or sent as null) keeps the stored value.
        const sheet: Record<string, unknown> = { ...existing, id };
        for (const [key, value] of Object.entries(incoming)) {
          if (key !== "id" && value != null) sheet[key] = value;
        }
        merged[id] = sheet;
      }
      const nextSheets = normalizeSheets(order, merged);
      const title = typeof s.title === "string" ? s.title.slice(0, 200) || DEFAULT_TITLE : meta.title;
      const structureChanged = title !== meta.title || JSON.stringify([order, nextSheets]) !== JSON.stringify([meta.sheetOrder, meta.sheets]);
      if (structureChanged) {
        // An edit that shrinks an oversized meta is let through.
        const length = JSON.stringify({ ...meta, sheetOrder: order, sheets: nextSheets, title }).length;
        if (length > MAX_META_CHARACTERS && length > JSON.stringify(meta).length) return rejected("too-large");
        nextStructure = { sheetOrder: order, sheets: nextSheets, title };
      }
    }
    const sheets = nextStructure?.sheets ?? meta.sheets;

    // --- Whole-sheet cell replacement (sort, clear, insert/delete) ------
    const replacements: { id: string; cells: CellMap; floor: number }[] = [];
    // A sheet replaced twice in one operation keeps the last replacement.
    const lastReplacements = new Map((operation.sheetReplacements || []).map((rep) => [String(rep.sheetId || ""), rep]));
    for (const [id, rep] of lastReplacements) {
      if (!sheets[id]) continue;
      const stored = await this.loadCells(id);
      if (rep.baseVersions && !sameVersions(stored, rep.baseVersions)) return rejected("stale");
      let floor = 0;
      for (const cell of Object.values(stored)) floor = Math.max(floor, cell.version);
      replacements.push({ id, cells: sanitizeCellMap(rep.cells), floor });
    }

    // --- Per-cell operations (optimistic concurrency) --------------------
    const upserts: CellUpsert[] = [];
    const deletes: CellDeletion[] = [];
    const conflicts: CellConflict[] = [];
    const bySheet = new Map<string, CellOp[]>();
    for (const op of operation.cellOps || []) {
      const sid = String(op.sheetId || "");
      if (!bySheet.has(sid)) bySheet.set(sid, []);
      bySheet.get(sid)!.push(op);
    }
    const editedCells = new Map<string, CellMap>();
    for (const [sheetId, ops] of bySheet) {
      if (!sheets[sheetId]) continue;
      const cells = replacements.find((rep) => rep.id === sheetId)?.cells ?? await this.loadCells(sheetId);
      let dirty = false;
      for (const op of ops) {
        const ref = String(op.ref || "");
        if (!/^[A-Z]+[0-9]+$/.test(ref)) continue;
        const cur = cells[ref];
        const base = Number(op.baseVersion || 0);
        const isDelete = op.value == null && op.fmt == null;
        if (isDelete) {
          if (!cur) continue;
          if (cur.version !== base) { conflicts.push({ sheetId, ref, cell: cur }); continue; }
          delete cells[ref];
          deletes.push({ sheetId, ref });
          dirty = true;
        } else {
          if (cur && cur.version !== base) { conflicts.push({ sheetId, ref, cell: cur }); continue; }
          const next: Cell = {
            value: op.value == null ? "" : String(op.value).slice(0, 8192),
            fmt: sanitizeFmt(op.fmt),
            version: (cur?.version || 0) + 1,
          };
          cells[ref] = next;
          upserts.push({ sheetId, ref, cell: next });
          dirty = true;
        }
      }
      if (dirty) editedCells.set(sheetId, cells);
    }

    // Every cell of a replaced sheet moves past every version the sheet held, so a cell op based on
    // the old layout -- a peer's edit sent before the replacement reached it -- conflicts instead
    // of landing on whatever cell the replacement moved to its coordinate. The cell ops above were
    // checked against the versions the replacement carried, since the sender queued them on top of
    // it; the cells they wrote are bumped in place, so the upserts reported carry the new version.
    // A collaborator's edit to a position the replacement emptied still lands; telling it from a
    // new cell would take a version for every position, not just every cell.
    for (const rep of replacements) {
      let floor = rep.floor;
      for (const cell of Object.values(rep.cells)) floor = Math.max(floor, cell.version);
      for (const cell of Object.values(rep.cells)) cell.version = floor + 1;
    }

    // A client that sent a structure is told what the server holds, changed or not, since the
    // server may have kept less than it sent (a sheet it would not restore, the last one it would
    // not delete): the client diffs its next update against this.
    const currentStructure = (): { structure?: Structure } =>
      operation.structure ? { structure: { sheetOrder: meta.sheetOrder, sheets: meta.sheets, title: meta.title } } : {};
    if (!nextStructure && !replacements.length && !upserts.length && !deletes.length) {
      return { result: { status: conflicts.length ? "conflict" : "unchanged", revision: meta.revision, conflicts, ...currentStructure() } };
    }

    // --- Commit ------------------------------------------------------------
    // In one transaction, so a write that fails leaves storage as it was.
    const writes = new Map<string, CellMap>();
    const removedSheets: string[] = [];
    if (nextStructure) {
      for (const id of nextStructure.sheetOrder) if (!meta.sheets[id]) writes.set(id, {});
      for (const id of meta.sheetOrder) if (!nextStructure.sheets[id]) removedSheets.push(id);
      Object.assign(meta, nextStructure);
    }
    for (const [id, cells] of editedCells) writes.set(id, cells);
    for (const rep of replacements) writes.set(rep.id, rep.cells);
    meta.revision += 1;
    meta.lastModified = Date.now();
    await this.ctx.storage.transaction(async (txn) => {
      for (const id of removedSheets) await txn.delete("cells:" + id);
      for (const [id, cells] of writes) await txn.put("cells:" + id, cells);
      await txn.put("meta", meta);
    });

    // The structure goes out only when it changed, and a replaced sheet goes out whole (below),
    // with the versions the server assigned.
    const event: OperationEvent = {
      type: "operation",
      senderId: operation.senderId,
      revision: meta.revision,
      upserts,
      deletes,
      replacedSheets: replacements.map((rep) => rep.id),
      lastModified: meta.lastModified,
    };
    if (nextStructure) event.structure = { sheetOrder: meta.sheetOrder, sheets: meta.sheets, title: meta.title };
    if (replacements.length) event.replacedCells = Object.fromEntries(replacements.map((rep) => [rep.id, rep.cells]));
    return { result: { status: conflicts.length ? "conflict" : "applied", ...event, ...currentStructure(), conflicts }, event };
  }

  // --- Presence & subscription ------------------------------------------
  async subscribe(callback: SubscriberCallbacks, client: Partial<CollaboratorInfo> = {}): Promise<SheetsDocument> {
    const info: CollaboratorInfo = {
      clientId: String(client.clientId || ""),
      name: String(client.name || "Guest").slice(0, 40),
      color: String(client.color || "#e1632e"),
    };
    // Registering and snapshotting inside the queue means the subscriber sees
    // every operation committed after its snapshot, and none before it. The
    // registry duplicates the callback only once the snapshot exists, so a
    // failed read leaves nothing to dispose; it then seeds the newcomer with
    // everyone here and announces it, once add() has returned.
    return this.mutations.run(async () => {
      const document = await this.assembleDocument(await this.loadMeta());
      this.subscribers.add(callback, info);
      return document;
    });
  }

  updatePresence(presence: Partial<PresenceUpdate>): void {
    this.broadcastPresence({
      type: "cursor",
      clientId: String(presence.clientId || ""),
      name: String(presence.name || "Guest").slice(0, 40),
      color: String(presence.color || "#e1632e"),
      sheetId: presence.sheetId ? String(presence.sheetId) : null,
      r1: int(presence.r1), c1: int(presence.c1),
      r2: int(presence.r2), c2: int(presence.c2),
      at: Date.now(),
    });
  }

  leavePresence(clientId: string): void {
    this.broadcastPresence({ type: "leave", clientId: String(clientId || ""), at: Date.now() });
  }

  // Delivery is the registry's: best-effort and not awaited, so a callback
  // that fails is dropped, and one that never settles holds up nothing but
  // its own client.
  broadcast(event: OperationEvent): void {
    this.subscribers.broadcast((each) => each.operation(event));
  }

  broadcastPresence(event: SheetsPresenceEvent): void {
    this.subscribers.broadcast((each) => each.presence(event));
  }
}

// --- Sanitizers -------------------------------------------------------------
function int(v: unknown): number { const n = Math.round(Number(v)); return Number.isFinite(n) ? Math.max(0, n) : 0; }
function clampInt(v: unknown, lo: number, hi: number, dflt: number): number {
  const n = Math.round(Number(v));
  if (!Number.isFinite(n)) return dflt;
  return Math.max(lo, Math.min(hi, n));
}

// Sizes of the first `count` rows or columns.
function sanitizeDims(dims: unknown, count: number): Dims {
  const out: Dims = {};
  if (dims && typeof dims === "object") {
    for (const [k, v] of Object.entries(dims as Record<string, unknown>)) {
      if (!/^\d+$/.test(k) || Number(k) >= count) continue;
      const n = Math.round(Number(v));
      if (Number.isFinite(n) && n >= 8 && n <= 2000) out[k] = n;
    }
  }
  return out;
}

function sheetMeta(s: Partial<SheetMeta> & Pick<SheetMeta, "id">): SheetMeta {
  const rows = clampInt(s.rows, 1, 50000, DEFAULT_ROWS);
  const cols = clampInt(s.cols, 1, 702, DEFAULT_COLS);
  return {
    id: String(s.id),
    name: String(s.name || "Sheet").slice(0, 60),
    rows,
    cols,
    colWidths: sanitizeDims(s.colWidths, cols),
    rowHeights: sanitizeDims(s.rowHeights, rows),
    frozenRows: clampInt(s.frozenRows, 0, 50, 0),
    frozenCols: clampInt(s.frozenCols, 0, 50, 0),
  };
}

// The sheets named by `order`, sanitized.
function normalizeSheets(order: string[], sheets: Record<string, Partial<SheetMeta>>): Record<string, SheetMeta> {
  const out: Record<string, SheetMeta> = {};
  for (const id of order) out[id] = sheetMeta({ ...sheets[id], id });
  return out;
}

// The sheet order after `update`. A sheet the client lists that the server lacks is kept only when
// the update adds it, so a sheet a collaborator deleted is not brought back empty; one the server
// holds that the client leaves out is kept unless the client removed it, so a sheet a collaborator
// added meanwhile is not deleted. Those go last, as does an added sheet the order leaves out.
function mergeSheetOrder(meta: DocumentMeta, update: StructureUpdate): string[] {
  const ids = (list: unknown) => new Set(Array.isArray(list) ? list.map(String) : []);
  const removed = ids(update.removedSheets), added = ids(update.addedSheets);
  if (!Array.isArray(update.sheetOrder) && !removed.size && !added.size) return meta.sheetOrder;
  const listed = Array.isArray(update.sheetOrder) ? update.sheetOrder.map(String) : meta.sheetOrder;
  const order = [...new Set(listed)].filter((id) => !removed.has(id) && (meta.sheets[id] || added.has(id)));
  for (const id of [...meta.sheetOrder, ...added]) if (!removed.has(id) && !order.includes(id)) order.push(id);
  // The last sheet cannot be removed.
  return order.length ? order : meta.sheetOrder;
}

// Whether `cells` holds exactly the cells `versions` lists, at those versions.
function sameVersions(cells: CellMap, versions: Record<string, number>): boolean {
  const refs = Object.keys(cells);
  return refs.length === Object.keys(versions).length && refs.every((ref) => cells[ref].version === versions[ref]);
}

// The stored `meta` is one storage value, capped at 2 MB; a string costs up to two bytes per
// character there. Every sheet's metadata lives in it, so this bounds how much a workbook holds.
const MAX_META_CHARACTERS = 900000;

const FMT_KEYS = new Set(["b", "i", "u", "s", "c", "bg", "a", "nf", "d", "fs", "wrap"]);
function isFmtKey(k: string): k is keyof CellFmt { return FMT_KEYS.has(k); }
function sanitizeFmt(fmt: unknown): CellFmt | null {
  if (!fmt || typeof fmt !== "object") return null;
  const out: CellFmt = {};
  for (const [k, v] of Object.entries(fmt as Record<string, unknown>)) {
    if (!isFmtKey(k) || v == null || v === false || v === "") continue;
    // A colour is a string or nothing: `RegExp.test` would stringify an array like `["#abc"]` into
    // a match and store the array itself.
    if (k === "c" || k === "bg") { if (typeof v === "string" && /^#[0-9a-fA-F]{3,8}$/.test(v)) out[k] = v; }
    else if (k === "a") { if (v === "l" || v === "c" || v === "r") out[k] = v; }
    else if (k === "nf") { out[k] = String(v).slice(0, 20); }
    else if (k === "d") { const n = Math.round(Number(v)); if (n >= 0 && n <= 10) out[k] = n; }
    else if (k === "fs") { const n = Math.round(Number(v)); if (n >= 6 && n <= 96) out[k] = n; }
    else out[k] = true;
  }
  return Object.keys(out).length ? out : null;
}

const MAX_CELL_VERSION = 2 ** 40;
function sanitizeCellMap(map: unknown): CellMap {
  const out: CellMap = {};
  if (!map || typeof map !== "object") return out;
  let count = 0;
  for (const [ref, cell] of Object.entries(map as Record<string, Partial<Record<keyof Cell, unknown>> | null | undefined>)) {
    if (!/^[A-Z]+[0-9]+$/.test(ref) || count++ > 200000) continue;
    out[ref] = {
      value: cell?.value == null ? "" : String(cell.value).slice(0, 8192),
      fmt: sanitizeFmt(cell?.fmt),
      // Bounded, so `version + 1` always moves on (a client sends versions in a replacement).
      version: Math.min(MAX_CELL_VERSION, Math.max(1, Math.round(Number(cell?.version)) || 1)),
    };
  }
  return out;
}

const CSV_FORMAT_PREFIX = "csv:";
const MAX_CSV_SHEETS = 31; // The platform allows 32 formats; one is the workbook.
const MAX_EXPORT_ID_LENGTH = 128;
const XLSX_FORMAT: ExportFormat = {
  id: "xlsx",
  label: "Excel Workbook",
  mode: "server",
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  fileExtension: ".xlsx",
};

// Sheet ids are client-chosen, so duplicates and over-long ids are possible in
// stored structure. Either would fail format validation and disable every export.
function csvSheetIds(document: SheetsDocument): string[] {
  const ids = document.sheetOrder.filter((id) => (CSV_FORMAT_PREFIX + id).length <= MAX_EXPORT_ID_LENGTH);
  return Array.from(new Set(ids)).slice(0, MAX_CSV_SHEETS);
}

export class ExportHandler extends WorkerEntrypoint {
  async getExportFormats(gadget: GadgetStub): Promise<ExportFormat[]> {
    const document = await gadget.getDocument();
    const sheetIds = csvSheetIds(document);
    return [XLSX_FORMAT, ...sheetIds.map((sheetId): ExportFormat => ({
      id: CSV_FORMAT_PREFIX + sheetId,
      label: sheetIds.length === 1 ? "CSV" : "CSV (" + document.sheets[sheetId].name + ")",
      mode: "server",
      contentType: "text/csv",
      fileExtension: ".csv",
    }))];
  }

  async export(gadget: GadgetStub, id: string): Promise<ReadableStream<Uint8Array>> {
    if (id === XLSX_FORMAT.id) {
      const document = await gadget.getDocument();
      return workbookToXlsx(document);
    }
    if (!id.startsWith(CSV_FORMAT_PREFIX)) {
      throw new Error("Unsupported spreadsheet export format: " + id);
    }
    const document = await gadget.getDocument();
    const sheetId = id.slice(CSV_FORMAT_PREFIX.length);
    if (!csvSheetIds(document).includes(sheetId)) {
      throw new Error("The selected worksheet is unavailable for CSV export.");
    }
    return new Response(workbookSheetToCsv(document, sheetId)).body!;
  }
}

function workbookSheetToCsv(document: SheetsDocument, sheetId: string): string {
  const cells = document.cells?.[sheetId] || {};
  let maxRow = -1;
  let maxColumn = -1;
  for (const [ref, cell] of Object.entries(cells)) {
    if (cell?.value == null || cell.value === "") continue;
    const position = parseCsvCellRef(ref);
    if (!position) continue;
    maxRow = Math.max(maxRow, position.row);
    maxColumn = Math.max(maxColumn, position.column);
  }
  if (maxRow < 0 || maxColumn < 0) return "";

  const rows: string[] = [];
  for (let row = 0; row <= maxRow; ++row) {
    const fields: string[] = [];
    for (let column = 0; column <= maxColumn; ++column) {
      const value = cells[csvCellRef(row, column)]?.value ?? "";
      fields.push(escapeCsvField(value));
    }
    rows.push(fields.join(","));
  }
  return rows.join("\r\n") + "\r\n";
}

function parseCsvCellRef(ref: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(ref);
  if (!match) return null;
  let column = 0;
  for (const char of match[1]) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function csvCellRef(row: number, column: number): string {
  let letters = "";
  for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
  }
  return letters + (row + 1);
}

function escapeCsvField(value: string): string {
  const text = String(value);
  return /[",\r\n]/.test(text) ? "\"" + text.replace(/\"/g, "\"\"") + "\"" : text;
}
