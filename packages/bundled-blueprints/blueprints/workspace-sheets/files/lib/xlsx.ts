import { createZip, type ZipEntry } from "./zip.ts";
import { CHART_BOUNDS, CHART_TYPES } from "./limits.ts";
import type { CellFmt, CellMap, ChartType, Dims, SheetMeta, SheetsDocument } from "./protocol.ts";

const encoder = new TextEncoder();
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;
const DEFAULT_ROWS = 100;
const DEFAULT_COLUMNS = 26;
const DEFAULT_ROW_PIXELS = 24;
const DEFAULT_COLUMN_PIXELS = 92;
const MAX_FONTS = 512;
const MAX_FILLS = 256;
const MAX_CELL_FORMATS = 65490;
const MAX_FORMULA_CHARACTERS = 8192;
const TEXT_CHUNK_SIZE = 64 * 1024;
// Excel's limits: series per chart, hyperlinks per sheet, a hyperlink target's length, and a
// note's text. A chart's extra series and the sheet's extra links are left out of the export
// (they stay in the workbook itself); a longer note is cut at the limit.
const MAX_SERIES = 255;
const MAX_HYPERLINKS = 65530;
const MAX_HYPERLINK_CHARACTERS = 2079;
const MAX_COMMENT_CHARACTERS = 32767;
const EMU_PER_PIXEL = 9525;
// The grid positions charts over its scroll area, so stored coordinates include the row-header
// column and the column-header row.
const ROW_HEADER_PIXELS = 44;
const COLUMN_HEADER_PIXELS = 22;
const CHART_COLORS = ["E1632E", "3478C7", "1F9D77", "8B5FBF", "C49324", "C4566A"];
// The grid draws links in this colour, underlined.
const HYPERLINK_FMT: CellFmt = {c: "#1967d2", u: true};
const FUTURE_FUNCTIONS = new Set([
  "CONCAT", "DAYS", "IFNA", "IFS", "SWITCH", "TEXTJOIN", "UNICHAR", "UNICODE", "XOR",
]);

/** A worksheet as `sourceSheets` reads it out of the document; `assignSheetNames` then gives it
 * its Excel-safe name. */
interface SheetDraft {
  id: string;
  sourceName: string;
  metadata: Partial<SheetMeta>;
  sourceCells: CellMap;
  name?: string;
}

/** A draft once `assignSheetNames` has run. */
type NamedSheet = SheetDraft & { name: string };

/** A worksheet once `prepareWorkbook` has filled in everything the XML needs. */
interface PreparedSheet {
  name: string;
  metadata: Partial<SheetMeta>;
  rows: number;
  columns: number;
  frozenRows: number;
  frozenColumns: number;
  columnWidths: Dimension[];
  rowHeights: Dimension[];
  cells: PreparedCell[];
  hyperlinks: PreparedHyperlink[];
  filter: PreparedFilter | null;
  /** One-based rows the filter hides, ascending. */
  hiddenRows: number[];
  drawing: PreparedDrawing | null;
  comments: PreparedComments | null;
}

/** A literal URL cell to link, and the worksheet relationship that holds its target. */
interface PreparedHyperlink {
  reference: string;
  row: number;
  column: number;
  target: string;
  relationshipId: string;
}

/** One `filterColumn` of an autofilter: a column outside the table only hides its button. */
type FilterColumn =
  | { offset: number; hiddenButton: true }
  | { offset: number; hiddenButton?: undefined; values: string[]; blank: boolean };

/** What a sheet's filter exports as: the autofilter, its defined name, and the rows it hides. */
interface PreparedFilter {
  ref: string;
  definedName: string;
  columns: FilterColumn[];
  hiddenRows: number[];
  sort: { ref: string; column: string; descending: boolean } | null;
}

/** A filter column's selection: the grid's tokens, and the Excel values they stand for. */
interface FilterCriterion {
  tokens: Set<string>;
  values: string[];
  blank: boolean;
  selectsNothing: boolean;
}

/** A one-based range of cells. */
interface CellRange {
  firstRow: number;
  lastRow: number;
  firstColumn: number;
  lastColumn: number;
}

/** A drawing anchor along one axis: the zero-based cell and the EMU offset into it. */
interface AnchorPosition {
  index: number;
  offset: number;
}

/** A chart ready to write. Rows and columns are one-based; `headerRow` and `categoryColumn` are
 * `null` when the chart has no series names or no category labels. */
interface PreparedChart {
  index: number;
  type: ChartType;
  title: string;
  xAxisTitle: string;
  yAxisTitle: string;
  legend: boolean;
  headerRow: number | null;
  categoryColumn: number | null;
  firstRow: number;
  lastRow: number;
  columns: number[];
  column: AnchorPosition;
  row: AnchorPosition;
  extent: { cx: number; cy: number };
}

/** A sheet's drawing part and the charts it anchors. */
interface PreparedDrawing {
  index: number;
  relationshipId: string;
  charts: PreparedChart[];
}

/** One note to write, at a one-based position. */
interface PreparedComment {
  reference: string;
  row: number;
  column: number;
  text: string;
}

/** A sheet's comments part, its VML drawing, and the first VML shape-id block it uses. */
interface PreparedComments {
  index: number;
  vmlRelationshipId: string;
  relationshipId: string;
  firstShapeBlock: number;
  list: PreparedComment[];
}

/** How many of each numbered part the workbook holds. */
interface PartCounts {
  drawings: number;
  charts: number;
  comments: number;
  vmlShapeBlocks: number;
}

/** A column width or row height to write: the zero-based index and the converted value. */
interface Dimension {
  index: number;
  value: string;
}

/** A non-empty cell ready to write: its position, raw text and style id. */
interface PreparedCell {
  reference: string;
  row: number;
  column: number;
  value: string;
  style: number;
}

/** A sheet or function reference recognized in a formula: where it ends and what to write instead. */
interface FormulaRewrite {
  end: number;
  text: string;
}

/** A cell's text as Excel should read it. */
type ParsedCellValue =
  | { type: "text" | "formula" | "blank"; value: string }
  | { type: "boolean"; value: boolean }
  | { type: "number"; value: number };

interface Font {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string | null;
  size: number | null;
}

interface Fill {
  color?: string;
  gray125?: boolean;
}

interface Alignment {
  horizontal: string | null;
  wrap: boolean;
}

interface CellFormat {
  fontId: number;
  fillId: number;
  numberFormatId: number;
  alignmentId: number;
}

function spreadsheetXml(value: unknown, attribute = false): string {
  const input = String(value).replace(/_x[0-9a-f]{4}_/gi, (match) => "_x005F_" + match.slice(1));
  let clean = "";
  for (let i = 0; i < input.length; ++i) {
    const code = input.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const low = input.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) clean += input[i] + input[++i];
      else clean += "_xFFFD_";
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      clean += "_xFFFD_";
    } else if (code === 13) {
      clean += "_x000D_";
    } else if (code === 9 || code === 10 ||
        (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd)) {
      clean += input[i];
    } else {
      clean += `_x${code.toString(16).toUpperCase().padStart(4, "0")}_`;
    }
  }
  clean = clean.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  if (attribute) clean = clean.replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  return clean;
}

function formulaXml(value: string): string {
  const input = String(value);
  let clean = "";
  for (let i = 0; i < input.length; ++i) {
    const code = input.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const low = input.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) clean += input[i] + input[++i];
      else clean += String.fromCharCode(0xfffd);
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      clean += String.fromCharCode(0xfffd);
    } else if (code === 9 || code === 10 || code === 13 ||
        (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd)) {
      clean += input[i];
    } else {
      clean += String.fromCharCode(0xfffd);
    }
  }
  return clean.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/\r/g, "&#13;");
}

function xmlAttribute(value: string): string {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

// DrawingML text has no `_xHHHH_` escapes: characters XML 1.0 forbids are dropped instead.
function plainXml(value: string): string {
  return String(value).replace(/[\0-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]/g, "")
    .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, "\ufffd")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Encodes a string generator into ~64 KiB byte chunks. Cell-sized chunks would make the ZIP's
// CompressionStream the bottleneck. `highWaterMark: 0` keeps generation lazy until the archive
// reaches this part.
function textStream(generator: Generator<string, void, unknown>): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      const parts: string[] = [];
      let length = 0;
      while (length < TEXT_CHUNK_SIZE) {
        const result = generator.next();
        if (result.done) {
          if (parts.length) controller.enqueue(encoder.encode(parts.join("")));
          controller.close();
          return;
        }
        parts.push(result.value);
        length += result.value.length;
      }
      controller.enqueue(encoder.encode(parts.join("")));
    },
    cancel(reason) {
      generator.return(reason);
    },
  }, {highWaterMark: 0});
}

function count(value: unknown, fallback: number, maximum: number): number {
  const number = Math.round(Number(value));
  if (!Number.isFinite(number)) return fallback;
  return Math.max(1, Math.min(maximum, number));
}

function frozenCount(value: unknown, maximum: number): number {
  const number = Math.round(Number(value));
  if (!Number.isFinite(number)) return 0;
  return Math.max(0, Math.min(50, maximum, number));
}

function truncateSheetName(value: string, length: number): string {
  const input = value.slice(0, length);
  let result = "";
  for (let i = 0; i < input.length; ++i) {
    const code = input.charCodeAt(i);
    if (code < 32 || (code >= 127 && code <= 159)) {
      result += "_";
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const low = input.charCodeAt(i + 1);
      if (low >= 0xdc00 && low <= 0xdfff) result += input[i] + input[++i];
      else result += "_";
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      result += "_";
    } else {
      result += input[i];
    }
  }
  return result;
}

function safeSheetName(value: unknown): string {
  let name = String(value ?? "").replace(/[:\\/?*\[\]]/g, "_").trim();
  name = truncateSheetName(name, 31);
  if (name.startsWith("'")) name = "_" + name.slice(1);
  if (name.endsWith("'")) name = name.slice(0, -1) + "_";
  if (name.toLowerCase() === "history") name += "_";
  return name || "Sheet";
}

function assignSheetNames(sheets: SheetDraft[]): asserts sheets is NamedSheet[] {
  const used = new Set<string>();
  // Next unused suffix per truncated stem (keyed with the suffix's digit count, since the stem
  // shrinks to make room), so N same-named sheets take O(N) probes rather than O(N²).
  const nextSuffixes = new Map<string, number>();
  for (const sheet of sheets) {
    const base = safeSheetName(sheet.sourceName);
    let name = base;
    let suffix = 2;
    while (used.has(name.toLowerCase())) {
      const digits = String(suffix).length;
      const stem = truncateSheetName(base, 28 - digits);
      const key = `${stem.toLowerCase()}|${digits}`;
      const next = nextSuffixes.get(key) ?? suffix;
      if (next > suffix) {
        suffix = next;
        continue;
      }
      name = `${stem} (${suffix})`;
      nextSuffixes.set(key, ++suffix);
    }
    used.add(name.toLowerCase());
    sheet.name = name;
  }
}

function parseCellReference(reference: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(reference);
  if (!match) return null;
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
    if (column > MAX_COLUMNS) return null;
  }
  const row = Number(match[2]);
  if (!Number.isSafeInteger(row) || row > MAX_ROWS) return null;
  return {row, column};
}

function columnName(column: number): string {
  let name = "";
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + (value - 1) % 26) + name;
  }
  return name;
}

function pixelDimension(value: unknown): number | null {
  const pixels = Math.round(Number(value));
  return Number.isFinite(pixels) && pixels >= 8 && pixels <= 2000 ? pixels : null;
}

function rowPoints(pixels: number): string {
  return String(Math.min(409, Math.round(pixels * 75) / 100));
}

function columnWidth(pixels: number): string {
  return String(Math.min(255, Math.round(Math.max(0, (pixels - 5) / 7) * 256) / 256));
}

function dimensions(source: Dims | undefined, maximum: number, convert: (pixels: number) => string): Dimension[] {
  const result: Dimension[] = [];
  if (!source || typeof source !== "object") return result;
  for (const [key, value] of Object.entries(source)) {
    if (!/^(0|[1-9]\d*)$/.test(key)) continue;
    const index = Number(key);
    const pixels = pixelDimension(value);
    if (!Number.isSafeInteger(index) || index < 0 || index >= maximum || pixels == null) continue;
    result.push({index, value: convert(pixels)});
  }
  result.sort((a, b) => a.index - b.index);
  return result;
}

function xlsxColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const hex = value.slice(1);
  if (!value.startsWith("#") || ![3, 4, 6, 8].includes(hex.length) || !/^[0-9a-f]+$/i.test(hex)) return null;
  if (hex.length === 3) return "FF" + Array.from(hex, character => character + character).join("").toUpperCase();
  if (hex.length === 4) {
    const [r, g, b, a] = Array.from(hex, character => character + character);
    return (a + r + g + b).toUpperCase();
  }
  if (hex.length === 6) return "FF" + hex.toUpperCase();
  return (hex.slice(6) + hex.slice(0, 6)).toUpperCase();
}

function decimals(fmt: CellFmt): number | null {
  if (fmt?.d == null) return null;
  const value = Math.round(Number(fmt?.d));
  return Number.isFinite(value) && value >= 0 && value <= 10 ? value : null;
}

function decimalPattern(value: number): string {
  return value ? "." + "0".repeat(value) : "";
}

class Styles {
  declare fonts: Font[];
  declare fontIds: Map<string, number>;
  declare fills: Array<Fill | null>;
  declare fillIds: Map<string, number>;
  declare numberFormats: Array<{ id: number; code: string }>;
  declare numberFormatIds: Map<string, number>;
  declare alignments: Array<Alignment | null>;
  declare alignmentIds: Map<string, number>;
  declare cellFormats: CellFormat[];
  declare cellFormatIds: Map<string, number>;

  constructor() {
    this.fonts = [{size: 11}];
    this.fontIds = new Map();
    this.fills = [null, {gray125: true}];
    this.fillIds = new Map();
    this.numberFormats = [];
    this.numberFormatIds = new Map();
    this.alignments = [null];
    this.alignmentIds = new Map();
    this.cellFormats = [{fontId: 0, fillId: 0, numberFormatId: 0, alignmentId: 0}];
    this.cellFormatIds = new Map();
  }

  font(fmt: CellFmt): number {
    const color = xlsxColor(fmt?.c);
    const pixels = Math.round(Number(fmt?.fs));
    // The grid renders `fs` in CSS pixels; Excel font sizes are points.
    const size = Number.isFinite(pixels) && pixels >= 6 && pixels <= 96 ? pixels * 0.75 : null;
    const font: Font = {
      bold: Boolean(fmt?.b), italic: Boolean(fmt?.i), underline: Boolean(fmt?.u),
      strike: Boolean(fmt?.s), color, size,
    };
    if (!font.bold && !font.italic && !font.underline && !font.strike && !font.color && !font.size) return 0;
    const key = JSON.stringify(font);
    let id = this.fontIds.get(key);
    if (id == null) {
      if (this.fonts.length >= MAX_FONTS) throw new Error("XLSX font count exceeds Excel's limit of 512.");
      id = this.fonts.length;
      this.fontIds.set(key, id);
      this.fonts.push(font);
    }
    return id;
  }

  fill(fmt: CellFmt): number {
    const color = xlsxColor(fmt?.bg);
    if (!color) return 0;
    let id = this.fillIds.get(color);
    if (id == null) {
      if (this.fills.length >= MAX_FILLS) throw new Error("XLSX fill count exceeds Excel's limit of 256.");
      id = this.fills.length;
      this.fillIds.set(color, id);
      this.fills.push({color});
    }
    return id;
  }

  customNumberFormat(code: string): number {
    let id = this.numberFormatIds.get(code);
    if (id == null) {
      if (164 + this.numberFormats.length > 0xffff) {
        throw new Error("XLSX number format count exceeds the format ID limit of 65,535.");
      }
      id = 164 + this.numberFormats.length;
      this.numberFormatIds.set(code, id);
      this.numberFormats.push({id, code});
    }
    return id;
  }

  numberFormat(fmt: CellFmt): number {
    const places = decimals(fmt);
    const name = fmt?.nf;
    if (name === "text") return 49;
    if (name === "integer") return this.customNumberFormat("#,##0");
    if (name === "number") return this.customNumberFormat("#,##0" + decimalPattern(places ?? 2));
    if (name === "currency") {
      const pattern = '"$"#,##0' + decimalPattern(places ?? 2);
      return this.customNumberFormat(pattern + ";-" + pattern);
    }
    if (name === "percent") return this.customNumberFormat("#,##0" + decimalPattern(places ?? 2) + "%");
    if (name === "scientific") return this.customNumberFormat("0" + decimalPattern(places ?? 2) + "E+00");
    if (name === "date") return this.customNumberFormat("mm/dd/yyyy");
    if (name === "time") return this.customNumberFormat("h:mm:ss AM/PM");
    if (name === "datetime") return this.customNumberFormat("mm/dd/yyyy h:mm:ss AM/PM");
    if (name != null) return 0;
    return places == null ? 0 : this.customNumberFormat("0" + decimalPattern(places));
  }

  alignment(fmt: CellFmt): number {
    const horizontal = fmt?.a === "l" ? "left" : fmt?.a === "c" ? "center" : fmt?.a === "r" ? "right" : null;
    const wrap = Boolean(fmt?.wrap);
    if (!horizontal && !wrap) return 0;
    const key = `${horizontal || ""}|${wrap}`;
    let id = this.alignmentIds.get(key);
    if (id == null) {
      id = this.alignments.length;
      this.alignmentIds.set(key, id);
      this.alignments.push({horizontal, wrap});
    }
    return id;
  }

  style(fmt: CellFmt | null | undefined): number {
    if (!fmt || typeof fmt !== "object") return 0;
    const cellFormat: CellFormat = {
      fontId: this.font(fmt), fillId: this.fill(fmt), numberFormatId: this.numberFormat(fmt),
      alignmentId: this.alignment(fmt),
    };
    if (!cellFormat.fontId && !cellFormat.fillId && !cellFormat.numberFormatId && !cellFormat.alignmentId) return 0;
    const key = `${cellFormat.fontId}|${cellFormat.fillId}|${cellFormat.numberFormatId}|${cellFormat.alignmentId}`;
    let id = this.cellFormatIds.get(key);
    if (id == null) {
      if (this.cellFormats.length >= MAX_CELL_FORMATS) {
        throw new Error("XLSX cell format count exceeds Excel's limit of 65,490.");
      }
      id = this.cellFormats.length;
      this.cellFormatIds.set(key, id);
      this.cellFormats.push(cellFormat);
    }
    return id;
  }
}

function sourceSheets(document: SheetsDocument): NamedSheet[] {
  const result: SheetDraft[] = [];
  const seen = new Set<string>();
  const order: string[] = Array.isArray(document?.sheetOrder) ? document.sheetOrder : [];
  const sheetMap: Record<string, SheetMeta> = document?.sheets && typeof document.sheets === "object" ? document.sheets : {};
  const cellMap: Record<string, CellMap> = document?.cells && typeof document.cells === "object" ? document.cells : {};
  for (const rawId of order) {
    const id = String(rawId);
    if (seen.has(id)) continue;
    seen.add(id);
    const metadata = sheetMap[id];
    if (!metadata || typeof metadata !== "object") continue;
    result.push({
      id,
      sourceName: typeof metadata.name === "string" ? metadata.name : "Sheet",
      metadata,
      sourceCells: cellMap[id] && typeof cellMap[id] === "object" ? cellMap[id] : {},
    });
  }
  if (!result.length) result.push({id: "", sourceName: "Sheet", metadata: {}, sourceCells: {}});
  assignSheetNames(result);
  return result;
}

function integerIn(value: unknown, minimum: number, maximum: number): number | null {
  const number = Number(value);
  return Number.isInteger(number) && number >= minimum && number <= maximum ? number : null;
}

function parseRangeReference(value: unknown): CellRange | null {
  const match = /^([A-Z]+[1-9]\d*)(?::([A-Z]+[1-9]\d*))?$/.exec(typeof value === "string" ? value : "");
  const first = match && parseCellReference(match[1]);
  const last = match && parseCellReference(match[2] ?? match[1]);
  if (!first || !last) return null;
  return {
    firstRow: Math.min(first.row, last.row), lastRow: Math.max(first.row, last.row),
    firstColumn: Math.min(first.column, last.column), lastColumn: Math.max(first.column, last.column),
  };
}

// One-based columns and rows.
function absoluteReference(sheetName: string, firstColumn: number, firstRow: number, lastColumn: number, lastRow: number): string {
  const first = `$${columnName(firstColumn)}$${firstRow}`;
  const last = `$${columnName(lastColumn)}$${lastRow}`;
  return `'${sheetName.replace(/'/g, "''")}'!${first}${first === last ? "" : ":" + last}`;
}

// The grid's filter criteria are tokens of its computed values: `s:` text, `n:` number, `b:1`/`b:0`
// booleans, `e:` an error code, `z:` blank, and `x:__none__` when nothing was selected.
function filterCriterion(token: unknown): { blank: true } | { value: string } | null {
  if (typeof token !== "string") return null;
  const kind = token.slice(0, 2);
  const text = token.slice(2);
  if (kind === "z:") return {blank: true};
  if (kind === "b:") return {value: text === "1" ? "TRUE" : "FALSE"};
  return kind === "s:" || kind === "n:" || kind === "e:" ? {value: text} : null;
}

// The token the grid derives from a literal cell, or null for a formula, whose value it computes.
function literalFilterToken(value: string): string | null {
  if (value[0] === "=") return null;
  const parsed = literalCellValue(value);
  if (parsed.type === "blank") return "z:";
  if (parsed.type === "boolean") return parsed.value ? "b:1" : "b:0";
  return (parsed.type === "number" ? "n:" : "s:") + String(parsed.value);
}

// Sorting in the grid reorders the stored cells, so only the header row, the criteria and the sort
// indicator need exporting. Rows are hidden by the literal values in the criteria columns; a row
// whose criteria cell holds a formula stays visible, since formulas are not evaluated here - unless
// the column selects nothing (`x:` tokens only), which no value can pass.
function prepareFilter(sheet: NamedSheet): PreparedFilter | null {
  const filter = sheet.metadata.filter;
  if (!filter || typeof filter !== "object") return null;
  const headerRow = integerIn(filter.row, 0, MAX_ROWS - 1);
  const endRow = headerRow == null ? null : integerIn(filter.endRow, headerRow, MAX_ROWS - 1);
  const columns = Array.isArray(filter.columns)
    ? [...new Set(filter.columns.map((column) => integerIn(column, 0, MAX_COLUMNS - 1)))]
      .filter((column): column is number => column != null).sort((a, b) => a - b)
    : [];
  if (headerRow == null || endRow == null || !columns.length) return null;
  const first = columns[0];
  const last = columns[columns.length - 1];
  const criteria = new Map<number, FilterCriterion>();
  const source: Record<string, unknown> = filter.criteria && typeof filter.criteria === "object" ? filter.criteria : {};
  for (const [key, tokens] of Object.entries(source)) {
    const column = /^(0|[1-9]\d*)$/.test(key) ? Number(key) : null;
    if (column == null || !columns.includes(column) || !Array.isArray(tokens) || !tokens.length) continue;
    const strings = tokens.filter((token): token is string => typeof token === "string");
    const criterion: FilterCriterion = {
      tokens: new Set(strings), values: [], blank: false,
      selectsNothing: strings.length === tokens.length && strings.every((token) => token.startsWith("x:")),
    };
    for (const token of strings) {
      const parsed = filterCriterion(token);
      if (parsed && "blank" in parsed) criterion.blank = true;
      else if (parsed) criterion.values.push(parsed.value);
    }
    criteria.set(column, criterion);
  }
  const filterColumns: FilterColumn[] = [];
  for (let column = first; column <= last; ++column) {
    const criterion = criteria.get(column);
    if (!columns.includes(column)) filterColumns.push({offset: column - first, hiddenButton: true});
    else if (criterion && (criterion.values.length || criterion.blank)) {
      filterColumns.push({offset: column - first, values: criterion.values, blank: criterion.blank});
    }
  }
  const hiddenRows: number[] = [];
  if (criteria.size) {
    for (let row = headerRow + 1; row <= endRow; ++row) {
      for (const [column, criterion] of criteria) {
        const cell = sheet.sourceCells[columnName(column + 1) + (row + 1)];
        const value = cell && typeof cell === "object" && cell.value != null ? String(cell.value) : "";
        const token = literalFilterToken(value);
        if (criterion.selectsNothing || (token != null && !criterion.tokens.has(token))) {
          hiddenRows.push(row + 1);
          break;
        }
      }
    }
  }
  const sortSource = filter.sort && typeof filter.sort === "object" ? filter.sort : null;
  const sortColumn = sortSource ? integerIn(sortSource.column, 0, MAX_COLUMNS - 1) : null;
  const direction = sortSource?.direction;
  const sort = sortColumn != null && columns.includes(sortColumn) && endRow > headerRow &&
      (direction === "asc" || direction === "desc")
    ? {
        ref: `${columnName(first + 1)}${headerRow + 2}:${columnName(last + 1)}${endRow + 1}`,
        column: `${columnName(sortColumn + 1)}${headerRow + 2}:${columnName(sortColumn + 1)}${endRow + 1}`,
        descending: direction === "desc",
      }
    : null;
  return {
    ref: `${columnName(first + 1)}${headerRow + 1}:${columnName(last + 1)}${endRow + 1}`,
    definedName: absoluteReference(sheet.name, first + 1, headerRow + 1, last + 1, endRow + 1),
    columns: filterColumns,
    hiddenRows,
    sort,
  };
}

// Resolved comments are hidden in the grid, so only open ones export. Excel allows one note per
// cell, so several comments on a cell are joined (and cut at Excel's note length).
function prepareComments(sheet: NamedSheet): PreparedComment[] {
  const source: unknown[] = Array.isArray(sheet.metadata.comments) ? sheet.metadata.comments : [];
  const byReference = new Map<string, PreparedComment>();
  for (const value of source) {
    if (!value || typeof value !== "object") continue;
    const comment = value as Record<string, unknown>;
    if (comment.resolved === true || typeof comment.text !== "string" || !comment.text.trim()) continue;
    const reference = typeof comment.ref === "string" ? comment.ref : "";
    const position = parseCellReference(reference);
    if (!position) continue;
    const existing = byReference.get(reference);
    if (existing) existing.text = (existing.text + "\n\n" + comment.text).slice(0, MAX_COMMENT_CHARACTERS);
    else byReference.set(reference, {reference, ...position, text: comment.text.slice(0, MAX_COMMENT_CHARACTERS)});
  }
  return [...byReference.values()].sort((a, b) => a.row - b.row || a.column - b.column);
}

// Walks the grid's pixel sizes to the cell under `pixels`, for a drawing anchor.
function anchorPosition(pixels: number, maxPixels: number, sizes: Dims | undefined, defaultPixels: number, maximum: number): AnchorPosition {
  const source: Dims = sizes && typeof sizes === "object" ? sizes : {};
  let index = 0;
  // A position past the grid's bounds is malformed and lands at the origin.
  let remaining = Number.isFinite(pixels) && pixels <= maxPixels ? Math.max(0, Math.round(pixels)) : 0;
  for (;;) {
    const size = pixelDimension(source[index]) ?? defaultPixels;
    if (remaining < size || index >= maximum - 1) return {index, offset: Math.min(remaining, size - 1) * EMU_PER_PIXEL};
    remaining -= size;
    ++index;
  }
}

// The grid's chartData() drops a column with no numeric value before building its series and
// skips filter-hidden rows. Only literals can be judged here, so a visible cell makes its column
// usable when it holds a number or a formula.
function usableSeriesColumns(cells: PreparedCell[], hiddenRows: number[], range: CellRange): number[] {
  const hidden = new Set(hiddenRows);
  const usable = new Set<number>();
  for (const cell of cells) {
    if (cell.column < range.firstColumn || cell.column > range.lastColumn ||
        cell.row < range.firstRow || cell.row > range.lastRow || hidden.has(cell.row)) continue;
    if (cell.value[0] === "=" || literalCellValue(cell.value).type === "number") usable.add(cell.column);
  }
  return [...usable].sort((a, b) => a - b);
}

// Mirrors the grid's chartData(): with header and label rows on, each remaining column is a series
// named by its first cell, categorized by the first column. Charts without a usable range are
// skipped, as the grid draws only a placeholder for them.
function prepareCharts(metadata: Partial<SheetMeta>, cells: PreparedCell[], hiddenRows: number[], parts: PartCounts): PreparedChart[] {
  const source: unknown[] = Array.isArray(metadata.charts) ? metadata.charts : [];
  const charts: PreparedChart[] = [];
  for (const value of source) {
    if (!value || typeof value !== "object") continue;
    const chart = value as Record<string, unknown>;
    const range = parseRangeReference(chart.range);
    if (!range) continue;
    const firstRowHeaders = chart.firstRowHeaders !== false;
    const firstColumnLabels = chart.firstColLabels !== false;
    const dataRow = range.firstRow + (firstRowHeaders && range.lastRow > range.firstRow ? 1 : 0);
    const seriesColumn = range.firstColumn + (firstColumnLabels && range.lastColumn > range.firstColumn ? 1 : 0);
    if (dataRow > range.lastRow || seriesColumn > range.lastColumn) continue;
    const type = CHART_TYPES.includes(chart.type as ChartType) ? chart.type as ChartType : "line";
    const series = usableSeriesColumns(cells, hiddenRows, {...range, firstRow: dataRow, firstColumn: seriesColumn});
    // Excel draws at most this many series; the grid's own chart keeps every one.
    const columns = series.slice(0, type === "pie" ? 1 : MAX_SERIES);
    if (!columns.length) continue;
    const width = count(chart.width, 520, CHART_BOUNDS.maxWidth);
    const height = count(chart.height, 320, CHART_BOUNDS.maxHeight);
    charts.push({
      index: ++parts.charts,
      type,
      title: typeof chart.title === "string" ? chart.title : "",
      xAxisTitle: typeof chart.xAxisTitle === "string" ? chart.xAxisTitle : "",
      yAxisTitle: typeof chart.yAxisTitle === "string" ? chart.yAxisTitle : "",
      legend: chart.legend !== false,
      headerRow: firstRowHeaders && range.lastRow > range.firstRow ? range.firstRow : null,
      categoryColumn: seriesColumn > range.firstColumn ? range.firstColumn : null,
      firstRow: dataRow,
      lastRow: range.lastRow,
      columns,
      column: anchorPosition(Number(chart.x ?? 96) - ROW_HEADER_PIXELS, CHART_BOUNDS.maxX, metadata.colWidths, DEFAULT_COLUMN_PIXELS, MAX_COLUMNS),
      row: anchorPosition(Number(chart.y ?? 44) - COLUMN_HEADER_PIXELS, CHART_BOUNDS.maxY, metadata.rowHeights, DEFAULT_ROW_PIXELS, MAX_ROWS),
      extent: {cx: width * EMU_PER_PIXEL, cy: height * EMU_PER_PIXEL},
    });
  }
  return charts;
}

// The grid links any literal whose text is an absolute http(s) URL, unless formatted as text.
// A target longer than Excel accepts stays plain text in the export.
function hyperlinkTarget(value: string, fmt: CellFmt | null | undefined): string | null {
  const text = value[0] === "'" ? value.slice(1) : value;
  if (fmt?.nf === "text" || !/^\s*https?:/i.test(text)) return null;
  try {
    const url = new URL(text);
    const external = url.protocol === "http:" || url.protocol === "https:";
    return external && url.href.length <= MAX_HYPERLINK_CHARACTERS ? url.href : null;
  } catch {
    return null;
  }
}

function prepareWorkbook(document: SheetsDocument): { sheets: PreparedSheet[]; styles: Styles; formulaNames: Map<string, string>; parts: PartCounts } {
  const drafts = sourceSheets(document);
  const formulaNames = new Map<string, string>();
  for (const sheet of drafts) {
    const key = sheet.sourceName.toLowerCase();
    if (!formulaNames.has(key)) formulaNames.set(key, sheet.name);
  }
  const styles = new Styles();
  const parts: PartCounts = {drawings: 0, charts: 0, comments: 0, vmlShapeBlocks: 0};
  const sheets = drafts.map((draft): PreparedSheet => {
    const rows = count(draft.metadata.rows, DEFAULT_ROWS, MAX_ROWS);
    const columns = count(draft.metadata.cols, DEFAULT_COLUMNS, MAX_COLUMNS);
    const cells: PreparedCell[] = [];
    const hyperlinks: PreparedHyperlink[] = [];
    for (const [reference, sourceCell] of Object.entries(draft.sourceCells)) {
      const position = parseCellReference(reference);
      if (!position || !sourceCell || typeof sourceCell !== "object") continue;
      const value = sourceCell.value == null ? "" : String(sourceCell.value);
      const target = hyperlinks.length < MAX_HYPERLINKS ? hyperlinkTarget(value, sourceCell.fmt) : null;
      if (target) hyperlinks.push({reference, ...position, target, relationshipId: `rId${hyperlinks.length + 1}`});
      const style = styles.style(target ? {...(sourceCell.fmt ?? {}), ...HYPERLINK_FMT} : sourceCell.fmt);
      if (value === "" && !style) continue;
      cells.push({reference, ...position, value, style});
    }
    cells.sort((a, b) => a.row - b.row || a.column - b.column);
    hyperlinks.sort((a, b) => a.row - b.row || a.column - b.column);
    const filter = prepareFilter(draft);
    const hiddenRows = filter?.hiddenRows ?? [];

    // Relationship ids after the hyperlinks' own.
    let relationships = hyperlinks.length;
    const charts = prepareCharts(draft.metadata, cells, hiddenRows, parts);
    const drawing = charts.length
      ? {index: ++parts.drawings, relationshipId: `rId${++relationships}`, charts}
      : null;
    const list = prepareComments(draft);
    let comments: PreparedComments | null = null;
    if (list.length) {
      comments = {
        index: ++parts.comments,
        vmlRelationshipId: `rId${++relationships}`,
        relationshipId: `rId${++relationships}`,
        // VML shape ids are allotted in blocks of 1024 per `o:idmap` entry, unique across the workbook.
        firstShapeBlock: parts.vmlShapeBlocks + 1,
        list,
      };
      parts.vmlShapeBlocks += Math.ceil(list.length / 1024);
    }
    return {
      name: draft.name,
      metadata: draft.metadata,
      rows,
      columns,
      frozenRows: frozenCount(draft.metadata.frozenRows, rows),
      frozenColumns: frozenCount(draft.metadata.frozenCols, columns),
      columnWidths: dimensions(draft.metadata.colWidths, columns, columnWidth),
      rowHeights: dimensions(draft.metadata.rowHeights, rows, rowPoints),
      cells,
      hyperlinks,
      filter,
      hiddenRows,
      drawing,
      comments,
    };
  });
  return {sheets, styles, formulaNames, parts};
}

function formulaReferenceAt(formula: string, offset: number): boolean {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)/.exec(formula.slice(offset));
  if (!match) return false;
  let column = 0;
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  if (column > MAX_COLUMNS || Number(match[2]) > MAX_ROWS) return false;
  const next = formula[offset + match[0].length];
  return !next || !/[A-Za-z0-9_$]/.test(next);
}

function quotedSheetReference(formula: string, offset: number, names: Map<string, string>): FormulaRewrite | null {
  const nameParts: string[] = [];
  for (let i = offset + 1; i < formula.length; ++i) {
    if (formula[i] !== "'") {
      nameParts.push(formula[i]);
      continue;
    }
    if (formula[i + 1] === "'") {
      nameParts.push("'");
      ++i;
      continue;
    }
    const quoteEnd = i + 1;
    const hasBang = formula[quoteEnd] === "!";
    const end = quoteEnd + (hasBang ? 1 : 0);
    const text = formula.slice(offset, end);
    const name = nameParts.join("");
    const normalized = names.get(name.toLowerCase());
    const malformed = offset > 0 && /[A-Za-z0-9_.$]/.test(formula[offset - 1]);
    const external = formula[offset - 1] === "]" || (!normalized && /\[[^\]]*\]/.test(name));
    if (!hasBang || !formulaReferenceAt(formula, end) || malformed || external ||
        isThreeDimensionalReference(formula, offset)) return {end, text};
    return normalized
      ? {end, text: `'${normalized.replace(/'/g, "''")}'!`}
      : {end, text};
  }
  return null;
}

function unquotedSheetReference(formula: string, offset: number, names: Map<string, string>): FormulaRewrite | null {
  if (!/[A-Za-z_$]/.test(formula[offset]) ||
      (offset > 0 && /[A-Za-z0-9_.$]/.test(formula[offset - 1])) ||
      formula[offset - 1] === "]" || isThreeDimensionalReference(formula, offset)) return null;
  let end = offset + 1;
  while (end < formula.length && /[A-Za-z0-9_.$]/.test(formula[end])) ++end;
  if (formula[end] !== "!" || !formulaReferenceAt(formula, end + 1)) return null;
  const name = formula.slice(offset, end);
  const normalized = names.get(name.toLowerCase());
  if (!normalized) return null;
  if (normalized.toLowerCase() === name.toLowerCase()) {
    return {end: end + 1, text: formula.slice(offset, end + 1)};
  }
  return {end: end + 1, text: `'${normalized.replace(/'/g, "''")}'!`};
}

function isThreeDimensionalReference(formula: string, offset: number): boolean {
  if (formula[offset - 1] !== ":") return false;
  let start = offset - 2;
  while (start >= 0 && /[A-Za-z0-9_$]/.test(formula[start])) --start;
  const preceding = formula.slice(start + 1, offset - 1);
  return !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(preceding);
}

// Recognizes a function call at `offset`. The grid's tokenizer discards whitespace, so it accepts
// `SUM (1)`; in Excel that space is the intersection operator, so the gap is dropped here.
function formulaFunctionAt(formula: string, offset: number): FormulaRewrite | null {
  if (!/[A-Za-z_]/.test(formula[offset]) ||
      (offset > 0 && /[A-Za-z0-9_.$!]/.test(formula[offset - 1]))) return null;
  let end = offset + 1;
  while (end < formula.length && /[A-Za-z0-9_.]/.test(formula[end])) ++end;
  let parenthesis = end;
  while (parenthesis < formula.length && /\s/.test(formula[parenthesis])) ++parenthesis;
  if (formula[parenthesis] !== "(") return null;
  const name = formula.slice(offset, end).toUpperCase();
  if (FUTURE_FUNCTIONS.has(name)) return {end: parenthesis, text: "_xlfn." + name};
  if (name === "ERRORTYPE") return {end: parenthesis, text: "ERROR.TYPE"};
  return parenthesis > end ? {end: parenthesis, text: formula.slice(offset, end)} : null;
}

// Mirrors the grid's tokenizer: inside a string a doubled quote is a literal quote. Emits the Excel form (double quotes, doubled inside) or null when unterminated.
function stringLiteralAt(formula: string, offset: number): FormulaRewrite | null {
  const quote = formula[offset];
  let text = "";
  for (let i = offset + 1; i < formula.length; ++i) {
    const character = formula[i];
    const escaped = character === quote && formula[i + 1] === quote;
    if (escaped) {
      text += quote;
      ++i;
    } else if (character === quote) {
      return {end: i + 1, text: `"${text.replace(/"/g, '""')}"`};
    } else {
      text += character;
    }
  }
  return null;
}

// The grid reads a single-quoted run as a sheet name only when its closing quote is followed by
// `!`; any other run is a string literal. A run followed by `:` is also left alone here so an
// Excel 3-D reference such as 'Jan':'Mar'!A1 survives verbatim (see isThreeDimensionalReference).
function quotedSheetNameAt(formula: string, offset: number): boolean {
  for (let i = offset + 1; i < formula.length; ++i) {
    if (formula[i] !== "'") continue;
    if (formula[i + 1] === "'") {
      ++i;
      continue;
    }
    return formula[i + 1] === "!" || formula[i + 1] === ":";
  }
  return false;
}

// Rewrites strings, sheet names and function names for Excel. Returns null when the formula is
// unbalanced (unterminated string or quoted name, mismatched parentheses or brackets): the grid's
// parser tolerates those, but one such `<f>` makes Excel report the whole workbook as damaged.
function rewriteFormula(formula: string, names: Map<string, string>): string | null {
  const result: string[] = [];
  let parentheses = 0;
  let structuredReferenceDepth = 0;
  for (let i = 0; i < formula.length;) {
    const character = formula[i];
    if (character === '"' ||
        (character === "'" && !structuredReferenceDepth && !quotedSheetNameAt(formula, i))) {
      const literal = stringLiteralAt(formula, i);
      if (!literal) return null;
      result.push(literal.text);
      i = literal.end;
      continue;
    }
    let apostrophes = 0;
    if (structuredReferenceDepth && (character === "[" || character === "]")) {
      for (let j = i - 1; formula[j] === "'"; --j) ++apostrophes;
    }
    const escapedBracket = apostrophes % 2 === 1;
    if (character === "[" && !escapedBracket) ++structuredReferenceDepth;
    else if (character === "]" && !escapedBracket && --structuredReferenceDepth < 0) return null;
    if (!structuredReferenceDepth) {
      if (character === "(") ++parentheses;
      else if (character === ")" && --parentheses < 0) return null;
      const reference = character === "'"
        ? quotedSheetReference(formula, i, names)
        : formulaFunctionAt(formula, i) || unquotedSheetReference(formula, i, names);
      if (character === "'" && !reference) return null;
      if (reference) {
        result.push(reference.text);
        i = reference.end;
        continue;
      }
    }
    result.push(character);
    ++i;
  }
  return parentheses || structuredReferenceDepth ? null : result.join("");
}

// A literal (non-formula) cell's text as Excel should read it.
function literalCellValue(value: string): ParsedCellValue {
  if (value[0] === "'") return {type: "text", value: value.slice(1)};
  const trimmed = value.trim();
  if (trimmed === "") return {type: "blank", value: ""};
  if (/^(TRUE|FALSE)$/i.test(trimmed)) return {type: "boolean", value: /^true$/i.test(trimmed)};
  if (/^[-+]?\$?[\d,]*\.?\d+%?$/.test(trimmed) && /\d/.test(trimmed)) {
    const negative = trimmed.startsWith("-");
    const cleaned = trimmed.replace(/[$,+%-]/g, "");
    let number = Number(cleaned);
    if (Number.isFinite(number)) {
      if (trimmed.endsWith("%")) number /= 100;
      return {type: "number", value: negative ? -number : number};
    }
  }
  return {type: "text", value};
}

function parsedCellValue(value: string, formulaNames: Map<string, string>): ParsedCellValue {
  if (value[0] === "=") {
    const formula = rewriteFormula(value.slice(1), formulaNames);
    // Excel also rejects empty formulas and those over its length limit; keep the stored text.
    return formula && formula.trim() && formula.length < MAX_FORMULA_CHARACTERS
      ? {type: "formula", value: formula}
      : {type: "text", value};
  }
  return literalCellValue(value);
}

function cellXml(cell: PreparedCell, formulaNames: Map<string, string>): string {
  const style = cell.style ? ` s="${cell.style}"` : "";
  if (cell.value === "") return `<c r="${cell.reference}"${style}/>`;
  const parsed = parsedCellValue(cell.value, formulaNames);
  if (parsed.type === "blank") return `<c r="${cell.reference}"${style}/>`;
  if (parsed.type === "formula") return `<c r="${cell.reference}"${style}><f>${formulaXml(parsed.value)}</f></c>`;
  if (parsed.type === "boolean") return `<c r="${cell.reference}"${style} t="b"><v>${parsed.value ? 1 : 0}</v></c>`;
  if (parsed.type === "number") return `<c r="${cell.reference}"${style}><v>${String(parsed.value)}</v></c>`;
  return `<c r="${cell.reference}"${style} t="inlineStr"><is><t xml:space="preserve">${spreadsheetXml(parsed.value)}</t></is></c>`;
}

function frozenPane(sheet: PreparedSheet): string {
  const rows = sheet.frozenRows;
  const columns = sheet.frozenColumns;
  if (!rows && !columns) return "";
  const attributes: string[] = [];
  if (columns) attributes.push(`xSplit="${columns}"`);
  if (rows) attributes.push(`ySplit="${rows}"`);
  attributes.push(`topLeftCell="${columnName(columns + 1)}${rows + 1}"`);
  attributes.push(`activePane="${rows && columns ? "bottomRight" : rows ? "bottomLeft" : "topRight"}"`);
  attributes.push('state="frozen"');
  return `<pane ${attributes.join(" ")}/>`;
}

function worksheetDimension(cells: PreparedCell[]): string {
  if (!cells.length) return "A1";
  let minRow = MAX_ROWS, minColumn = MAX_COLUMNS, maxRow = 1, maxColumn = 1;
  for (const cell of cells) {
    minRow = Math.min(minRow, cell.row);
    minColumn = Math.min(minColumn, cell.column);
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  const first = columnName(minColumn) + minRow;
  const last = columnName(maxColumn) + maxRow;
  return first === last ? first : first + ":" + last;
}

function autoFilterXml(filter: PreparedFilter): string {
  let xml = `<autoFilter ref="${filter.ref}">`;
  for (const column of filter.columns) {
    if (!("values" in column)) {
      xml += `<filterColumn colId="${column.offset}" hiddenButton="1"/>`;
      continue;
    }
    xml += `<filterColumn colId="${column.offset}"><filters${column.blank ? ' blank="1"' : ""}>`;
    for (const value of column.values) xml += `<filter val="${spreadsheetXml(value, true)}"/>`;
    xml += "</filters></filterColumn>";
  }
  // A sort applied through the filter dropdown lives inside the autoFilter, as Excel writes it.
  if (filter.sort) {
    xml += `<sortState ref="${filter.sort.ref}"><sortCondition ref="${filter.sort.column}"` +
      `${filter.sort.descending ? ' descending="1"' : ""}/></sortState>`;
  }
  return xml + "</autoFilter>";
}

function* worksheetXml(sheet: PreparedSheet, formulaNames: Map<string, string>): Generator<string, void, unknown> {
  yield `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`;
  yield `<dimension ref="${worksheetDimension(sheet.cells)}"/>`;
  yield `<sheetViews><sheetView workbookViewId="0">${frozenPane(sheet)}</sheetView></sheetViews>`;
  yield `<sheetFormatPr defaultColWidth="${columnWidth(DEFAULT_COLUMN_PIXELS)}" defaultRowHeight="${rowPoints(DEFAULT_ROW_PIXELS)}"/>`;
  if (sheet.columnWidths.length) {
    yield "<cols>";
    for (const width of sheet.columnWidths) {
      yield `<col min="${width.index + 1}" max="${width.index + 1}" width="${width.value}" customWidth="1"/>`;
    }
    yield "</cols>";
  }
  yield "<sheetData>";
  // Rows are written in order wherever a cell, a custom height or a hidden (filtered) row falls.
  let cellIndex = 0;
  let heightIndex = 0;
  let hiddenIndex = 0;
  for (;;) {
    const cellRow = sheet.cells[cellIndex]?.row ?? Infinity;
    const heightRow = (sheet.rowHeights[heightIndex]?.index ?? Infinity) + 1;
    const hiddenRow = sheet.hiddenRows[hiddenIndex] ?? Infinity;
    const row = Math.min(cellRow, heightRow, hiddenRow);
    if (row === Infinity) break;
    const height = heightRow === row ? sheet.rowHeights[heightIndex++] : null;
    const hidden = hiddenRow === row;
    if (hidden) ++hiddenIndex;
    yield `<row r="${row}"${height ? ` ht="${height.value}" customHeight="1"` : ""}${hidden ? ' hidden="1"' : ""}>`;
    while (sheet.cells[cellIndex]?.row === row) yield cellXml(sheet.cells[cellIndex++], formulaNames);
    yield "</row>";
  }
  yield "</sheetData>";
  if (sheet.filter) yield autoFilterXml(sheet.filter);
  if (sheet.hyperlinks.length) {
    yield "<hyperlinks>";
    for (const link of sheet.hyperlinks) yield `<hyperlink ref="${link.reference}" r:id="${link.relationshipId}"/>`;
    yield "</hyperlinks>";
  }
  if (sheet.drawing) yield `<drawing r:id="${sheet.drawing.relationshipId}"/>`;
  if (sheet.comments) yield `<legacyDrawing r:id="${sheet.comments.vmlRelationshipId}"/>`;
  yield "</worksheet>";
}

function worksheetRelationships(sheet: PreparedSheet): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_REL_NS}">`;
  for (const link of sheet.hyperlinks) {
    xml += `<Relationship Id="${link.relationshipId}" Type="${REL_NS}/hyperlink" Target="${xmlAttribute(link.target)}" TargetMode="External"/>`;
  }
  if (sheet.drawing) {
    xml += `<Relationship Id="${sheet.drawing.relationshipId}" Type="${REL_NS}/drawing" Target="../drawings/drawing${sheet.drawing.index}.xml"/>`;
  }
  if (sheet.comments) {
    xml += `<Relationship Id="${sheet.comments.vmlRelationshipId}" Type="${REL_NS}/vmlDrawing" Target="../drawings/vmlDrawing${sheet.comments.index}.vml"/>`;
    xml += `<Relationship Id="${sheet.comments.relationshipId}" Type="${REL_NS}/comments" Target="../comments${sheet.comments.index}.xml"/>`;
  }
  return xml + "</Relationships>";
}

function* commentsXml(comments: PreparedComments): Generator<string, void, unknown> {
  yield `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="${MAIN_NS}">`;
  yield "<authors><author>Workspace Sheets</author></authors><commentList>";
  for (const comment of comments.list) {
    yield `<comment ref="${comment.reference}" authorId="0"><text><t xml:space="preserve">${spreadsheetXml(comment.text)}</t></text></comment>`;
  }
  yield "</commentList></comments>";
}

// Excel shows a legacy note only through its VML shape; the anchor places the hidden popup beside
// the cell and Excel sizes it on open.
function* vmlDrawingXml(comments: PreparedComments): Generator<string, void, unknown> {
  const blocks = Array.from({length: Math.ceil(comments.list.length / 1024)}, (_, i) => comments.firstShapeBlock + i);
  yield '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">';
  yield `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${blocks.join(",")}"/></o:shapelayout>`;
  yield '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
  for (let i = 0; i < comments.list.length; ++i) {
    const comment = comments.list[i];
    const row = comment.row - 1;
    const column = comment.column - 1;
    const anchor = [
      Math.min(column + 1, MAX_COLUMNS - 1), 15, Math.max(0, row - 1), 10,
      Math.min(column + 3, MAX_COLUMNS - 1), 15, Math.min(row + 3, MAX_ROWS - 1), 4,
    ];
    yield `<v:shape id="_x0000_s${comments.firstShapeBlock * 1024 + i + 1}" type="#_x0000_t202" ` +
      `style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" ` +
      'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>' +
      '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor.join(", ")}</x:Anchor>` +
      `<x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${column}</x:Column></x:ClientData></v:shape>`;
  }
  yield "</xml>";
}

function drawingXml(drawing: PreparedDrawing): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWING_NS}">`;
  drawing.charts.forEach((chart, i) => {
    xml += `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.column.index}</xdr:col><xdr:colOff>${chart.column.offset}</xdr:colOff>` +
      `<xdr:row>${chart.row.index}</xdr:row><xdr:rowOff>${chart.row.offset}</xdr:rowOff></xdr:from>` +
      `<xdr:ext cx="${chart.extent.cx}" cy="${chart.extent.cy}"/>` +
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      `<a:graphic><a:graphicData uri="${CHART_NS}"><c:chart xmlns:c="${CHART_NS}" xmlns:r="${REL_NS}" r:id="rId${i + 1}"/></a:graphicData></a:graphic>` +
      "</xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>";
  });
  return xml + "</xdr:wsDr>";
}

function drawingRelationships(drawing: PreparedDrawing): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_REL_NS}">`;
  drawing.charts.forEach((chart, i) => {
    xml += `<Relationship Id="rId${i + 1}" Type="${REL_NS}/chart" Target="../charts/chart${chart.index}.xml"/>`;
  });
  return xml + "</Relationships>";
}

function chartTitleXml(text: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>${plainXml(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function solidFill(color: string, alpha?: number): string {
  const rgb = alpha ? `<a:srgbClr val="${color}"><a:alpha val="${alpha}"/></a:srgbClr>` : `<a:srgbClr val="${color}"/>`;
  return `<a:solidFill>${rgb}</a:solidFill>`;
}

function chartSeriesXml(chart: PreparedChart, sheetName: string, index: number): string {
  const column = chart.columns[index];
  const color = CHART_COLORS[index % CHART_COLORS.length];
  const reference = (firstRow: number, lastRow: number) => plainXml(absoluteReference(sheetName, column, firstRow, column, lastRow));
  let xml = `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`;
  xml += chart.headerRow
    ? `<c:tx><c:strRef><c:f>${reference(chart.headerRow, chart.headerRow)}</c:f></c:strRef></c:tx>`
    : `<c:tx><c:v>${columnName(column)}</c:v></c:tx>`;
  if (chart.type === "line") {
    xml += `<c:spPr><a:ln w="28575" cap="rnd">${solidFill(color)}<a:round/></a:ln></c:spPr>`;
    xml += `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${solidFill(color)}<a:ln>${solidFill(color)}</a:ln></c:spPr></c:marker>`;
  } else if (chart.type === "area") {
    xml += `<c:spPr>${solidFill(color, 18000)}<a:ln w="28575">${solidFill(color)}</a:ln></c:spPr>`;
  } else if (chart.type === "stackedBar") {
    xml += `<c:spPr>${solidFill(color)}</c:spPr><c:invertIfNegative val="0"/>`;
  } else {
    // Slice colours cycle through the palette; past this many slices Excel's own variation applies.
    const points = Math.min(chart.lastRow - chart.firstRow + 1, 64);
    for (let point = 0; point < points; ++point) {
      xml += `<c:dPt><c:idx val="${point}"/><c:bubble3D val="0"/><c:spPr>${solidFill(CHART_COLORS[point % CHART_COLORS.length])}` +
        `<a:ln w="19050">${solidFill("FFFFFF")}</a:ln></c:spPr></c:dPt>`;
    }
  }
  if (chart.categoryColumn) {
    const categories = absoluteReference(sheetName, chart.categoryColumn, chart.firstRow, chart.categoryColumn, chart.lastRow);
    xml += `<c:cat><c:strRef><c:f>${plainXml(categories)}</c:f></c:strRef></c:cat>`;
  }
  xml += `<c:val><c:numRef><c:f>${reference(chart.firstRow, chart.lastRow)}</c:f></c:numRef></c:val>`;
  if (chart.type === "line") xml += '<c:smooth val="0"/>';
  return xml + "</c:ser>";
}

function chartAxisXml(kind: "catAx" | "valAx", id: number, crossId: number, position: string, title: string,
    {reversed = false, crossesMax = false}: { reversed?: boolean; crossesMax?: boolean } = {}): string {
  return `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="${reversed ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${position}"/>${kind === "valAx" ? "<c:majorGridlines/>" : ""}${title ? chartTitleXml(title) : ""}` +
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    `<c:crossAx val="${crossId}"/><c:crosses val="${crossesMax ? "max" : "autoZero"}"/>` +
    (kind === "catAx" ? '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>' : '<c:crossBetween val="between"/>') +
    `</c:${kind}>`;
}

// Series references point at the worksheet without cached values; Excel reads them on open, like
// formulas. Palette and layout follow the grid's SVG renderer.
function chartXml(chart: PreparedChart, sheetName: string): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}">`;
  xml += '<c:roundedCorners val="0"/><c:chart>';
  xml += chart.title ? chartTitleXml(chart.title) + '<c:autoTitleDeleted val="0"/>' : '<c:autoTitleDeleted val="1"/>';
  xml += "<c:plotArea><c:layout/>";
  const series = chart.columns.map((_, index) => chartSeriesXml(chart, sheetName, index)).join("");
  if (chart.type === "pie") {
    xml += `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
  } else {
    const axes = '<c:axId val="1"/><c:axId val="2"/>';
    if (chart.type === "stackedBar") {
      xml += `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${series}<c:gapWidth val="56"/><c:overlap val="100"/>${axes}</c:barChart>`;
      // Categories run top-down as in the grid, which puts the value axis at the category maximum.
      xml += chartAxisXml("catAx", 1, 2, "l", chart.yAxisTitle, {reversed: true}) +
        chartAxisXml("valAx", 2, 1, "b", chart.xAxisTitle, {crossesMax: true});
    } else {
      xml += chart.type === "area"
        ? `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>`
        : `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>`;
      xml += chartAxisXml("catAx", 1, 2, "b", chart.xAxisTitle) + chartAxisXml("valAx", 2, 1, "l", chart.yAxisTitle);
    }
  }
  xml += "</c:plotArea>";
  if (chart.legend) xml += '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>';
  return xml + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
}

function* stylesXml(styles: Styles): Generator<string, void, unknown> {
  yield `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${MAIN_NS}">`;
  if (styles.numberFormats.length) {
    yield `<numFmts count="${styles.numberFormats.length}">`;
    for (const format of styles.numberFormats) yield `<numFmt numFmtId="${format.id}" formatCode="${xmlAttribute(format.code)}"/>`;
    yield "</numFmts>";
  }
  yield `<fonts count="${styles.fonts.length}">`;
  for (const font of styles.fonts) {
    yield "<font>";
    if (font.bold) yield "<b/>";
    if (font.italic) yield "<i/>";
    if (font.underline) yield "<u/>";
    if (font.strike) yield "<strike/>";
    yield `<sz val="${font.size || 11}"/>`;
    if (font.color) yield `<color rgb="${font.color}"/>`;
    yield '<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
  }
  yield "</fonts>";
  yield `<fills count="${styles.fills.length}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>`;
  for (let i = 2; i < styles.fills.length; ++i) {
    yield `<fill><patternFill patternType="solid"><fgColor rgb="${styles.fills[i]!.color}"/><bgColor indexed="64"/></patternFill></fill>`;
  }
  yield "</fills>";
  yield '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
  yield `<cellXfs count="${styles.cellFormats.length}">`;
  for (const format of styles.cellFormats) {
    const alignment = styles.alignments[format.alignmentId];
    let attributes = `numFmtId="${format.numberFormatId}" fontId="${format.fontId}" fillId="${format.fillId}" borderId="0" xfId="0"`;
    if (format.numberFormatId) attributes += ' applyNumberFormat="1"';
    if (format.fontId) attributes += ' applyFont="1"';
    if (format.fillId) attributes += ' applyFill="1"';
    if (alignment) attributes += ' applyAlignment="1"';
    if (!alignment) {
      yield `<xf ${attributes}/>`;
      continue;
    }
    const alignmentAttributes: string[] = [];
    if (alignment.horizontal) alignmentAttributes.push(`horizontal="${alignment.horizontal}"`);
    if (alignment.wrap) alignmentAttributes.push('wrapText="1"');
    yield `<xf ${attributes}><alignment ${alignmentAttributes.join(" ")}/></xf>`;
  }
  yield '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
}

function contentTypes(sheets: PreparedSheet[], parts: PartCounts): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
  xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  xml += '<Default Extension="xml" ContentType="application/xml"/>';
  if (parts.comments) xml += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
  xml += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  xml += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
  for (let i = 1; i <= sheets.length; ++i) {
    xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }
  for (let i = 1; i <= parts.comments; ++i) {
    xml += `<Override PartName="/xl/comments${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
  }
  for (let i = 1; i <= parts.drawings; ++i) {
    xml += `<Override PartName="/xl/drawings/drawing${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
  }
  for (let i = 1; i <= parts.charts; ++i) {
    xml += `<Override PartName="/xl/charts/chart${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
  }
  return xml + "</Types>";
}

function workbookXml(sheets: PreparedSheet[]): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`;
  xml += "<bookViews><workbookView/></bookViews><sheets>";
  for (let i = 0; i < sheets.length; ++i) {
    xml += `<sheet name="${spreadsheetXml(sheets[i].name, true)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  }
  xml += "</sheets>";
  // Excel records each sheet's autofilter range under this reserved hidden name.
  const filtered = sheets.flatMap((sheet, index) => sheet.filter ? [{index, filter: sheet.filter}] : []);
  if (filtered.length) {
    xml += "<definedNames>";
    for (const {index, filter} of filtered) {
      xml += `<definedName name="_xlnm._FilterDatabase" localSheetId="${index}" hidden="1">${plainXml(filter.definedName)}</definedName>`;
    }
    xml += "</definedNames>";
  }
  // Formulas are written without cached results, so ask for one full recalculation on open.
  return xml + '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>';
}

function workbookRelationships(sheetCount: number): string {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_REL_NS}">`;
  for (let i = 1; i <= sheetCount; ++i) {
    xml += `<Relationship Id="rId${i}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${i}.xml"/>`;
  }
  return xml + `<Relationship Id="rId${sheetCount + 1}" Type="${REL_NS}/styles" Target="styles.xml"/></Relationships>`;
}

// A worksheet's part, plus its relationships, drawing, charts, comments and VML when it has any.
function sheetParts(sheet: PreparedSheet, index: number, formulaNames: Map<string, string>): ZipEntry[] {
  const entries: ZipEntry[] = [{name: `xl/worksheets/sheet${index + 1}.xml`, data: textStream(worksheetXml(sheet, formulaNames))}];
  if (sheet.hyperlinks.length || sheet.drawing || sheet.comments) {
    entries.push({name: `xl/worksheets/_rels/sheet${index + 1}.xml.rels`, data: worksheetRelationships(sheet)});
  }
  if (sheet.drawing) {
    entries.push({name: `xl/drawings/drawing${sheet.drawing.index}.xml`, data: drawingXml(sheet.drawing)});
    entries.push({name: `xl/drawings/_rels/drawing${sheet.drawing.index}.xml.rels`, data: drawingRelationships(sheet.drawing)});
    for (const chart of sheet.drawing.charts) {
      entries.push({name: `xl/charts/chart${chart.index}.xml`, data: chartXml(chart, sheet.name)});
    }
  }
  if (sheet.comments) {
    entries.push({name: `xl/comments${sheet.comments.index}.xml`, data: textStream(commentsXml(sheet.comments))});
    entries.push({name: `xl/drawings/vmlDrawing${sheet.comments.index}.vml`, data: textStream(vmlDrawingXml(sheet.comments))});
  }
  return entries;
}

/** Streams `document` (a complete `Gadget.getDocument()` snapshot) as an XLSX workbook. */
export function workbookToXlsx(document: SheetsDocument): ReadableStream<Uint8Array> {
  const {sheets, styles, formulaNames, parts} = prepareWorkbook(document);
  const entries: ZipEntry[] = [
    {name: "[Content_Types].xml", data: contentTypes(sheets, parts)},
    {name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PACKAGE_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
    {name: "xl/workbook.xml", data: workbookXml(sheets)},
    {name: "xl/_rels/workbook.xml.rels", data: workbookRelationships(sheets.length)},
    {name: "xl/styles.xml", data: textStream(stylesXml(styles))},
    ...sheets.flatMap((sheet, i) => sheetParts(sheet, i, formulaNames)),
  ];
  return createZip(entries);
}
