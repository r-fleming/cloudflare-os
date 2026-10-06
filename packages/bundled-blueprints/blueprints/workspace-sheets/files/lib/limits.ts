// Limits and enumerations for a workbook's stored metadata. The server enforces them; they live
// here so the client and the exporter can apply the same ones.
import type { ChartType, PivotAggregate } from "./protocol.ts";

/** The largest sheet: rows and columns. */
export const MAX_SHEET_ROWS = 50000;
export const MAX_SHEET_COLS = 702;

/** Every chart type, for validation. */
export const CHART_TYPES: readonly ChartType[] = ["line", "pie", "area", "stackedBar"];

/** Every pivot aggregate, for validation. */
export const PIVOT_AGGREGATES: readonly PivotAggregate[] = ["sum", "count", "average", "min", "max"];

/** The most charts a sheet holds. */
export const MAX_CHARTS_PER_SHEET = 50;

/** Where a chart may sit and how large it may be, in grid pixels. */
export const CHART_BOUNDS = {
  minX: 48, maxX: 100000, minY: 28, maxY: 2000000,
  minWidth: 280, maxWidth: 1200, minHeight: 200, maxHeight: 900,
} as const;

/**
 * The most cells a chart or pivot range may cover. Clients walk these ranges cell by cell on every
 * render, so a larger one would hang everyone who opens the workbook.
 */
export const MAX_RANGE_CELLS = 200000;

/** The longest comment, in characters. */
export const MAX_COMMENT_LENGTH = 4000;
