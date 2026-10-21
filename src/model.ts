// A workbook in memory: sheets of cells, each holding a value or a formula
// and a style (an index into the workbook's styles, as in an .xlsx file).
// Cells are never changed in place; a change puts a new cell object in the
// map, so undo can keep cheap copies of the maps.

import type { CfBlock } from "./conditional";
import type { Validation } from "./validation";
import type { Styles } from "./xlsx/styles";

/** An Excel error value, like #DIV/0! */
export class CellError {
  constructor(public code: "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#SPILL!" | "#CALC!") {}
  toString() {
    return this.code;
  }
}

export type Scalar = number | string | boolean | CellError | null;

export interface Cell {
  /** The value typed in, for a cell without a formula. */
  v?: Scalar;
  /** The formula, without its "=". */
  f?: string;
  /** The style: an index into the workbook's cell formats. */
  s?: number;
}

/** A block of cells, rows and columns counted from 0, both ends included. */
export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** Rows and columns in one number, for map keys. */
export const key = (r: number, c: number) => r * MAX_COLS + c;
export const rowOf = (k: number) => Math.floor(k / MAX_COLS);
export const colOf = (k: number) => k % MAX_COLS;

export interface Filter {
  range: Range;
  /** For each column of the range (by offset) with a filter: the values shown. */
  shown: Map<number, Set<string>>;
}

export type ChartKind = "column" | "bar" | "line" | "pie";

/** A chart on a sheet: its data block (first row names the series, first column the categories) and where it sits. */
export interface Chart {
  id: string;
  kind: ChartKind;
  data: Range;
  title: string;
  /** The cell its top-left corner is in, the offset into that cell, and its size, in pixels at 100%. */
  at: { r: number; c: number; dx: number; dy: number; w: number; h: number };
}

export interface Sheet {
  name: string;
  cells: Map<number, Cell>;
  /** Column widths and row heights in pixels, where not the default. */
  colWidths: Map<number, number>;
  rowHeights: Map<number, number>;
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
  merges: Range[];
  /** Rows and columns kept in view at the top and left. */
  freeze: { rows: number; cols: number };
  filter: Filter | null;
  charts: Chart[];
  /** Conditional formatting: rules for blocks of cells. */
  conditional: CfBlock[];
  /** Data validation: drop-down lists and limits on what can be typed. */
  validations: Validation[];
  defaultColWidth: number;
  defaultRowHeight: number;
  /** For a sheet read from an .xlsx file: its part in the file and its XML, so what isn't edited here is kept. */
  origin?: {
    path: string;
    xml: string;
    /** Its drawing (charts, pictures): the part, its XML without the charts shown here, and the chart parts those were. */
    drawing?: { path: string; xml: string; owned: string[] };
  };
}

export interface Workbook {
  sheets: Sheet[];
  styles: Styles;
  /** Dates count from 1904 instead of 1900 (old Mac files). */
  date1904: boolean;
  /** Named ranges: name (upper case) → formula text, like Sheet1!$A$1:$B$4. */
  names: Map<string, string>;
  /** The .xlsx package the workbook was read from (a JSZip), kept for everything not edited here. */
  source?: unknown;
}

export function newSheet(name: string): Sheet {
  return {
    name,
    cells: new Map(),
    colWidths: new Map(),
    rowHeights: new Map(),
    hiddenCols: new Set(),
    hiddenRows: new Set(),
    merges: [],
    freeze: { rows: 0, cols: 0 },
    filter: null,
    charts: [],
    conditional: [],
    validations: [],
    defaultColWidth: 64,
    defaultRowHeight: 20,
  };
}

// ---- addresses ----------------------------------------------------------------------

/** 0 → A, 25 → Z, 26 → AA. */
export function colName(c: number): string {
  let name = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** A → 0, AA → 26; -1 if not a column. */
export function colIndex(name: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(name)) return -1;
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1 < MAX_COLS ? n - 1 : -1;
}

export const address = (r: number, c: number) => `${colName(c)}${r + 1}`;

export function rangeAddress(range: Range): string {
  const a = address(range.r1, range.c1);
  return range.r1 === range.r2 && range.c1 === range.c2 ? a : `${a}:${address(range.r2, range.c2)}`;
}

/** "B3" → {r: 2, c: 1}; null if it isn't a cell address. */
export function parseAddress(text: string): { r: number; c: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = Number(m[2]) - 1;
  return c < 0 || r < 0 || r >= MAX_ROWS ? null : { r, c };
}

/** "A1:C4" or "B2" → a range; null if neither. */
export function parseRange(text: string): Range | null {
  const [a, b] = text.split(":");
  const p = parseAddress(a);
  const q = b === undefined ? p : parseAddress(b);
  if (!p || !q) return null;
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
}

export const inRange = (range: Range, r: number, c: number) => r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2;

/** The merged block a cell belongs to, if any. */
export const mergeAt = (sheet: Sheet, r: number, c: number) => sheet.merges.find((m) => inRange(m, r, c));

/** The rows and columns that have something in them (from A1). */
export function usedRange(sheet: Sheet): Range | null {
  let r2 = -1;
  let c2 = -1;
  for (const k of sheet.cells.keys()) {
    r2 = Math.max(r2, rowOf(k));
    c2 = Math.max(c2, colOf(k));
  }
  return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
}

/** A block with another taken out: what's left, in up to four pieces. */
export function minusRange(g: Range, cut: Range): Range[] {
  if (g.r2 < cut.r1 || g.r1 > cut.r2 || g.c2 < cut.c1 || g.c1 > cut.c2) return [g];
  const out: Range[] = [];
  if (g.r1 < cut.r1) out.push({ ...g, r2: cut.r1 - 1 });
  if (g.r2 > cut.r2) out.push({ ...g, r1: cut.r2 + 1 });
  const r1 = Math.max(g.r1, cut.r1);
  const r2 = Math.min(g.r2, cut.r2);
  if (g.c1 < cut.c1) out.push({ r1, r2, c1: g.c1, c2: cut.c1 - 1 });
  if (g.c2 > cut.c2) out.push({ r1, r2, c1: cut.c2 + 1, c2: g.c2 });
  return out;
}
