// Table references (Excel's structured references): Sales[Amount] is the
// Amount column's data rows, Sales[#All] the whole table with its header and
// totals, Sales[[#Totals],[Amount]] one cell of the totals row, and [@Amount]
// the Amount cell in the formula's own row.

import { type Range, type SheetTable, type Workbook } from "./model";
import { columnNames } from "./xlsx/tables";

/** The table's totals rows, as its XML says (none for tables made here). */
// (In the <table> tag itself: a file's part starts with the <?xml?> declaration.)
export const totalsRows = (table: SheetTable) => Number(/<table\b[^>]*\btotalsRowCount="(\d+)"/.exec(table.xml ?? "")?.[1] ?? 0);

/** The table a cell is in. */
export function tableAt(book: Workbook, sheet: number, r: number, c: number): SheetTable | undefined {
  return book.sheets[sheet]?.tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2);
}

/** Where a table reference points: its sheet and cells; null when there is no such table or column. */
export function resolveTable(book: Workbook, name: string, spec: string, home: { sheet: number; row: number; col: number }): { sheet: number; range: Range } | null {
  let sheet = home.sheet;
  let table: SheetTable | undefined;
  if (name) {
    sheet = book.sheets.findIndex((s) => s.tables.some((t) => t.name.toLowerCase() === name.toLowerCase()));
    table = book.sheets[sheet]?.tables.find((t) => t.name.toLowerCase() === name.toLowerCase());
  } else table = tableAt(book, home.sheet, home.row, home.col);
  if (!table) return null;
  const inner = spec.slice(1, -1).trim();
  // [[#Totals],[Amount]], [[Jan]:[Mar]], or one item: Amount, #All, @Amount.
  const items = inner.startsWith("[") ? [...inner.matchAll(/\[((?:'.|[^\]'])*)\]|(:)/g)].map((m) => m[2] ?? m[1]) : [inner];
  const areas = new Set<string>();
  const columns: string[] = [];
  for (let item of items) {
    if (item === ":") continue;
    item = item.trim();
    if (item.startsWith("@")) {
      areas.add("#this row");
      item = item.slice(1).trim();
      if (item.startsWith("[")) item = item.slice(1, -1);
    }
    if (item.startsWith("#")) areas.add(item.toLowerCase());
    else if (item) columns.push(item.replace(/'(.)/g, "$1"));
  }
  const { r1, c1, r2, c2 } = table.range;
  const totals = totalsRows(table);
  const rows: [number, number][] = [];
  if (!areas.size || areas.has("#data")) rows.push([r1 + 1, r2 - totals]);
  if (areas.has("#all")) rows.push([r1, r2]);
  if (areas.has("#headers")) rows.push([r1, r1]);
  if (areas.has("#totals")) {
    if (!totals) return null;
    rows.push([r2 - totals + 1, r2]);
  }
  if (areas.has("#this row")) {
    if (home.row < r1 + 1 || home.row > r2 - totals || home.sheet !== sheet) return null;
    rows.push([home.row, home.row]);
  }
  const names = columnNames(book.sheets[sheet], table).map((n) => n.toLowerCase());
  const cols = columns.map((col) => names.indexOf(col.toLowerCase()));
  if (cols.some((i) => i < 0)) return null;
  const from = cols.length ? c1 + Math.min(...cols) : c1;
  const to = cols.length ? c1 + Math.max(...cols) : c2;
  return { sheet, range: { r1: Math.min(...rows.map((x) => x[0])), c1: from, r2: Math.max(...rows.map((x) => x[1])), c2: to } };
}

/** Office's accent colours, in the order Excel's table styles use them (the first is black). */
const ACCENTS = ["#000000", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

/** A colour mixed with white: 0.8 is four parts white to one of the colour. */
function tint(hex: string, white: number): string {
  const n = parseInt(hex.slice(1), 16);
  const mix = (v: number) => Math.round(v + (255 - v) * white).toString(16).padStart(2, "0");
  return `#${mix(n >> 16)}${mix((n >> 8) & 255)}${mix(n & 255)}`.toUpperCase();
}

/**
 * How a table's style colours a cell, roughly as Excel draws its built-in styles: the header in
 * the style's colour with white bold text, and every other row lightly tinted.
 * shortcut: one look per style family; the exact borders and column stripes of each style aren't drawn.
 */
export function tableLook(sheet: { tables: SheetTable[] }, r: number, c: number): { fill: string | null; color?: string; bold?: boolean } | null {
  const table = sheet.tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2);
  const m = table && /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(table.style);
  if (!table || !m) return null;
  const n = Number(m[2]);
  const accent = ACCENTS[(n - 1) % 7];
  const header = !/headerRowCount="0"/.test(table.xml ?? "") && r === table.range.r1;
  const totals = r > table.range.r2 - totalsRows(table);
  if (header || totals) {
    // Light 1–7 only underline their header; the rest fill it.
    if (m[1] === "Light" && n <= 7) return { fill: null, bold: true };
    return { fill: m[1] === "Dark" ? accent : header ? accent : tint(accent, 0.6), color: header ? "#FFFFFF" : undefined, bold: true };
  }
  const band = (r - table.range.r1) % 2 === 1;
  return { fill: band ? tint(accent, m[1] === "Dark" ? 0.6 : 0.8) : null };
}
