// Reads an Excel workbook (.xlsx, .xlsm): sheets, values, formulas, cell
// formats, column widths and row heights, merged cells, frozen panes and
// filters. The package itself is kept, so writing it back keeps everything
// DuckSheets doesn't edit (charts, comments, conditional formatting, …).

import JSZip from "jszip";
import { FormulaError, parse, print, shifted } from "../formula/parse";
import { CellError, MAX_COLS, MAX_ROWS, key, newSheet, parseRange, type Cell, type Range, type Scalar, type Sheet, type Workbook } from "../model";
import { readStyles, readTheme } from "./styles";
import { readCharts } from "./charts";
import { readConditional } from "../conditional";
import { ARRAY_FUNCTIONS } from "../calc";
import { readValidations } from "../validation";
import { readNotes } from "./notes";
import { readLinks } from "./links";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const xml = (text: string) => new DOMParser().parseFromString(text, "application/xml");
const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] ?? null;
const relId = (el: Element) => el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ?? el.getAttribute("r:id");

/** A target in a .rels file, as a path in the package. */
export function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const piece of target.split("/")) {
    if (piece === "..") parts.pop();
    else if (piece !== ".") parts.push(piece);
  }
  return parts.join("/");
}

export const relsPath = (part: string) => {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
};

/** The relationships of a part: id → { type, target path }. */
export async function readRels(zip: JSZip, part: string): Promise<Map<string, { type: string; path: string }>> {
  const out = new Map<string, { type: string; path: string }>();
  const text = await zip.file(relsPath(part))?.async("string");
  if (!text) return out;
  for (const rel of Array.from(xml(text).getElementsByTagNameNS("*", "Relationship"))) {
    if (rel.getAttribute("TargetMode") === "External") continue;
    out.set(rel.getAttribute("Id")!, { type: rel.getAttribute("Type") ?? "", path: resolvePath(part, rel.getAttribute("Target") ?? "") });
  }
  return out;
}

/** Excel's column width (in characters of Calibri 11) in pixels, and back. */
export const widthToPx = (width: number) => Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7);
export const pxToWidth = (px: number) => Math.trunc(((px - 0) / 7) * 256) / 256;

export class NotAWorkbook extends Error {}

export async function readXlsx(bytes: Uint8Array): Promise<Workbook> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new NotAWorkbook("This file isn't an Excel workbook (.xlsx). It may be damaged, or an old .xls file.");
  }
  // The workbook part, from the package's own relationships.
  const root = await readRels(zip, "");
  const main = [...root.values()].find((r) => r.type.endsWith("/officeDocument"))?.path ?? "xl/workbook.xml";
  const workbookText = await zip.file(main)?.async("string");
  if (!workbookText) throw new NotAWorkbook("This file has no workbook in it.");
  const wbDoc = xml(workbookText);
  const rels = await readRels(zip, main);
  const byType = (suffix: string) => [...rels.values()].find((r) => r.type.endsWith(suffix))?.path;

  const theme = readTheme((await zip.file(byType("/theme") ?? "")?.async("string")) ?? null);
  const styles = readStyles((await zip.file(byType("/styles") ?? "")?.async("string")) ?? null, theme);
  const strings = readSharedStrings((await zip.file(byType("/sharedStrings") ?? "")?.async("string")) ?? null);
  const workbookPr = wbDoc.getElementsByTagNameNS("*", "workbookPr")[0];
  const date1904 = workbookPr?.getAttribute("date1904") === "1" || workbookPr?.getAttribute("date1904") === "true";

  const book: Workbook = { sheets: [], styles, date1904, names: new Map(), source: zip };
  for (const el of Array.from(wbDoc.getElementsByTagNameNS("*", "definedName"))) {
    const name = el.getAttribute("name") ?? "";
    // Excel's own names (print areas, filter ranges) and names for one sheet only aren't offered.
    if (!name || name.startsWith("_xlnm.") || el.getAttribute("localSheetId") !== null) continue;
    book.names.set(name.toUpperCase(), el.textContent ?? "");
  }

  for (const el of Array.from(wbDoc.getElementsByTagNameNS("*", "sheet"))) {
    const rel = rels.get(relId(el) ?? "");
    if (!rel || !rel.type.endsWith("/worksheet")) continue; // chart sheets aren't shown yet
    const text = await zip.file(rel.path)?.async("string");
    if (!text) continue;
    const sheet = readSheet(el.getAttribute("name") ?? `Sheet${book.sheets.length + 1}`, text, strings, styles.theme);
    sheet.origin = { path: rel.path, xml: text };
    await readCharts(zip, sheet);
    sheet.notes = await readNotes(zip, rel.path);
    sheet.links = await readLinks(zip, rel.path, new DOMParser().parseFromString(text, "application/xml").documentElement);
    book.sheets.push(sheet);
  }
  if (!book.sheets.length) throw new NotAWorkbook("This workbook has no sheets DuckSheets can show.");
  return book;
}

function readSharedStrings(text: string | null): string[] {
  if (!text) return [];
  return Array.from(xml(text).getElementsByTagNameNS(MAIN, "si")).map((si) =>
    // Plain text, or the runs of formatted text joined (their formatting isn't kept yet).
    Array.from(si.getElementsByTagNameNS(MAIN, "t"))
      .filter((t) => t.parentElement?.localName !== "rPh")
      .map((t) => t.textContent ?? "")
      .join(""),
  );
}

/** Reads "B12" quickly, without checks: cell references in files are well formed. */
function cellPos(ref: string): [number, number] {
  let c = 0;
  let i = 0;
  for (; i < ref.length && ref.charCodeAt(i) >= 65; i++) c = c * 26 + (ref.charCodeAt(i) - 64);
  return [Number(ref.slice(i)) - 1, c - 1];
}

const ERROR_CODES = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"]);

/** A formula from a file: _xlfn. prefixes off, in DuckSheets' own spelling; kept as it was if it can't be read. */
function formulaText(text: string): string {
  try {
    return print(parse(text));
  } catch (error) {
    if (!(error instanceof FormulaError)) throw error;
    return text;
  }
}

export function readSheet(name: string, text: string, strings: string[], theme: string[] = []): Sheet {
  const doc = xml(text);
  const sheet = newSheet(name);
  const ws = doc.documentElement;

  const format = kid(ws, "sheetFormatPr");
  if (format?.getAttribute("defaultRowHeight")) sheet.defaultRowHeight = Math.round((Number(format.getAttribute("defaultRowHeight")) * 4) / 3);
  if (format?.getAttribute("defaultColWidth")) sheet.defaultColWidth = widthToPx(Number(format.getAttribute("defaultColWidth")));
  else if (format?.getAttribute("baseColWidth")) sheet.defaultColWidth = widthToPx(Number(format.getAttribute("baseColWidth")) + 0.7109375);

  for (const col of kids(kid(ws, "cols"), "col")) {
    const min = Number(col.getAttribute("min")) - 1;
    const max = Math.min(Number(col.getAttribute("max")) - 1, MAX_COLS - 1);
    const width = col.getAttribute("width");
    const hidden = col.getAttribute("hidden") === "1" || col.getAttribute("hidden") === "true";
    // A width for every column to the end is the sheet's default width.
    if (max - min > 500) {
      if (width && min === 0) sheet.defaultColWidth = widthToPx(Number(width));
      continue;
    }
    for (let c = min; c <= max; c++) {
      if (width && col.getAttribute("customWidth") !== "0") sheet.colWidths.set(c, widthToPx(Number(width)));
      if (hidden) sheet.hiddenCols.add(c);
    }
  }

  // Shared formulas: the first cell holds the formula; the others say which one, and move it by their offset.
  const shared = new Map<string, { formula: string; r: number; c: number }>();
  /** Blocks filled by dynamic array formulas: the values Excel saved in them are worked out again here. */
  const spills: { range: Range; r: number; c: number }[] = [];
  let nextRow = 0;
  for (const row of kids(kid(ws, "sheetData"), "row")) {
    const r = row.getAttribute("r") ? Number(row.getAttribute("r")) - 1 : nextRow;
    nextRow = r + 1;
    if (r >= MAX_ROWS) continue;
    if (row.getAttribute("customHeight") === "1" || row.getAttribute("customHeight") === "true") sheet.rowHeights.set(r, Math.round((Number(row.getAttribute("ht")) * 4) / 3));
    if (row.getAttribute("hidden") === "1" || row.getAttribute("hidden") === "true") sheet.hiddenRows.add(r);
    let nextCol = 0;
    for (const c of kids(row, "c")) {
      const ref = c.getAttribute("r");
      const [, col] = ref ? cellPos(ref) : [r, nextCol];
      nextCol = col + 1;
      const type = c.getAttribute("t") ?? "n";
      const vText = kid(c, "v")?.textContent;
      const s = c.getAttribute("s");
      const cell: Cell = {};
      if (s && s !== "0") cell.s = Number(s);
      let value: Scalar = null;
      if (type === "s") value = strings[Number(vText)] ?? "";
      else if (type === "inlineStr") value = Array.from(c.getElementsByTagNameNS(MAIN, "t")).map((t) => t.textContent ?? "").join("");
      else if (type === "b") value = vText === "1";
      else if (type === "e") value = ERROR_CODES.has(vText ?? "") ? new CellError(vText as CellError["code"]) : new CellError("#VALUE!");
      else if (type === "str") value = vText ?? "";
      else if (vText !== null && vText !== undefined && vText !== "") value = Number(vText);
      const f = kid(c, "f");
      if (f) {
        const si = f.getAttribute("si");
        const fText = f.textContent ?? "";
        if (f.getAttribute("t") === "shared" && si !== null) {
          if (fText) {
            shared.set(si, { formula: fText, r, c: col });
            cell.f = formulaText(fText);
          } else {
            const master = shared.get(si);
            if (master) {
              try {
                cell.f = print(shifted(parse(master.formula), r - master.r, col - master.c));
              } catch {
                cell.f = master.formula;
              }
            }
          }
        } else if (fText) cell.f = formulaText(fText);
        const block = f.getAttribute("t") === "array" ? parseRange(f.getAttribute("ref") ?? "") : null;
        if (block && cell.f && ARRAY_FUNCTIONS.test(cell.f)) spills.push({ range: block, r, c: col });
      }
      if (cell.f === undefined && value !== null) cell.v = value;
      if (cell.f !== undefined || cell.v !== undefined || cell.s !== undefined) sheet.cells.set(key(r, col), cell);
    }
  }

  for (const { range, r, c } of spills)
    for (let i = range.r1; i <= range.r2; i++)
      for (let j = range.c1; j <= range.c2; j++) {
        const k = key(i, j);
        const cell = sheet.cells.get(k);
        if ((i === r && j === c) || !cell || cell.f !== undefined) continue;
        if (cell.s !== undefined) sheet.cells.set(k, { s: cell.s });
        else sheet.cells.delete(k);
      }

  for (const m of kids(kid(ws, "mergeCells"), "mergeCell")) {
    const range = parseRange(m.getAttribute("ref") ?? "");
    if (range) sheet.merges.push(range);
  }

  const pane = kid(kid(kid(ws, "sheetViews"), "sheetView"), "pane");
  if (pane && (pane.getAttribute("state") === "frozen" || pane.getAttribute("state") === "frozenSplit")) {
    sheet.freeze = { rows: Number(pane.getAttribute("ySplit") ?? 0), cols: Number(pane.getAttribute("xSplit") ?? 0) };
  }

  const filter = kid(ws, "autoFilter");
  const filterRange = filter && parseRange(filter.getAttribute("ref") ?? "");
  if (filterRange) {
    sheet.filter = { range: filterRange, shown: new Map() };
    for (const fc of kids(filter, "filterColumn")) {
      const filters = kid(fc, "filters");
      if (!filters) continue;
      const shown = new Set(kids(filters, "filter").map((x) => x.getAttribute("val") ?? ""));
      if (filters.getAttribute("blank") === "1") shown.add("");
      sheet.filter.shown.set(Number(fc.getAttribute("colId")), shown);
    }
  }
  sheet.conditional = readConditional(ws, theme);
  sheet.validations = readValidations(ws);
  return sheet;
}

