// Writes a workbook as .xlsx. A workbook read from a file is written back
// into its own package: the sheets' data, the cell formats and the sheet list
// are rebuilt, and every other part (charts, pictures, comments, conditional
// formatting, print settings, …) is kept as it was.

import JSZip from "jszip";
import type { Calculator } from "../calc";
import { parse, print } from "../formula/parse";
import { CellError, address, colOf, inRange, key, rangeAddress, rowOf, usedRange, type Range, type Scalar, type Sheet, type SheetTable, type Workbook } from "../model";
import { formatValue } from "../numfmt";
import { pxToWidth, readRels, relsPath, resolvePath, siText } from "./read";
import { lookOf, writeStyles } from "./styles";
import { EMPTY_DRAWING, anchorXml, chartXml, pictureAnchorXml } from "./charts";
import { chartData } from "../chartdata";
import { conditionalXml } from "../conditional";
import { ARRAY_FUNCTIONS } from "../calc";
import { validationsXml } from "../validation";
import { writeNotes } from "./notes";
import { writeLinks } from "./links";
import { writeTables } from "./tables";
import { totalsRows } from "../structured";
import { PIVOT_TYPE, pivotAreas, writePivots } from "./pivots";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const TYPE = {
  sheet: `${REL}/worksheet`,
  styles: `${REL}/styles`,
  strings: `${REL}/sharedStrings`,
  calcChain: `${REL}/calcChain`,
  metadata: `${REL}/sheetMetadata`,
};
const CONTENT = {
  workbook: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  sheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  styles: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  strings: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  metadata: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
};

/** Excel's cell metadata for dynamic arrays: cells with cm="1" spill. */
const DYNAMIC_ARRAY_METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
// Characters XML can't hold at all; Excel writes them as _x001F_.
const escText = (s: string) => esc(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).padStart(4, "0").toUpperCase()}_`);
const xml = (text: string) => new DOMParser().parseFromString(text, "application/xml");
const serialize = (doc: Document) => {
  const text = new XMLSerializer().serializeToString(doc);
  return text.startsWith("<?xml") ? text : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${text}`;
};

/** The order Excel requires for a worksheet's parts. */
const SHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
  "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks",
  "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
];

/** Puts a part in a worksheet at its proper place (replacing what was there), or removes it (xmlText null). */
function setPart(doc: Document, name: string, xmlText: string | null) {
  const ws = doc.documentElement;
  const parsed = xmlText === null ? null : xml(`<x xmlns="${MAIN}" xmlns:r="${REL}">${xmlText}</x>`);
  // XML that doesn't parse would drop the part: the file's own copy is kept instead.
  if (parsed?.getElementsByTagName("parsererror").length) {
    console.warn(`DuckSheets: the sheet's ${name} couldn't be written; the file's own is kept`);
    return;
  }
  for (const el of Array.from(ws.children)) if (el.localName === name) ws.removeChild(el);
  if (!parsed) return;
  const at = SHEET_ORDER.indexOf(name);
  const after = Array.from(ws.children).find((el) => SHEET_ORDER.indexOf(el.localName) > at);
  // Some parts come several times in a row (conditionalFormatting).
  for (const fragment of Array.from(parsed.documentElement.children)) ws.insertBefore(doc.importNode(fragment, true), after ?? null);
}

class SharedStrings {
  list: { text: string; rich?: string }[] = [];
  private index = new Map<string, number>();
  /** A text's number in the list; `rich` is its runs of formatted text, as read from a file. */
  add(s: string, rich?: string): number {
    const id = rich ? `\u0000${rich}` : s;
    let i = this.index.get(id);
    if (i === undefined) {
      i = this.list.length;
      this.list.push({ text: s, rich });
      this.index.set(id, i);
    }
    return i;
  }
  xml(): string {
    const items = this.list.map(({ text: s, rich }) => (rich ? `<si>${rich}</si>` : `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ""}>${escText(s)}</t></si>`)).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${MAIN}" count="${this.list.length}" uniqueCount="${this.list.length}">${items}</sst>`;
  }
}

/** The text of a shared string's runs (without phonetic guides). */
const richText = (runs: string) => siText(xml(`<si xmlns="${MAIN}">${runs}</si>`).documentElement);

function valueXml(v: Scalar, strings: SharedStrings, formula: boolean): { t: string; v: string } | null {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return { t: "", v: String(v) };
  if (typeof v === "boolean") return { t: ' t="b"', v: v ? "1" : "0" };
  if (v instanceof CellError) return { t: ' t="e"', v: v.code };
  return formula ? { t: ' t="str"', v: escText(v) } : { t: ' t="s"', v: String(strings.add(v)) };
}

function sheetDataXml(sheet: Sheet, index: number, calc: Calculator, strings: SharedStrings): string {
  // Dynamic array formulas: the block each fills, and the values they put in the cells around them.
  const spillRefs = new Map<number, string>();
  const spilled = new Map<number, Scalar>();
  for (const [k, cell] of sheet.cells) {
    if (cell.f === undefined || !(cell.dyn || ARRAY_FUNCTIONS.test(cell.f))) continue;
    // A dynamic array from Excel that gives one value stays one, in its own cell.
    const block = calc.spillRange(index, rowOf(k), colOf(k)) ?? (cell.dyn ? { r1: rowOf(k), c1: colOf(k), r2: rowOf(k), c2: colOf(k) } : undefined);
    if (!block) continue;
    spillRefs.set(k, rangeAddress(block));
    for (let r = block.r1; r <= block.r2; r++) for (let c = block.c1; c <= block.c2; c++) if (key(r, c) !== k) spilled.set(key(r, c), calc.value(index, r, c));
  }
  const rows = new Map<number, number[]>();
  for (const k of new Set([...sheet.cells.keys(), ...spilled.keys()])) {
    const r = rowOf(k);
    let list = rows.get(r);
    if (!list) rows.set(r, (list = []));
    list.push(k);
  }
  for (const r of [...sheet.rowHeights.keys(), ...sheet.hiddenRows, ...(sheet.rowAttrs?.keys() ?? [])]) if (!rows.has(r)) rows.set(r, []);
  const out: string[] = [];
  for (const r of [...rows.keys()].sort((a, b) => a - b)) {
    const height = sheet.rowHeights.get(r);
    const attrs = `${height !== undefined ? ` ht="${Math.round(height * 0.75 * 100) / 100}" customHeight="1"` : ""}${sheet.hiddenRows.has(r) ? ' hidden="1"' : ""}${sheet.rowAttrs?.get(r) ?? ""}`;
    const cells = rows
      .get(r)!
      .sort((a, b) => a - b)
      .map((k) => {
        const cell = sheet.cells.get(k) ?? {};
        const ref = address(r, colOf(k));
        const s = cell.s ? ` s="${cell.s}"` : "";
        if (spilled.has(k)) {
          const v = valueXml(spilled.get(k)!, strings, true);
          return v ? `<c r="${ref}"${s}${v.t}><v>${v.v}</v></c>` : `<c r="${ref}"${s}/>`;
        }
        if (cell.f !== undefined) {
          let f = cell.f;
          try {
            f = print(parse(cell.f), true);
          } catch {
            // Kept as typed.
          }
          const cached = valueXml(calc.value(index, r, colOf(k)), strings, true);
          // A spilling formula is a dynamic array (cm="1" points at its entry in metadata.xml).
          const spill = spillRefs.get(k);
          const fXml = spill ? `<f t="array" ref="${spill}">${escText(f)}</f>` : `<f>${escText(f)}</f>`;
          return `<c r="${ref}"${s}${cached?.t ?? ""}${spill ? ' cm="1"' : ""}>${fXml}${cached ? `<v>${cached.v}</v>` : ""}</c>`;
        }
        // Text with formatted parts keeps them while it is still the same text.
        if (cell.rich && typeof cell.v === "string" && richText(cell.rich) === cell.v) return `<c r="${ref}"${s} t="s"><v>${strings.add(cell.v, cell.rich)}</v></c>`;
        const v = valueXml(cell.v ?? null, strings, false);
        return v ? `<c r="${ref}"${s}${v.t}><v>${v.v}</v></c>` : `<c r="${ref}"${s}/>`;
      })
      .join("");
    out.push(`<row r="${r + 1}"${attrs}>${cells}</row>`);
  }
  return `<sheetData>${out.join("")}</sheetData>`;
}

function colsXml(sheet: Sheet): string | null {
  const columns = [...new Set([...sheet.colWidths.keys(), ...sheet.hiddenCols, ...(sheet.colAttrs?.keys() ?? [])])].sort((a, b) => a - b);
  if (!columns.length) return null;
  const parts: string[] = [];
  // Neighbours with the same width, hiding and other settings are written as one.
  const look = (c: number) => `${sheet.colWidths.get(c) ?? ""}|${sheet.hiddenCols.has(c)}|${sheet.colAttrs?.get(c) ?? ""}`;
  for (let i = 0; i < columns.length; ) {
    const c = columns[i];
    const width = sheet.colWidths.get(c) ?? sheet.defaultColWidth;
    const hidden = sheet.hiddenCols.has(c);
    let j = i;
    while (j + 1 < columns.length && columns[j + 1] === columns[j] + 1 && look(columns[j + 1]) === look(c)) j++;
    parts.push(`<col min="${c + 1}" max="${columns[j] + 1}" width="${pxToWidth(width)}"${sheet.colWidths.has(c) ? ' customWidth="1"' : ""}${hidden ? ' hidden="1"' : ""}${sheet.colAttrs?.get(c) ?? ""}/>`);
    i = j + 1;
  }
  return `<cols>${parts.join("")}</cols>`;
}

function paneXml(sheet: Sheet): string {
  const { rows, cols } = sheet.freeze;
  if (!rows && !cols) return "";
  const active = rows && cols ? "bottomRight" : rows ? "bottomLeft" : "topRight";
  return `<pane${cols ? ` xSplit="${cols}"` : ""}${rows ? ` ySplit="${rows}"` : ""} topLeftCell="${address(rows, cols)}" activePane="${active}" state="frozen"/>`;
}

function filterXml(sheet: Sheet): string | null {
  if (!sheet.filter) return null;
  const lists = [...sheet.filter.shown].map(([col, shown]): [number, string] => {
    const values = [...shown].filter((v) => v !== "");
    return [col, `<filters${shown.has("") ? ' blank="1"' : ""}>${values.map((v) => `<filter val="${esc(v)}"/>`).join("")}</filters>`];
  });
  const rules = [...(sheet.filter.custom ?? [])].map(([col, rule]): [number, string] => [
    col,
    rule.kind === "top"
      ? `<top10${rule.top ? "" : ' top="0"'}${rule.percent ? ' percent="1"' : ""} val="${rule.count}"/>`
      : `<customFilters${rule.and ? ' and="1"' : ""}>${rule.conditions.map((c) => `<customFilter${c.op === "equal" ? "" : ` operator="${c.op}"`} val="${esc(c.val)}"/>`).join("")}</customFilters>`,
  ]);
  const columns = [...lists, ...rules].sort(([a], [b]) => a - b).map(([col, xml]) => `<filterColumn colId="${col}">${xml}</filterColumn>`);
  return `<autoFilter ref="${rangeAddress(sheet.filter.range)}">${columns.join("")}</autoFilter>`;
}

/** Whether the sheet's filter is the table's own: on its rows, above any totals. */
const sameRange = (a: Range, t: SheetTable) => a.r1 === t.range.r1 && a.c1 === t.range.c1 && a.r2 === t.range.r2 - totalsRows(t) && a.c2 === t.range.c2;

function sheetXml(sheet: Sheet, index: number, calc: Calculator, strings: SharedStrings, selected: boolean, drawing: string | null, legacy: string | null | undefined, links: string | null, tables: string | null | undefined): string {
  const used = usedRange(sheet);
  const doc = xml(sheet.origin?.xml ?? `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  setPart(doc, "dimension", `<dimension ref="${used ? rangeAddress(used) : "A1"}"/>`);
  // The sheet view: its frozen pane (selections are dropped, as they name the panes).
  const views = Array.from(doc.documentElement.children).find((el) => el.localName === "sheetViews");
  const view = views && Array.from(views.children).find((el) => el.localName === "sheetView");
  if (view) {
    for (const el of Array.from(view.children)) if (el.localName === "pane" || el.localName === "selection") view.removeChild(el);
    const pane = paneXml(sheet);
    if (pane) view.insertBefore(doc.importNode(xml(`<x xmlns="${MAIN}">${pane}</x>`).documentElement.firstElementChild!, true), view.firstChild);
    if (!selected) view.removeAttribute("tabSelected");
  } else setPart(doc, "sheetViews", `<sheetViews><sheetView${selected ? ' tabSelected="1"' : ""} workbookViewId="0">${paneXml(sheet)}</sheetView></sheetViews>`);
  const format = Array.from(doc.documentElement.children).find((el) => el.localName === "sheetFormatPr");
  if (!format) setPart(doc, "sheetFormatPr", `<sheetFormatPr defaultRowHeight="${Math.round(sheet.defaultRowHeight * 0.75 * 100) / 100}"/>`);
  setPart(doc, "cols", colsXml(sheet));
  setPart(doc, "sheetData", sheetDataXml(sheet, index, calc, strings));
  // A table's filter is written in the table; Excel won't have it twice.
  setPart(doc, "autoFilter", sheet.tables.some((t) => sheet.filter && sameRange(sheet.filter.range, t)) ? null : filterXml(sheet));
  if (tables !== undefined) setPart(doc, "tableParts", tables);
  // A drawing (for charts) the sheet didn't have before.
  if (drawing && !Array.from(doc.documentElement.children).some((el) => el.localName === "drawing")) setPart(doc, "drawing", `<drawing r:id="${drawing}"/>`);
  setPart(doc, "conditionalFormatting", conditionalXml(sheet.conditional));
  setPart(doc, "dataValidations", validationsXml(sheet.validations));
  setPart(doc, "hyperlinks", links);
  if (legacy !== undefined) setPart(doc, "legacyDrawing", legacy ? `<legacyDrawing r:id="${legacy}"/>` : null);
  setPart(doc, "mergeCells", sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rangeAddress(m)}"/>`).join("")}</mergeCells>` : null);
  return serialize(doc);
}

// ---- charts --------------------------------------------------------------------------------------

/** A relationship target: `to` as seen from the folder `from` is in. */
function relative(from: string, to: string): string {
  const a = from.split("/").slice(0, -1);
  const b = to.split("/");
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return [...a.slice(i).map(() => ".."), ...b.slice(i)].join("/");
}

const emptyRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL}"></Relationships>`;

/**
 * Writes a sheet's drawing: what was in it and isn't shown here, kept as it was,
 * and every chart, as new chart parts. Returns the relationship id the sheet
 * needs for a drawing it didn't have, or null.
 */
async function writeDrawing(zip: JSZip, book: Workbook, calc: Calculator, index: number, sheetPath: string, override: (path: string, type: string | null) => void): Promise<string | null> {
  const sheet = book.sheets[index];
  const kept = sheet.origin?.drawing;
  if (!sheet.charts.length && !kept) return null;
  let drawingPath = kept?.path;
  if (!drawingPath) {
    let k = 1;
    while (zip.file(`xl/drawings/drawing${k}.xml`)) k++;
    drawingPath = `xl/drawings/drawing${k}.xml`;
  }
  const drawingDoc = xml(kept?.xml ?? EMPTY_DRAWING);
  const drawingRelsPath = relsPath(drawingPath);
  const drawingRels = xml((await zip.file(drawingRelsPath)?.async("string")) ?? emptyRels);
  const relRoot = drawingRels.documentElement;
  // The charts read from the file are written again from what's shown here.
  for (const path of kept?.owned ?? []) {
    zip.remove(path);
    zip.remove(relsPath(path));
    override(path, null);
    for (const el of Array.from(relRoot.children)) if (resolvePath(drawingPath, el.getAttribute("Target") ?? "") === path) relRoot.removeChild(el);
  }
  let nextRel = Math.max(0, ...Array.from(relRoot.children).map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1;
  let nextShape = Math.max(1, ...Array.from(drawingDoc.getElementsByTagNameNS("*", "cNvPr")).map((el) => Number(el.getAttribute("id")) || 0)) + 1;
  let k = 1;
  let m = 1;
  for (const chart of sheet.charts) {
    if (chart.kind === "picture") {
      // The picture's bytes in xl/media, linked from the drawing.
      const match = /^data:image\/(png|jpeg|gif|bmp);base64,(.*)$/.exec(chart.src ?? "");
      if (!match) continue;
      const ext = match[1] === "jpeg" ? "jpg" : match[1];
      while (zip.file(`xl/media/duckpicture${m}.${ext}`)) m++;
      const mediaPath = `xl/media/duckpicture${m}.${ext}`;
      zip.file(mediaPath, match[2], { base64: true });
      override(mediaPath, `image/${match[1]}`);
      const picRel = `rId${nextRel++}`;
      const rel = drawingRels.createElementNS(PKG_REL, "Relationship");
      rel.setAttribute("Id", picRel);
      rel.setAttribute("Type", `${REL}/image`);
      rel.setAttribute("Target", relative(drawingPath, mediaPath));
      relRoot.appendChild(rel);
      const anchor = xml(`<x xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${pictureAnchorXml(chart, sheet, picRel, nextShape++)}</x>`).documentElement.firstElementChild!;
      drawingDoc.documentElement.appendChild(drawingDoc.importNode(anchor, true));
      continue;
    }
    while (zip.file(`xl/charts/chart${k}.xml`)) k++;
    const chartPath = `xl/charts/chart${k}.xml`;
    const source = {
      value: (r: number, c: number, s: number) => calc.value(s, r, c),
      shown: (r: number, c: number, s: number) => formatValue(calc.value(s, r, c), lookOf(book.styles, book.sheets[s].cells.get(key(r, c))?.s).numFmt, book.date1904).text,
      cell: (r: number, c: number, s: number) => book.sheets[s].cells.get(key(r, c)),
      book,
    };
    zip.file(chartPath, chartXml(chart, sheet.name, chartData(source, chart, index)));
    override(chartPath, CONTENT.chart);
    const relId = `rId${nextRel++}`;
    const rel = drawingRels.createElementNS(PKG_REL, "Relationship");
    rel.setAttribute("Id", relId);
    rel.setAttribute("Type", `${REL}/chart`);
    rel.setAttribute("Target", relative(drawingPath, chartPath));
    relRoot.appendChild(rel);
    const anchor = xml(`<x xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorXml(chart, sheet, relId, nextShape++)}</x>`).documentElement.firstElementChild!;
    drawingDoc.documentElement.appendChild(drawingDoc.importNode(anchor, true));
  }
  zip.file(drawingPath, serialize(drawingDoc));
  zip.file(drawingRelsPath, serialize(drawingRels));
  override(drawingPath, CONTENT.drawing);
  if (kept) return null;
  // The sheet's own relationships gain the drawing.
  const sheetRelsPath = relsPath(sheetPath);
  const sheetRels = xml((await zip.file(sheetRelsPath)?.async("string")) ?? emptyRels);
  const root = sheetRels.documentElement;
  const id = `rId${Math.max(0, ...Array.from(root.children).map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1}`;
  const rel = sheetRels.createElementNS(PKG_REL, "Relationship");
  rel.setAttribute("Id", id);
  rel.setAttribute("Type", `${REL}/drawing`);
  rel.setAttribute("Target", relative(sheetPath, drawingPath));
  root.appendChild(rel);
  zip.file(sheetRelsPath, serialize(sheetRels));
  return id;
}

// ---- the package ---------------------------------------------------------------------------------

const STARTER: Record<string, string> = {
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${CONTENT.workbook}"/></Types>`,
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${MAIN}" xmlns:r="${REL}"><workbookPr/><bookViews><workbookView/></bookViews><sheets/><calcPr calcId="191029"/></workbook>`,
  "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL}"></Relationships>`,
};

/** The workbook as .xlsx bytes. `active` is the sheet shown when the file is opened. */
export async function writeXlsx(book: Workbook, calc: Calculator, active = 0): Promise<Uint8Array> {
  const zip = book.source instanceof JSZip ? await JSZip.loadAsync(await (book.source as JSZip).generateAsync({ type: "uint8array" })) : new JSZip();
  for (const [path, text] of Object.entries(STARTER)) if (!zip.file(path)) zip.file(path, text);
  const root = await readRels(zip, "");
  const main = [...root.values()].find((r) => r.type.endsWith("/officeDocument"))?.path ?? "xl/workbook.xml";
  const relsFile = relsPath(main);
  const wbDoc = xml(await zip.file(main)!.async("string"));
  const relsDoc = xml((await zip.file(relsFile)?.async("string")) ?? STARTER["xl/_rels/workbook.xml.rels"]);
  const typesDoc = xml(await zip.file("[Content_Types].xml")!.async("string"));
  const dir = main.slice(0, main.lastIndexOf("/") + 1);
  const relsRoot = relsDoc.documentElement;
  const relEls = () => Array.from(relsRoot.children).filter((el) => el.localName === "Relationship");
  const target = (path: string) => (path.startsWith(dir) ? path.slice(dir.length) : `/${path}`);
  let nextRel = Math.max(0, ...relEls().map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1;
  const addRel = (type: string, path: string, each = false) => {
    const existing = relEls().find((el) => el.getAttribute("Type") === type && ((type !== TYPE.sheet && !each) || el.getAttribute("Target") === target(path)));
    if (existing) return existing.getAttribute("Id")!;
    const el = relsDoc.createElementNS(PKG_REL, "Relationship");
    const id = `rId${nextRel++}`;
    el.setAttribute("Id", id);
    el.setAttribute("Type", type);
    el.setAttribute("Target", target(path));
    relsRoot.appendChild(el);
    return id;
  };
  const override = (path: string, contentType: string | null) => {
    const name = `/${path}`;
    for (const el of Array.from(typesDoc.documentElement.children)) if (el.localName === "Override" && el.getAttribute("PartName") === name) typesDoc.documentElement.removeChild(el);
    if (!contentType) return;
    const el = typesDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override");
    el.setAttribute("PartName", name);
    el.setAttribute("ContentType", contentType);
    typesDoc.documentElement.appendChild(el);
  };

  // The calculation chain lists formula cells as they were; Excel rebuilds it.
  for (const el of relEls()) {
    if (el.getAttribute("Type") !== TYPE.calcChain) continue;
    const path = resolvePath(main, el.getAttribute("Target") ?? "");
    zip.remove(path);
    override(path, null);
    relsRoot.removeChild(el);
  }

  // Sheets: each one's part (its own, or a new one), then the list in the workbook.
  const strings = new SharedStrings();
  const used = new Set(book.sheets.map((s) => s.origin?.path).filter(Boolean));
  let n = 1;
  const sheetEls: string[] = [];
  let tableId = 1;
  const pivotCaches: { path: string; id: number }[] = [];
  // Cache ids go on from the ones a file already has.
  let cacheId = Math.max(0, ...Array.from(wbDoc.getElementsByTagNameNS("*", "pivotCache")).map((el) => Number(el.getAttribute("cacheId")) || 0)) + 1;
  const keepPaths = new Set<string>();
  for (const [i, sheet] of book.sheets.entries()) {
    let path = sheet.origin?.path;
    if (!path) {
      while (used.has(`${dir}worksheets/sheet${n}.xml`) || zip.file(`${dir}worksheets/sheet${n}.xml`)) n++;
      path = `${dir}worksheets/sheet${n}.xml`;
      used.add(path);
    }
    keepPaths.add(path);
    const drawing = await writeDrawing(zip, book, calc, i, path, override);
    const legacy = await writeNotes(zip, sheet, path, override);
    const links = await writeLinks(zip, sheet, path);
    const tables = await writeTables(zip, sheet, path, override, () => tableId++, (t) => (sheet.filter && sameRange(sheet.filter.range, t) ? filterXml(sheet) : null));
    pivotCaches.push(...(await writePivots(zip, book, calc, i, path, override, () => cacheId++)));
    // A PivotTable's cells are values in the file; DuckSheets keeps its formulas.
    const areas = pivotAreas(book, i);
    const out = areas.length
      ? { ...sheet, cells: new Map([...sheet.cells].map(([k, cell]) => [k, cell.f !== undefined && areas.some((a) => inRange(a, rowOf(k), colOf(k))) ? { s: cell.s, v: calc.value(i, rowOf(k), colOf(k)) } : cell])) }
      : sheet;
    zip.file(path, sheetXml(out, i, calc, strings, i === active, drawing, legacy, links, tables));
    override(path, CONTENT.sheet);
    const id = addRel(TYPE.sheet, path);
    sheetEls.push(`<sheet name="${esc(sheet.name)}" sheetId="${i + 1}"${sheet.state ? ` state="${sheet.state}"` : ""} r:id="${id}"/>`);
  }
  // Sheets deleted here go, with their relationship.
  for (const el of relEls()) {
    if (el.getAttribute("Type") !== TYPE.sheet) continue;
    const path = resolvePath(main, el.getAttribute("Target") ?? "");
    if (keepPaths.has(path)) continue;
    zip.remove(path);
    override(path, null);
    relsRoot.removeChild(el);
  }
  const wb = wbDoc.documentElement;
  const sheetsEl = Array.from(wb.children).find((el) => el.localName === "sheets")!;
  const fresh = xml(`<x xmlns="${MAIN}" xmlns:r="${REL}"><sheets>${sheetEls.join("")}</sheets></x>`).documentElement.firstElementChild!;
  wb.replaceChild(wbDoc.importNode(fresh, true), sheetsEl);
  // Named ranges for one sheet point at it by position, which may have changed; those for sheets gone go.
  for (const el of Array.from(wbDoc.getElementsByTagNameNS("*", "definedName"))) {
    const local = el.getAttribute("localSheetId");
    if (local !== null && Number(local) >= book.sheets.length) el.parentNode!.removeChild(el);
  }
  const view = wbDoc.getElementsByTagNameNS("*", "workbookView")[0];
  if (view) {
    view.setAttribute("activeTab", String(active));
    view.removeAttribute("firstSheet");
  }
  // Excel works everything out again on opening, so values from other programs don't linger.
  let calcPr = Array.from(wb.children).find((el) => el.localName === "calcPr");
  if (!calcPr) {
    calcPr = wbDoc.createElementNS(MAIN, "calcPr");
    wb.appendChild(calcPr);
  }
  calcPr.setAttribute("fullCalcOnLoad", "1");
  if (pivotCaches.length) {
    let list = Array.from(wb.children).find((el) => el.localName === "pivotCaches");
    if (!list) {
      list = wbDoc.createElementNS(MAIN, "pivotCaches");
      // In Excel's order: after the calculation settings and views, before the rest.
      const after = [...Array.from(wb.children)].reverse().find((el) => ["calcPr", "oleSize", "customWorkbookViews"].includes(el.localName))!;
      wb.insertBefore(list, after.nextSibling);
    }
    for (const cache of pivotCaches) {
      const el = wbDoc.createElementNS(MAIN, "pivotCache");
      el.setAttribute("cacheId", String(cache.id));
      el.setAttributeNS(REL, "r:id", addRel(PIVOT_TYPE.cache, cache.path, true));
      list.appendChild(el);
    }
  }
  if (book.date1904) {
    const pr = Array.from(wb.children).find((el) => el.localName === "workbookPr");
    pr?.setAttribute("date1904", "1");
  }

  // Notes' boxes are VML drawings, which need their type listed.
  if (book.sheets.some((s) => s.notes.size) && !Array.from(typesDoc.documentElement.children).some((el) => el.localName === "Default" && el.getAttribute("Extension") === "vml")) {
    const el = typesDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Default");
    el.setAttribute("Extension", "vml");
    el.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing");
    typesDoc.documentElement.insertBefore(el, typesDoc.documentElement.firstChild);
  }

  // Dynamic arrays need Excel's metadata saying cm="1" means "a dynamic array"; a file's own is kept.
  const spilling = book.sheets.some((s, i) => [...s.cells].some(([k, cell]) => cell.f !== undefined && (cell.dyn || (ARRAY_FUNCTIONS.test(cell.f) && !!calc.spillRange(i, rowOf(k), colOf(k))))));
  if (spilling && !relEls().some((el) => el.getAttribute("Type") === TYPE.metadata)) {
    const path = `${dir}metadata.xml`;
    zip.file(path, DYNAMIC_ARRAY_METADATA);
    override(path, CONTENT.metadata);
    addRel(TYPE.metadata, path);
  }

  const stylesPath = `${dir}styles.xml`;
  const existingStyles = relEls().find((el) => el.getAttribute("Type") === TYPE.styles);
  const stylesTarget = existingStyles ? resolvePath(main, existingStyles.getAttribute("Target") ?? "") : stylesPath;
  zip.file(stylesTarget, writeStyles(book.styles));
  override(stylesTarget, CONTENT.styles);
  addRel(TYPE.styles, stylesTarget);

  const stringsRel = relEls().find((el) => el.getAttribute("Type") === TYPE.strings);
  const stringsPath = stringsRel ? resolvePath(main, stringsRel.getAttribute("Target") ?? "") : `${dir}sharedStrings.xml`;
  zip.file(stringsPath, strings.xml());
  override(stringsPath, CONTENT.strings);
  addRel(TYPE.strings, stringsPath);

  zip.file(main, serialize(wbDoc));
  zip.file(relsFile, serialize(relsDoc));
  zip.file("[Content_Types].xml", serialize(typesDoc));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** A value as it's shown, for CSV and copying. */
export const shown = (v: Scalar, numFmt: string, date1904: boolean) => formatValue(v, numFmt, date1904).text;
