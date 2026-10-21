// Printing (Ctrl+P), and saving as PDF through the printer list: the sheet's
// cells laid out as a table with their fonts, fills, borders and alignment,
// without gridlines, as Excel prints by default; the selection only, when
// more than one cell is selected.

import { ARRAY_FUNCTIONS } from "./calc";
import { selRange, type Doc } from "./doc";
import { colOf, key, mergeAt, rowOf, usedRange, type Range } from "./model";
import { lookOf, type BorderSide } from "./xlsx/styles";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const BORDER: Record<string, string> = {
  thin: "1px solid",
  hair: "1px solid",
  medium: "2px solid",
  thick: "3px solid",
  dashed: "1px dashed",
  mediumDashed: "2px dashed",
  dotted: "1px dotted",
  double: "3px double",
};

const side = (b: BorderSide | undefined) => (b ? `${BORDER[b.style] ?? "1px solid"} ${b.color ?? "#000"}` : "");

/** What gets printed: the selection if it's more than one cell, else everything used (spilled results too). */
function area(doc: Doc): Range | null {
  const sel = selRange(doc.sel);
  if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) return sel;
  const used = usedRange(doc.sheet);
  if (!used) return null;
  for (const [k, cell] of doc.sheet.cells) {
    if (cell.f === undefined || !ARRAY_FUNCTIONS.test(cell.f)) continue;
    const spill = doc.calc.spillRange(doc.active, rowOf(k), colOf(k));
    if (spill) {
      used.r2 = Math.max(used.r2, spill.r2);
      used.c2 = Math.max(used.c2, spill.c2);
    }
  }
  return used;
}

export function printHtml(doc: Doc): string | null {
  const range = area(doc);
  if (!range) return null;
  const sheet = doc.sheet;
  const styles = doc.book.styles;
  const cols = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i).filter((c) => !sheet.hiddenCols.has(c));
  const rows = Array.from({ length: range.r2 - range.r1 + 1 }, (_, i) => range.r1 + i).filter((r) => !sheet.hiddenRows.has(r));
  const width = (c: number) => sheet.colWidths.get(c) ?? sheet.defaultColWidth;
  const out: string[] = [`<colgroup>${cols.map((c) => `<col style="width:${width(c)}px">`).join("")}</colgroup>`];
  for (const r of rows) {
    const tds: string[] = [];
    for (const c of cols) {
      const m = mergeAt(sheet, r, c);
      if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const cell = sheet.cells.get(key(r, c));
      const look = lookOf(styles, cell?.s);
      const value = doc.value(r, c);
      const numeric = typeof value === "number";
      const f = look.font;
      const css = [
        `font:${f.italic ? "italic " : ""}${f.bold ? "bold " : ""}${f.size}pt "${f.name}",sans-serif`,
        f.color ? `color:${f.color}` : "",
        f.underline || f.strike ? `text-decoration:${f.underline ? "underline " : ""}${f.strike ? "line-through" : ""}` : "",
        look.fill ? `background:${look.fill}` : "",
        `text-align:${look.h && look.h !== "general" ? look.h : numeric ? "right" : "left"}`,
        `vertical-align:${look.v === "center" ? "middle" : look.v ?? "bottom"}`,
        look.wrap ? "white-space:pre-wrap" : "white-space:pre",
        look.indent ? `padding-left:${3 + look.indent * 9}px` : "",
        ...(["left", "right", "top", "bottom"] as const).map((s) => (look.border[s] ? `border-${s}:${side(look.border[s])}` : "")),
      ].filter(Boolean);
      const span = m ? `${m.r2 > m.r1 ? ` rowspan="${rows.filter((x) => x >= m.r1 && x <= m.r2).length}"` : ""}${m.c2 > m.c1 ? ` colspan="${cols.filter((x) => x >= m.c1 && x <= m.c2).length}"` : ""}` : "";
      tds.push(`<td${span} style='${css.join(";")}'>${esc(doc.shown(r, c))}</td>`);
    }
    const height = sheet.rowHeights.get(r) ?? sheet.defaultRowHeight;
    out.push(`<tr style="height:${height}px">${tds.join("")}</tr>`);
  }
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(doc.name)}</title><style>
@page { margin: 18mm 15mm; }
body { margin: 0; }
table { border-collapse: collapse; table-layout: fixed; }
td { padding: 1px 3px; overflow: hidden; }
tr { break-inside: avoid; }
</style></head><body><table>${out.join("")}</table></body></html>`;
}

/** Opens Windows' print dialog for the sheet; "Microsoft Print to PDF" there saves a PDF. */
export function printSheet(doc: Doc): boolean {
  const html = printHtml(doc);
  if (!html) return false;
  const frame = document.createElement("iframe");
  frame.className = "print-frame";
  frame.setAttribute("aria-hidden", "true");
  document.body.append(frame);
  const page = frame.contentDocument!;
  page.open();
  page.write(html);
  page.close();
  setTimeout(() => {
    frame.contentWindow!.focus();
    frame.contentWindow!.print();
    setTimeout(() => frame.remove(), 1000);
  }, 100);
  return true;
}
