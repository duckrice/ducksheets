// Excel tables (Insert → Table, Format as Table) in .xlsx files: each is a
// part of its own (xl/tables/tableN.xml) naming its block of cells, its
// column names and its style. Tables read from a file keep their XML, so
// what DuckSheets doesn't edit (totals, column formulas) stays as it was.

import type JSZip from "jszip";
import { key, parseRange, rangeAddress, type Sheet, type SheetTable } from "../model";
import { readRels, relsPath } from "./read";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const TABLE_TYPE = `${REL}/table`;
const TABLE_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export async function readTables(zip: JSZip, sheetPath: string, ws: Element): Promise<SheetTable[]> {
  const parts = Array.from(ws.children).find((el) => el.localName === "tableParts");
  if (!parts) return [];
  const rels = await readRels(zip, sheetPath);
  const tables: SheetTable[] = [];
  for (const part of Array.from(parts.children)) {
    const rel = rels.get(part.getAttributeNS(REL, "id") ?? part.getAttribute("r:id") ?? "");
    const text = rel && (await zip.file(rel.path)?.async("string"));
    if (!text) continue;
    const table = new DOMParser().parseFromString(text, "application/xml").documentElement;
    const range = parseRange(table.getAttribute("ref") ?? "");
    if (!range) continue;
    const style = Array.from(table.children).find((el) => el.localName === "tableStyleInfo");
    tables.push({ name: table.getAttribute("displayName") ?? table.getAttribute("name") ?? `Table${tables.length + 1}`, range, style: style?.getAttribute("name") ?? "", xml: text });
  }
  return tables;
}

/** The text of each header cell, as Excel needs the column names: none empty, none the same. */
export function columnNames(sheet: Sheet, table: SheetTable): string[] {
  const names: string[] = [];
  const taken = new Set<string>();
  for (let c = table.range.c1; c <= table.range.c2; c++) {
    const v = sheet.cells.get(key(table.range.r1, c))?.v;
    const base = v === undefined || v === null || v === "" ? `Column${c - table.range.c1 + 1}` : String(v);
    let name = base;
    for (let n = 2; taken.has(name.toLowerCase()); n++) name = `${base}${n}`;
    taken.add(name.toLowerCase());
    names.push(name);
  }
  return names;
}

/** The table's XML: its own (read from the file) brought up to date, or a new one. */
function tableXml(sheet: Sheet, table: SheetTable, id: number, filter: string | null): string {
  const ref = rangeAddress(table.range);
  const names = columnNames(sheet, table);
  const columns = `<tableColumns count="${names.length}">${names.map((n, i) => `<tableColumn id="${i + 1}" name="${esc(n)}"/>`).join("")}</tableColumns>`;
  const style = `<tableStyleInfo name="${esc(table.style)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>`;
  if (!table.xml) return `<table xmlns="${MAIN}" id="${id}" name="${esc(table.name)}" displayName="${esc(table.name)}" ref="${ref}" totalsRowShown="0">${filter ?? ""}${columns}${style}</table>`;
  const doc = new DOMParser().parseFromString(table.xml, "application/xml");
  const root = doc.documentElement;
  root.setAttribute("id", String(id));
  root.setAttribute("ref", ref);
  root.setAttribute("name", table.name);
  root.setAttribute("displayName", table.name);
  const child = (name: string) => Array.from(root.children).find((el) => el.localName === name);
  const swap = (name: string, text: string | null, before: string[]) => {
    const old = child(name);
    const fresh = text && doc.importNode(new DOMParser().parseFromString(`<x xmlns="${MAIN}">${text}</x>`, "application/xml").documentElement.firstElementChild!, true);
    if (old && fresh) root.replaceChild(fresh, old);
    else if (old) root.removeChild(old);
    else if (fresh) root.insertBefore(fresh, before.map(child).find(Boolean) ?? null);
  };
  // Without a header row the columns aren't named by cells.
  if (root.getAttribute("headerRowCount") !== "0") {
    const old = child("tableColumns");
    const kids = old ? Array.from(old.children) : [];
    if (old && kids.length === names.length) kids.forEach((el, i) => el.setAttribute("name", names[i]));
    else swap("tableColumns", columns, ["tableStyleInfo", "extLst"]);
  }
  swap("autoFilter", filter, ["sortState", "tableColumns"]);
  return new XMLSerializer().serializeToString(doc);
}

/**
 * Writes the sheet's tables as parts of their own (replacing the ones it had) and returns its
 * <tableParts>, or null; undefined if the sheet never had tables. `ids` hands out table ids, which
 * must differ across the workbook. `filter` is the autoFilter for a table that has the sheet's filter.
 */
export async function writeTables(
  zip: JSZip,
  sheet: Sheet,
  sheetPath: string,
  override: (path: string, type: string | null) => void,
  ids: () => number,
  filterFor: (table: SheetTable) => string | null,
): Promise<string | null | undefined> {
  const relsFile = relsPath(sheetPath);
  const text = await zip.file(relsFile)?.async("string");
  const had = !!text && text.includes(TABLE_TYPE);
  if (!had && !sheet.tables.length) return undefined;
  const doc = new DOMParser().parseFromString(text ?? `<Relationships xmlns="${PKG_REL}"/>`, "application/xml");
  const root = doc.documentElement;
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf("/worksheets/") + 1);
  const oldRels = await readRels(zip, sheetPath);
  for (const el of Array.from(root.children)) {
    if (el.getAttribute("Type") !== TABLE_TYPE) continue;
    const old = oldRels.get(el.getAttribute("Id") ?? "");
    if (old) {
      zip.remove(old.path);
      override(old.path, null);
    }
    root.removeChild(el);
  }
  let next = Math.max(0, ...Array.from(root.children).map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1;
  const refs: string[] = [];
  let n = 1;
  for (const table of sheet.tables) {
    while (zip.file(`${dir}tables/table${n}.xml`)) n++;
    const path = `${dir}tables/table${n}.xml`;
    zip.file(path, tableXml(sheet, table, ids(), filterFor(table)));
    override(path, TABLE_CONTENT);
    const id = `rId${next++}`;
    const rel = doc.createElementNS(PKG_REL, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", TABLE_TYPE);
    rel.setAttribute("Target", `../tables/table${n}.xml`);
    root.appendChild(rel);
    refs.push(`<tablePart r:id="${id}"/>`);
  }
  zip.file(relsFile, new XMLSerializer().serializeToString(doc));
  return refs.length ? `<tableParts count="${refs.length}">${refs.join("")}</tableParts>` : null;
}
