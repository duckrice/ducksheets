// Summaries made with Summarize, saved as Excel PivotTables. In DuckSheets a summary is formulas
// (SUMIFS and friends), so it follows the data; in the file it is also a PivotTable over the same
// data, which Excel builds again when it opens the file (refreshOnLoad, no saved records). What
// it summarizes is read back from the formulas, so renamed sheets and moved rows are followed.

import type JSZip from "jszip";
import type { Calculator } from "../calc";
import { colName, key, type Scalar, type Workbook } from "../model";
import { relsPath } from "./read";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const PIVOT_TYPE = { table: `${REL}/pivotTable`, cache: `${REL}/pivotCacheDefinition` };
export const PIVOT_CONTENT = {
  table: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
  cache: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
};
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const FNS: Record<string, { subtotal: string | null; label: string }> = {
  SUMIFS: { subtotal: null, label: "Sum" },
  COUNTIFS: { subtotal: "count", label: "Count" },
  AVERAGEIFS: { subtotal: "average", label: "Average" },
  MAXIFS: { subtotal: "max", label: "Max" },
  MINIFS: { subtotal: "min", label: "Min" },
};

/** A summary as found on a sheet: where it is, what it reads, and how. */
export interface Summary {
  /** Its block on the sheet: heading row, one row per group, and the total. */
  r: number;
  c: number;
  end: number;
  source: string;
  /** The data with its heading row, on the source sheet. */
  r1: number;
  r2: number;
  c1: number;
  c2: number;
  rowField: number;
  valueField: number;
  fn: keyof typeof FNS;
}

const REF = /('(?:[^']|'')+'|[^'!,()]+)!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/g;
const colIndex = (letters: string) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** The summary whose heading is at (r, c) of sheet `index`, or null if its formulas no longer make one. */
export function findSummary(book: Workbook, index: number, r: number, c: number): Summary | null {
  const cells = book.sheets[index].cells;
  const first = cells.get(key(r + 1, c + 1))?.f;
  const m = first && /^(\w+)\((.*)\)$/s.exec(first);
  if (!m || !FNS[m[1].toUpperCase()]) return null;
  const fn = m[1].toUpperCase();
  const refs = [...m[2].matchAll(REF)].map((x) => ({ sheet: x[1].replace(/^'|'$/g, "").replace(/''/g, "'"), c1: colIndex(x[2]), r1: Number(x[3]) - 1, c2: colIndex(x[4]), r2: Number(x[5]) - 1 }));
  const by = fn === "COUNTIFS" ? refs[0] : refs[1];
  const of = refs[0];
  if (!by || by.c1 !== by.c2 || of.c1 !== of.c2 || by.sheet !== of.sheet || by.r1 !== of.r1 || by.r2 !== of.r2 || by.r1 < 1) return null;
  // Every group's row the same formula over the same data, down to the total.
  let end = r + 1;
  while (cells.get(key(end, c + 1))?.f?.toUpperCase().startsWith(`${fn}(`) && cells.get(key(end, c + 1))!.f!.includes(m[2].slice(0, m[2].indexOf(","))) && cells.get(key(end, c))?.v !== "Total") end++;
  if (cells.get(key(end, c))?.v !== "Total" || end === r + 1) return null;
  const sheet = book.sheets.findIndex((s) => s.name.toLowerCase() === by.sheet.toLowerCase());
  if (sheet < 0) return null;
  const c1 = Math.min(by.c1, of.c1);
  return { r, c, end, source: book.sheets[sheet].name, r1: by.r1 - 1, r2: by.r2, c1, c2: Math.max(by.c1, of.c1), rowField: by.c1 - c1, valueField: of.c1 - c1, fn: fn as keyof typeof FNS };
}

/** The cache definition and table parts for one summary, Excel's way. */
export function pivotXml(book: Workbook, calc: Calculator, index: number, s: Summary, cacheId: number, name: string): { cache: string; table: string } {
  const src = book.sheets.findIndex((sh) => sh.name === s.source);
  const value = (r: number, c: number): Scalar => {
    const v = calc.value(src, r, c);
    return typeof v === "object" ? null : v;
  };
  const shown = (r: number, c: number): Scalar => {
    const v = calc.value(index, r, c);
    return typeof v === "object" ? null : v;
  };
  // Field names: the headings, each once, as Excel needs.
  const taken = new Set<string>();
  const names = Array.from({ length: s.c2 - s.c1 + 1 }, (_, i) => {
    const base = String(value(s.r1, s.c1 + i) ?? "").trim() || `Column${i + 1}`;
    let n = base;
    for (let k = 2; taken.has(n.toLowerCase()); k++) n = `${base}${k}`;
    taken.add(n.toLowerCase());
    return n;
  });
  const groups = Array.from({ length: s.end - s.r - 1 }, (_, i) => shown(s.r + 1 + i, s.c));
  const item = (v: Scalar) => (typeof v === "number" ? `<n v="${v}"/>` : typeof v === "boolean" ? `<b v="${v ? 1 : 0}"/>` : `<s v="${esc(String(v ?? ""))}"/>`);
  const numbers = (field: number) => {
    const list: number[] = [];
    for (let r = s.r1 + 1; r <= s.r2; r++) {
      const v = value(r, s.c1 + field);
      if (typeof v === "number") list.push(v);
    }
    return list;
  };
  const fields = names.map((n, i) => {
    if (i === s.rowField) return `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems count="${groups.length}">${groups.map(item).join("")}</sharedItems></cacheField>`;
    const nums = numbers(i);
    const all = s.r2 - s.r1;
    // shortcut: a column of only numbers is described as such; anything else as general, which Excel accepts as it refreshes.
    return nums.length === all && nums.length
      ? `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"/></cacheField>`
      : `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems/></cacheField>`;
  });
  const ref = `${colName(s.c1)}${s.r1 + 1}:${colName(s.c2)}${s.r2 + 1}`;
  const cache =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${REL}" saveData="0" refreshOnLoad="1" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="0">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${ref}" sheet="${esc(s.source)}"/></cacheSource><cacheFields count="${names.length}">${fields.join("")}</cacheFields></pivotCacheDefinition>`;
  const pivotFields = names.map((_, i) =>
    i === s.rowField
      ? `<pivotField axis="axisRow" showAll="0"${i === s.valueField ? ' dataField="1"' : ""}><items count="${groups.length + 1}">${groups.map((_, k) => `<item x="${k}"/>`).join("")}<item t="default"/></items></pivotField>`
      : `<pivotField${i === s.valueField ? ' dataField="1"' : ""} showAll="0"/>`,
  );
  const { subtotal } = FNS[s.fn];
  const caption = String(shown(s.r, s.c + 1) ?? `${FNS[s.fn].label} of ${names[s.valueField]}`);
  const location = `${colName(s.c)}${s.r + 1}:${colName(s.c + 1)}${s.end + 1}`;
  const table =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${MAIN}" name="${esc(name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" grandTotalCaption="Total" rowHeaderCaption="${esc(String(shown(s.r, s.c) ?? names[s.rowField]))}" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${location}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>` +
    `<pivotFields count="${names.length}">${pivotFields.join("")}</pivotFields>` +
    `<rowFields count="1"><field x="${s.rowField}"/></rowFields>` +
    `<rowItems count="${groups.length + 1}">${groups.map((_, k) => (k ? `<i><x v="${k}"/></i>` : "<i><x/></i>")).join("")}<i t="grand"><x/></i></rowItems>` +
    `<colItems count="1"><i/></colItems>` +
    `<dataFields count="1"><dataField name="${esc(caption)}" fld="${s.valueField}"${subtotal ? ` subtotal="${subtotal}"` : ""} baseField="0" baseItem="0"/></dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
  return { cache, table };
}

/**
 * Writes the sheet's summaries as PivotTables: their parts, and a relationship from the sheet to
 * each. Returns, for the workbook, each cache's part path and id. `cacheIds` hands out ids.
 */
export async function writePivots(
  zip: JSZip,
  book: Workbook,
  calc: Calculator,
  index: number,
  sheetPath: string,
  override: (path: string, type: string | null) => void,
  cacheIds: () => number,
): Promise<{ path: string; id: number }[]> {
  const found = (book.sheets[index].pivots ?? []).map((p) => findSummary(book, index, p.r, p.c)).filter((s): s is Summary => !!s);
  if (!found.length) return [];
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf("/worksheets/") + 1);
  const free = (make: (n: number) => string) => {
    let n = 1;
    while (zip.file(make(n))) n++;
    return make(n);
  };
  const relsFile = relsPath(sheetPath);
  let rels = (await zip.file(relsFile)?.async("string")) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const caches: { path: string; id: number }[] = [];
  found.forEach((s, i) => {
    const id = cacheIds();
    const { cache, table } = pivotXml(book, calc, index, s, id, `PivotTable${id}`);
    const cachePath = free((n) => `${dir}pivotCache/pivotCacheDefinition${n}.xml`);
    zip.file(cachePath, cache);
    override(cachePath, PIVOT_CONTENT.cache);
    const tablePath = free((n) => `${dir}pivotTables/pivotTable${n}.xml`);
    zip.file(tablePath, table);
    override(tablePath, PIVOT_CONTENT.table);
    const cacheName = cachePath.slice(cachePath.lastIndexOf("/") + 1);
    zip.file(relsPath(tablePath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${PIVOT_TYPE.cache}" Target="../pivotCache/${cacheName}"/></Relationships>`);
    const tableName = tablePath.slice(tablePath.lastIndexOf("/") + 1);
    rels = rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="rIdPivot${i + 1}" Type="${PIVOT_TYPE.table}" Target="../pivotTables/${tableName}"/></Relationships>`);
    caches.push({ path: cachePath, id });
  });
  zip.file(relsFile, rels);
  return caches;
}

/** The summaries' blocks on a sheet, whose cells go in the file as values: a PivotTable holds no formulas. */
export function pivotAreas(book: Workbook, index: number): { r1: number; c1: number; r2: number; c2: number }[] {
  return (book.sheets[index].pivots ?? [])
    .map((p) => findSummary(book, index, p.r, p.c))
    .filter((s): s is Summary => !!s)
    .map((s) => ({ r1: s.r, c1: s.c, r2: s.end, c2: s.c + 1 }));
}
