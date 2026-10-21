// Data validation, as in Excel: a cell can only take a value from a list
// (shown as a drop-down beside the cell), or a whole or decimal number within
// limits. Read from and written to .xlsx files; kinds DuckSheets doesn't
// check (dates, text length, custom formulas) are kept unchanged.

import type { Calculator } from "./calc";
import type { CellOp } from "./conditional";
import { CellError, minusRange, parseRange, rangeAddress, type Range, type Scalar, type Sheet } from "./model";

export type Validation =
  /** `f`: the list, as written in the file: "a,b,c" (quoted) or a range like $D$1:$D$5 or Sheet2!$A$1:$A$9. */
  | { ranges: Range[]; t: "list"; f: string }
  | { ranges: Range[]; t: "whole" | "decimal"; op: CellOp; f: string[] }
  | { ranges: Range[]; t: "other"; xml: string };

const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function ranges(sqref: string): Range[] {
  return sqref
    .split(/\s+/)
    .map((part) => parseRange(part.replace(/\$/g, "")))
    .filter((r): r is Range => !!r);
}

export function readValidations(ws: Element): Validation[] {
  const out: Validation[] = [];
  for (const dv of kids(kids(ws, "dataValidations")[0], "dataValidation")) {
    const where = ranges(dv.getAttribute("sqref") ?? "");
    if (!where.length) continue;
    const type = dv.getAttribute("type");
    const f = kids(dv, "formula1").concat(kids(dv, "formula2")).map((e) => e.textContent ?? "");
    // Only the plain kinds are checked here; anything with messages or other settings of its own is kept as it was.
    const plain = !dv.getAttribute("promptTitle") && !dv.getAttribute("prompt") && !dv.getAttribute("errorTitle") && !dv.getAttribute("error") && dv.getAttribute("errorStyle") === null;
    if (type === "list" && f[0] && plain) out.push({ ranges: where, t: "list", f: f[0] });
    else if ((type === "whole" || type === "decimal") && f[0] && plain) out.push({ ranges: where, t: type, op: (dv.getAttribute("operator") ?? "between") as CellOp, f });
    else out.push({ ranges: where, t: "other", xml: new XMLSerializer().serializeToString(dv).replace(/\sxmlns(:\w+)?="[^"]*"/g, "") });
  }
  return out;
}

export function validationsXml(list: Validation[]): string | null {
  const items = list
    .filter((v) => v.ranges.length)
    .map((v) => {
      const sqref = v.ranges.map(rangeAddress).join(" ");
      if (v.t === "other") return v.xml.replace(/\ssqref="[^"]*"/, ` sqref="${sqref}"`);
      const f = v.t === "list" ? [v.f] : v.f;
      const op = v.t === "list" || v.op === "between" ? "" : ` operator="${v.op}"`;
      return `<dataValidation type="${v.t}"${op} allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${sqref}">${f.map((x, i) => `<formula${i + 1}>${esc(x)}</formula${i + 1}>`).join("")}</dataValidation>`;
    });
  return items.length ? `<dataValidations count="${items.length}">${items.join("")}</dataValidations>` : null;
}

/** The rule for a cell (the last one given wins, as rules for the same cells replace each other). */
export function validationAt(sheet: Sheet, r: number, c: number): Validation | undefined {
  for (let i = sheet.validations.length - 1; i >= 0; i--) {
    const v = sheet.validations[i];
    if (v.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) return v;
  }
  return undefined;
}

/** A list's choices: the quoted items, or the values in its range. */
export function listItems(v: Extract<Validation, { t: "list" }>, sheet: number, calc: Calculator): string[] {
  const f = v.f.trim();
  if (f.startsWith('"')) return f.slice(1, f.endsWith('"') ? -1 : undefined).split(",").map((s) => s.trim()).filter(Boolean);
  const m = /^(?:'?([^'!]+)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(f);
  const named = calc.book.names.get(f.toUpperCase());
  if (!m && named) return listItems({ ...v, f: named }, sheet, calc);
  if (!m) return [];
  const s = m[1] ? calc.book.sheets.findIndex((x) => x.name.toLowerCase() === m[1].toLowerCase()) : sheet;
  const range = parseRange(m[2].replace(/\$/g, ""));
  if (s < 0 || !range) return [];
  const out: string[] = [];
  for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 1000); r++)
    for (let c = range.c1; c <= range.c2; c++) {
      const value = calc.value(s, r, c);
      if (value !== null && value !== "" && !(value instanceof CellError)) out.push(String(value));
    }
  return out;
}

/** Whether a value is allowed in a cell with this rule; an empty cell always is. */
export function allowed(v: Validation, value: Scalar, sheet: number, calc: Calculator): boolean {
  if (value === null || value === "") return true;
  if (v.t === "other") return true;
  if (v.t === "list") return listItems(v, sheet, calc).some((item) => item.toLowerCase() === String(value).toLowerCase());
  if (typeof value !== "number") return false;
  if (v.t === "whole" && !Number.isInteger(value)) return false;
  const bound = (f: string | undefined) => {
    const n = f === undefined ? NaN : Number(calc.formulaValue(f, sheet, 0, 0));
    return Number.isNaN(n) ? 0 : n;
  };
  const a = bound(v.f[0]);
  const b = bound(v.f[1]);
  switch (v.op) {
    case "between":
      return value >= Math.min(a, b) && value <= Math.max(a, b);
    case "notBetween":
      return value < Math.min(a, b) || value > Math.max(a, b);
    case "equal":
      return value === a;
    case "notEqual":
      return value !== a;
    case "greaterThan":
      return value > a;
    case "lessThan":
      return value < a;
    case "greaterThanOrEqual":
      return value >= a;
    case "lessThanOrEqual":
      return value <= a;
  }
}

/** What a rule asks for, in words, for the message when a value doesn't fit. */
export function describe(v: Validation): string {
  if (v.t === "list") return "one of the values in the list";
  if (v.t === "other") return "";
  const kind = v.t === "whole" ? "a whole number" : "a number";
  const [a, b] = v.f;
  const words: Record<CellOp, string> = {
    between: `between ${a} and ${b}`,
    notBetween: `not between ${a} and ${b}`,
    equal: `equal to ${a}`,
    notEqual: `not equal to ${a}`,
    greaterThan: `greater than ${a}`,
    lessThan: `less than ${a}`,
    greaterThanOrEqual: `${a} or more`,
    lessThanOrEqual: `${a} or less`,
  };
  return `${kind} ${words[v.op]}`;
}

/** The rules without `cut` in their ranges. */
export function withoutValidation(list: Validation[], cut: Range): Validation[] {
  return list.map((v) => ({ ...v, ranges: v.ranges.flatMap((g) => minusRange(g, cut)) })).filter((v) => v.ranges.length);
}
