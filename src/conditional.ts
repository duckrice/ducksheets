// Conditional formatting, as in Excel: rules that colour cells by their
// values (greater than, between, text that contains, duplicates, top 10,
// above average, or a formula), colour scales and data bars. Read from and
// written to .xlsx files the way Excel keeps them; rules DuckSheets doesn't
// show (icon sets and the like) are kept unchanged.

import type { Calculator } from "./calc";
import { parse, print, shifted } from "./formula/parse";
import { CellError, colOf, minusRange, parseRange, rangeAddress, rowOf, type Range, type Scalar, type Sheet } from "./model";
import { colorOf } from "./xlsx/styles";

export type CellOp = "greaterThan" | "lessThan" | "between" | "notBetween" | "equal" | "notEqual" | "greaterThanOrEqual" | "lessThanOrEqual";

export type CfRule =
  /** Compares the cell's value with one value (two for between): numbers, "text" or formulas, as written in the file. */
  | { t: "cell"; op: CellOp; f: string[]; dxf: number }
  | { t: "text"; text: string; dxf: number }
  | { t: "duplicate" | "unique"; dxf: number }
  | { t: "top"; n: number; bottom: boolean; percent: boolean; dxf: number }
  | { t: "average"; below: boolean; dxf: number }
  /** True where the formula (written for the block's top-left cell) is true. */
  | { t: "formula"; f: string; dxf: number }
  /** Lowest to highest: two colours, or three with the middle at the 50th percentile. */
  | { t: "scale"; colors: string[] }
  | { t: "bar"; color: string }
  /** A rule shown only by Excel, kept as it was. */
  | { t: "other"; xml: string };

export interface CfBlock {
  ranges: Range[];
  /** Most important first. */
  rules: CfRule[];
}

/** What a dxf (a rule's formatting) sets. */
export interface CfStyle {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

/** What the rules do to one cell. */
export interface CfEffect extends CfStyle {
  /** A data bar: how far across the cell (0 to 1), and its colour. */
  bar?: { frac: number; color: string };
}

// ---- formats (dxfs) --------------------------------------------------------------------------

const hex = (c: string) => `FF${c.replace("#", "").toUpperCase()}`;

export function dxfXml(style: CfStyle): string {
  const font = style.color || style.bold || style.italic ? `<font>${style.bold ? "<b/>" : ""}${style.italic ? "<i/>" : ""}${style.color ? `<color rgb="${hex(style.color)}"/>` : ""}</font>` : "";
  const fill = style.fill ? `<fill><patternFill><bgColor rgb="${hex(style.fill)}"/></patternFill></fill>` : "";
  return `<dxf>${font}${fill}</dxf>`;
}

export function readDxf(xml: string, theme: string[]): CfStyle {
  const el = new DOMParser().parseFromString(xml.replace(/^<dxf\b/, '<dxf xmlns:x="x"'), "application/xml").documentElement;
  const find = (parent: Element | null | undefined, name: string) => (parent ? Array.from(parent.children).find((c) => c.localName === name) : undefined);
  const font = find(el, "font");
  const pattern = find(find(el, "fill"), "patternFill");
  const out: CfStyle = {};
  const fill = colorOf(find(pattern, "bgColor"), theme) ?? colorOf(find(pattern, "fgColor"), theme);
  if (fill) out.fill = fill;
  const color = colorOf(find(font, "color"), theme);
  if (color) out.color = color;
  const on = (name: string) => {
    const e = find(font, name);
    return !!e && e.getAttribute("val") !== "0" && e.getAttribute("val") !== "false";
  };
  if (on("b")) out.bold = true;
  if (on("i")) out.italic = true;
  return out;
}

/** Excel's ready-made highlight formats. */
export const PRESETS: { name: string; style: CfStyle }[] = [
  { name: "Light red fill with dark red text", style: { fill: "#FFC7CE", color: "#9C0006" } },
  { name: "Yellow fill with dark yellow text", style: { fill: "#FFEB9C", color: "#9C5700" } },
  { name: "Green fill with dark green text", style: { fill: "#C6EFCE", color: "#006100" } },
  { name: "Light red fill", style: { fill: "#FFC7CE" } },
  { name: "Red text", style: { color: "#9C0006" } },
  { name: "Bold", style: { bold: true } },
];

// ---- reading and writing ----------------------------------------------------------------------

const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] ?? null;

function ranges(sqref: string): Range[] {
  return sqref
    .split(/\s+/)
    .map((part) => parseRange(part.replace(/\$/g, "")))
    .filter((r): r is Range => !!r);
}

export function readConditional(ws: Element, theme: string[]): CfBlock[] {
  const blocks: CfBlock[] = [];
  for (const cf of kids(ws, "conditionalFormatting")) {
    const block: CfBlock = { ranges: ranges(cf.getAttribute("sqref") ?? ""), rules: [] };
    if (!block.ranges.length) continue;
    const rules = kids(cf, "cfRule").sort((a, b) => Number(a.getAttribute("priority") ?? 0) - Number(b.getAttribute("priority") ?? 0));
    for (const rule of rules) block.rules.push(readRule(rule, theme));
    blocks.push(block);
  }
  return blocks;
}

function readRule(rule: Element, theme: string[]): CfRule {
  const type = rule.getAttribute("type");
  const dxf = Number(rule.getAttribute("dxfId") ?? -1);
  const formulas = kids(rule, "formula").map((f) => f.textContent ?? "");
  const other = (): CfRule => ({ t: "other", xml: new XMLSerializer().serializeToString(rule) });
  const plainEnds = (el: Element | null) => {
    const vos = kids(el, "cfvo").map((v) => `${v.getAttribute("type")}:${v.getAttribute("val") ?? ""}`);
    return vos.length >= 2 && vos[0] === "min:" && vos[vos.length - 1] === "max:" && (vos.length === 2 || vos[1] === "percentile:50");
  };
  switch (type) {
    case "cellIs":
      if (dxf < 0 || !formulas.length) return other();
      return { t: "cell", op: (rule.getAttribute("operator") ?? "equal") as CellOp, f: formulas, dxf };
    case "containsText":
      return dxf < 0 ? other() : { t: "text", text: rule.getAttribute("text") ?? "", dxf };
    case "duplicateValues":
    case "uniqueValues":
      return dxf < 0 ? other() : { t: type === "duplicateValues" ? "duplicate" : "unique", dxf };
    case "top10":
      return dxf < 0 ? other() : { t: "top", n: Number(rule.getAttribute("rank") ?? 10), bottom: rule.getAttribute("bottom") === "1", percent: rule.getAttribute("percent") === "1", dxf };
    case "aboveAverage":
      if (dxf < 0 || rule.getAttribute("equalAverage") === "1" || rule.getAttribute("stdDev")) return other();
      return { t: "average", below: rule.getAttribute("aboveAverage") === "0", dxf };
    case "expression":
      return dxf < 0 || !formulas.length ? other() : { t: "formula", f: formulas[0], dxf };
    case "colorScale": {
      const scale = kid(rule, "colorScale");
      const colors = kids(scale, "color").map((c) => colorOf(c, theme) ?? "#FFFFFF");
      return plainEnds(scale) && colors.length === kids(scale, "cfvo").length ? { t: "scale", colors } : other();
    }
    case "dataBar": {
      const bar = kid(rule, "dataBar");
      // Excel 2010 bars keep more in an extension the file links by id; those stay as they are.
      return plainEnds(bar) && !kid(rule, "extLst") ? { t: "bar", color: colorOf(kid(bar, "color"), theme) ?? "#638EC6" } : other();
    }
    default:
      return other();
  }
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** The sheet's conditionalFormatting elements, numbering the rules' priorities in order; null if there are none. */
export function conditionalXml(blocks: CfBlock[]): string | null {
  let priority = 0;
  const out = blocks
    .filter((b) => b.ranges.length && b.rules.length)
    .map((block) => {
      const first = block.ranges[0];
      const at = rangeAddress({ r1: first.r1, c1: first.c1, r2: first.r1, c2: first.c1 });
      const rules = block.rules.map((rule) => {
        const p = ++priority;
        const formula = (f: string) => `<formula>${esc(f)}</formula>`;
        switch (rule.t) {
          case "cell":
            return `<cfRule type="cellIs" dxfId="${rule.dxf}" priority="${p}" operator="${rule.op}">${rule.f.map(formula).join("")}</cfRule>`;
          case "text":
            return `<cfRule type="containsText" dxfId="${rule.dxf}" priority="${p}" operator="containsText" text="${esc(rule.text)}">${formula(`NOT(ISERROR(SEARCH("${rule.text.replace(/"/g, '""')}",${at})))`)}</cfRule>`;
          case "duplicate":
          case "unique":
            return `<cfRule type="${rule.t === "duplicate" ? "duplicateValues" : "uniqueValues"}" dxfId="${rule.dxf}" priority="${p}"/>`;
          case "top":
            return `<cfRule type="top10" dxfId="${rule.dxf}" priority="${p}" rank="${rule.n}"${rule.bottom ? ' bottom="1"' : ""}${rule.percent ? ' percent="1"' : ""}/>`;
          case "average":
            return `<cfRule type="aboveAverage" dxfId="${rule.dxf}" priority="${p}"${rule.below ? ' aboveAverage="0"' : ""}/>`;
          case "formula":
            return `<cfRule type="expression" dxfId="${rule.dxf}" priority="${p}">${formula(rule.f)}</cfRule>`;
          case "scale": {
            const vos = rule.colors.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
            return `<cfRule type="colorScale" priority="${p}"><colorScale>${vos}${rule.colors.map((c) => `<color rgb="${hex(c)}"/>`).join("")}</colorScale></cfRule>`;
          }
          case "bar":
            return `<cfRule type="dataBar" priority="${p}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${hex(rule.color)}"/></dataBar></cfRule>`;
          case "other":
            return rule.xml.replace(/\spriority="\d+"/, ` priority="${p}"`).replace(/\sxmlns(:\w+)?="[^"]*"/g, "");
        }
      });
      return `<conditionalFormatting sqref="${block.ranges.map(rangeAddress).join(" ")}">${rules.join("")}</conditionalFormatting>`;
    });
  return out.length ? out.join("") : null;
}

// ---- working out what each cell looks like -------------------------------------------------------

/** Cells of a block that exist (ranges like whole columns are cut to what's used). */
function cellsIn(sheet: Sheet, block: CfBlock): number[] {
  const out: number[] = [];
  for (const k of sheet.cells.keys()) {
    const r = rowOf(k);
    const c = colOf(k);
    if (block.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) out.push(k);
  }
  return out;
}

const isNum = (v: Scalar): v is number => typeof v === "number";
const lower = (v: Scalar) => (typeof v === "string" ? v.toLowerCase() : v);

function compare(a: Scalar, b: Scalar): number {
  if (isNum(a) && isNum(b)) return a - b;
  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b, undefined, { sensitivity: "base" });
  // Numbers come before text, as when Excel compares.
  return isNum(a) ? -1 : isNum(b) ? 1 : 0;
}

function mix(a: string, b: string, t: number): string {
  const p = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16);
  return `#${[0, 1, 2].map((i) => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, "0")).join("")}`;
}

/** What every rule on a sheet does to its cells, keyed like the sheet's cells. */
export function conditionalEffects(sheet: Sheet, index: number, calc: Calculator, dxfs: CfStyle[]): Map<number, CfEffect> {
  const effects = new Map<number, CfEffect>();
  const style = (k: number) => {
    let e = effects.get(k);
    if (!e) effects.set(k, (e = {}));
    return e;
  };
  /** A higher rule's formatting wins: only what isn't set yet is set. */
  const apply = (k: number, s: CfStyle | undefined) => {
    if (!s) return;
    const e = style(k);
    for (const name of ["fill", "color", "bold", "italic"] as const) if (s[name] !== undefined && e[name] === undefined) (e as Record<string, unknown>)[name] = s[name];
  };
  for (const block of sheet.conditional) {
    if (!block.rules.length) continue;
    const keys = cellsIn(sheet, block);
    if (!keys.length) continue;
    const values = new Map(keys.map((k) => [k, calc.value(index, rowOf(k), colOf(k))] as const));
    const numbers = [...values.values()].filter(isNum).sort((a, b) => a - b);
    const top = block.ranges[0];
    for (const rule of block.rules) {
      if (rule.t === "other") continue;
      if (rule.t === "scale" || rule.t === "bar") {
        if (!numbers.length) continue;
        const min = numbers[0];
        const max = numbers[numbers.length - 1];
        const mid = percentile(numbers, 0.5);
        for (const [k, v] of values) {
          if (!isNum(v)) continue;
          if (rule.t === "bar") {
            if (!style(k).bar) style(k).bar = { frac: max === min ? 1 : 0.1 + (0.9 * (v - min)) / (max - min), color: rule.color };
          } else if (style(k).fill === undefined) {
            const [a, b, t] =
              rule.colors.length === 3
                ? v <= mid
                  ? [rule.colors[0], rule.colors[1], mid === min ? 1 : (v - min) / (mid - min)]
                  : [rule.colors[1], rule.colors[2], max === mid ? 1 : (v - mid) / (max - mid)]
                : [rule.colors[0], rule.colors[1], max === min ? 1 : (v - min) / (max - min)];
            style(k).fill = mix(a, b, t);
          }
        }
        continue;
      }
      const look = dxfs[rule.dxf];
      let test: (v: Scalar, k: number) => boolean;
      switch (rule.t) {
        case "cell": {
          const at = (f: string, k: number) => calc.formulaValue(moved(f, rowOf(k) - top.r1, colOf(k) - top.c1), index, rowOf(k), colOf(k));
          test = (v, k) => {
            if (v === null || v instanceof CellError) return false;
            const a = compare(lower(v), lower(at(rule.f[0], k)));
            const b = rule.f[1] !== undefined ? compare(lower(v), lower(at(rule.f[1], k))) : 0;
            switch (rule.op) {
              case "greaterThan":
                return a > 0;
              case "lessThan":
                return a < 0;
              case "greaterThanOrEqual":
                return a >= 0;
              case "lessThanOrEqual":
                return a <= 0;
              case "equal":
                return a === 0;
              case "notEqual":
                return a !== 0;
              case "between":
                return (a >= 0 && b <= 0) || (a <= 0 && b >= 0);
              case "notBetween":
                return !((a >= 0 && b <= 0) || (a <= 0 && b >= 0));
            }
          };
          break;
        }
        case "text":
          test = (v) => v !== null && !(v instanceof CellError) && String(v).toLowerCase().includes(rule.text.toLowerCase());
          break;
        case "duplicate":
        case "unique": {
          const counts = new Map<unknown, number>();
          for (const v of values.values()) if (v !== null && v !== "") counts.set(lower(v), (counts.get(lower(v)) ?? 0) + 1);
          test = (v) => v !== null && v !== "" && (rule.t === "duplicate" ? counts.get(lower(v))! > 1 : counts.get(lower(v)) === 1);
          break;
        }
        case "top": {
          const count = rule.percent ? Math.max(1, Math.floor((numbers.length * rule.n) / 100)) : rule.n;
          const sorted = rule.bottom ? numbers : [...numbers].reverse();
          const edge = sorted[Math.min(count, sorted.length) - 1];
          test = (v) => isNum(v) && edge !== undefined && (rule.bottom ? v <= edge : v >= edge);
          break;
        }
        case "average": {
          const mean = numbers.reduce((s, n) => s + n, 0) / (numbers.length || 1);
          test = (v) => isNum(v) && (rule.below ? v < mean : v > mean);
          break;
        }
        case "formula":
          test = (_v, k) => {
            const result = calc.formulaValue(moved(rule.f, rowOf(k) - top.r1, colOf(k) - top.c1), index, rowOf(k), colOf(k));
            return result === true || (isNum(result) && result !== 0);
          };
          break;
      }
      for (const [k, v] of values) if (test(v, k)) apply(k, look);
    }
  }
  for (const [k, e] of effects) if (!Object.keys(e).length) effects.delete(k);
  return effects;
}

function percentile(sorted: number[], p: number): number {
  const at = (sorted.length - 1) * p;
  const lo = Math.floor(at);
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (at - lo);
}

/** A formula as written for the block's top-left cell, moved to a cell (dr, dc) away. */
function moved(f: string, dr: number, dc: number): string {
  if (!dr && !dc) return f;
  try {
    return print(shifted(parse(f), dr, dc), true);
  } catch {
    return f;
  }
}

/** The blocks with `cut` taken out of their ranges (a range partly inside keeps the rest, in up to four pieces). */
export function withoutRules(blocks: CfBlock[], cut: Range): CfBlock[] {
  return blocks.map((b) => ({ ...b, ranges: b.ranges.flatMap((g) => minusRange(g, cut)) })).filter((b) => b.ranges.length);
}
