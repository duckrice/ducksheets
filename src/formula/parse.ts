// Formulas: reading them into a tree, writing the tree back as text, and
// moving the references in it when cells are copied or rows and columns are
// inserted or deleted. Follows Excel: unary minus binds tighter than ^
// (=-2^2 is 4), ^ goes left to right, % applies to what's just before it.
// Arguments may be separated by , or ; (Norwegian Excel uses ;).

import { MAX_COLS, MAX_ROWS, colIndex, colName } from "../model";

export interface CellRef {
  r: number;
  c: number;
  /** $ before the row / the column. */
  ar: boolean;
  ac: boolean;
}

export type Node =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "bool"; v: boolean }
  | { t: "err"; v: string }
  | ({ t: "ref"; sheet?: string } & CellRef)
  /** whole: "cols" for A:C, "rows" for 1:3. */
  | { t: "range"; sheet?: string; a: CellRef; b: CellRef; whole?: "cols" | "rows" }
  | { t: "name"; name: string }
  | { t: "fn"; name: string; args: Node[] }
  | { t: "un"; op: "-" | "+"; e: Node }
  | { t: "pct"; e: Node }
  | { t: "bin"; op: string; l: Node; r: Node }
  | { t: "paren"; e: Node };

export class FormulaError extends Error {}

// ---- reading -------------------------------------------------------------------------

type Token =
  | { k: "num"; v: number }
  | { k: "str"; v: string }
  | { k: "err"; v: string }
  | { k: "word"; v: string; sheet?: string }
  | { k: "op"; v: string }
  | { k: "("; } | { k: ")"; } | { k: "sep" } | { k: "end" };

const ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA"];

function tokens(text: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    const rest = text.slice(i);
    let m: RegExpExecArray | null;
    if ((m = /^(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(rest)) && !/^\d+:/.test(rest)) {
      out.push({ k: "num", v: Number(m[0]) });
      i += m[0].length;
    } else if (ch === '"') {
      let s = "";
      i++;
      for (;;) {
        if (i >= text.length) throw new FormulaError("A text in quotes isn't closed");
        if (text[i] === '"') {
          if (text[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += text[i++];
      }
      out.push({ k: "str", v: s });
    } else if (ch === "#") {
      const e = ERRORS.find((x) => rest.toUpperCase().startsWith(x));
      if (!e) throw new FormulaError(`Unknown error value ${rest.slice(0, 8)}`);
      out.push({ k: "err", v: e });
      i += e.length;
    } else if (ch === "'") {
      // 'Sheet name'!A1
      let name = "";
      i++;
      for (;;) {
        if (i >= text.length) throw new FormulaError("A sheet name in quotes isn't closed");
        if (text[i] === "'") {
          if (text[i + 1] === "'") {
            name += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        name += text[i++];
      }
      if (text[i] !== "!") throw new FormulaError("Expected ! after the sheet name");
      i++;
      const ref = /^[$A-Za-z0-9_.:]+/.exec(text.slice(i));
      if (!ref) throw new FormulaError("Expected a cell after the sheet name");
      out.push({ k: "word", v: ref[0], sheet: name });
      i += ref[0].length;
    } else if ((m = /^\$?\d+:\$?\d+/.exec(rest))) {
      out.push({ k: "word", v: m[0] });
      i += m[0].length;
    } else if ((m = /^[A-Za-z_\\$][A-Za-z0-9_.\\$]*(:\$?[A-Za-z]+\$?\d*)?/.exec(rest))) {
      let word = m[0];
      i += word.length;
      if (text[i] === "!") {
        // Sheet1!A1
        i++;
        const ref = /^[$A-Za-z0-9_.:]+/.exec(text.slice(i));
        if (!ref) throw new FormulaError("Expected a cell after the sheet name");
        out.push({ k: "word", v: ref[0], sheet: word });
        i += ref[0].length;
        continue;
      }
      // A function name ending in a range-like bit (LOG10) is still a function.
      if (word.includes(":") && text[i] === "(") throw new FormulaError(`Unexpected ${word}`);
      out.push({ k: "word", v: word });
    } else if ((m = /^(<=|>=|<>|[-+*/^&=<>%])/.exec(rest))) {
      out.push({ k: "op", v: m[0] });
      i += m[0].length;
    } else if (ch === "(" || ch === ")") {
      out.push({ k: ch });
      i++;
    } else if (ch === "," || ch === ";") {
      out.push({ k: "sep" });
      i++;
    } else if (ch === "{") {
      throw new FormulaError("Arrays in braces aren't supported yet");
    } else throw new FormulaError(`Unexpected "${ch}"`);
  }
  out.push({ k: "end" });
  return out;
}

const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/;
const ROWS = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})$/;

function cellRef(text: string): CellRef | null {
  const m = CELL.exec(text);
  if (!m) return null;
  const c = colIndex(m[2]);
  const r = Number(m[4]) - 1;
  if (c < 0 || r < 0 || r >= MAX_ROWS) return null;
  return { r, c, ar: m[3] === "$", ac: m[1] === "$" };
}

/** A word as a reference (A1, A1:B2, A:B, 1:3), or null if it isn't one. */
function reference(word: string, sheet?: string): Node | null {
  const single = cellRef(word);
  if (single) return { t: "ref", ...(sheet !== undefined ? { sheet } : {}), ...single };
  const [x, y] = word.split(":");
  if (y !== undefined) {
    const a = cellRef(x);
    const b = cellRef(y);
    if (a && b) return { t: "range", ...(sheet !== undefined ? { sheet } : {}), a, b };
  }
  let m = COLS.exec(word);
  if (m) {
    const c1 = colIndex(m[2]);
    const c2 = colIndex(m[4]);
    if (c1 >= 0 && c2 >= 0) return { t: "range", ...(sheet !== undefined ? { sheet } : {}), whole: "cols", a: { r: 0, c: c1, ar: true, ac: m[1] === "$" }, b: { r: MAX_ROWS - 1, c: c2, ar: true, ac: m[3] === "$" } };
  }
  m = ROWS.exec(word);
  if (m) {
    const r1 = Number(m[2]) - 1;
    const r2 = Number(m[4]) - 1;
    if (r1 >= 0 && r2 >= 0) return { t: "range", ...(sheet !== undefined ? { sheet } : {}), whole: "rows", a: { r: r1, c: 0, ar: m[1] === "$", ac: true }, b: { r: r2, c: MAX_COLS - 1, ar: m[3] === "$", ac: true } };
  }
  return null;
}

const BINARY: Record<string, number> = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };
const UNARY = 6;

/** Reads a formula (without its "="). Throws FormulaError when it doesn't make sense. */
export function parse(text: string): Node {
  const list = tokens(text);
  let pos = 0;
  const peek = () => list[pos];
  const next = () => list[pos++];

  function primary(): Node {
    const tok = next();
    switch (tok.k) {
      case "num":
        return { t: "num", v: tok.v };
      case "str":
        return { t: "str", v: tok.v };
      case "err":
        return { t: "err", v: tok.v };
      case "(": {
        const e = expr(0);
        if (next().k !== ")") throw new FormulaError("A bracket isn't closed");
        return { t: "paren", e };
      }
      case "op":
        if (tok.v === "-" || tok.v === "+") return { t: "un", op: tok.v, e: expr(UNARY) };
        throw new FormulaError(`Unexpected ${tok.v}`);
      case "word": {
        if (tok.sheet === undefined && peek().k === "(") {
          next();
          // Excel writes newer functions as _xlfn.NAME, and FILTER and SORT as _xlfn._xlws.NAME.
          const name = tok.v.toUpperCase().replace(/^(?:_XLFN\.)?(?:_XLWS\.)?/, "");
          const args: Node[] = [];
          if (peek().k !== ")") {
            for (;;) {
              // An empty argument, as in IF(A1,,1), counts as missing.
              args.push(peek().k === "sep" || peek().k === ")" ? { t: "name", name: "" } : expr(0));
              const sep = next();
              if (sep.k === ")") break;
              if (sep.k !== "sep") throw new FormulaError(`Expected , or ) in ${name}`);
            }
          } else next();
          return { t: "fn", name, args };
        }
        const ref = reference(tok.v, tok.sheet);
        if (ref) return ref;
        if (tok.sheet !== undefined) {
          if (tok.v.toUpperCase() === "#REF!") return { t: "err", v: "#REF!" };
          throw new FormulaError(`${tok.v} isn't a cell`);
        }
        const upper = tok.v.toUpperCase();
        if (upper === "TRUE" || upper === "FALSE") return { t: "bool", v: upper === "TRUE" };
        return { t: "name", name: tok.v };
      }
      default:
        throw new FormulaError(tok.k === "end" ? "The formula stops too early" : `Unexpected ${tok.k === "sep" ? "separator" : tok.k}`);
    }
  }

  function expr(min: number): Node {
    let left = primary();
    for (;;) {
      const tok = peek();
      if (tok.k === "op" && tok.v === "%") {
        next();
        left = { t: "pct", e: left };
        continue;
      }
      const prec = tok.k === "op" ? BINARY[tok.v] : undefined;
      // Stopping at the same level makes ^ and the rest go left to right.
      if (prec === undefined || prec <= min || tok.k !== "op") return left;
      next();
      left = { t: "bin", op: tok.v, l: left, r: expr(prec) };
    }
  }

  const tree = expr(0);
  if (peek().k !== "end") throw new FormulaError("There is something extra at the end");
  return tree;
}

// ---- writing ---------------------------------------------------------------------------

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !CELL.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

const refText = (r: CellRef) => `${r.ac ? "$" : ""}${colName(r.c)}${r.ar ? "$" : ""}${r.r + 1}`;

/** Functions newer than Excel 2007, written with Excel's _xlfn. prefix in files. */
export const NEW_FUNCTIONS = new Set([
  "XLOOKUP", "XMATCH", "IFS", "SWITCH", "CONCAT", "TEXTJOIN", "MAXIFS", "MINIFS", "IFNA", "XOR", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P",
  "DAYS", "ISOWEEKNUM", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "LET", "CEILING.MATH", "FLOOR.MATH", "MODE.SNGL", "RANK.EQ", "PERCENTILE.INC", "QUARTILE.INC",
]);

/** The tree as formula text (without "="). `forFile` adds the prefixes Excel's files use. */
export function print(node: Node, forFile = false): string {
  const p = (n: Node) => print(n, forFile);
  switch (node.t) {
    case "num":
      return String(node.v).replace("e+", "E+").replace("e-", "E-");
    case "str":
      return `"${node.v.replace(/"/g, '""')}"`;
    case "bool":
      return node.v ? "TRUE" : "FALSE";
    case "err":
      return node.v;
    case "ref":
      return `${node.sheet !== undefined ? `${quoteSheet(node.sheet)}!` : ""}${refText(node)}`;
    case "range": {
      const prefix = node.sheet !== undefined ? `${quoteSheet(node.sheet)}!` : "";
      if (node.whole === "cols") return `${prefix}${node.a.ac ? "$" : ""}${colName(node.a.c)}:${node.b.ac ? "$" : ""}${colName(node.b.c)}`;
      if (node.whole === "rows") return `${prefix}${node.a.ar ? "$" : ""}${node.a.r + 1}:${node.b.ar ? "$" : ""}${node.b.r + 1}`;
      return `${prefix}${refText(node.a)}:${refText(node.b)}`;
    }
    case "name":
      return node.name;
    case "fn":
      return `${forFile && NEW_FUNCTIONS.has(node.name) ? (node.name === "FILTER" || node.name === "SORT" ? "_xlfn._xlws." : "_xlfn.") : ""}${node.name}(${node.args.map(p).join(",")})`;
    case "un":
      return `${node.op}${p(node.e)}`;
    case "pct":
      return `${p(node.e)}%`;
    case "bin":
      return `${p(node.l)}${node.op}${p(node.r)}`;
    case "paren":
      return `(${p(node.e)})`;
  }
}

// ---- moving references -------------------------------------------------------------------

/** Rebuilds the tree, replacing each reference with what `fn` returns. */
export function mapRefs(node: Node, fn: (ref: Extract<Node, { t: "ref" | "range" }>) => Node): Node {
  switch (node.t) {
    case "ref":
    case "range":
      return fn(node);
    case "fn":
      return { ...node, args: node.args.map((a) => mapRefs(a, fn)) };
    case "un":
    case "pct":
    case "paren":
      return { ...node, e: mapRefs(node.e, fn) };
    case "bin":
      return { ...node, l: mapRefs(node.l, fn), r: mapRefs(node.r, fn) };
    default:
      return node;
  }
}

const REF_ERROR: Node = { t: "err", v: "#REF!" };

/** The formula as copied by (dr, dc): relative references move with it, $ ones stay. */
export function shifted(node: Node, dr: number, dc: number): Node {
  const move = (ref: CellRef): CellRef | null => {
    const r = ref.ar ? ref.r : ref.r + dr;
    const c = ref.ac ? ref.c : ref.c + dc;
    return r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS ? null : { ...ref, r, c };
  };
  return mapRefs(node, (ref) => {
    if (ref.t === "ref") {
      const moved = move(ref);
      return moved ? { ...ref, ...moved } : REF_ERROR;
    }
    const a = ref.whole === "rows" ? { ...ref.a } : move(ref.a);
    const b = ref.whole === "rows" ? { ...ref.b } : move(ref.b);
    if (ref.whole === "cols") return a && b ? { ...ref, a: { ...a, r: 0 }, b: { ...b, r: MAX_ROWS - 1 } } : REF_ERROR;
    if (ref.whole === "rows") {
      const ra = move({ ...ref.a, c: 0, ac: true });
      const rb = move({ ...ref.b, c: 0, ac: true });
      return ra && rb ? { ...ref, a: { ...ref.a, r: ra.r }, b: { ...ref.b, r: rb.r } } : REF_ERROR;
    }
    return a && b ? { ...ref, a, b } : REF_ERROR;
  });
}

/**
 * The formula after rows (axis "r") or columns ("c") were inserted (count > 0)
 * or deleted (count < 0) at `at` on sheet `target`. `home` is the sheet the
 * formula is on, for references without a sheet name.
 */
export function afterInsertDelete(node: Node, home: string, target: string, axis: "r" | "c", at: number, count: number): Node {
  const onTarget = (sheet?: string) => (sheet ?? home).toLowerCase() === target.toLowerCase();
  // Where an index goes; null when it was deleted.
  const place = (i: number): number | null => {
    if (i < at) return i;
    if (count > 0) return i + count;
    return i < at - count ? null : i + count;
  };
  return mapRefs(node, (ref) => {
    if (!onTarget(ref.sheet)) return ref;
    if (ref.t === "ref") {
      const i = place(ref[axis]);
      return i === null ? REF_ERROR : { ...ref, [axis]: i };
    }
    // Whole rows aren't moved by inserting columns, and the other way round.
    if ((ref.whole === "cols" && axis === "r") || (ref.whole === "rows" && axis === "c")) return ref;
    let lo = ref.a[axis];
    let hi = ref.b[axis];
    if (count > 0) {
      lo = lo >= at ? lo + count : lo;
      hi = hi >= at ? hi + count : hi;
    } else {
      const end = at - count;
      if (lo >= at && hi < end) return REF_ERROR;
      lo = lo < at ? lo : lo < end ? at : lo + count;
      hi = hi < at ? hi : hi < end ? at - 1 : hi + count;
      if (hi < lo) return REF_ERROR;
    }
    return { ...ref, a: { ...ref.a, [axis]: lo }, b: { ...ref.b, [axis]: hi } };
  });
}

/** The formula with a sheet renamed (references to it carry the new name). */
export function renamedSheet(node: Node, from: string, to: string): Node {
  return mapRefs(node, (ref) => (ref.sheet !== undefined && ref.sheet.toLowerCase() === from.toLowerCase() ? { ...ref, sheet: to } : ref));
}
