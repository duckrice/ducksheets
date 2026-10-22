// Working out a formula's value, with Excel's rules: text that looks like a
// number counts as one in arithmetic, TRUE is 1, text compares without regard
// to case, errors pass through, and a range used where one value is wanted
// gives the cell in the formula's own row or column.

import { CellError, MAX_COLS, MAX_ROWS, type Scalar } from "../model";
import { formatValue, parseInput } from "../numfmt";
import { parse, type Node } from "./parse";

export class RangeValue {
  constructor(
    public sheet: number,
    public r1: number,
    public c1: number,
    public r2: number,
    public c2: number,
  ) {}
  get rows() {
    return this.r2 - this.r1 + 1;
  }
  get cols() {
    return this.c2 - this.c1 + 1;
  }
}

/** Values worked out by a formula rather than read from cells, row by row (FILTER, SORT, A1:A9>5 inside FILTER…). */
export class ArrayValue {
  constructor(public values: Scalar[][]) {}
  get rows() {
    return this.values.length;
  }
  get cols() {
    return this.values[0]?.length ?? 0;
  }
}

export type Value = Scalar | RangeValue | ArrayValue;

export interface EvalContext {
  /** Where the formula is. */
  sheet: number;
  row: number;
  col: number;
  /** A cell's value (worked out if it holds a formula). */
  cell(sheet: number, r: number, c: number): Scalar;
  /** The index of a sheet by name, or -1. */
  sheetIndex(name: string): number;
  /** The last row and column with something in them, to keep whole-column ranges short. */
  extent(sheet: number): { rows: number; cols: number };
  /** A named range's formula text, if there is such a name. */
  name(name: string): string | undefined;
  date1904: boolean;
  /** Inside functions that take arrays (FILTER, SORT, SUMPRODUCT…): A1:A9>5 compares every cell, giving an array. */
  arrays?: boolean;
  /** Names given with LET, by name in capitals. */
  local?: Map<string, Value>;
  /** Where a table reference points, or null if nowhere. */
  table?(name: string, spec: string): { sheet: number; r1: number; c1: number; r2: number; c2: number } | null;
}

export const err = (code: CellError["code"]) => new CellError(code);
export const isErr = (v: unknown): v is CellError => v instanceof CellError;

export function evaluate(node: Node, ctx: EvalContext): Value {
  switch (node.t) {
    case "num":
      return node.v;
    case "str":
      return node.v;
    case "bool":
      return node.v;
    case "err":
      return err(node.v as CellError["code"]);
    case "paren":
      return evaluate(node.e, ctx);
    case "ref": {
      const sheet = node.sheet === undefined ? ctx.sheet : ctx.sheetIndex(node.sheet);
      if (sheet < 0) return err("#REF!");
      return new RangeValue(sheet, node.r, node.c, node.r, node.c);
    }
    case "range": {
      const sheet = node.sheet === undefined ? ctx.sheet : ctx.sheetIndex(node.sheet);
      if (sheet < 0) return err("#REF!");
      let r2 = Math.max(node.a.r, node.b.r);
      let c2 = Math.max(node.a.c, node.b.c);
      // Whole columns and rows stop where the sheet's data does.
      const extent = ctx.extent(sheet);
      if (r2 === MAX_ROWS - 1) r2 = Math.max(Math.min(node.a.r, node.b.r), extent.rows - 1);
      if (c2 === MAX_COLS - 1) c2 = Math.max(Math.min(node.a.c, node.b.c), extent.cols - 1);
      return new RangeValue(sheet, Math.min(node.a.r, node.b.r), Math.min(node.a.c, node.b.c), r2, c2);
    }
    case "table": {
      const at = ctx.table?.(node.table, node.spec);
      return at ? new RangeValue(at.sheet, at.r1, at.c1, at.r2, at.c2) : err("#REF!");
    }
    case "name": {
      if (!node.name) return null; // a missing argument
      const own = ctx.local?.get(node.name.toUpperCase());
      if (own !== undefined) return own;
      const text = ctx.name(node.name.toUpperCase());
      if (text === undefined) return err("#NAME?");
      try {
        return evaluate(parse(text.replace(/^=/, "")), ctx);
      } catch {
        return err("#NAME?");
      }
    }
    case "fn": {
      const fn = FUNCTIONS[node.name];
      if (!fn) return err("#NAME?");
      return fn(node.args, ctx);
    }
    case "un": {
      const e = evaluate(node.e, ctx);
      if (spreads(e, ctx)) return elementwise(node.op === "-" ? "*" : "+", e, node.op === "-" ? -1 : 0, ctx);
      const v = toNumber(scalar(e, ctx));
      return isErr(v) ? v : node.op === "-" ? -v : v;
    }
    case "pct": {
      const v = toNumber(scalar(evaluate(node.e, ctx), ctx));
      return isErr(v) ? v : v / 100;
    }
    case "bin": {
      const l = evaluate(node.l, ctx);
      const r = evaluate(node.r, ctx);
      if (spreads(l, ctx) || spreads(r, ctx)) return elementwise(node.op, l, r, ctx);
      return binary(node.op, scalar(l, ctx), scalar(r, ctx));
    }
  }
}

/** Whether an operator works cell by cell on this value: an array always, a range only where arrays are expected. */
function spreads(v: Value, ctx: EvalContext): boolean {
  return v instanceof ArrayValue ? v.rows * v.cols > 1 : !!ctx.arrays && v instanceof RangeValue && v.rows * v.cols > 1;
}

/** An operator applied cell by cell; a single row or column (or value) is used for every row or column, as in Excel. */
function elementwise(op: string, l: Value, r: Value, ctx: EvalContext): ArrayValue {
  const a = toGrid(l, ctx);
  const b = toGrid(r, ctx);
  const rows = Math.max(a.length, b.length);
  const cols = Math.max(a[0].length, b[0].length);
  const at = (g: Scalar[][], i: number, j: number): Scalar => {
    const row = g.length === 1 ? g[0] : g[i];
    if (!row) return err("#N/A");
    const v = row.length === 1 ? row[0] : row[j];
    return v === undefined ? err("#N/A") : v;
  };
  return new ArrayValue(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => binary(op, at(a, i, j), at(b, i, j)))));
}

/** Any value as rows of values. */
export function toGrid(v: Value, ctx: EvalContext): Scalar[][] {
  if (v instanceof ArrayValue) return v.values;
  if (v instanceof RangeValue) return Array.from({ length: v.rows }, (_, i) => Array.from({ length: v.cols }, (_, j) => ctx.cell(v.sheet, v.r1 + i, v.c1 + j)));
  return [[v]];
}

function binary(op: string, a: Scalar, b: Scalar): Scalar {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (op === "&") return toText(a) + toText(b);
  if (["=", "<>", "<", ">", "<=", ">="].includes(op)) {
    const c = compare(a, b);
    return { "=": c === 0, "<>": c !== 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0 }[op as "="];
  }
  const x = toNumber(a);
  const y = toNumber(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  switch (op) {
    case "+":
      return x + y;
    case "-":
      return x - y;
    case "*":
      return x * y;
    case "/":
      return y === 0 ? err("#DIV/0!") : x / y;
    case "^": {
      const r = x ** y;
      return Number.isFinite(r) ? r : err("#NUM!");
    }
  }
  return err("#VALUE!");
}

// ---- conversions ----------------------------------------------------------------------------

/** One value from what may be a range: the cell in the formula's row or column. */
export function scalar(v: Value, ctx: EvalContext): Scalar {
  if (v instanceof ArrayValue) return v.values[0]?.[0] ?? null;
  if (!(v instanceof RangeValue)) return v;
  if (v.rows === 1 && v.cols === 1) return ctx.cell(v.sheet, v.r1, v.c1);
  if (v.cols === 1 && ctx.row >= v.r1 && ctx.row <= v.r2) return ctx.cell(v.sheet, ctx.row, v.c1);
  if (v.rows === 1 && ctx.col >= v.c1 && ctx.col <= v.c2) return ctx.cell(v.sheet, v.r1, ctx.col);
  return err("#VALUE!");
}

export function toNumber(v: Scalar): number | CellError {
  if (v === null) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isErr(v)) return v;
  const typed = parseInput(v);
  return typeof typed.value === "number" ? typed.value : err("#VALUE!");
}

export function toText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isErr(v)) return v.code;
  return formatValue(v, "General").text;
}

export function toBool(v: Scalar): boolean | CellError {
  if (v === null) return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (isErr(v)) return v;
  const upper = v.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return err("#VALUE!");
}

/** Excel's order: numbers, then text, then FALSE, TRUE; an empty cell is like 0, "" or FALSE. */
export function compare(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3);
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "string" && typeof b === "string") {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === "boolean" && typeof b === "boolean") return Number(a) - Number(b);
  return 0;
}

/** Every cell value in a range, row by row. */
export function* cellsOf(range: RangeValue | ArrayValue, ctx: EvalContext): Generator<Scalar> {
  if (range instanceof ArrayValue) {
    for (const row of range.values) yield* row;
    return;
  }
  for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) yield ctx.cell(range.sheet, r, c);
}

export type Fn = (args: Node[], ctx: EvalContext) => Value;

// Filled in by functions.ts, which needs the helpers above.
export const FUNCTIONS: Record<string, Fn> = {};
