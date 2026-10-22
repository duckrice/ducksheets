// Works out the value of every formula in a workbook. Values are worked out
// when asked for and kept until something changes; then everything is
// worked out again.
// shortcut: recalculates the whole workbook after each change; fine for tens of thousands of formulas, track dependencies if bigger sheets get slow.

import "./formula/functions";
import { ArrayValue, evaluate, scalar, type EvalContext } from "./formula/eval";
import { resolveTable } from "./structured";
import { FormulaError, parse, type Node } from "./formula/parse";
import { CellError, address, colOf, key, rowOf, type Range, type Scalar, type Workbook } from "./model";

/** Formulas whose results can fill more than one cell. */
export const ARRAY_FUNCTIONS = /\b(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|TRANSPOSE)\s*\(/i;

export class Calculator {
  private memo = new Map<number, Map<number, Scalar>>();
  private busy = new Set<string>();
  private trees = new Map<string, Node | FormulaError>();
  private extents = new Map<number, { rows: number; cols: number }>();
  /** Values spilled by array formulas (FILTER, SORT…) into the cells around them, by sheet and cell. */
  private spills = new Map<number, Map<number, Scalar>>();
  /** The block each spilling formula fills, by sheet and the formula's cell. */
  private spillRanges = new Map<number, Map<number, Range>>();
  private spillsFound = false;
  /** A cell in a circular reference, if there is one ("Sheet1!A1"). */
  circular: string | null = null;

  constructor(public book: Workbook) {}

  /** Forgets all worked-out values, after a change. */
  reset() {
    this.memo.clear();
    this.extents.clear();
    this.spills.clear();
    this.spillRanges.clear();
    this.spillsFound = false;
    this.circular = null;
  }

  /** The formula as a tree, or the error reading it. Kept, as the same formulas come back often. */
  tree(formula: string): Node | FormulaError {
    let t = this.trees.get(formula);
    if (!t) {
      try {
        t = parse(formula);
      } catch (error) {
        t = error instanceof FormulaError ? error : new FormulaError(String(error));
      }
      if (this.trees.size > 50000) this.trees.clear();
      this.trees.set(formula, t);
    }
    return t;
  }

  /** A cell's value: what was typed, or what its formula gives. */
  value(sheet: number, r: number, c: number): Scalar {
    const cell = this.book.sheets[sheet]?.cells.get(key(r, c));
    if (!cell || (cell.f === undefined && cell.v === undefined)) return this.spilled(sheet, r, c);
    if (cell.f === undefined) return cell.v ?? null;
    let memo = this.memo.get(sheet);
    if (!memo) this.memo.set(sheet, (memo = new Map()));
    const k = key(r, c);
    if (memo.has(k)) return memo.get(k)!;
    const id = `${sheet}:${k}`;
    if (this.busy.has(id)) {
      // A circular reference: Excel shows 0 and a warning.
      this.circular ??= `${this.book.sheets[sheet].name}!${address(r, c)}`;
      return 0;
    }
    this.busy.add(id);
    let result: Scalar;
    try {
      const tree = this.tree(cell.f);
      if (tree instanceof FormulaError) result = new CellError("#NAME?");
      else {
        const ctx = this.context(sheet, r, c);
        const raw = evaluate(tree, ctx);
        result = raw instanceof ArrayValue && raw.rows * raw.cols > 1 ? this.spill(sheet, r, c, raw) : scalar(raw, ctx);
      }
    } catch (error) {
      // A chain too long to follow at once (stack overflow) or a bug: never take the app down.
      console.warn(`Couldn't work out ${this.book.sheets[sheet].name}!${r + 1}:${c + 1}`, error);
      result = new CellError("#CALC!");
    } finally {
      this.busy.delete(id);
    }
    if (typeof result === "number" && !Number.isFinite(result)) result = new CellError("#NUM!");
    memo.set(k, result);
    return result;
  }

  /** Puts an array formula's values in the cells from (r, c) on; #SPILL! if something is in the way, as in Excel. */
  private spill(sheet: number, r: number, c: number, array: ArrayValue): Scalar {
    const cells = this.book.sheets[sheet].cells;
    for (let i = 0; i < array.rows; i++)
      for (let j = 0; j < array.cols; j++) {
        const other = (i || j) && cells.get(key(r + i, c + j));
        if (other && (other.v !== undefined || other.f !== undefined)) return new CellError("#SPILL!");
      }
    let map = this.spills.get(sheet);
    if (!map) this.spills.set(sheet, (map = new Map()));
    for (let i = 0; i < array.rows; i++) for (let j = 0; j < array.cols; j++) if (i || j) map.set(key(r + i, c + j), array.values[i][j]);
    let ranges = this.spillRanges.get(sheet);
    if (!ranges) this.spillRanges.set(sheet, (ranges = new Map()));
    ranges.set(key(r, c), { r1: r, c1: c, r2: r + array.rows - 1, c2: c + array.cols - 1 });
    return array.values[0][0];
  }

  /** The value an array formula spilled into an empty cell, if any. Formulas that may spill are worked out first. */
  private spilled(sheet: number, r: number, c: number): Scalar {
    if (!this.spillsFound) {
      this.spillsFound = true;
      this.book.sheets.forEach((s, i) => {
        for (const [k, cell] of s.cells) if (cell.f !== undefined && ARRAY_FUNCTIONS.test(cell.f)) this.value(i, rowOf(k), colOf(k));
      });
    }
    return this.spills.get(sheet)?.get(key(r, c)) ?? null;
  }

  /** The block a formula's results fill, if it spills. */
  spillRange(sheet: number, r: number, c: number): Range | undefined {
    this.value(sheet, r, c);
    return this.spillRanges.get(sheet)?.get(key(r, c));
  }

  /** Where a cell's HYPERLINK formula points, if it has one. */
  hyperlink(sheet: number, r: number, c: number): string | null {
    const f = this.book.sheets[sheet]?.cells.get(key(r, c))?.f;
    if (!f || !/^\s*HYPERLINK\s*\(/i.test(f)) return null;
    const tree = this.tree(f);
    if (tree instanceof FormulaError || tree.t !== "fn" || !tree.args[0]) return null;
    const ctx = this.context(sheet, r, c);
    const v = scalar(evaluate(tree.args[0], ctx), ctx);
    return typeof v === "string" && v ? v : null;
  }

  /** What a formula (without its "=") gives at a cell, as if it were there; for rules like conditional formatting. */
  formulaValue(formula: string, sheet: number, r: number, c: number): Scalar {
    const tree = this.tree(formula);
    if (tree instanceof FormulaError) return new CellError("#NAME?");
    try {
      const ctx = this.context(sheet, r, c);
      return scalar(evaluate(tree, ctx), ctx);
    } catch {
      return new CellError("#CALC!");
    }
  }

  /** Works out every formula, top to bottom, so long chains are worked out a step at a time. */
  all() {
    this.book.sheets.forEach((sheet, i) => {
      const keys = [...sheet.cells.entries()].filter(([, cell]) => cell.f !== undefined).map(([k]) => k).sort((a, b) => a - b);
      for (const k of keys) this.value(i, rowOf(k), colOf(k));
    });
  }

  private context(sheet: number, row: number, col: number): EvalContext {
    return {
      sheet,
      row,
      col,
      cell: (s, r, c) => this.value(s, r, c),
      sheetIndex: (name) => this.book.sheets.findIndex((x) => x.name.toLowerCase() === name.toLowerCase()),
      extent: (s) => {
        let e = this.extents.get(s);
        if (!e) {
          let rows = 0;
          let cols = 0;
          for (const k of this.book.sheets[s]?.cells.keys() ?? []) {
            rows = Math.max(rows, rowOf(k) + 1);
            cols = Math.max(cols, colOf(k) + 1);
          }
          this.extents.set(s, (e = { rows, cols }));
        }
        return e;
      },
      name: (name) => this.book.names.get(name),
      table: (name, spec) => {
        const at = resolveTable(this.book, name, spec, { sheet, row, col });
        return at && { sheet: at.sheet, ...at.range };
      },
      date1904: this.book.date1904,
    };
  }
}
