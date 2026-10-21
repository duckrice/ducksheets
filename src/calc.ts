// Works out the value of every formula in a workbook. Values are worked out
// when asked for and kept until something changes; then everything is
// worked out again.
// shortcut: recalculates the whole workbook after each change; fine for tens of thousands of formulas, track dependencies if bigger sheets get slow.

import "./formula/functions";
import { evaluate, scalar, type EvalContext } from "./formula/eval";
import { FormulaError, parse, type Node } from "./formula/parse";
import { CellError, address, colOf, key, rowOf, type Scalar, type Workbook } from "./model";

export class Calculator {
  private memo = new Map<number, Map<number, Scalar>>();
  private busy = new Set<string>();
  private trees = new Map<string, Node | FormulaError>();
  private extents = new Map<number, { rows: number; cols: number }>();
  /** A cell in a circular reference, if there is one ("Sheet1!A1"). */
  circular: string | null = null;

  constructor(public book: Workbook) {}

  /** Forgets all worked-out values, after a change. */
  reset() {
    this.memo.clear();
    this.extents.clear();
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
    if (!cell) return null;
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
        result = scalar(evaluate(tree, ctx), ctx);
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
      date1904: this.book.date1904,
    };
  }
}
