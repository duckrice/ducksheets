// The open workbook and everything done to it. Each change is one step that
// can be undone: before it, the sheets' maps are copied (cells themselves are
// never changed in place, so the copies are cheap), and undo puts them back.

import { Calculator } from "./calc";
import { afterInsertDelete, mapRefs, quoteSheet, parse, print, renamedColumn, renamedSheet, shifted, type Node } from "./formula/parse";
import { columnNames } from "./xlsx/tables";
import { compare } from "./formula/eval";
import { MAX_COLS, MAX_ROWS, colName, colOf, inRange, key, newSheet, rowOf, type Cell, type Range, type Scalar, type Sheet, type SheetTable, type Workbook } from "./model";
import { editText, formatValue, isDateFormat, locale, parseInput } from "./numfmt";
import { defaultStyles, lookOf, withStyle, type StyleChange } from "./xlsx/styles";
import { dxfXml, withoutRules, type CfRule, type CfStyle } from "./conditional";
import { withoutValidation, type Validation } from "./validation";
import type { Note } from "./xlsx/notes";

export interface Selection {
  /** The active cell. */
  r: number;
  c: number;
  /** The other corner of the selected block. */
  ar: number;
  ac: number;
}

export const selRange = (s: Selection): Range => ({ r1: Math.min(s.r, s.ar), c1: Math.min(s.c, s.ac), r2: Math.max(s.r, s.ar), c2: Math.max(s.c, s.ac) });

interface Snapshot {
  sheets: Sheet[];
  names: Map<string, string>;
  active: number;
  sel: Selection;
}

function cloneSheet(s: Sheet): Sheet {
  return {
    ...s,
    cells: new Map(s.cells),
    colWidths: new Map(s.colWidths),
    rowHeights: new Map(s.rowHeights),
    hiddenCols: new Set(s.hiddenCols),
    hiddenRows: new Set(s.hiddenRows),
    merges: s.merges.map((m) => ({ ...m })),
    freeze: { ...s.freeze },
    filter: s.filter && { range: { ...s.filter.range }, shown: new Map([...s.filter.shown].map(([k, v]) => [k, new Set(v)])) },
    charts: [...s.charts],
    conditional: [...s.conditional],
    validations: [...s.validations],
    notes: new Map(s.notes),
    links: new Map(s.links),
    tables: s.tables.map((t) => ({ ...t, range: { ...t.range } })),
  };
}

/** What Copy keeps, to paste formulas and formats as well as values. */
export interface Clip {
  rows: number;
  cols: number;
  /** Row by row; undefined for empty cells. */
  cells: (Cell | undefined)[][];
  /** Where it was copied from, for moving formulas and for cut. */
  sheet: number;
  r: number;
  c: number;
  cut: boolean;
}

export function newWorkbook(): Workbook {
  return { sheets: [newSheet("Sheet1")], styles: defaultStyles(), date1904: false, names: new Map() };
}

/** What a summary works out for each group. */
export type SummaryFn = "sum" | "count" | "average" | "max" | "min";

export class Doc {
  calc: Calculator;
  active = 0;
  sel: Selection = { r: 0, c: 0, ar: 0, ac: 0 };
  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
  /** Changes since the last save. */
  dirty = false;
  private listeners = new Set<() => void>();

  constructor(
    public book: Workbook,
    public name = "Book1.xlsx",
    public path: string | null = null,
  ) {
    this.calc = new Calculator(book);
  }

  get sheet(): Sheet {
    return this.book.sheets[this.active];
  }

  onChange(listener: () => void) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  emit() {
    for (const listener of this.listeners) listener();
  }

  // ---- undo -------------------------------------------------------------------------------

  private snapshot(): Snapshot {
    return { sheets: this.book.sheets.map(cloneSheet), names: new Map(this.book.names), active: this.active, sel: { ...this.sel } };
  }

  /** Makes a change as one step that can be undone. */
  change(fn: () => void) {
    const before = this.snapshot();
    fn();
    this.undoStack.push(before);
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack = [];
    this.dirty = true;
    this.calc.reset();
    this.emit();
  }

  private restore(s: Snapshot) {
    this.book.sheets = s.sheets;
    this.book.names = s.names;
    this.active = Math.min(s.active, s.sheets.length - 1);
    this.sel = s.sel;
    this.dirty = true;
    this.calc.reset();
    this.emit();
  }

  undo() {
    const s = this.undoStack.pop();
    if (!s) return;
    this.redoStack.push(this.snapshot());
    this.restore(s);
  }

  redo() {
    const s = this.redoStack.pop();
    if (!s) return;
    this.undoStack.push(this.snapshot());
    this.restore(s);
  }

  get canUndo() {
    return this.undoStack.length > 0;
  }

  get canRedo() {
    return this.redoStack.length > 0;
  }

  // ---- reading -----------------------------------------------------------------------------

  value(r: number, c: number, sheet = this.active): Scalar {
    return this.calc.value(sheet, r, c);
  }

  cell(r: number, c: number, sheet = this.active): Cell | undefined {
    return this.book.sheets[sheet].cells.get(key(r, c));
  }

  /** What the formula bar shows for a cell: its formula, or its value as typed. */
  inputText(r: number, c: number): string {
    const cell = this.cell(r, c);
    if (!cell) return "";
    if (cell.f !== undefined) return `=${cell.f}`;
    return editText(cell.v ?? null, lookOf(this.book.styles, cell.s).numFmt, this.book.date1904);
  }

  /** A cell's value as it's shown. */
  shown(r: number, c: number, sheet = this.active): string {
    const cell = this.book.sheets[sheet].cells.get(key(r, c));
    return formatValue(this.value(r, c, sheet), lookOf(this.book.styles, cell?.s).numFmt, this.book.date1904).text;
  }

  // ---- typing --------------------------------------------------------------------------------

  /** Sets a cell from what was typed: a formula (starting with =), or a value read as Excel reads it. */
  private put(sheet: Sheet, r: number, c: number, text: string) {
    const k = key(r, c);
    const old = sheet.cells.get(k);
    const s = old?.s;
    if (text === "") {
      if (s) sheet.cells.set(k, { s });
      else sheet.cells.delete(k);
      return;
    }
    if (text.startsWith("=") && text.length > 1) {
      let f = text.slice(1);
      try {
        // Excel's own spelling: function names in capitals, , between arguments.
        f = print(parse(f));
      } catch {
        // Kept as typed; it shows #NAME? until it's mended.
      }
      sheet.cells.set(k, { f, ...(s ? { s } : {}) });
      return;
    }
    const typed = parseInput(text, this.book.date1904);
    // Typing 10% into a General cell makes it a percentage, as in Excel.
    const general = lookOf(this.book.styles, s).numFmt === "General";
    const style = typed.format && general ? withStyle(this.book.styles, s, { numFmt: typed.format }) : s;
    // Typing a line break wraps the text.
    const wrapped = typeof typed.value === "string" && typed.value.includes("\n") && !lookOf(this.book.styles, style).wrap ? withStyle(this.book.styles, style, { wrap: true }) : style;
    sheet.cells.set(k, { v: typed.value, ...(wrapped ? { s: wrapped } : {}) });
  }

  setInput(r: number, c: number, text: string) {
    this.change(() => {
      const table = this.sheet.tables.find((t) => t.range.r1 === r && c >= t.range.c1 && c <= t.range.c2 && !/headerRowCount="0"/.test(t.xml ?? ""));
      const before = table && columnNames(this.sheet, table)[c - table.range.c1];
      this.put(this.sheet, r, c, text);
      this.widen(c, c, r, r);
      // Renaming a table's column renames it in the formulas that use it, as in Excel.
      const after = table && columnNames(this.sheet, table)[c - table.range.c1];
      if (table && before && after && before !== after) this.renameColumn(table, before, after);
    });
  }

  private renameColumn(table: SheetTable, from: string, to: string) {
    this.book.sheets.forEach((sheet, i) => {
      for (const [k, cell] of sheet.cells) {
        if (cell.f === undefined || !cell.f.includes("[")) continue;
        const here = i === this.active && inRange(table.range, rowOf(k), colOf(k));
        try {
          const next = print(renamedColumn(parse(cell.f), table.name, from, to, here));
          if (next !== cell.f) sheet.cells.set(k, { ...cell, f: next });
        } catch {
          // A formula that can't be read is left as it is.
        }
      }
    });
  }

  /** Measures the width a column needs; set by the grid, which knows the fonts. */
  fit: ((c: number) => number) | null = null;

  /**
   * As in Excel: a column at its default width grows to show numbers and dates
   * in full (rather than ###), and a row at its default height grows for bigger text.
   */
  private widen(c1: number, c2: number, r1: number, r2: number) {
    const sheet = this.sheet;
    if (this.fit && c2 - c1 < 50) {
      for (let c = c1; c <= c2; c++) {
        if (sheet.colWidths.has(c)) continue;
        let numbers = false;
        for (let r = r1; r <= Math.min(r2, r1 + 2000) && !numbers; r++) numbers = typeof this.value(r, c) === "number";
        if (!numbers) continue;
        const needed = this.fit(c);
        if (needed > sheet.defaultColWidth) sheet.colWidths.set(c, needed);
      }
    }
    if (r2 - r1 < 5000) {
      for (let r = r1; r <= r2; r++) {
        if (sheet.rowHeights.has(r)) continue;
        let tallest = 0;
        for (let c = c1; c <= Math.min(c2, c1 + 200); c++) {
          const cell = sheet.cells.get(key(r, c));
          if (cell?.s) tallest = Math.max(tallest, lookOf(this.book.styles, cell.s).font.size);
        }
        // Excel gives 11-point text a 15-point row: 1.36 times the font size.
        const needed = Math.ceil(((tallest * 4) / 3) * 1.36);
        if (needed > sheet.defaultRowHeight) sheet.rowHeights.set(r, needed);
      }
    }
  }

  /** Ctrl+Enter: the same input in every selected cell (formulas move with each cell). */
  fillInput(range: Range, text: string) {
    this.change(() => {
      for (let r = range.r1; r <= range.r2; r++)
        for (let c = range.c1; c <= range.c2; c++) {
          if (text.startsWith("=")) {
            try {
              this.put(this.sheet, r, c, `=${print(shifted(parse(text.slice(1)), r - range.r1, c - range.c1))}`);
              continue;
            } catch {
              // fall through: kept as typed
            }
          }
          this.put(this.sheet, r, c, text);
        }
    });
  }

  clear(range: Range, what: "all" | "contents" | "formats" = "contents") {
    this.change(() => {
      const cells = this.sheet.cells;
      for (const [k, cell] of [...cells]) {
        if (!inRange(range, rowOf(k), colOf(k))) continue;
        if (what === "all" || (what === "contents" && !cell.s) || (what === "formats" && cell.f === undefined && cell.v === undefined)) cells.delete(k);
        else if (what === "contents") cells.set(k, { s: cell.s });
        else {
          const { s: _s, ...rest } = cell;
          cells.set(k, rest);
        }
      }
      if (what !== "contents") {
        this.sheet.merges = this.sheet.merges.filter((m) => !(m.r1 >= range.r1 && m.r2 <= range.r2 && m.c1 >= range.c1 && m.c2 <= range.c2));
        this.sheet.conditional = withoutRules(this.sheet.conditional, range);
      }
    });
  }

  // ---- conditional formatting -----------------------------------------------------------------

  /** Adds a rule for a block of cells, above the rules already there (as Excel does). */
  addRule(range: Range, rule: CfRule) {
    this.change(() => {
      this.sheet.conditional = [{ ranges: [range], rules: [rule] }, ...this.sheet.conditional];
    });
  }

  /** A format for rules: its index in the workbook's list, added if it isn't there yet. */
  dxf(style: CfStyle): number {
    const xml = dxfXml(style);
    const list = this.book.styles.dxfs;
    const i = list.indexOf(xml);
    return i >= 0 ? i : list.push(xml) - 1;
  }

  /** Sets what may be typed in a block of cells (a list, number limits), or (null) takes that off. */
  setValidation(range: Range, rule: Validation | null) {
    this.change(() => {
      const rest = withoutValidation(this.sheet.validations, range);
      this.sheet.validations = rule ? [...rest, { ...rule, ranges: [range] }] : rest;
    });
  }

  /** Puts a link on a cell (and, if given, the text it shows), or (empty) takes it off. */
  setLink(r: number, c: number, link: string, text?: string) {
    this.change(() => {
      if (text !== undefined && text !== "") this.put(this.sheet, r, c, text);
      else if (link && !this.sheet.cells.get(key(r, c))?.v && !this.sheet.cells.get(key(r, c))?.f) this.put(this.sheet, r, c, link.replace(/^#/, "").replace(/^mailto:/, ""));
      if (link) this.sheet.links.set(key(r, c), link);
      else this.sheet.links.delete(key(r, c));
    });
  }

  /** Puts a note on a cell, or (empty text) takes it off. */
  setNote(r: number, c: number, text: string, author: string) {
    this.change(() => {
      if (text.trim()) this.sheet.notes.set(key(r, c), { text, author });
      else this.sheet.notes.delete(key(r, c));
    });
  }

  /** Takes the rules off a block of cells, or (null) the whole sheet. */
  clearRules(range: Range | null) {
    this.change(() => {
      this.sheet.conditional = range ? withoutRules(this.sheet.conditional, range) : [];
    });
  }

  // ---- formatting -----------------------------------------------------------------------------

  /** Applies a format change to every cell of the range (empty ones too, so typing there keeps it). */
  format(range: Range, change: StyleChange | ((s: number | undefined, r: number, c: number) => StyleChange)) {
    const r2 = Math.min(range.r2, this.lastRow(range));
    const c2 = Math.min(range.c2, this.lastCol(range));
    this.change(() => {
      const memo = new Map<string, number>();
      for (let r = range.r1; r <= r2; r++)
        for (let c = range.c1; c <= c2; c++) {
          const k = key(r, c);
          const cell = this.sheet.cells.get(k);
          const ch = typeof change === "function" ? change(cell?.s, r, c) : change;
          const id = `${cell?.s ?? 0}|${JSON.stringify(ch)}`;
          let s = memo.get(id);
          if (s === undefined) memo.set(id, (s = withStyle(this.book.styles, cell?.s, ch)));
          if (s) this.sheet.cells.set(k, { ...cell, s });
          else if (cell && (cell.f !== undefined || cell.v !== undefined)) {
            const { s: _s, ...rest } = cell;
            this.sheet.cells.set(k, rest);
          } else this.sheet.cells.delete(k);
        }
      this.widen(range.c1, c2, range.r1, r2);
    });
  }

  /** Whole columns or rows reach only as far as the data does, so formatting them stays quick. */
  private lastRow(range: Range) {
    if (range.r2 < MAX_ROWS - 1) return range.r2;
    let last = range.r1;
    for (const k of this.sheet.cells.keys()) last = Math.max(last, rowOf(k));
    return last;
  }

  private lastCol(range: Range) {
    if (range.c2 < MAX_COLS - 1) return range.c2;
    let last = range.c1;
    for (const k of this.sheet.cells.keys()) last = Math.max(last, colOf(k));
    return last;
  }

  /** Merges the range, keeping the top-left value (as Excel does), or unmerges what's in it. */
  toggleMerge(range: Range, center = true) {
    const inside = this.sheet.merges.filter((m) => m.r1 >= range.r1 && m.r2 <= range.r2 && m.c1 >= range.c1 && m.c2 <= range.c2);
    this.change(() => {
      if (inside.length) {
        this.sheet.merges = this.sheet.merges.filter((m) => !inside.includes(m));
        return;
      }
      if (range.r1 === range.r2 && range.c1 === range.c2) return;
      this.sheet.merges = this.sheet.merges.filter((m) => m.r2 < range.r1 || m.r1 > range.r2 || m.c2 < range.c1 || m.c1 > range.c2);
      for (const k of [...this.sheet.cells.keys()]) {
        const r = rowOf(k);
        const c = colOf(k);
        if (inRange(range, r, c) && (r !== range.r1 || c !== range.c1)) this.sheet.cells.delete(k);
      }
      this.sheet.merges.push({ ...range });
      if (center) {
        const k = key(range.r1, range.c1);
        const cell = this.sheet.cells.get(k);
        this.sheet.cells.set(k, { ...cell, s: withStyle(this.book.styles, cell?.s, { h: "center" }) });
      }
    });
  }

  setColWidth(cols: number[], px: number) {
    this.change(() => cols.forEach((c) => this.sheet.colWidths.set(c, Math.max(0, Math.round(px)))));
  }

  setRowHeight(rows: number[], px: number) {
    this.change(() => rows.forEach((r) => this.sheet.rowHeights.set(r, Math.max(0, Math.round(px)))));
  }

  setHidden(axis: "r" | "c", indexes: number[], hidden: boolean) {
    this.change(() => {
      const set = axis === "r" ? this.sheet.hiddenRows : this.sheet.hiddenCols;
      for (const i of indexes) {
        if (hidden) set.add(i);
        else set.delete(i);
      }
    });
  }

  freeze(rows: number, cols: number) {
    this.change(() => (this.sheet.freeze = { rows, cols }));
  }

  // ---- rows and columns ------------------------------------------------------------------------

  /** Every formula in the workbook, rewritten by `fn` (given the formula's own sheet name). */
  private rewriteFormulas(fn: (node: Node, home: string) => Node) {
    for (const sheet of this.book.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (cell.f === undefined) continue;
        try {
          const next = print(fn(parse(cell.f), sheet.name));
          if (next !== cell.f) sheet.cells.set(k, { ...cell, f: next });
        } catch {
          // A formula that can't be read is left as it is.
        }
      }
    }
    for (const [name, text] of this.book.names) {
      try {
        this.book.names.set(name, print(fn(parse(text.replace(/^=/, "")), this.sheet.name)));
      } catch {
        // as above
      }
    }
  }

  /** Inserts (count > 0) or deletes (count < 0) rows (axis "r") or columns ("c") at `at`. */
  insertDelete(axis: "r" | "c", at: number, count: number) {
    const sheet = this.sheet;
    const name = sheet.name;
    this.change(() => {
      const end = at - count; // for deleting: the first index after the deleted ones
      const move = (i: number): number | null => (i < at ? i : count > 0 ? i + count : i < end ? null : i + count);
      const cells = new Map<number, Cell>();
      for (const [k, cell] of sheet.cells) {
        const r = rowOf(k);
        const c = colOf(k);
        const i = move(axis === "r" ? r : c);
        if (i === null || i >= (axis === "r" ? MAX_ROWS : MAX_COLS)) continue;
        cells.set(axis === "r" ? key(i, c) : key(r, i), cell);
      }
      sheet.cells = cells;
      const notes = new Map<number, Note>();
      for (const [k, note] of sheet.notes) {
        const r = rowOf(k);
        const c = colOf(k);
        const i = move(axis === "r" ? r : c);
        if (i !== null) notes.set(axis === "r" ? key(i, c) : key(r, i), note);
      }
      sheet.notes = notes;
      const links = new Map<number, string>();
      for (const [k, link] of sheet.links) {
        const r = rowOf(k);
        const c = colOf(k);
        const i = move(axis === "r" ? r : c);
        if (i !== null) links.set(axis === "r" ? key(i, c) : key(r, i), link);
      }
      sheet.links = links;
      const sizes = axis === "r" ? sheet.rowHeights : sheet.colWidths;
      const moved = new Map<number, number>();
      for (const [i, size] of sizes) {
        const j = move(i);
        if (j !== null) moved.set(j, size);
      }
      if (axis === "r") sheet.rowHeights = moved;
      else sheet.colWidths = moved;
      const hidden = axis === "r" ? sheet.hiddenRows : sheet.hiddenCols;
      const movedHidden = new Set([...hidden].map(move).filter((i): i is number => i !== null));
      if (axis === "r") sheet.hiddenRows = movedHidden;
      else sheet.hiddenCols = movedHidden;
      const [lo, hi] = axis === "r" ? (["r1", "r2"] as const) : (["c1", "c2"] as const);
      const shiftRange = (m: Range): Range | null => {
        let a = m[lo];
        let b = m[hi];
        if (count > 0) {
          a = a >= at ? a + count : a;
          b = b >= at ? b + count : b;
        } else {
          if (a >= at && b < end) return null;
          a = a < at ? a : a < end ? at : a + count;
          b = b < at ? b : b < end ? at - 1 : b + count;
        }
        return { ...m, [lo]: a, [hi]: b };
      };
      sheet.merges = sheet.merges.map(shiftRange).filter((m): m is Range => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));
      sheet.conditional = sheet.conditional.map((b) => ({ ...b, ranges: b.ranges.map(shiftRange).filter((g): g is Range => !!g) })).filter((b) => b.ranges.length);
      sheet.validations = sheet.validations.map((v) => ({ ...v, ranges: v.ranges.map(shiftRange).filter((g): g is Range => !!g) })).filter((v) => v.ranges.length);
      sheet.tables = sheet.tables.flatMap((t) => {
        const range = shiftRange(t.range);
        return range ? [{ ...t, range }] : [];
      });
      if (sheet.filter) {
        const range = shiftRange(sheet.filter.range);
        sheet.filter = range ? { ...sheet.filter, range } : null;
      }
      // Charts keep showing their data, and sit where they did; a chart whose data is all deleted goes.
      sheet.charts = sheet.charts.flatMap((chart) => {
        const data = shiftRange(chart.data);
        if (!data) return [];
        const corner = axis === "r" ? chart.at.r : chart.at.c;
        const placed = move(corner) ?? at;
        return [{ ...chart, data, at: { ...chart.at, [axis]: placed } }];
      });
      this.rewriteFormulas((node, home) => afterInsertDelete(node, home, name, axis, at, count));
    });
  }

  // ---- sheets ---------------------------------------------------------------------------------

  addSheet() {
    this.change(() => {
      let n = this.book.sheets.length + 1;
      while (this.book.sheets.some((s) => s.name.toLowerCase() === `sheet${n}`)) n++;
      this.book.sheets.splice(this.active + 1, 0, newSheet(`Sheet${n}`));
      this.active++;
      this.sel = { r: 0, c: 0, ar: 0, ac: 0 };
    });
  }

  /**
   * A summary of the data in `range` (headers in its first row) on a new sheet, as a simple
   * PivotTable: each value of column `by` once, A to Z, with the sum (or count, average, …) of
   * column `of` for it, and a total. Formulas, so it follows changes to the data.
   */
  summarize(range: Range, by: number, of: number, fn: SummaryFn) {
    const source = this.sheet;
    const top = range.r1 + 1;
    const column = (c: number) => `${quoteSheet(source.name)}!${colName(c)}${top + 1}:${colName(c)}${range.r2 + 1}`;
    const keys = new Map<string, Scalar>();
    for (let r = top; r <= range.r2; r++) {
      const v = this.value(r, by);
      if (v === null || v === "" || typeof v === "object") continue;
      const id = typeof v === "string" ? v.toLowerCase() : `${typeof v}:${v}`;
      if (!keys.has(id)) keys.set(id, v);
    }
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
    const groups = [...keys.values()].sort((a, b) => (typeof a === typeof b ? (typeof a === "string" ? collator.compare(a, b as string) : Number(a) - Number(b)) : typeof a === "string" ? 1 : -1));
    const label = (c: number) => String(this.value(range.r1, c) ?? `Column ${colName(c)}`);
    const [byName, ofName] = [label(by), label(of)];
    const ifs = { sum: "SUMIFS", count: "COUNTIFS", average: "AVERAGEIFS", max: "MAXIFS", min: "MINIFS" }[fn];
    const total = { sum: "SUM", count: "COUNTA", average: "AVERAGE", max: "MAX", min: "MIN" }[fn];
    this.addSheet();
    this.change(() => {
      const sheet = this.sheet;
      let name = "Summary";
      for (let n = 2; this.book.sheets.some((s) => s !== sheet && s.name.toLowerCase() === name.toLowerCase()); n++) name = `Summary ${n}`;
      sheet.name = name;
      this.put(sheet, 0, 0, byName);
      this.put(sheet, 0, 1, `${fn[0].toUpperCase()}${fn.slice(1)} of ${ofName}`);
      groups.forEach((g, i) => {
        // shortcut: a group whose text has * or ? matches like a wildcard, as in Excel's own SUMIFS.
        this.put(sheet, i + 1, 0, typeof g === "string" ? `'${g}` : String(g === true ? "TRUE" : g === false ? "FALSE" : g));
        const criteria = `${column(by)},A${i + 2}`;
        this.put(sheet, i + 1, 1, fn === "count" ? `=COUNTIFS(${criteria})` : `=${ifs}(${column(of)},${criteria})`);
      });
      const end = groups.length + 1;
      this.put(sheet, end, 0, "Total");
      this.put(sheet, end, 1, `=${total}(${fn === "count" ? column(by) : column(of)})`);
    });
    const header = { bold: true, color: "#FFFFFF", fill: "#4472C4" };
    this.format({ r1: 0, c1: 0, r2: 0, c2: 1 }, () => header);
    this.format({ r1: groups.length + 1, c1: 0, r2: groups.length + 1, c2: 1 }, () => ({ bold: true, border: { top: { style: "thin", color: "#4472C4" } } }));
    this.change(() => this.widen(0, 1, 0, groups.length + 1));
  }

  /** Renames a sheet; formulas pointing at it follow. Returns an error message if the name can't be used. */
  renameSheet(index: number, to: string): string | null {
    const name = to.trim();
    if (!name) return "A sheet needs a name.";
    if (name.length > 31) return "Sheet names can be at most 31 characters long.";
    if (/[\\/?*[\]:]/.test(name)) return "Sheet names can't contain \\ / ? * [ ] or :";
    if (this.book.sheets.some((s, i) => i !== index && s.name.toLowerCase() === name.toLowerCase())) return "There is already a sheet with that name.";
    const from = this.book.sheets[index].name;
    this.change(() => {
      this.book.sheets[index].name = name;
      this.rewriteFormulas((node) => renamedSheet(node, from, name));
    });
    return null;
  }

  deleteSheet(index: number) {
    if (this.book.sheets.length < 2) return;
    const gone = this.book.sheets[index].name.toLowerCase();
    this.change(() => {
      this.book.sheets.splice(index, 1);
      this.active = Math.min(this.active, this.book.sheets.length - 1);
      // References to the deleted sheet become #REF!, as in Excel.
      this.rewriteFormulas((node) => mapSheetGone(node, gone));
    });
  }

  moveSheet(from: number, to: number) {
    if (from === to) return;
    this.change(() => {
      const [s] = this.book.sheets.splice(from, 1);
      this.book.sheets.splice(to, 0, s);
      this.active = to;
    });
  }

  // ---- sorting and filtering ----------------------------------------------------------------------

  /** The block of filled cells around a cell, as Excel's Ctrl+A and Sort find it. */
  region(r: number, c: number): Range {
    const filled = (rr: number, cc: number) => rr >= 0 && cc >= 0 && this.sheet.cells.has(key(rr, cc)) && this.value(rr, cc) !== null;
    const range = { r1: r, c1: c, r2: r, c2: c };
    for (let grew = true; grew; ) {
      grew = false;
      const tryGrow = (side: keyof Range, delta: number) => {
        const next = { ...range, [side]: range[side] + delta };
        const line =
          side === "r1" || side === "r2"
            ? Array.from({ length: range.c2 - range.c1 + 3 }, (_, i) => filled(next[side], range.c1 - 1 + i))
            : Array.from({ length: range.r2 - range.r1 + 3 }, (_, i) => filled(range.r1 - 1 + i, next[side]));
        if (line.some(Boolean) && next[side] >= 0) {
          Object.assign(range, next);
          grew = true;
        }
      };
      tryGrow("r1", -1);
      tryGrow("r2", 1);
      tryGrow("c1", -1);
      tryGrow("c2", 1);
    }
    return range;
  }

  /** Whether the first row of a range looks like headings: text over a column with other kinds of values, or bold. */
  hasHeader(range: Range): boolean {
    if (range.r2 <= range.r1) return false;
    for (let c = range.c1; c <= range.c2; c++) {
      const top = this.value(range.r1, c);
      const below = this.value(range.r1 + 1, c);
      if (lookOf(this.book.styles, this.cell(range.r1, c)?.s).font.bold && !lookOf(this.book.styles, this.cell(range.r1 + 1, c)?.s).font.bold) return true;
      if (typeof top === "string" && below !== null && typeof below !== "string") return true;
    }
    return false;
  }

  /** Sorts the rows of a range by one column; empty cells go last either way. */
  sort(range: Range, col: number, descending: boolean, header = this.hasHeader(range)) {
    const top = header ? range.r1 + 1 : range.r1;
    if (top >= range.r2) return;
    const rows = Array.from({ length: range.r2 - top + 1 }, (_, i) => top + i);
    const keys = new Map(rows.map((r) => [r, this.value(r, col)]));
    rows.sort((a, b) => {
      const x = keys.get(a)!;
      const y = keys.get(b)!;
      if (x === null || y === null) return x === null ? (y === null ? 0 : 1) : -1;
      const c = compare(x, y);
      return descending ? -c : c;
    });
    this.change(() => {
      const cells = this.sheet.cells;
      const old = new Map<number, Cell | undefined>();
      for (const r of rows) for (let c = range.c1; c <= range.c2; c++) old.set(key(r, c), cells.get(key(r, c)));
      rows.forEach((from, i) => {
        const to = top + i;
        for (let c = range.c1; c <= range.c2; c++) {
          const cell = old.get(key(from, c));
          const k = key(to, c);
          if (!cell) cells.delete(k);
          else cells.set(k, cell.f !== undefined ? { ...cell, f: moveFormula(cell.f, to - from, 0) } : cell);
        }
      });
    });
  }

  /**
   * Removes rows whose values in `cols` repeat an earlier row's (text compared without regard to
   * case), as Excel's Remove duplicates: the rows left move up and the bottom is cleared. Returns
   * how many rows were removed.
   */
  removeDuplicates(range: Range, cols: number[], header = this.hasHeader(range)): number {
    const top = header ? range.r1 + 1 : range.r1;
    const seen = new Set<string>();
    const keep: number[] = [];
    for (let r = top; r <= range.r2; r++) {
      const id = JSON.stringify(cols.map((c) => {
        const v = this.value(r, c);
        return typeof v === "string" ? v.toLowerCase() : v instanceof Object ? String(v) : v;
      }));
      if (seen.has(id)) continue;
      seen.add(id);
      keep.push(r);
    }
    const removed = range.r2 - top + 1 - keep.length;
    if (!removed) return 0;
    this.change(() => {
      const cells = this.sheet.cells;
      const old = new Map<number, Cell | undefined>();
      for (let r = top; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) old.set(key(r, c), cells.get(key(r, c)));
      for (let i = 0; i <= range.r2 - top; i++) {
        const from = keep[i];
        for (let c = range.c1; c <= range.c2; c++) {
          const k = key(top + i, c);
          const cell = from === undefined ? undefined : old.get(key(from, c));
          if (!cell) cells.delete(k);
          else cells.set(k, cell.f !== undefined ? { ...cell, f: moveFormula(cell.f, top + i - from, 0) } : cell);
        }
      }
    });
    return removed;
  }

  /**
   * Goal Seek, as Excel's: finds the number for `input` that makes the formula in `target` give
   * `goal`, and puts it in (one step to undo). Returns the number, or null if none was found.
   */
  goalSeek(target: { r: number; c: number }, goal: number, input: { r: number; c: number }): number | null {
    const cells = this.sheet.cells;
    const k = key(input.r, input.c);
    const original = cells.get(k);
    const at = (x: number) => {
      cells.set(k, { ...(original?.s !== undefined ? { s: original.s } : {}), v: x });
      this.calc.reset();
      const v = this.calc.value(this.active, target.r, target.c);
      return typeof v === "number" ? v - goal : NaN;
    };
    // The secant method from the number there now, as Excel does; then a wider search if that fails.
    let x0 = typeof original?.v === "number" ? original.v : 0;
    let x1 = x0 === 0 ? 1 : x0 * 1.01;
    let f0 = at(x0);
    let found: number | null = Math.abs(f0) < 1e-9 ? x0 : null;
    for (let i = 0; i < 100 && found === null && Number.isFinite(f0); i++) {
      const f1 = at(x1);
      if (!Number.isFinite(f1)) break;
      if (Math.abs(f1) < 1e-9 * Math.max(1, Math.abs(goal))) found = x1;
      else if (f1 === f0) break;
      else {
        const next = x1 - (f1 * (x1 - x0)) / (f1 - f0);
        [x0, f0, x1] = [x1, f1, next];
      }
    }
    // Put the cell back as it was, then (if a number was found) put it in as one step.
    if (original) cells.set(k, original);
    else cells.delete(k);
    this.calc.reset();
    if (found === null) return null;
    const value = Math.abs(found - Math.round(found)) < 1e-9 ? Math.round(found) : found;
    this.change(() => cells.set(k, { ...(original?.s !== undefined ? { s: original.s } : {}), v: value }));
    return value;
  }

  /** Splits a column's text at a separator into the columns to its right, each part read as if typed. */
  textToColumns(range: Range, separator: string) {
    if (!separator) return;
    this.change(() => {
      for (let r = range.r1; r <= range.r2; r++) {
        const cell = this.sheet.cells.get(key(r, range.c1));
        if (!cell || cell.f !== undefined || typeof cell.v !== "string") continue;
        const parts = separator === " " ? cell.v.trim().split(/\s+/) : cell.v.split(separator);
        parts.forEach((part, j) => this.put(this.sheet, r, range.c1 + j, part.trim()));
      }
    });
  }

  /** Format painter: the formats copied (row by row), repeated over the cells painted, contents kept. */
  paintFormats(source: (number | undefined)[][], range: Range) {
    const rows = source.length;
    const cols = source[0]?.length ?? 0;
    if (!rows || !cols) return;
    this.change(() => {
      for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 5000); r++)
        for (let c = range.c1; c <= Math.min(range.c2, range.c1 + 500); c++) {
          const k = key(r, c);
          const s = source[(r - range.r1) % rows][(c - range.c1) % cols];
          const { s: _old, ...content } = this.sheet.cells.get(k) ?? {};
          const next = { ...content, ...(s !== undefined ? { s } : {}) };
          if (Object.keys(next).length) this.sheet.cells.set(k, next);
          else this.sheet.cells.delete(k);
        }
    });
  }

  /**
   * Format as table, as Excel's quick styles: a coloured header row in white bold, rows striped in
   * a light shade, thin lines, and filter buttons on the header.
   */
  formatAsTable(range: Range, header: string, band: string, style = "TableStyleMedium2") {
    const last = Math.min(range.r2, this.lastRow(range));
    // An Excel table needs a header row of names, and no other table in the way.
    const block = { ...range, r2: Math.max(last, range.r1 + 1) };
    const overlaps = (t: { range: Range }) => !(t.range.r2 < block.r1 || t.range.r1 > block.r2 || t.range.c2 < block.c1 || t.range.c1 > block.c2);
    const line = { style: "thin", color: header };
    this.format({ ...range, r2: last }, (_s, r) =>
      r === range.r1
        ? { bold: true, color: "#FFFFFF", fill: header, border: { bottom: line } }
        : { fill: (r - range.r1) % 2 === 1 ? band : null, border: { bottom: { style: "thin", color: band } } },
    );
    this.change(() => {
      const sheet = this.sheet;
      for (let c = block.c1; c <= block.c2; c++) {
        const v = this.value(block.r1, c);
        if (v === null || v === "") this.put(sheet, block.r1, c, `Column${c - block.c1 + 1}`);
        else if (typeof v !== "string") this.put(sheet, block.r1, c, `'${this.shown(block.r1, c)}`);
      }
      const taken = new Set(this.book.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())));
      let n = 1;
      while (taken.has(`table${n}`) || this.book.names.has(`TABLE${n}`)) n++;
      sheet.tables = [...sheet.tables.filter((t) => !overlaps(t)), { name: `Table${n}`, range: block, style }];
      sheet.filter = { range: block, shown: new Map() };
    });
  }

  /** Turns filtering on for the data around the selection, or off. */
  toggleFilter() {
    this.change(() => {
      if (this.sheet.filter) {
        this.sheet.filter = null;
        this.sheet.hiddenRows.clear();
        return;
      }
      const range = selRange(this.sel);
      const area = range.r1 === range.r2 && range.c1 === range.c2 ? this.region(range.r1, range.c1) : range;
      this.sheet.filter = { range: area, shown: new Map() };
    });
  }

  /** Shows only the rows whose value in column `col` (counted from the filter's first column) is in `shown`; null shows all. */
  setFilter(col: number, shown: Set<string> | null) {
    const filter = this.sheet.filter;
    if (!filter) return;
    this.change(() => {
      if (shown) filter.shown.set(col, shown);
      else filter.shown.delete(col);
      this.applyFilter();
    });
  }

  private applyFilter() {
    const filter = this.sheet.filter!;
    const { range } = filter;
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const visible = [...filter.shown].every(([col, shown]) => shown.has(this.shown(r, range.c1 + col)));
      if (visible) this.sheet.hiddenRows.delete(r);
      else this.sheet.hiddenRows.add(r);
    }
  }

  // ---- filling, copying and pasting --------------------------------------------------------------

  /** Drags the selection's contents over `target` (the fill handle): series where it sees one, copies otherwise. */
  fill(source: Range, target: Range) {
    const down = target.r2 > source.r2;
    const up = target.r1 < source.r1;
    const right = target.c2 > source.c2;
    const vertical = down || up;
    const left = target.c1 < source.c1;
    if (!down && !up && !right && !left) return;
    this.change(() => {
      const lines = vertical ? source.c2 - source.c1 + 1 : source.r2 - source.r1 + 1;
      const length = vertical ? source.r2 - source.r1 + 1 : source.c2 - source.c1 + 1;
      for (let line = 0; line < lines; line++) {
        const at = (i: number) => (vertical ? { r: source.r1 + i, c: source.c1 + line } : { r: source.r1 + line, c: source.c1 + i });
        const cells = Array.from({ length }, (_, i) => this.cell(at(i).r, at(i).c));
        const isDate = isDateFormat(lookOf(this.book.styles, cells[0]?.s).numFmt);
        const series = seriesOf(cells.map((cell) => (cell?.f === undefined ? (cell?.v ?? null) : null)), isDate);
        const count = vertical ? (down ? target.r2 - source.r2 : source.r1 - target.r1) : right ? target.c2 - source.c2 : source.c1 - target.c1;
        for (let n = 1; n <= count; n++) {
          const step = down || right ? length - 1 + n : -n;
          const src = cells[(((down || right ? n - 1 : -n) % length) + length) % length];
          const r = vertical ? source.r1 + step : source.r1 + line;
          const c = vertical ? source.c1 + line : source.c1 + step;
          const k = key(r, c);
          if (!src) {
            this.sheet.cells.delete(k);
            continue;
          }
          const from = at((((down || right ? n - 1 : -n) % length) + length) % length);
          if (src.f !== undefined) this.sheet.cells.set(k, { ...src, f: moveFormula(src.f, r - from.r, c - from.c) });
          else this.sheet.cells.set(k, { ...src, v: series ? series(step) : src.v });
        }
      }
    });
  }

  copy(range: Range, cut = false): Clip {
    const cells = Array.from({ length: range.r2 - range.r1 + 1 }, (_, i) => Array.from({ length: range.c2 - range.c1 + 1 }, (_, j) => this.cell(range.r1 + i, range.c1 + j)));
    return { rows: cells.length, cols: cells[0].length, cells, sheet: this.active, r: range.r1, c: range.c1, cut };
  }

  /** The block as text for other programs: tab-separated rows of what's shown. */
  copyText(range: Range): { tsv: string; html: string } {
    const rows: string[][] = [];
    for (let r = range.r1; r <= range.r2; r++) {
      if (this.sheet.hiddenRows.has(r)) continue;
      const row: string[] = [];
      for (let c = range.c1; c <= range.c2; c++) row.push(this.shown(r, c));
      rows.push(row);
    }
    const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    return {
      tsv: rows.map((row) => row.map((v) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join("\t")).join("\r\n"),
      html: `<table>${rows.map((row) => `<tr>${row.map((v) => `<td>${esc(v)}</td>`).join("")}</tr>`).join("")}</table>`,
    };
  }

  /** Pastes what was copied in DuckSheets: formulas move with the cells, formats come along. */
  /**
   * Pastes what was copied: everything, or (Paste special) only the values, the formulas or the
   * formats, the latter keeping the formats or contents already there; turned on its side if asked.
   */
  paste(clip: Clip, at: { r: number; c: number }, how: { what?: "all" | "values" | "formulas" | "formats"; transpose?: boolean } = {}) {
    const what = how.what ?? "all";
    this.change(() => {
      const source = this.book.sheets[clip.sheet];
      // Values as they are now, before a cut takes the cells away.
      const values = what === "values" ? clip.cells.map((row, i) => row.map((_, j) => this.calc.value(clip.sheet, clip.r + i, clip.c + j))) : null;
      if (clip.cut && what === "all" && source) for (let i = 0; i < clip.rows; i++) for (let j = 0; j < clip.cols; j++) source.cells.delete(key(clip.r + i, clip.c + j));
      for (let i = 0; i < clip.rows; i++)
        for (let j = 0; j < clip.cols; j++) {
          const [ti, tj] = how.transpose ? [j, i] : [i, j];
          const k = key(at.r + ti, at.c + tj);
          const cell = clip.cells[i][j];
          const old = this.sheet.cells.get(k);
          const keep = old?.s !== undefined ? { s: old.s } : {};
          // A cut moves formulas as they are; a copy moves their relative references.
          const [dr, dc] = [at.r + ti - clip.r - i, at.c + tj - clip.c - j];
          const moved = cell?.f !== undefined && !clip.cut ? { ...cell, f: how.transpose ? transposedFormula(cell.f, clip, at, dr, dc) : moveFormula(cell.f, dr, dc) } : cell;
          let next: Cell | undefined;
          if (what === "all") next = moved;
          else if (what === "values") next = values![i][j] === null ? undefined : { v: values![i][j], ...keep };
          else if (what === "formulas") next = moved && (moved.f !== undefined || moved.v !== undefined) ? { ...(moved.f !== undefined ? { f: moved.f } : { v: moved.v }), ...keep } : undefined;
          else {
            const { s: _s, ...content } = old ?? {};
            next = { ...content, ...(cell?.s !== undefined ? { s: cell.s } : {}) };
          }
          if (next && Object.keys(next).length) this.sheet.cells.set(k, next);
          else if (what === "all" || !old?.s || what === "formats") this.sheet.cells.delete(k);
          else this.sheet.cells.set(k, { s: old.s });
        }
      const [rows, cols] = how.transpose ? [clip.cols, clip.rows] : [clip.rows, clip.cols];
      this.sel = { r: at.r, c: at.c, ar: at.r + rows - 1, ac: at.c + cols - 1 };
    });
  }

  /** Pastes text from another program: tab-separated rows, each value read as if typed. */
  pasteText(text: string, at: { r: number; c: number }) {
    const rows = text.replace(/\r?\n$/, "").split(/\r?\n/).map((line) => line.split("\t").map((v) => v.replace(/^"([\s\S]*)"$/, (_, inner: string) => inner.replace(/""/g, '"'))));
    this.change(() => {
      rows.forEach((row, i) => row.forEach((v, j) => this.put(this.sheet, at.r + i, at.c + j, v)));
      this.sel = { r: at.r, c: at.c, ar: at.r + rows.length - 1, ac: at.c + Math.max(...rows.map((r) => r.length)) - 1 };
    });
  }
}

/** A formula pasted on its side: references into the copied block turn with it, others move as in a copy. */
function transposedFormula(f: string, clip: Clip, at: { r: number; c: number }, dr: number, dc: number): string {
  const inside = (r: number, c: number) => r >= clip.r && r < clip.r + clip.rows && c >= clip.c && c < clip.c + clip.cols;
  try {
    return print(mapRefs(parse(f), (ref) => (ref.t === "ref" && ref.sheet === undefined && !ref.ar && !ref.ac && inside(ref.r, ref.c) ? { ...ref, r: at.r + ref.c - clip.c, c: at.c + ref.r - clip.r } : shifted(ref, dr, dc))));
  } catch {
    return f;
  }
}

/** A formula as copied by (dr, dc), or as it was if it can't be read. */
export function moveFormula(f: string, dr: number, dc: number): string {
  try {
    return print(shifted(parse(f), dr, dc));
  } catch {
    return f;
  }
}

function mapSheetGone(node: Node, gone: string): Node {
  const walk = (n: Node): Node => {
    if ((n.t === "ref" || n.t === "range") && n.sheet !== undefined && n.sheet.toLowerCase() === gone) return { t: "err", v: "#REF!" };
    if (n.t === "fn") return { ...n, args: n.args.map(walk) };
    if (n.t === "un" || n.t === "pct" || n.t === "paren") return { ...n, e: walk(n.e) };
    if (n.t === "bin") return { ...n, l: walk(n.l), r: walk(n.r) };
    return n;
  };
  return walk(node);
}

// ---- series for the fill handle ------------------------------------------------------------------

const lists = (): string[][] => {
  const en = {
    months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    days: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  };
  const mondayFirst = (days: string[]) => [...days.slice(1), days[0]];
  return [en.months, en.months.map((m) => m.slice(0, 3)), en.days, en.days.map((d) => d.slice(0, 3)), locale.months, locale.monthsShort, mondayFirst(locale.days), mondayFirst(locale.daysShort)];
};

/**
 * How the fill handle continues a run of values: numbers by their step (one
 * number is copied, as in Excel), text ending in a number counts on
 * (Week 1, Week 2), and month and day names go round. Null means copy.
 */
export function seriesOf(values: Scalar[], dates = false): ((i: number) => Scalar) | null {
  if (!values.length || values.some((v) => v === null)) return null;
  if (values.every((v) => typeof v === "number")) {
    const ns = values as number[];
    // One number is copied; one date counts on a day at a time, as in Excel.
    if (ns.length < 2) return dates ? (i) => ns[0] + i : null;
    const step = (ns[ns.length - 1] - ns[0]) / (ns.length - 1);
    if (ns.some((n, i) => Math.abs(n - (ns[0] + step * i)) > 1e-9)) return null;
    return (i) => Math.round((ns[0] + step * i) * 1e10) / 1e10;
  }
  if (values.every((v) => typeof v === "string")) {
    const texts = values as string[];
    for (const list of lists()) {
      const at = texts.map((t) => list.findIndex((x) => x.toLowerCase() === t.toLowerCase()));
      if (at.some((i) => i < 0)) continue;
      const step = at.length > 1 ? at[1] - at[0] : 1;
      const style = (t: string, word: string) => (t === t.toUpperCase() ? word.toUpperCase() : t === t.toLowerCase() ? word.toLowerCase() : word);
      return (i) => style(texts[0], list[(((at[0] + step * i) % list.length) + list.length) % list.length]);
    }
    const parts = texts.map((t) => /^(.*?)(\d+)(\D*)$/.exec(t));
    if (parts.every((p) => p && p[1] === parts[0]![1] && p[3] === parts[0]![3])) {
      const ns = parts.map((p) => Number(p![2]));
      const step = ns.length > 1 ? ns[1] - ns[0] : 1;
      const width = parts[0]![2].length;
      return (i) => `${parts[0]![1]}${String(Math.max(0, ns[0] + step * i)).padStart(parts[0]![2].startsWith("0") ? width : 0, "0")}${parts[0]![3]}`;
    }
  }
  return null;
}
