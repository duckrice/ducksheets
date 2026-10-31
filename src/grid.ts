// The sheet on screen: drawn on a canvas, so large sheets scroll smoothly.
// Excel's ways: text spills into empty cells to the right, numbers too wide
// show ###, numbers sit right and text left, frozen rows and columns stay put.
// A transparent scrolling layer on top takes the mouse and gives scrollbars.

import { Doc, selRange } from "./doc";
import { CellError, MAX_COLS, MAX_ROWS, colName, inRange, key, mergeAt, type Range } from "./model";
import { columnChars, formatValue } from "./numfmt";
import { lookOf, type BorderSide } from "./xlsx/styles";
import { conditionalEffects, readDxf, type CfEffect } from "./conditional";
import { tableLook } from "./structured";

const HEADER_H = 22;
const MIN_HEADER_W = 42;
const ACCENT = "#15803d";
const GRID_LINE = "#e2e5e9";

/** Positions along rows or columns, with most at the default size. */
class Axis {
  private keys: number[] = [];
  private prefix: number[] = [];

  constructor(
    private def: number,
    private sizes: Map<number, number>,
    private hidden: Set<number>,
  ) {
    this.keys = [...new Set([...sizes.keys(), ...hidden])].sort((a, b) => a - b);
    let sum = 0;
    for (const k of this.keys) {
      this.prefix.push(sum);
      sum += this.size(k) - def;
    }
    this.prefix.push(sum);
  }

  size(i: number): number {
    return this.hidden.has(i) ? 0 : (this.sizes.get(i) ?? this.def);
  }

  /** Where index i starts, from 0. */
  start(i: number): number {
    let lo = 0;
    let hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.def + this.prefix[lo];
  }

  /** The index at a position. */
  at(pos: number, max: number): number {
    let lo = 0;
    let hi = max - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.start(mid) <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

export interface GridHooks {
  /** Starts editing the active cell, with a typed character or the cell's text. */
  edit(initial?: string): void;
  selectionChanged(): void;
  /** Whether the editor wants clicks to insert references (a formula being typed). */
  pointing(): boolean;
  point(range: Range, sheet: number): void;
  /** Puts in what's being typed, if anything (a click elsewhere does that, as in Excel). */
  finishEditing(): void;
  /** A click (not a hold) on a cell: follows its link, if it has one. */
  follow(r: number, c: number): void;
  contextMenu(x: number, y: number, on: "cell" | "row" | "col"): void;
  filterMenu(col: number, x: number, y: number): void;
  /** Two fingers pinched the sheet bigger or smaller. */
  zoom(zoom: number): void;
}

type Drag =
  | { kind: "select"; r: number; c: number }
  | { kind: "rows"; r: number }
  | { kind: "cols"; c: number }
  | { kind: "fill"; source: Range; target: Range }
  | { kind: "resize-col"; c: number; start: number; size: number }
  | { kind: "resize-row"; r: number; start: number; size: number }
  | { kind: "point"; r: number; c: number };

export class Grid {
  el: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private scroller: HTMLElement;
  private sizer: HTMLElement;
  zoom = 1;
  private rows!: Axis;
  private cols!: Axis;
  private extentRows = 100;
  private extentCols = 26;
  private headerW = MIN_HEADER_W;
  private frame = 0;
  private drag: Drag | null = null;
  /** Coloured outlines of the cells a formula being edited refers to. */
  highlights: { range: Range; color: string }[] = [];
  /** The dashed outline around copied cells. */
  marquee: Range | null = null;
  private measures = new Map<string, number>();

  constructor(
    private doc: Doc,
    readonly hooks: GridHooks,
  ) {
    this.canvas = document.createElement("canvas");
    this.canvas.className = "grid-canvas";
    this.ctx = this.canvas.getContext("2d")!;
    this.sizer = document.createElement("div");
    this.sizer.className = "grid-sizer";
    this.scroller = document.createElement("div");
    this.scroller.className = "grid-scroller";
    this.scroller.append(this.sizer);
    this.el = document.createElement("div");
    this.el.className = "grid";
    this.el.tabIndex = 0;
    this.el.append(this.canvas, this.scroller);
    new ResizeObserver(() => this.layout()).observe(this.el);
    this.scroller.addEventListener("scroll", () => {
      this.growToScroll();
      this.render();
    });
    this.scroller.addEventListener("pointerdown", (e) => this.down(e));
    this.scroller.addEventListener("pointermove", (e) => this.move(e));
    this.scroller.addEventListener("pointerup", (e) => this.up(e));
    this.scroller.addEventListener("pointercancel", (e) => this.fingers.has(e.pointerId) && this.fingerUp(e));
    this.scroller.addEventListener("dblclick", (e) => this.doubleClick(e));
    this.scroller.addEventListener("contextmenu", (e) => this.context(e));
    this.refresh();
  }

  /** What conditional formatting does to the sheet's cells, worked out when first needed after a change. */
  private cf: Map<number, CfEffect> | null = null;
  /** Ctrl+`: formulas shown in the cells instead of their results. */
  showFormulas = false;

  get effects(): Map<number, CfEffect> {
    if (!this.cf) {
      const sheet = this.doc.sheet;
      const { styles } = this.doc.book;
      this.cf = sheet.conditional.length ? conditionalEffects(sheet, this.doc.active, this.doc.calc, styles.dxfs.map((x) => readDxf(x, styles.theme))) : new Map();
    }
    return this.cf;
  }

  // ---- sizes and positions ---------------------------------------------------------------------

  /** Picks up changes to the sheet: sizes, hidden rows, how far it reaches. */
  refresh() {
    this.cf = null;
    const sheet = this.doc.sheet;
    this.rows = new Axis(sheet.defaultRowHeight, sheet.rowHeights, sheet.hiddenRows);
    this.cols = new Axis(sheet.defaultColWidth, sheet.colWidths, sheet.hiddenCols);
    let usedR = 0;
    let usedC = 0;
    for (const k of sheet.cells.keys()) {
      usedR = Math.max(usedR, Math.floor(k / MAX_COLS) + 1);
      usedC = Math.max(usedC, (k % MAX_COLS) + 1);
    }
    this.extentRows = Math.min(MAX_ROWS, Math.max(this.extentRows, usedR + 50, this.doc.sel.r + 20));
    this.extentCols = Math.min(MAX_COLS, Math.max(this.extentCols, usedC + 10, this.doc.sel.c + 5));
    this.layout();
  }

  private layout() {
    const z = this.zoom;
    const digits = String(Math.min(MAX_ROWS, this.extentRows)).length;
    this.headerW = Math.max(MIN_HEADER_W, 14 + digits * 8);
    this.sizer.style.width = `${(this.headerW + this.cols.start(this.extentCols)) * z}px`;
    this.sizer.style.height = `${(HEADER_H + this.rows.start(this.extentRows)) * z}px`;
    const dpr = window.devicePixelRatio || 1;
    const { clientWidth: w, clientHeight: h } = this.el;
    if (this.canvas.width !== Math.round(w * dpr) || this.canvas.height !== Math.round(h * dpr)) {
      this.canvas.width = Math.round(w * dpr);
      this.canvas.height = Math.round(h * dpr);
      this.canvas.style.width = `${w}px`;
      this.canvas.style.height = `${h}px`;
    }
    this.render();
  }

  /** Scrolling near the end makes the sheet longer, as in Excel. */
  private growToScroll() {
    const z = this.zoom;
    const bottom = (this.scroller.scrollTop + this.scroller.clientHeight) / z;
    const right = (this.scroller.scrollLeft + this.scroller.clientWidth) / z;
    let grew = false;
    if (bottom > this.rows.start(this.extentRows) - 300 && this.extentRows < MAX_ROWS) {
      this.extentRows = Math.min(MAX_ROWS, this.extentRows + 200);
      grew = true;
    }
    if (right > this.cols.start(this.extentCols) - 300 && this.extentCols < MAX_COLS) {
      this.extentCols = Math.min(MAX_COLS, this.extentCols + 20);
      grew = true;
    }
    if (grew) this.layout();
  }

  private get frozen() {
    const f = this.doc.sheet.freeze;
    return { rows: f.rows, cols: f.cols, h: this.rows.start(f.rows), w: this.cols.start(f.cols) };
  }

  /** Where a row starts on screen, in unzoomed pixels (frozen rows don't scroll). */
  private y(r: number): number {
    const f = this.frozen;
    const top = this.rows.start(r);
    return r < f.rows ? HEADER_H + top : HEADER_H + f.h + (top - f.h) - this.scroller.scrollTop / this.zoom;
  }

  private x(c: number): number {
    const f = this.frozen;
    const left = this.cols.start(c);
    return c < f.cols ? this.headerW + left : this.headerW + f.w + (left - f.w) - this.scroller.scrollLeft / this.zoom;
  }

  /** The row and column under a point (unzoomed pixels inside the grid). */
  private cellAt(px: number, py: number): { r: number; c: number } {
    const f = this.frozen;
    const yy = py - HEADER_H;
    const xx = px - this.headerW;
    const r = yy < f.h ? this.rows.at(Math.max(0, yy), this.extentRows) : this.rows.at(yy + this.scroller.scrollTop / this.zoom, this.extentRows);
    const c = xx < f.w ? this.cols.at(Math.max(0, xx), this.extentCols) : this.cols.at(xx + this.scroller.scrollLeft / this.zoom, this.extentCols);
    return { r: Math.max(0, r), c: Math.max(0, c) };
  }

  /** Where the cells start, below and right of the headers, in CSS pixels. */
  get origin(): { x: number; y: number } {
    return { x: this.headerW * this.zoom, y: HEADER_H * this.zoom };
  }

  /** The cell under a point given in CSS pixels inside the grid (for placing charts). */
  cellUnder(x: number, y: number): { r: number; c: number } {
    return this.cellAt(Math.max(this.headerW + 1, x / this.zoom), Math.max(HEADER_H + 1, y / this.zoom));
  }

  /** A cell's box on screen in CSS pixels (merged cells as one), for the editor. */
  cellBox(r: number, c: number): { x: number; y: number; w: number; h: number } {
    const m = mergeAt(this.doc.sheet, r, c) ?? { r1: r, c1: c, r2: r, c2: c };
    const z = this.zoom;
    const x = this.x(m.c1);
    const y = this.y(m.r1);
    return { x: x * z, y: y * z, w: (this.x(m.c2) + this.cols.size(m.c2) - x) * z, h: (this.y(m.r2) + this.rows.size(m.r2) - y) * z };
  }

  /** Scrolls so a cell can be seen. */
  reveal(r: number, c: number) {
    const f = this.frozen;
    const z = this.zoom;
    const viewH = this.scroller.clientHeight / z - HEADER_H - f.h;
    const viewW = this.scroller.clientWidth / z - this.headerW - f.w;
    if (r >= f.rows) {
      const top = this.rows.start(r) - f.h;
      const bottom = top + this.rows.size(r);
      const st = this.scroller.scrollTop / z;
      if (top < st) this.scroller.scrollTop = top * z;
      else if (bottom > st + viewH) this.scroller.scrollTop = (bottom - viewH) * z;
    }
    if (c >= f.cols) {
      const left = this.cols.start(c) - f.w;
      const right = left + this.cols.size(c);
      const sl = this.scroller.scrollLeft / z;
      if (left < sl) this.scroller.scrollLeft = left * z;
      else if (right > sl + viewW) this.scroller.scrollLeft = (right - viewW) * z;
    }
    if (r + 20 > this.extentRows || c + 5 > this.extentCols) this.refresh();
  }

  /** How many rows fit in the view, for Page Up and Page Down. */
  get pageRows(): number {
    return Math.max(1, Math.floor((this.scroller.clientHeight / this.zoom - HEADER_H) / this.doc.sheet.defaultRowHeight) - 1);
  }

  setZoom(zoom: number) {
    this.zoom = Math.min(4, Math.max(0.25, Math.round(zoom * 100) / 100));
    this.measures.clear();
    this.layout();
  }

  // ---- drawing --------------------------------------------------------------------------------

  render() {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.draw();
    });
  }

  private draw() {
    const ctx = this.ctx;
    const dpr = window.devicePixelRatio || 1;
    const z = this.zoom;
    ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
    const w = this.el.clientWidth / z;
    const h = this.el.clientHeight / z;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, w, h);
    const f = this.frozen;
    const sy = this.scroller.scrollTop / z;
    const sx = this.scroller.scrollLeft / z;
    // The four panes: frozen corner, frozen rows, frozen columns, and the part that scrolls.
    const rowSpans = [
      { r0: 0, r1: f.rows - 1, top: HEADER_H, bottom: HEADER_H + f.h },
      { r0: this.rows.at(f.h + sy, this.extentRows), r1: this.rows.at(f.h + sy + h, this.extentRows), top: HEADER_H + f.h, bottom: h },
    ].filter((s) => s.r1 >= s.r0 && s.bottom > s.top);
    const colSpans = [
      { c0: 0, c1: f.cols - 1, left: this.headerW, right: this.headerW + f.w },
      { c0: this.cols.at(f.w + sx, this.extentCols), c1: this.cols.at(f.w + sx + w, this.extentCols), left: this.headerW + f.w, right: w },
    ].filter((s) => s.c1 >= s.c0 && s.right > s.left);
    for (const rs of rowSpans)
      for (const cs of colSpans) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cs.left, rs.top, cs.right - cs.left, rs.bottom - rs.top);
        ctx.clip();
        this.drawCells(rs.r0, rs.r1, cs.c0, cs.c1);
        this.drawSelection();
        ctx.restore();
      }
    this.drawHeaders(rowSpans, colSpans, w, h);
    // The lines that show where the frozen part ends.
    ctx.strokeStyle = "#9aa1a9";
    ctx.lineWidth = 1;
    if (f.rows) this.line(0, HEADER_H + f.h, w, HEADER_H + f.h);
    if (f.cols) this.line(this.headerW + f.w, 0, this.headerW + f.w, h);
  }

  private line(x1: number, y1: number, x2: number, y2: number) {
    const ctx = this.ctx;
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  }

  private drawCells(r0: number, r1: number, c0: number, c1: number) {
    const ctx = this.ctx;
    const sheet = this.doc.sheet;
    const styles = this.doc.book.styles;
    const visibleRows: number[] = [];
    for (let r = r0; r <= r1; r++) if (!sheet.hiddenRows.has(r)) visibleRows.push(r);
    const visibleCols: number[] = [];
    for (let c = c0; c <= c1; c++) if (!sheet.hiddenCols.has(c)) visibleCols.push(c);
    const left = this.x(c0);
    const right = this.x(c1) + this.cols.size(c1);
    const top = this.y(r0);
    const bottom = this.y(r1) + this.rows.size(r1);

    // Fills, then data bars, from the cell formats and conditional formatting.
    const effects = this.effects;
    for (const r of visibleRows)
      for (const c of visibleCols) {
        const cell = sheet.cells.get(key(r, c));
        const inTable = sheet.tables.length ? tableLook(sheet, r, c) : null;
        if (!cell && !inTable) continue;
        const effect = effects.get(key(r, c));
        const fill = effect?.fill ?? (cell?.s ? lookOf(styles, cell.s).fill : null) ?? inTable?.fill;
        if (fill) {
          ctx.fillStyle = fill;
          ctx.fillRect(this.x(c), this.y(r), this.cols.size(c), this.rows.size(r));
        }
        if (effect?.bar) {
          ctx.globalAlpha = 0.75;
          ctx.fillStyle = effect.bar.color;
          ctx.fillRect(this.x(c) + 2, this.y(r) + 3, Math.max(1, (this.cols.size(c) - 4) * effect.bar.frac), this.rows.size(r) - 6);
          ctx.globalAlpha = 1;
        }
        if (effect?.icon) {
          // At the left of the cell, as Excel draws them: red for the lowest third, yellow, green for the highest.
          const { kind, level } = effect.icon;
          const colour = level >= 0.67 ? "#3DA45C" : level >= 0.33 ? "#F5B800" : "#D9433A";
          const size = Math.min(12, this.rows.size(r) - 6);
          const x = this.x(c) + 3, y = this.y(r) + (this.rows.size(r) - size) / 2;
          ctx.fillStyle = colour;
          ctx.beginPath();
          if (kind === "arrow") {
            // Up, across or down, as the value is high, middling or low.
            const dir = level >= 0.67 ? -1 : level >= 0.33 ? 0 : 1;
            const cx = x + size / 2, cy = y + size / 2;
            if (dir === 0) ctx.moveTo(x + size, cy), ctx.lineTo(x + size / 2, y), ctx.lineTo(x + size / 2, y + size);
            else ctx.moveTo(cx, dir < 0 ? y : y + size), ctx.lineTo(x + size, cy), ctx.lineTo(x, cy);
          } else if (kind === "flag") ctx.rect(x + 2, y, size - 3, size / 2);
          else ctx.arc(x + size / 2, y + size / 2, size / 2, 0, Math.PI * 2);
          ctx.fill();
        }
      }
    // Gridlines.
    ctx.strokeStyle = GRID_LINE;
    ctx.lineWidth = 1 / this.zoom;
    ctx.beginPath();
    for (const r of visibleRows) {
      const y = Math.round((this.y(r) + this.rows.size(r)) * this.zoom) / this.zoom - 0.5 / this.zoom;
      ctx.moveTo(left, y);
      ctx.lineTo(right, y);
    }
    for (const c of visibleCols) {
      const x = Math.round((this.x(c) + this.cols.size(c)) * this.zoom) / this.zoom - 0.5 / this.zoom;
      ctx.moveTo(x, top);
      ctx.lineTo(x, bottom);
    }
    ctx.stroke();

    // Merged cells cover the gridlines inside them.
    const merges = sheet.merges.filter((m) => m.r2 >= r0 && m.r1 <= r1 && m.c2 >= c0 && m.c1 <= c1);
    for (const m of merges) {
      const box = this.box(m);
      const cell = sheet.cells.get(key(m.r1, m.c1));
      ctx.fillStyle = lookOf(styles, cell?.s).fill ?? "#ffffff";
      ctx.fillRect(box.x, box.y, box.w - 1 / this.zoom, box.h - 1 / this.zoom);
    }

    // Text.
    for (const r of visibleRows) {
      for (const c of visibleCols) {
        const m = mergeAt(sheet, r, c);
        if (m && (m.r1 !== r || m.c1 !== c)) continue;
        this.drawText(r, c, m, c1);
      }
    }
    // Merged cells that start above or left of the view still show their text.
    for (const m of merges) if (m.r1 < r0 || m.c1 < c0) this.drawText(m.r1, m.c1, m, c1);

    // A red corner on cells with a note, as in Excel.
    ctx.fillStyle = "#d32f2f";
    for (const k of sheet.notes.keys()) {
      const r = Math.floor(k / MAX_COLS);
      const c = k % MAX_COLS;
      if (r < r0 || r > r1 || c < c0 || c > c1 || sheet.hiddenRows.has(r) || sheet.hiddenCols.has(c)) continue;
      const box = this.cellBox(r, c);
      const right = (box.x + box.w) / this.zoom;
      const top = box.y / this.zoom;
      ctx.beginPath();
      ctx.moveTo(right - 7, top);
      ctx.lineTo(right - 1, top);
      ctx.lineTo(right - 1, top + 6);
      ctx.fill();
    }

    // Borders.
    for (const r of visibleRows)
      for (const c of visibleCols) {
        const cell = sheet.cells.get(key(r, c));
        if (!cell?.s) continue;
        const b = lookOf(styles, cell.s).border;
        if (!b.top && !b.bottom && !b.left && !b.right) continue;
        const x = this.x(c);
        const y = this.y(r);
        const w = this.cols.size(c);
        const h = this.rows.size(r);
        if (b.top) this.border(b.top, x, y, x + w, y);
        if (b.bottom) this.border(b.bottom, x, y + h, x + w, y + h);
        if (b.left) this.border(b.left, x, y, x, y + h);
        if (b.right) this.border(b.right, x + w, y, x + w, y + h);
      }

    // Filter buttons on the heading row.
    const filter = sheet.filter;
    if (filter && filter.range.r1 >= r0 && filter.range.r1 <= r1) {
      for (let c = Math.max(c0, filter.range.c1); c <= Math.min(c1, filter.range.c2); c++) {
        const bx = this.x(c) + this.cols.size(c) - 17;
        const by = this.y(filter.range.r1) + this.rows.size(filter.range.r1) - 17;
        const active = filter.shown.has(c - filter.range.c1) || !!filter.custom?.has(c - filter.range.c1);
        ctx.fillStyle = active ? "#dcfce7" : "#f3f4f6";
        ctx.strokeStyle = active ? ACCENT : "#9aa1a9";
        ctx.lineWidth = 1;
        ctx.fillRect(bx, by, 15, 15);
        ctx.strokeRect(bx + 0.5, by + 0.5, 14, 14);
        ctx.fillStyle = active ? ACCENT : "#4b5563";
        ctx.beginPath();
        ctx.moveTo(bx + 4, by + 6);
        ctx.lineTo(bx + 11, by + 6);
        ctx.lineTo(bx + 7.5, by + 10);
        ctx.fill();
      }
    }
  }

  private box(m: Range) {
    const x = this.x(m.c1);
    const y = this.y(m.r1);
    return { x, y, w: this.x(m.c2) + this.cols.size(m.c2) - x, h: this.y(m.r2) + this.rows.size(m.r2) - y };
  }

  private border(side: BorderSide, x1: number, y1: number, x2: number, y2: number) {
    const ctx = this.ctx;
    const width = /thick/.test(side.style) ? 3 : /medium/.test(side.style) ? 2 : 1;
    ctx.strokeStyle = side.color ?? "#000000";
    ctx.lineWidth = width;
    ctx.setLineDash(/dash/i.test(side.style) ? [4, 2] : /dot/i.test(side.style) || side.style === "hair" ? [1, 1] : []);
    if (side.style === "double") {
      ctx.lineWidth = 1;
      const dx = y1 === y2 ? 0 : 1;
      const dy = y1 === y2 ? 1 : 0;
      this.line(x1 - dx, y1 - dy, x2 - dx, y2 - dy);
      this.line(x1 + dx, y1 + dy, x2 + dx, y2 + dy);
    } else {
      const off = width % 2 ? 0.5 : 0;
      this.line(x1 - (y1 === y2 ? 0 : off), y1 - (y1 === y2 ? off : 0), x2 - (y1 === y2 ? 0 : off), y2 - (y1 === y2 ? off : 0));
    }
    ctx.setLineDash([]);
  }

  private measure(font: string, text: string): number {
    const id = `${font}|${text}`;
    let w = this.measures.get(id);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.measures.size > 20000) this.measures.clear();
      this.measures.set(id, w);
    }
    return w;
  }

  private drawText(r: number, c: number, merge: Range | undefined, lastCol: number) {
    const sheet = this.doc.sheet;
    const cell = sheet.cells.get(key(r, c));
    const value = this.doc.value(r, c);
    if (value === null) return;
    const look = lookOf(this.doc.book.styles, cell?.s);
    // A "General" number fits its cell, as in Excel (fewer decimals in a narrow column).
    const chars = columnChars(merge ? this.x(merge.c2) + this.cols.size(merge.c2) - this.x(merge.c1) : this.cols.size(c));
    const shown = this.showFormulas && cell?.f !== undefined ? { text: `=${cell.f}` } as ReturnType<typeof formatValue> : formatValue(value, look.numFmt, this.doc.book.date1904, chars);
    if (!shown.text) return;
    const ctx = this.ctx;
    const effect = this.effects.get(key(r, c));
    // Links show blue and underlined, as Excel's Hyperlink style.
    const linked = sheet.links.has(key(r, c)) || (!!cell?.f && /^\s*HYPERLINK\s*\(/i.test(cell.f));
    // A table's header is white and bold, unless the cell has its own colour.
    const inTable = sheet.tables.length ? tableLook(sheet, r, c) : null;
    const base = inTable ? { ...look.font, bold: look.font.bold || !!inTable.bold, color: inTable.color && (!look.font.color || look.font.color === "#000000") && !look.fill ? inTable.color : look.font.color } : look.font;
    const styled = effect ? { ...base, ...(effect.bold ? { bold: true } : {}), ...(effect.italic ? { italic: true } : {}) } : base;
    const font = linked ? { ...styled, underline: true, color: styled.color && styled.color !== "#000000" ? styled.color : "#0563C1" } : styled;
    if (linked && !effect?.color) shown.color = font.color;
    if (effect?.color) shown.color = effect.color;
    const size = (font.size * 4) / 3;
    const css = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${size}px "${font.name}", "Segoe UI", sans-serif`;
    const box = merge ? this.box(merge) : { x: this.x(c), y: this.y(r), w: this.cols.size(c), h: this.rows.size(r) };
    const numeric = typeof value === "number";
    const align = look.h && look.h !== "general" ? look.h : numeric ? "right" : typeof value === "boolean" || value instanceof CellError ? "center" : "left";
    const pad = 3 + (look.indent ?? 0) * 9;
    let text = shown.text;
    let width = this.measure(css, text);
    let clipRight = box.x + box.w;
    let clipLeft = box.x;
    if (look.wrap && !numeric) return this.drawWrapped(text, css, size, box, look, shown.color ?? font.color ?? "#000000", pad, align);
    if (width > box.w - pad * 2) {
      if (numeric) {
        // A number that doesn't fit shows ###, as in Excel.
        const hash = this.measure(css, "#");
        text = "#".repeat(Math.max(1, Math.floor((box.w - 4) / hash)));
        width = this.measure(css, text);
      } else if (!merge && align === "left") {
        // Text spills into the empty cells to its right.
        let cc = c + 1;
        while (clipRight - box.x < width + pad * 2 && cc <= lastCol + 20 && cc < MAX_COLS) {
          if (this.doc.value(r, cc) !== null || mergeAt(sheet, r, cc)) break;
          clipRight += this.cols.size(cc);
          cc++;
        }
        if (clipRight > box.x + box.w) {
          // Cover the gridlines the text runs over.
          ctx.fillStyle = look.fill ?? "#ffffff";
          ctx.fillRect(box.x + box.w - 1, box.y, Math.min(clipRight - box.x - box.w, width + pad * 2 - box.w + 2), box.h - 1);
        }
      } else if (!merge && align === "center") {
        clipLeft = box.x - (width - box.w) / 2 - pad;
        clipRight = box.x + box.w + (width - box.w) / 2 + pad;
      }
    }
    const tx = align === "right" ? box.x + box.w - pad - width : align === "center" ? box.x + (box.w - width) / 2 : box.x + pad;
    const ty = look.v === "top" ? box.y + size + 2 : look.v === "center" ? box.y + box.h / 2 + size * 0.35 : box.y + box.h - 4;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clipLeft, box.y, clipRight - clipLeft, box.h);
    ctx.clip();
    ctx.font = css;
    ctx.fillStyle = shown.color ?? font.color ?? "#000000";
    ctx.textBaseline = "alphabetic";
    ctx.fillText(text, tx, ty);
    this.decorate(font, tx, ty, width, size);
    ctx.restore();
  }

  private decorate(font: { underline?: boolean; strike?: boolean }, x: number, y: number, width: number, size: number) {
    const ctx = this.ctx;
    ctx.strokeStyle = ctx.fillStyle as string;
    ctx.lineWidth = Math.max(1, size / 14);
    if (font.underline) this.line(x, y + 2, x + width, y + 2);
    if (font.strike) this.line(x, y - size * 0.3, x + width, y - size * 0.3);
  }

  private drawWrapped(text: string, css: string, size: number, box: { x: number; y: number; w: number; h: number }, look: { v?: string; font: { underline?: boolean; strike?: boolean } }, color: string, pad: number, align: string) {
    const ctx = this.ctx;
    const max = box.w - pad * 2;
    const lines: string[] = [];
    for (const para of text.split("\n")) {
      let line = "";
      for (const word of para.split(" ")) {
        const next = line ? `${line} ${word}` : word;
        if (line && this.measure(css, next) > max) {
          lines.push(line);
          line = word;
        } else line = next;
      }
      lines.push(line);
    }
    const lineH = size * 1.25;
    const total = lines.length * lineH;
    let y = look.v === "top" ? box.y + size + 2 : look.v === "center" ? box.y + (box.h - total) / 2 + size : box.y + box.h - 4 - total + lineH;
    ctx.save();
    ctx.beginPath();
    ctx.rect(box.x, box.y, box.w, box.h);
    ctx.clip();
    ctx.font = css;
    ctx.fillStyle = color;
    for (const line of lines) {
      const w = this.measure(css, line);
      const x = align === "right" ? box.x + box.w - pad - w : align === "center" ? box.x + (box.w - w) / 2 : box.x + pad;
      ctx.fillText(line, x, y);
      this.decorate(look.font, x, y, w, size);
      y += lineH;
    }
    ctx.restore();
  }

  /** The selection, references being pointed at, and the copy outline. */
  private drawSelection() {
    const ctx = this.ctx;
    const sheet = this.doc.sheet;
    const range = this.expanded(selRange(this.doc.sel));
    const b = this.box(range);
    const z = this.zoom;
    ctx.fillStyle = "rgba(21, 128, 61, 0.12)";
    ctx.fillRect(b.x, b.y, b.w, b.h);
    // The active cell stays clear.
    const active = mergeAt(sheet, this.doc.sel.r, this.doc.sel.c) ?? { r1: this.doc.sel.r, c1: this.doc.sel.c, r2: this.doc.sel.r, c2: this.doc.sel.c };
    const a = this.box(active);
    const fill = lookOf(this.doc.book.styles, sheet.cells.get(key(active.r1, active.c1))?.s).fill;
    if (range.r1 !== range.r2 || range.c1 !== range.c2) {
      // The active cell stays clear of the tint: its background and text again.
      ctx.save();
      ctx.beginPath();
      ctx.rect(a.x, a.y, a.w - 1 / z, a.h - 1 / z);
      ctx.clip();
      ctx.fillStyle = fill ?? "#ffffff";
      ctx.fillRect(a.x, a.y, a.w, a.h);
      this.drawText(active.r1, active.c1, mergeAt(sheet, active.r1, active.c1), active.c2);
      ctx.restore();
    }
    ctx.strokeStyle = ACCENT;
    ctx.lineWidth = 2;
    ctx.strokeRect(b.x, b.y, b.w - 1 / z, b.h - 1 / z);
    // The fill handle.
    ctx.fillStyle = ACCENT;
    ctx.fillRect(b.x + b.w - 4, b.y + b.h - 4, 6, 6);
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 1;
    ctx.strokeRect(b.x + b.w - 4.5, b.y + b.h - 4.5, 7, 7);

    if (this.drag?.kind === "fill") {
      const t = this.box(this.drag.target);
      ctx.strokeStyle = "#6b7280";
      ctx.setLineDash([3, 3]);
      ctx.strokeRect(t.x + 0.5, t.y + 0.5, t.w - 1, t.h - 1);
      ctx.setLineDash([]);
    }
    for (const h of this.highlights) {
      const hb = this.box(h.range);
      ctx.strokeStyle = h.color;
      ctx.lineWidth = 2;
      ctx.strokeRect(hb.x + 1, hb.y + 1, hb.w - 2, hb.h - 2);
      ctx.fillStyle = `${h.color}22`;
      ctx.fillRect(hb.x, hb.y, hb.w, hb.h);
    }
    if (this.marquee) {
      const mb = this.box(this.marquee);
      ctx.strokeStyle = ACCENT;
      ctx.lineWidth = 1.5;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(mb.x + 1, mb.y + 1, mb.w - 2, mb.h - 2);
      ctx.setLineDash([]);
    }
  }

  /** A range grown to take in any merged cells it cuts through, as Excel's selection does. */
  expanded(range: Range): Range {
    const out = { ...range };
    for (let grew = true; grew; ) {
      grew = false;
      for (const m of this.doc.sheet.merges) {
        if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
        if (m.r1 < out.r1 || m.r2 > out.r2 || m.c1 < out.c1 || m.c2 > out.c2) {
          out.r1 = Math.min(out.r1, m.r1);
          out.r2 = Math.max(out.r2, m.r2);
          out.c1 = Math.min(out.c1, m.c1);
          out.c2 = Math.max(out.c2, m.c2);
          grew = true;
        }
      }
    }
    return out;
  }

  private drawHeaders(rowSpans: { r0: number; r1: number; top: number; bottom: number }[], colSpans: { c0: number; c1: number; left: number; right: number }[], w: number, h: number) {
    const ctx = this.ctx;
    const sel = selRange(this.doc.sel);
    const whole = { rows: sel.c1 === 0 && sel.c2 >= MAX_COLS - 1, cols: sel.r1 === 0 && sel.r2 >= MAX_ROWS - 1 };
    ctx.fillStyle = "#f3f4f6";
    ctx.fillRect(0, 0, w, HEADER_H);
    ctx.fillRect(0, 0, this.headerW, h);
    ctx.font = `13px "Segoe UI", sans-serif`;
    ctx.textBaseline = "middle";
    ctx.textAlign = "center";
    for (const cs of colSpans) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(cs.left, 0, cs.right - cs.left, HEADER_H);
      ctx.clip();
      for (let c = cs.c0; c <= cs.c1; c++) {
        const size = this.cols.size(c);
        if (!size) continue;
        const x = this.x(c);
        const on = c >= sel.c1 && c <= sel.c2;
        if (on) {
          ctx.fillStyle = whole.cols ? "#15803d" : "#d1e7d6";
          ctx.fillRect(x, 0, size, HEADER_H);
        }
        ctx.fillStyle = on ? (whole.cols ? "#ffffff" : "#14532d") : "#4b5563";
        ctx.fillText(colName(c), x + size / 2, HEADER_H / 2 + 1);
        ctx.strokeStyle = "#d4d8dd";
        ctx.lineWidth = 1;
        this.line(x + size - 0.5, 0, x + size - 0.5, HEADER_H);
        if (on && !whole.cols) {
          ctx.strokeStyle = ACCENT;
          ctx.lineWidth = 2;
          this.line(x, HEADER_H - 1, x + size, HEADER_H - 1);
        }
      }
      ctx.restore();
    }
    for (const rs of rowSpans) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, rs.top, this.headerW, rs.bottom - rs.top);
      ctx.clip();
      for (let r = rs.r0; r <= rs.r1; r++) {
        const size = this.rows.size(r);
        if (!size) continue;
        const y = this.y(r);
        const on = r >= sel.r1 && r <= sel.r2;
        if (on) {
          ctx.fillStyle = whole.rows ? "#15803d" : "#d1e7d6";
          ctx.fillRect(0, y, this.headerW, size);
        }
        ctx.fillStyle = on ? (whole.rows ? "#ffffff" : "#14532d") : "#4b5563";
        ctx.fillText(String(r + 1), this.headerW / 2, y + size / 2 + 1);
        ctx.strokeStyle = "#d4d8dd";
        ctx.lineWidth = 1;
        this.line(0, y + size - 0.5, this.headerW, y + size - 0.5);
        if (on && !whole.rows) {
          ctx.strokeStyle = ACCENT;
          ctx.lineWidth = 2;
          this.line(this.headerW - 1, y, this.headerW - 1, y + size);
        }
      }
      ctx.restore();
    }
    ctx.textAlign = "start";
    ctx.strokeStyle = "#c5cad1";
    ctx.lineWidth = 1;
    this.line(0, HEADER_H - 0.5, w, HEADER_H - 0.5);
    this.line(this.headerW - 0.5, 0, this.headerW - 0.5, h);
    // The corner: a triangle that selects everything.
    ctx.fillStyle = "#b8bec6";
    ctx.beginPath();
    ctx.moveTo(this.headerW - 4, HEADER_H - 4);
    ctx.lineTo(this.headerW - 4, HEADER_H - 14);
    ctx.lineTo(this.headerW - 14, HEADER_H - 4);
    ctx.fill();
  }

  // ---- the mouse ---------------------------------------------------------------------------------

  /** The cell under the mouse, if it's over a cell. */
  cellAtMouse(e: MouseEvent): { r: number; c: number } | null {
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    return hit.kind === "cell" ? { r: hit.r, c: hit.c } : null;
  }

  private local(e: MouseEvent) {
    const rect = this.el.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / this.zoom, y: (e.clientY - rect.top) / this.zoom };
  }

  /** What's under the mouse: a cell, a header, a header edge (to resize), the fill handle or a filter button. */
  private hit(x: number, y: number) {
    if (x < this.headerW && y < HEADER_H) return { kind: "corner" as const };
    if (y < HEADER_H) {
      const { c } = this.cellAt(x, HEADER_H + 1);
      const right = this.x(c) + this.cols.size(c);
      if (Math.abs(x - right) <= 4) return { kind: "col-edge" as const, c };
      const left = this.x(c);
      if (Math.abs(x - left) <= 4 && c > 0) return { kind: "col-edge" as const, c: c - 1 };
      return { kind: "col" as const, c };
    }
    if (x < this.headerW) {
      const { r } = this.cellAt(this.headerW + 1, y);
      const bottom = this.y(r) + this.rows.size(r);
      if (Math.abs(y - bottom) <= 3) return { kind: "row-edge" as const, r };
      if (Math.abs(y - this.y(r)) <= 3 && r > 0) return { kind: "row-edge" as const, r: r - 1 };
      return { kind: "row" as const, r };
    }
    const { r, c } = this.cellAt(x, y);
    const sel = this.box(this.expanded(selRange(this.doc.sel)));
    if (Math.abs(x - (sel.x + sel.w)) <= 5 && Math.abs(y - (sel.y + sel.h)) <= 5) return { kind: "fill" as const, r, c };
    const filter = this.doc.sheet.filter;
    if (filter && r === filter.range.r1 && c >= filter.range.c1 && c <= filter.range.c2) {
      const bx = this.x(c) + this.cols.size(c) - 17;
      const by = this.y(r) + this.rows.size(r) - 17;
      if (x >= bx && x <= bx + 15 && y >= by && y <= by + 15) return { kind: "filter" as const, r, c };
    }
    return { kind: "cell" as const, r, c };
  }

  private setSel(r: number, c: number, ar = r, ac = c) {
    this.doc.sel = { r, c, ar, ac };
    this.hooks.selectionChanged();
    this.render();
  }

  /** When the mouse button went down, to tell a click from a hold. */
  private downAt = 0;
  /** Fingers on the sheet: one drags it along (a tap selects), two pinch to zoom. */
  private fingers = new Map<number, { x: number; y: number }>();
  private pan: { x: number; y: number; top: number; left: number; moved: boolean; vx: number; vy: number; t: number } | null = null;
  private pinch: { distance: number; zoom: number } | null = null;
  private glide = 0;

  private spread() {
    const [a, b] = [...this.fingers.values()];
    return Math.hypot(a.x - b.x, a.y - b.y) || 1;
  }

  private fingerDown(e: PointerEvent) {
    cancelAnimationFrame(this.glide);
    this.fingers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    this.scroller.setPointerCapture(e.pointerId);
    e.preventDefault();
    if (this.fingers.size === 2) {
      this.pan = null;
      this.pinch = { distance: this.spread(), zoom: this.zoom };
    } else if (this.fingers.size === 1) {
      this.pan = { x: e.clientX, y: e.clientY, top: this.scroller.scrollTop, left: this.scroller.scrollLeft, moved: false, vx: 0, vy: 0, t: e.timeStamp };
    }
  }

  private fingerMove(e: PointerEvent) {
    const last = this.fingers.get(e.pointerId)!;
    this.fingers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    if (this.pinch && this.fingers.size === 2) {
      this.hooks.zoom(this.pinch.zoom * (this.spread() / this.pinch.distance));
      return;
    }
    const pan = this.pan;
    if (!pan) return;
    if (Math.hypot(e.clientX - pan.x, e.clientY - pan.y) > 8) pan.moved = true;
    if (!pan.moved) return;
    this.scroller.scrollTop = pan.top - (e.clientY - pan.y);
    this.scroller.scrollLeft = pan.left - (e.clientX - pan.x);
    // How fast the finger went at the end, to let the sheet glide on after it lets go.
    const dt = Math.max(1, e.timeStamp - pan.t);
    pan.vx = 0.8 * ((last.x - e.clientX) / dt) + 0.2 * pan.vx;
    pan.vy = 0.8 * ((last.y - e.clientY) / dt) + 0.2 * pan.vy;
    pan.t = e.timeStamp;
  }

  private fingerUp(e: PointerEvent) {
    this.fingers.delete(e.pointerId);
    if (this.scroller.hasPointerCapture(e.pointerId)) this.scroller.releasePointerCapture(e.pointerId);
    const pan = this.pan;
    if (this.pinch) {
      if (this.fingers.size < 2) this.pinch = null;
      return;
    }
    this.pan = null;
    if (!pan || e.type === "pointercancel") return;
    if (!pan.moved) return this.tap(e);
    let { vx, vy } = pan;
    let t = performance.now();
    const step = (now: number) => {
      const dt = now - t;
      t = now;
      this.scroller.scrollTop += vy * dt;
      this.scroller.scrollLeft += vx * dt;
      vx *= Math.pow(0.995, dt);
      vy *= Math.pow(0.995, dt);
      if (Math.hypot(vx, vy) > 0.02) this.glide = requestAnimationFrame(step);
    };
    if (Math.hypot(vx, vy) > 0.1) this.glide = requestAnimationFrame(step);
  }

  /** A tap with a finger selects, as a click does. */
  private tap(e: PointerEvent) {
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    this.el.focus({ preventScroll: true });
    this.hooks.finishEditing();
    if (hit.kind === "cell") {
      const m = mergeAt(this.doc.sheet, hit.r, hit.c);
      this.setSel(m?.r1 ?? hit.r, m?.c1 ?? hit.c);
    } else if (hit.kind === "row") this.setSel(hit.r, 0, hit.r, MAX_COLS - 1);
    else if (hit.kind === "col") this.setSel(0, hit.c, MAX_ROWS - 1, hit.c);
    else if (hit.kind === "filter") this.hooks.filterMenu(hit.c, e.clientX, e.clientY);
  }

  private down(e: PointerEvent) {
    if (e.pointerType === "touch") return this.fingerDown(e);
    if (e.button !== 0) return;
    this.downAt = Date.now();
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    const sel = this.doc.sel;
    this.el.focus({ preventScroll: true });
    // While a formula is being typed, clicking a cell puts its address in.
    if (this.hooks.pointing() && (hit.kind === "cell" || hit.kind === "fill")) {
      e.preventDefault();
      this.drag = { kind: "point", r: hit.r, c: hit.c };
      this.hooks.point({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c }, this.doc.active);
      this.scroller.setPointerCapture(e.pointerId);
      return;
    }
    // Clicking anywhere else on the sheet finishes what's being typed, as in Excel; otherwise
    // the keys would still go to the cell that was being typed in.
    this.hooks.finishEditing();
    switch (hit.kind) {
      case "corner":
        this.setSel(0, 0, MAX_ROWS - 1, MAX_COLS - 1);
        return;
      case "col-edge":
        this.drag = { kind: "resize-col", c: hit.c, start: x, size: this.cols.size(hit.c) };
        break;
      case "row-edge":
        this.drag = { kind: "resize-row", r: hit.r, start: y, size: this.rows.size(hit.r) };
        break;
      case "col":
        if (e.shiftKey) this.setSel(0, sel.c, MAX_ROWS - 1, hit.c);
        else this.setSel(0, hit.c, MAX_ROWS - 1, hit.c);
        this.drag = { kind: "cols", c: e.shiftKey ? sel.c : hit.c };
        break;
      case "row":
        if (e.shiftKey) this.setSel(sel.r, 0, hit.r, MAX_COLS - 1);
        else this.setSel(hit.r, 0, hit.r, MAX_COLS - 1);
        this.drag = { kind: "rows", r: e.shiftKey ? sel.r : hit.r };
        break;
      case "fill":
        this.drag = { kind: "fill", source: this.expanded(selRange(sel)), target: this.expanded(selRange(sel)) };
        break;
      case "filter":
        this.hooks.filterMenu(hit.c, e.clientX, e.clientY);
        return;
      case "cell":
        if (e.shiftKey) this.setSel(sel.r, sel.c, hit.r, hit.c);
        else {
          const m = mergeAt(this.doc.sheet, hit.r, hit.c);
          this.setSel(m?.r1 ?? hit.r, m?.c1 ?? hit.c);
        }
        this.drag = { kind: "select", r: e.shiftKey ? sel.r : hit.r, c: e.shiftKey ? sel.c : hit.c };
        break;
    }
    this.scroller.setPointerCapture(e.pointerId);
    e.preventDefault();
  }

  private move(e: PointerEvent) {
    if (this.fingers.has(e.pointerId)) return this.fingerMove(e);
    const { x, y } = this.local(e);
    if (!this.drag) {
      const hit = this.hit(x, y);
      this.scroller.style.cursor =
        hit.kind === "col-edge" ? "col-resize" : hit.kind === "row-edge" ? "row-resize" : hit.kind === "fill" ? "crosshair" : hit.kind === "col" ? "s-resize" : hit.kind === "row" ? "e-resize" : hit.kind === "filter" ? "pointer" : "cell";
      return;
    }
    const drag = this.drag;
    const { r, c } = this.cellAt(Math.max(this.headerW + 1, x), Math.max(HEADER_H + 1, y));
    switch (drag.kind) {
      case "select":
        this.setSel(drag.r, drag.c, r, c);
        this.reveal(r, c);
        break;
      case "rows":
        this.setSel(drag.r, 0, r, MAX_COLS - 1);
        break;
      case "cols":
        this.setSel(0, drag.c, MAX_ROWS - 1, c);
        break;
      case "point":
        this.hooks.point({ r1: Math.min(drag.r, r), c1: Math.min(drag.c, c), r2: Math.max(drag.r, r), c2: Math.max(drag.c, c) }, this.doc.active);
        break;
      case "fill": {
        // Along the direction the mouse moved furthest, as in Excel.
        const s = drag.source;
        const dr = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0;
        const dc = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0;
        drag.target = Math.abs(dr) >= Math.abs(dc) ? { ...s, r1: Math.min(s.r1, r), r2: Math.max(s.r2, r) } : { ...s, c1: Math.min(s.c1, c), c2: Math.max(s.c2, c) };
        this.render();
        break;
      }
      case "resize-col":
        this.liveSize("c", drag.c, Math.max(0, drag.size + x - drag.start));
        break;
      case "resize-row":
        this.liveSize("r", drag.r, Math.max(0, drag.size + y - drag.start));
        break;
    }
  }

  /** Shows a new size while dragging, before it's a change that can be undone. */
  private liveSize(axis: "r" | "c", i: number, size: number) {
    const sheet = this.doc.sheet;
    const map = new Map(axis === "r" ? sheet.rowHeights : sheet.colWidths).set(i, size);
    if (axis === "r") this.rows = new Axis(sheet.defaultRowHeight, map, sheet.hiddenRows);
    else this.cols = new Axis(sheet.defaultColWidth, map, sheet.hiddenCols);
    this.render();
  }

  private up(e: PointerEvent) {
    if (this.fingers.has(e.pointerId)) return this.fingerUp(e);
    const drag = this.drag;
    this.drag = null;
    if (this.scroller.hasPointerCapture(e.pointerId)) this.scroller.releasePointerCapture(e.pointerId);
    if (!drag) return;
    // As in Excel, a click on a link follows it; holding the button a moment just selects the cell.
    if (drag.kind === "select" && Date.now() - this.downAt < 500 && !e.shiftKey) {
      const s = this.doc.sel;
      if (s.r === s.ar && s.c === s.ac && s.r === drag.r && s.c === drag.c) this.hooks.follow(s.r, s.c);
    }
    const sel = selRange(this.doc.sel);
    if (drag.kind === "fill") {
      const s = drag.source;
      const t = drag.target;
      if (t.r1 !== s.r1 || t.r2 !== s.r2 || t.c1 !== s.c1 || t.c2 !== s.c2) {
        this.doc.fill(s, t);
        this.doc.sel = { r: t.r1, c: t.c1, ar: t.r2, ac: t.c2 };
        this.hooks.selectionChanged();
      }
      this.render();
    } else if (drag.kind === "resize-col") {
      const size = this.cols.size(drag.c);
      // Dragging the edge of a selected column sizes all the selected columns.
      const whole = sel.r1 === 0 && sel.r2 >= MAX_ROWS - 1 && drag.c >= sel.c1 && drag.c <= sel.c2;
      const cols = whole ? Array.from({ length: Math.min(sel.c2, sel.c1 + 500) - sel.c1 + 1 }, (_, i) => sel.c1 + i) : [drag.c];
      this.doc.setColWidth(cols, size);
    } else if (drag.kind === "resize-row") {
      const size = this.rows.size(drag.r);
      const whole = sel.c1 === 0 && sel.c2 >= MAX_COLS - 1 && drag.r >= sel.r1 && drag.r <= sel.r2;
      const rows = whole ? Array.from({ length: Math.min(sel.r2, sel.r1 + 5000) - sel.r1 + 1 }, (_, i) => sel.r1 + i) : [drag.r];
      this.doc.setRowHeight(rows, size);
    }
  }

  private doubleClick(e: MouseEvent) {
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    if (hit.kind === "col-edge") this.doc.setColWidth([hit.c], this.fitWidth(hit.c));
    else if (hit.kind === "row-edge") this.doc.setRowHeight([hit.r], this.doc.sheet.defaultRowHeight);
    else if (hit.kind === "cell") this.hooks.edit();
  }

  /** The width that fits a column's widest value (double-clicking its edge). */
  fitWidth(c: number): number {
    let widest = 0;
    for (const k of this.doc.sheet.cells.keys()) {
      if (k % MAX_COLS !== c) continue;
      const r = Math.floor(k / MAX_COLS);
      if (mergeAt(this.doc.sheet, r, c)) continue;
      const look = lookOf(this.doc.book.styles, this.doc.sheet.cells.get(k)?.s);
      const css = `${look.font.italic ? "italic " : ""}${look.font.bold ? "bold " : ""}${(look.font.size * 4) / 3}px "${look.font.name}", sans-serif`;
      for (const line of this.doc.shown(r, c).split("\n")) widest = Math.max(widest, this.measure(css, line));
    }
    return widest ? Math.ceil(widest + 10) : this.doc.sheet.defaultColWidth;
  }

  private context(e: MouseEvent) {
    e.preventDefault();
    const { x, y } = this.local(e);
    const hit = this.hit(x, y);
    const sel = selRange(this.doc.sel);
    if (hit.kind === "cell" && !inRange(sel, hit.r, hit.c)) this.setSel(hit.r, hit.c);
    if (hit.kind === "row" && !(hit.r >= sel.r1 && hit.r <= sel.r2 && sel.c1 === 0)) this.setSel(hit.r, 0, hit.r, MAX_COLS - 1);
    if (hit.kind === "col" && !(hit.c >= sel.c1 && hit.c <= sel.c2 && sel.r1 === 0)) this.setSel(0, hit.c, MAX_ROWS - 1, hit.c);
    this.hooks.contextMenu(e.clientX, e.clientY, hit.kind === "row" ? "row" : hit.kind === "col" ? "col" : "cell");
  }
}
