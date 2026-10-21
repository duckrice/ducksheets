// Typing into cells: over the cell itself, mirrored in the formula bar, as in
// Excel. While a formula is typed, clicking cells (or using the arrow keys)
// puts their addresses in, each reference outlined in its own colour on the
// sheet. F4 cycles A1 → $A$1 → A$1 → $A1. Function names are suggested.

import { selRange, type Doc } from "./doc";
import { FUNCTION_NAMES } from "./formula/functions";
import { quoteSheet } from "./formula/parse";
import type { Grid } from "./grid";
import { MAX_COLS, MAX_ROWS, address, colIndex, mergeAt, parseRange, rangeAddress, type Range } from "./model";
import { lookOf } from "./xlsx/styles";

const COLORS = ["#2563eb", "#dc2626", "#9333ea", "#ea580c", "#0891b2", "#be185d"];
/** After these, a cell address can come next. */
const BEFORE_REF = /[=(,;+\-*/^&<>:%]\s*$/;

export type Move = "down" | "up" | "right" | "left" | "none";

export interface EditorHooks {
  commit(r: number, c: number, text: string, move: Move, everywhere: boolean): void;
  /** Editing started or stopped (the toolbar dims what can't be used while typing). */
  changed(): void;
}

export class Editor {
  readonly bar: HTMLElement;
  private nameBox: HTMLInputElement;
  private barInput: HTMLTextAreaElement;
  private box: HTMLTextAreaElement;
  private list: HTMLElement;
  /** The cell being edited; "enter" means arrow keys move on (or point), "edit" means they move the caret. */
  editing: { r: number; c: number; sheet: number; mode: "enter" | "edit" } | null = null;
  /** The address last put in by pointing, so pointing again replaces it. */
  private pointed: { start: number; end: number; r: number; c: number } | null = null;
  private suggestions: string[] = [];
  private chosen = 0;

  constructor(
    private doc: Doc,
    private grid: Grid,
    private hooks: EditorHooks,
  ) {
    this.nameBox = Object.assign(document.createElement("input"), { className: "name-box", spellcheck: false, title: "Go to a cell: type an address like B12 and press Enter" });
    this.nameBox.setAttribute("aria-label", "Cell address");
    const fx = Object.assign(document.createElement("span"), { className: "fx", textContent: "fx" });
    this.barInput = Object.assign(document.createElement("textarea"), { className: "bar-input", rows: 1, spellcheck: false });
    this.barInput.setAttribute("aria-label", "Formula bar");
    this.bar = document.createElement("div");
    this.bar.className = "formula-bar";
    this.bar.append(this.nameBox, fx, this.barInput);
    this.box = Object.assign(document.createElement("textarea"), { className: "cell-editor", spellcheck: false, hidden: true });
    this.box.setAttribute("aria-label", "Cell");
    this.list = Object.assign(document.createElement("div"), { className: "suggestions", hidden: true });
    grid.el.append(this.box, this.list);

    for (const input of [this.box, this.barInput]) {
      input.addEventListener("input", () => this.typed(input));
      input.addEventListener("keydown", (e) => this.key(e, input));
      input.addEventListener("click", () => {
        if (this.editing) this.editing.mode = "edit";
        this.pointed = null;
      });
    }
    this.barInput.addEventListener("focus", () => {
      if (!this.editing) this.start(undefined, "edit", true);
    });
    this.nameBox.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        this.goTo(this.nameBox.value);
      } else if (e.key === "Escape") {
        this.show();
        this.grid.el.focus();
      }
    });
    this.nameBox.addEventListener("focus", () => this.nameBox.select());
  }

  /** The formula bar and name box for the selected cell, when not editing. */
  show() {
    const { r, c } = this.doc.sel;
    const range = selRange(this.doc.sel);
    this.nameBox.value = range.r2 - range.r1 >= MAX_ROWS - 1 || range.c2 - range.c1 >= MAX_COLS - 1 || (range.r1 === range.r2 && range.c1 === range.c2) ? address(r, c) : rangeAddress(range);
    if (!this.editing) this.barInput.value = this.doc.inputText(r, c);
  }

  get active() {
    return !!this.editing;
  }

  private get input(): HTMLTextAreaElement {
    return document.activeElement === this.barInput ? this.barInput : this.box;
  }

  private get text() {
    return this.input.value;
  }

  /** Starts editing the active cell: with a typed character (replacing what's there) or its current text. */
  start(initial?: string, mode: "enter" | "edit" = initial === undefined ? "edit" : "enter", inBar = false) {
    const { r, c } = this.doc.sel;
    const m = mergeAt(this.doc.sheet, r, c);
    const er = m?.r1 ?? r;
    const ec = m?.c1 ?? c;
    this.editing = { r: er, c: ec, sheet: this.doc.active, mode };
    this.pointed = null;
    const text = initial ?? this.doc.inputText(er, ec);
    this.box.value = text;
    this.barInput.value = text;
    this.place();
    this.box.hidden = false;
    const target = inBar ? this.barInput : this.box;
    target.focus();
    target.setSelectionRange(text.length, text.length);
    this.typed(target);
    this.hooks.changed();
  }

  /** Puts the box over the cell, in the cell's font and colours, so typing looks like it's in the cell; it grows to fit the text. */
  private place() {
    if (!this.editing) return;
    this.grid.reveal(this.editing.r, this.editing.c);
    const cell = this.grid.cellBox(this.editing.r, this.editing.c);
    const look = lookOf(this.doc.book.styles, this.doc.cell(this.editing.r, this.editing.c)?.s);
    const z = this.grid.zoom;
    // Inside the selection border, text where the cell shows it (at the bottom, 3 px in), no frame of its own.
    const size = ((look.font.size * 4) / 3) * z;
    Object.assign(this.box.style, {
      left: `${cell.x + 2}px`,
      top: `${cell.y + 2}px`,
      minWidth: `${Math.max(4, cell.w - 4)}px`,
      minHeight: `${Math.max(4, cell.h - 4)}px`,
      paddingTop: `${Math.max(0, cell.h - 5 - size * 1.25)}px`,
      font: `${look.font.italic ? "italic " : ""}${look.font.bold ? "bold " : ""}${size}px "${look.font.name}", "Segoe UI", sans-serif`,
      textAlign: look.h === "center" || look.h === "right" ? look.h : "left",
      background: look.fill ?? "#ffffff",
      color: look.font.color ?? "#000000",
    });
    this.fit();
  }

  private fit() {
    const box = this.box;
    box.style.width = "0px";
    box.style.height = "0px";
    const max = this.grid.el.clientWidth - box.offsetLeft - 4;
    box.style.width = `${Math.min(max, box.scrollWidth)}px`;
    box.style.height = `${box.scrollHeight}px`;
  }

  /** After typing: mirror the text, outline the references, suggest functions. */
  private typed(from: HTMLTextAreaElement) {
    const other = from === this.box ? this.barInput : this.box;
    if (other.value !== from.value) other.value = from.value;
    if (document.activeElement === from && this.pointed && from.selectionStart !== this.pointed.end) this.pointed = null;
    this.fit();
    this.outline();
    this.suggest(from);
  }

  private outline() {
    const text = this.text;
    this.grid.highlights = [];
    if (this.editing && text.startsWith("=")) {
      const seen = new Map<string, string>();
      const refs = text.slice(1).replace(/"[^"]*"/g, (s) => " ".repeat(s.length)).matchAll(/(?:(?:'[^']+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/g);
      for (const m of refs) {
        const [sheetPart, ref] = m[0].includes("!") ? [m[0].slice(0, m[0].lastIndexOf("!")).replace(/^'|'$/g, ""), m[0].slice(m[0].lastIndexOf("!") + 1)] : [null, m[0]];
        const sheetName = sheetPart ?? this.doc.book.sheets[this.editing.sheet].name;
        if (sheetName.toLowerCase() !== this.doc.sheet.name.toLowerCase()) continue;
        const range = parseRange(ref.replace(/\$/g, ""));
        if (!range) continue;
        const id = rangeAddress(range);
        if (!seen.has(id)) seen.set(id, COLORS[seen.size % COLORS.length]);
        this.grid.highlights.push({ range, color: seen.get(id)! });
      }
    }
    this.grid.render();
  }

  private suggest(input: HTMLTextAreaElement) {
    const before = input.value.slice(0, input.selectionStart);
    const m = /(^=|[=(,;+\-*/^&<>:]\s*)([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    this.suggestions = input.value.startsWith("=") && m ? FUNCTION_NAMES.filter((name) => name.startsWith(m[2].toUpperCase())).slice(0, 8) : [];
    if (this.suggestions.length === 1 && this.suggestions[0] === m?.[2].toUpperCase()) this.suggestions = [];
    this.chosen = 0;
    this.list.hidden = !this.suggestions.length || input !== this.box;
    if (this.list.hidden) return;
    this.list.replaceChildren(
      ...this.suggestions.map((name, i) => {
        const item = Object.assign(document.createElement("div"), { className: `suggestion${i === 0 ? " chosen" : ""}`, textContent: name });
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          this.chosen = i;
          this.acceptSuggestion();
        });
        return item;
      }),
    );
    this.list.style.left = this.box.style.left;
    this.list.style.top = `${this.box.offsetTop + this.box.offsetHeight + 2}px`;
  }

  private acceptSuggestion() {
    const input = this.input;
    const name = this.suggestions[this.chosen];
    const before = input.value.slice(0, input.selectionStart);
    const word = /[A-Za-z][A-Za-z0-9.]*$/.exec(before)![0];
    input.setRangeText(`${name}(`, input.selectionStart - word.length, input.selectionStart, "end");
    this.typed(input);
  }

  /** Whether clicking a cell now should put its address in the formula. */
  pointing(): boolean {
    if (!this.editing) return false;
    const input = this.input;
    if (!input.value.startsWith("=")) return false;
    if (this.pointed && input.selectionStart === this.pointed.end) return true;
    return input.selectionStart === input.selectionEnd && BEFORE_REF.test(input.value.slice(0, input.selectionStart));
  }

  /** Puts a cell or range address into the formula (replacing the one just pointed at). */
  point(range: Range, sheet: number) {
    if (!this.editing) return;
    const input = this.input;
    const own = sheet === this.editing.sheet;
    const text = `${own ? "" : `${quoteSheet(this.doc.book.sheets[sheet].name)}!`}${rangeAddress(range)}`;
    const start = this.pointed && input.selectionStart === this.pointed.end ? this.pointed.start : input.selectionStart;
    input.setRangeText(text, start, input.selectionEnd, "end");
    this.pointed = { start, end: start + text.length, r: range.r2, c: range.c2 };
    this.typed(input);
    input.focus();
  }

  /** F4: cycles the reference at the caret through A1, $A$1, A$1, $A1. */
  private toggleAbsolute(input: HTMLTextAreaElement) {
    const text = input.value;
    const at = input.selectionStart;
    const re = /\$?[A-Za-z]{1,3}\$?\d+/g;
    for (const m of text.matchAll(re)) {
      const start = m.index!;
      const end = start + m[0].length;
      if (at < start || at > end) continue;
      const bare = m[0].replace(/\$/g, "");
      const [, col, row] = /([A-Za-z]+)(\d+)/.exec(bare)!;
      const next = m[0] === bare ? `$${col}$${row}` : m[0] === `$${col}$${row}` ? `${col}$${row}` : m[0] === `${col}$${row}` ? `$${col}${row}` : bare;
      input.setRangeText(next, start, end, "end");
      this.typed(input);
      return;
    }
  }

  private key(e: KeyboardEvent, input: HTMLTextAreaElement) {
    if (!this.editing) return;
    if (!this.list.hidden && this.suggestions.length) {
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        this.chosen = (this.chosen + (e.key === "ArrowDown" ? 1 : this.suggestions.length - 1)) % this.suggestions.length;
        Array.from(this.list.children).forEach((el, i) => el.classList.toggle("chosen", i === this.chosen));
        return;
      }
      if (e.key === "Tab") {
        e.preventDefault();
        this.acceptSuggestion();
        return;
      }
    }
    switch (e.key) {
      case "Enter":
        e.preventDefault();
        if (e.altKey) {
          input.setRangeText("\n", input.selectionStart, input.selectionEnd, "end");
          this.typed(input);
        } else this.finish(e.shiftKey ? "up" : "down", e.ctrlKey);
        return;
      case "Tab":
        e.preventDefault();
        this.finish(e.shiftKey ? "left" : "right");
        return;
      case "Escape":
        e.preventDefault();
        this.cancel();
        return;
      case "F2":
        e.preventDefault();
        this.editing.mode = this.editing.mode === "edit" ? "enter" : "edit";
        return;
      case "F4":
        e.preventDefault();
        this.toggleAbsolute(input);
        return;
    }
    const arrows: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] };
    const step = arrows[e.key];
    if (!step || this.editing.mode === "edit" || input === this.barInput) return;
    e.preventDefault();
    if (this.pointing()) {
      // Arrow keys point at cells, from the one pointed at last (or the cell being edited).
      const from = this.pointed ?? { r: this.editing.r, c: this.editing.c };
      const r = Math.max(0, Math.min(MAX_ROWS - 1, from.r + step[0]));
      const c = Math.max(0, Math.min(MAX_COLS - 1, from.c + step[1]));
      this.point({ r1: r, c1: c, r2: r, c2: c }, this.doc.active);
      this.grid.reveal(r, c);
      return;
    }
    this.finish(e.key === "ArrowDown" ? "down" : e.key === "ArrowUp" ? "up" : e.key === "ArrowRight" ? "right" : "left");
  }

  /** Puts what was typed into the cell. */
  finish(move: Move = "none", everywhere = false) {
    const editing = this.editing;
    if (!editing) return;
    let text = this.text;
    // A formula left without its closing brackets gets them, as Excel does.
    if (text.startsWith("=")) {
      const open = (text.replace(/"[^"]*"/g, "").match(/\(/g) ?? []).length - (text.replace(/"[^"]*"/g, "").match(/\)/g) ?? []).length;
      if (open > 0) text += ")".repeat(open);
    }
    this.stop();
    if (this.doc.active !== editing.sheet) this.doc.active = editing.sheet;
    this.hooks.commit(editing.r, editing.c, text, move, everywhere);
  }

  cancel() {
    this.stop();
    this.show();
    this.grid.el.focus();
  }

  private stop() {
    this.editing = null;
    this.pointed = null;
    this.box.hidden = true;
    this.list.hidden = true;
    this.grid.highlights = [];
    this.grid.render();
    this.hooks.changed();
  }

  /** The name box: an address, a range, or a named range. */
  private goTo(text: string) {
    const value = text.trim();
    const named = this.doc.book.names.get(value.toUpperCase());
    const target = parseRange((named ?? value).replace(/^.*!/, "").replace(/\$/g, "")) ?? (colIndex(value) >= 0 ? { r1: 0, c1: colIndex(value), r2: MAX_ROWS - 1, c2: colIndex(value) } : null);
    if (!target) {
      this.nameBox.select();
      return;
    }
    this.doc.sel = { r: target.r1, c: target.c1, ar: target.r2, ac: target.c2 };
    this.grid.reveal(target.r1, target.c1);
    this.grid.render();
    this.show();
    this.grid.el.focus();
  }

  /** Keeps the box over its cell when the sheet scrolls or zooms. */
  follow() {
    if (this.editing && !this.box.hidden) this.place();
  }
}
