// DuckSheets: a spreadsheet that opens and saves Excel workbooks. This file
// puts the window together (toolbar, formula bar, sheet, tabs, status bar)
// and handles the keyboard, the clipboard and files.

import "./styles.css";
import { Doc, newWorkbook, selRange, type Clip } from "./doc";
import { Editor, type Move } from "./editor";
import { Grid } from "./grid";
import { h, isTyping } from "./dom";
import { icon } from "./icons";
import { MAX_COLS, MAX_ROWS, inRange, key, mergeAt, rangeAddress } from "./model";
import { dateToSerial, formatValue, locale } from "./numfmt";
import { createToolbar, numberFormats, sumRange } from "./toolbar";
import { readCsv, writeCsv } from "./csv";
import { readXlsx } from "./xlsx/read";
import { writeXlsx } from "./xlsx/write";
import { lookOf } from "./xlsx/styles";
import { askName, askSave, menu, popover, toast, type MenuItem } from "./ui";
import { baseName, onCloseRequested, onFilesOpened, pickFile, pickSavePath, readPath, setTitle, startupFiles, writePath, type OpenedFile } from "./platform";
import { offerUpdate } from "./updates";
import { insertChart, renderCharts } from "./charts";

const app = document.getElementById("app")!;
let doc: Doc;
let grid: Grid;
let editor: Editor;
let toolbar: ReturnType<typeof createToolbar>;
let tabs: HTMLElement;
let status: HTMLElement;
let summary: HTMLElement;
let zoomLabel: HTMLElement;
let clip: { data: Clip; text: string } | null = null;
let message = "";

// ---- putting the window together -------------------------------------------------------------

function mount(next: Doc) {
  doc = next;
  doc.fit = (c) => grid.fitWidth(c);
  grid = new Grid(doc, {
    edit: (initial) => editor.start(initial),
    selectionChanged: () => refreshSelection(),
    pointing: () => editor.pointing(),
    point: (range, sheet) => editor.point(range, sheet),
    contextMenu: (x, y, on) => contextMenu(x, y, on),
    filterMenu: (col, x, y) => filterMenu(col, x, y),
  });
  editor = new Editor(doc, grid, {
    commit: (r, c, text, move, everywhere) => {
      const range = selRange(doc.sel);
      if (everywhere && (range.r1 !== range.r2 || range.c1 !== range.c2)) doc.fillInput(range, text);
      else doc.setInput(r, c, text);
      moveAfterEntry(move);
      grid.el.focus({ preventScroll: true });
    },
    changed: () => app.classList.toggle("editing", editor.active),
  });
  toolbar = createToolbar(doc, grid, {
    newBook: () => void newBook(),
    open: () => void openFile(),
    save: () => void save(),
    autoSum,
    find: () => findReplace(false),
    chart: () => insertChart(doc, grid),
  });
  tabs = h("div", { class: "sheet-tabs", role: "tablist" });
  summary = h("span", { class: "summary" });
  zoomLabel = h("button", { type: "button", class: "tb zoom-label", title: "Back to 100% (Ctrl+scroll wheel zooms)", onclick: () => zoom(1) }, "100%");
  status = h("span", { class: "status-text" });
  const zoomBar = h(
    "div",
    { class: "zoom" },
    h("button", { type: "button", class: "tb", title: "Zoom out", "aria-label": "Zoom out", html: icon("zoomOut", 16), onclick: () => zoom(grid.zoom / 1.1) }),
    zoomLabel,
    h("button", { type: "button", class: "tb", title: "Zoom in", "aria-label": "Zoom in", html: icon("zoomIn", 16), onclick: () => zoom(grid.zoom * 1.1) }),
  );
  app.replaceChildren(toolbar.el, editor.bar, grid.el, h("div", { class: "bottom" }, tabs, h("div", { class: "statusbar" }, status, summary, zoomBar)));
  grid.el.addEventListener("keydown", keydown);
  // As in Excel, typing goes to the selected cell even after clicking a toolbar button.
  window.addEventListener("keydown", (e) => {
    const target = e.target as Element;
    if (e.defaultPrevented || grid.el.contains(target) || isTyping(target) || target.closest?.("dialog, .menu") || (target.tagName === "BUTTON" && (e.key === "Enter" || e.key === " "))) return;
    keydown(e);
  });
  grid.el.addEventListener("wheel", wheel, { passive: false });
  grid.el.addEventListener(
    "scroll",
    () => {
      editor.follow();
      renderCharts(doc, grid);
    },
    true,
  );
  doc.onChange(() => {
    grid.refresh();
    refreshSelection();
    showTabs();
    showTitle();
    renderCharts(doc, grid);
  });
  showTabs();
  showTitle();
  refreshSelection();
  renderCharts(doc, grid);
  grid.el.focus();
}

function refreshSelection() {
  editor.show();
  toolbar.update();
  showSummary();
  grid.render();
}

function showTitle() {
  setTitle(`${doc.dirty ? "• " : ""}${doc.name} – DuckSheets`);
}

function showStatus(text: string) {
  message = text;
  status.textContent = doc.calc.circular ? `Circular reference: ${doc.calc.circular}` : message;
  status.classList.toggle("warning", !!doc.calc.circular);
}

/** Average, count and sum of the selected cells, as in Excel's status bar. */
function showSummary() {
  const range = selRange(doc.sel);
  let count = 0;
  let numbers = 0;
  let sum = 0;
  if (range.r1 !== range.r2 || range.c1 !== range.c2) {
    for (const [k] of doc.sheet.cells) {
      const r = Math.floor(k / MAX_COLS);
      const c = k % MAX_COLS;
      if (!inRange(range, r, c) || doc.sheet.hiddenRows.has(r)) continue;
      const v = doc.value(r, c);
      if (v === null) continue;
      count++;
      if (typeof v === "number") {
        numbers++;
        sum += v;
      }
    }
  }
  const fmt = lookOf(doc.book.styles, doc.cell(doc.sel.r, doc.sel.c)?.s).numFmt;
  const show = (n: number) => formatValue(n, fmt === "General" || fmt === "@" ? "General" : fmt, doc.book.date1904).text;
  summary.textContent = numbers ? `Average: ${show(sum / numbers)}    Count: ${count}    Sum: ${show(sum)}` : count > 1 ? `Count: ${count}` : "";
  showStatus(message);
}

function showTabs() {
  tabs.replaceChildren(
    ...doc.book.sheets.map((sheet, i) => {
      const tab = h("button", { type: "button", class: `sheet-tab${i === doc.active ? " active" : ""}`, role: "tab", "aria-selected": String(i === doc.active), title: "Double-click to rename" }, sheet.name);
      tab.addEventListener("click", () => switchSheet(i));
      tab.addEventListener("dblclick", () => void renameSheet(i));
      tab.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        switchSheet(i);
        menu(tab, [
          { label: "Rename…", action: () => void renameSheet(i) },
          { label: "Insert a new sheet", action: () => doc.addSheet() },
          { label: "Move left", action: () => doc.moveSheet(i, i - 1), ...(i === 0 ? { items: [] } : {}) },
          { label: "Move right", action: () => doc.moveSheet(i, i + 1), ...(i === doc.book.sheets.length - 1 ? { items: [] } : {}) },
          { label: "Delete", danger: true, action: () => doc.deleteSheet(i), ...(doc.book.sheets.length < 2 ? { items: [] } : {}) },
        ]);
      });
      return tab;
    }),
    h("button", { type: "button", class: "tb add-sheet", title: "New sheet (Shift+F11)", "aria-label": "New sheet", html: icon("plus", 16), onclick: () => doc.addSheet() }),
  );
}

function switchSheet(i: number) {
  if (editor.active && !editor.pointing()) editor.finish();
  // While pointing in a formula, other sheets can be clicked through without losing the formula.
  if (i === doc.active) return;
  doc.active = i;
  doc.sel = { r: 0, c: 0, ar: 0, ac: 0 };
  grid.refresh();
  showTabs();
  refreshSelection();
  renderCharts(doc, grid);
}

async function renameSheet(i: number) {
  const name = await askName("Rename sheet", doc.book.sheets[i].name, "Rename");
  if (!name) return;
  const problem = doc.renameSheet(i, name);
  if (problem) toast(problem);
}

function zoom(z: number) {
  grid.setZoom(z);
  zoomLabel.textContent = `${Math.round(grid.zoom * 100)}%`;
  editor.follow();
  renderCharts(doc, grid);
}

// ---- moving around -------------------------------------------------------------------------------

function moveAfterEntry(move: Move) {
  const step: Record<Move, [number, number]> = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1], none: [0, 0] };
  const [dr, dc] = step[move];
  const range = selRange(doc.sel);
  // In a selected block, Enter and Tab move within it, as in Excel.
  if (range.r1 !== range.r2 || range.c1 !== range.c2) {
    let { r, c } = doc.sel;
    if (dr) {
      r += dr;
      if (r > range.r2) (r = range.r1), (c = c + 1 > range.c2 ? range.c1 : c + 1);
      if (r < range.r1) (r = range.r2), (c = c - 1 < range.c1 ? range.c2 : c - 1);
    } else if (dc) {
      c += dc;
      if (c > range.c2) (c = range.c1), (r = r + 1 > range.r2 ? range.r1 : r + 1);
      if (c < range.c1) (c = range.c2), (r = r - 1 < range.r1 ? range.r2 : r - 1);
    }
    doc.sel = { ...doc.sel, r, c };
  } else go(doc.sel.r + dr, doc.sel.c + dc);
  refreshSelection();
}

function go(r: number, c: number, extend = false) {
  r = Math.max(0, Math.min(MAX_ROWS - 1, r));
  c = Math.max(0, Math.min(MAX_COLS - 1, c));
  doc.sel = extend ? { ...doc.sel, ar: r, ac: c } : { r, c, ar: r, ac: c };
  grid.reveal(r, c);
  refreshSelection();
}

/** Ctrl+arrow: to the edge of the data, as in Excel. */
function jump(r: number, c: number, dr: number, dc: number): [number, number] {
  const filled = (rr: number, cc: number) => doc.value(rr, cc) !== null;
  const inside = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS;
  if (!inside(r + dr, c + dc)) return [r, c];
  let rr = r + dr;
  let cc = c + dc;
  if (filled(r, c) && filled(rr, cc)) {
    while (inside(rr + dr, cc + dc) && filled(rr + dr, cc + dc)) (rr += dr), (cc += dc);
    return [rr, cc];
  }
  // From an empty cell (or the end of a block): to the next filled one, or the sheet's edge.
  const lastRow = Math.max(0, ...[...doc.sheet.cells.keys()].map((k) => Math.floor(k / MAX_COLS)));
  const lastCol = Math.max(0, ...[...doc.sheet.cells.keys()].map((k) => k % MAX_COLS));
  while (inside(rr, cc) && !filled(rr, cc)) {
    if ((dr > 0 && rr > lastRow) || (dc > 0 && cc > lastCol)) return [dr ? MAX_ROWS - 1 : rr, dc ? MAX_COLS - 1 : cc];
    rr += dr;
    cc += dc;
  }
  return inside(rr, cc) ? [rr, cc] : [Math.max(0, Math.min(MAX_ROWS - 1, rr - dr)), Math.max(0, Math.min(MAX_COLS - 1, cc - dc))];
}

// ---- keys ------------------------------------------------------------------------------------------

function keydown(e: KeyboardEvent) {
  if (editor.active || isTyping(e.target as Element)) return;
  const mod = e.ctrlKey || e.metaKey;
  const range = selRange(doc.sel);
  const { r, c } = doc.sel;
  const arrows: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] };
  const done = () => e.preventDefault();

  if (arrows[e.key]) {
    const [dr, dc] = arrows[e.key];
    const from = e.shiftKey ? { r: doc.sel.ar, c: doc.sel.ac } : { r, c };
    // Out of a merged cell, from its far side.
    const m = !e.shiftKey && mergeAt(doc.sheet, r, c);
    const start = m ? { r: dr > 0 ? m.r2 : m.r1, c: dc > 0 ? m.c2 : m.c1 } : from;
    const [nr, nc] = mod ? jump(start.r, start.c, dr, dc) : [start.r + dr, start.c + dc];
    const skipHidden = (i: number, step: number, hidden: Set<number>) => {
      while (hidden.has(i) && i + step >= 0) i += step;
      return i;
    };
    go(skipHidden(nr, dr || 1, doc.sheet.hiddenRows), skipHidden(nc, dc || 1, doc.sheet.hiddenCols), e.shiftKey);
    return done();
  }
  switch (e.key) {
    case "Enter":
      moveAfterEntry(e.shiftKey ? "up" : "down");
      return done();
    case "Tab":
      moveAfterEntry(e.shiftKey ? "left" : "right");
      return done();
    case "Home":
      go(mod ? 0 : r, 0, e.shiftKey);
      return done();
    case "End":
      if (mod) {
        const used = [...doc.sheet.cells.keys()];
        go(Math.max(0, ...used.map((k) => Math.floor(k / MAX_COLS))), Math.max(0, ...used.map((k) => k % MAX_COLS)), e.shiftKey);
        return done();
      }
      return;
    case "PageDown":
    case "PageUp":
      if (mod) switchSheet(Math.max(0, Math.min(doc.book.sheets.length - 1, doc.active + (e.key === "PageDown" ? 1 : -1))));
      else go(r + (e.key === "PageDown" ? grid.pageRows : -grid.pageRows), c, e.shiftKey);
      return done();
    case "F2":
      editor.start();
      return done();
    case "F11":
      if (e.shiftKey) {
        doc.addSheet();
        return done();
      }
      return;
    case "Delete":
      doc.clear(range, "contents");
      return done();
    case "Backspace":
      doc.clear({ r1: r, c1: c, r2: r, c2: c }, "contents");
      editor.start("");
      return done();
    case "Escape":
      grid.marquee = null;
      clip = clip && { ...clip, data: { ...clip.data, cut: false } };
      grid.render();
      return done();
  }
  if (e.key === " " && mod && !e.shiftKey) {
    go(0, c);
    doc.sel = { r: 0, c: range.c1, ar: MAX_ROWS - 1, ac: range.c2 };
    refreshSelection();
    return done();
  }
  if (e.key === " " && e.shiftKey && !mod) {
    doc.sel = { r: range.r1, c: 0, ar: range.r2, ac: MAX_COLS - 1 };
    refreshSelection();
    return done();
  }
  if (e.key === "=" && e.altKey) {
    autoSum();
    return done();
  }
  if (mod && !e.altKey) {
    const k = e.key.toLowerCase();
    const look = lookOf(doc.book.styles, doc.cell(r, c)?.s);
    const formats = numberFormats();
    // Ctrl+Shift+4 / 5 / 3 / 1 by key position, which works on every keyboard layout.
    if (e.shiftKey && /^Digit[1345]$/.test(e.code)) {
      const code = { Digit1: formats[1][1], Digit3: formats[4][1], Digit4: formats[2][1], Digit5: "0%" }[e.code as "Digit1"];
      doc.format(grid.expanded(range), { numFmt: code });
      return done();
    }
    switch (k) {
      case "z":
        if (e.shiftKey) doc.redo();
        else doc.undo();
        return done();
      case "y":
        doc.redo();
        return done();
      case "b":
        doc.format(grid.expanded(range), { bold: !look.font.bold });
        return done();
      case "i":
        doc.format(grid.expanded(range), { italic: !look.font.italic });
        return done();
      case "u":
        doc.format(grid.expanded(range), { underline: !look.font.underline });
        return done();
      case "5":
        if (e.shiftKey) return;
        doc.format(grid.expanded(range), { strike: !look.font.strike });
        return done();
      case "a": {
        // First the block of data, then the whole sheet.
        const region = doc.region(r, c);
        const already = range.r1 === region.r1 && range.r2 === region.r2 && range.c1 === region.c1 && range.c2 === region.c2;
        const target = already || doc.value(r, c) === null ? { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 } : region;
        doc.sel = { r: target.r1, c: target.c1, ar: target.r2, ac: target.c2 };
        refreshSelection();
        return done();
      }
      case "d":
        if (range.r2 > range.r1) doc.fill({ ...range, r2: range.r1 }, range);
        else if (r > 0) doc.fill({ r1: r - 1, c1: range.c1, r2: r - 1, c2: range.c2 }, { ...range, r1: r - 1 });
        return done();
      case "r":
        if (range.c2 > range.c1) doc.fill({ ...range, c2: range.c1 }, range);
        else if (c > 0) doc.fill({ r1: range.r1, c1: c - 1, r2: range.r2, c2: c - 1 }, { ...range, c1: c - 1 });
        return done();
      case ";":
      case ":": {
        const now = new Date();
        if (e.shiftKey || k === ":") doc.setInput(r, c, `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`);
        else doc.setInput(r, c, formatValue(dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()), locale.dateCode).text);
        return done();
      }
      case "l":
        if (!e.shiftKey) return;
        doc.toggleFilter();
        return done();
      case "f":
        findReplace(false);
        return done();
      case "h":
        findReplace(true);
        return done();
      case "s":
        void save(e.shiftKey);
        return done();
      case "o":
        void openFile();
        return done();
      case "n":
        void newBook();
        return done();
      case "+":
        insertOrDelete(1);
        return done();
      case "-":
        insertOrDelete(-1);
        return done();
    }
    return;
  }
  // Typing starts entering a new value, replacing the cell's.
  if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !(e.altKey && !e.getModifierState("AltGraph"))) {
    e.preventDefault();
    editor.start(e.key);
  }
}

function wheel(e: WheelEvent) {
  if (e.ctrlKey) {
    e.preventDefault();
    zoom(grid.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
  } else if (e.altKey) {
    // Alt+scroll wheel makes the selected cells' text a size bigger or smaller.
    e.preventDefault();
    const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
    const now = lookOf(doc.book.styles, doc.cell(doc.sel.r, doc.sel.c)?.s).font.size;
    const next = e.deltaY < 0 ? sizes.find((s) => s > now) : [...sizes].reverse().find((s) => s < now);
    if (next) doc.format(grid.expanded(selRange(doc.sel)), { fontSize: next });
  }
}

/** Ctrl++ and Ctrl+-: whole rows or columns when they're selected, else ask which. */
function insertOrDelete(sign: 1 | -1) {
  const range = selRange(doc.sel);
  if (range.c1 === 0 && range.c2 >= MAX_COLS - 1) return doc.insertDelete("r", range.r1, sign * (range.r2 - range.r1 + 1));
  if (range.r1 === 0 && range.r2 >= MAX_ROWS - 1) return doc.insertDelete("c", range.c1, sign * (range.c2 - range.c1 + 1));
  const anchor = grid.el;
  menu(anchor, [
    { label: sign > 0 ? "Insert whole rows" : "Delete whole rows", action: () => doc.insertDelete("r", range.r1, sign * (range.r2 - range.r1 + 1)) },
    { label: sign > 0 ? "Insert whole columns" : "Delete whole columns", action: () => doc.insertDelete("c", range.c1, sign * (range.c2 - range.c1 + 1)) },
  ]);
}

function autoSum() {
  const range = selRange(doc.sel);
  if (range.r1 === range.r2 && range.c1 === range.c2) {
    const sum = sumRange(doc, range.r1, range.c1);
    editor.start(`=SUM(${sum ? rangeAddress(sum) : ""})`, "edit");
    return;
  }
  // A block of numbers: a total under each column.
  const below = range.r2 + 1;
  doc.change(() => {
    for (let c = range.c1; c <= range.c2; c++) doc.sheet.cells.set(key(below, c), { ...doc.cell(below, c), f: `SUM(${rangeAddress({ r1: range.r1, c1: c, r2: range.r2, c2: c })})` });
  });
}

// ---- the clipboard ----------------------------------------------------------------------------------

document.addEventListener("copy", (e) => copy(e, false));
document.addEventListener("cut", (e) => copy(e, true));
document.addEventListener("paste", (e) => {
  if (!doc || editor.active || isTyping(e.target as Element)) return;
  e.preventDefault();
  const text = e.clipboardData?.getData("text/plain") ?? "";
  const at = { r: selRange(doc.sel).r1, c: selRange(doc.sel).c1 };
  // Our own copy, if the clipboard still holds it: with formulas and formats.
  if (clip && text.replace(/\r/g, "") === clip.text.replace(/\r/g, "")) {
    doc.paste(clip.data, at);
    if (clip.data.cut) {
      clip = null;
      grid.marquee = null;
    }
  } else if (text) doc.pasteText(text, at);
  refreshSelection();
});

function copy(e: ClipboardEvent, cut: boolean) {
  if (!doc || editor.active || isTyping(e.target as Element)) return;
  e.preventDefault();
  const range = grid.expanded(selRange(doc.sel));
  if (range.r2 - range.r1 > 100000 || range.c2 - range.c1 > 1000) return toast("That's too much to copy at once; select just the cells with data.");
  const { tsv, html } = doc.copyText(range);
  e.clipboardData?.setData("text/plain", tsv);
  e.clipboardData?.setData("text/html", html);
  clip = { data: doc.copy(range, cut), text: tsv };
  grid.marquee = range;
  grid.render();
}

// ---- menus --------------------------------------------------------------------------------------------

function contextMenu(x: number, y: number, on: "cell" | "row" | "col") {
  const range = selRange(doc.sel);
  const rows = range.r2 - range.r1 + 1;
  const cols = range.c2 - range.c1 + 1;
  const exec = (command: "copy" | "cut") => () => document.execCommand(command);
  const items: MenuItem[] = [
    { label: "Cut", action: exec("cut") },
    { label: "Copy", action: exec("copy") },
    { label: "Paste", action: () => void navigator.clipboard.readText().then((text) => (clip && text.replace(/\r/g, "") === clip.text.replace(/\r/g, "") ? doc.paste(clip.data, { r: range.r1, c: range.c1 }) : doc.pasteText(text, { r: range.r1, c: range.c1 }))) },
  ];
  if (on !== "col") items.push({ label: rows > 1 ? `Insert ${rows} rows above` : "Insert a row above", action: () => doc.insertDelete("r", range.r1, rows) }, { label: rows > 1 ? `Delete ${rows} rows` : "Delete the row", action: () => doc.insertDelete("r", range.r1, -rows) });
  if (on !== "row") items.push({ label: cols > 1 ? `Insert ${cols} columns left` : "Insert a column left", action: () => doc.insertDelete("c", range.c1, cols) }, { label: cols > 1 ? `Delete ${cols} columns` : "Delete the column", action: () => doc.insertDelete("c", range.c1, -cols) });
  items.push({ label: "Clear contents", action: () => doc.clear(range, "contents") }, { label: "Clear formats", action: () => doc.clear(range, "formats") });
  if (on === "row") {
    const list = Array.from({ length: Math.min(rows, 5000) }, (_, i) => range.r1 + i);
    items.push(
      { label: "Row height…", action: () => void askName("Row height (pixels)", String(doc.sheet.rowHeights.get(range.r1) ?? doc.sheet.defaultRowHeight), "OK").then((v) => v && Number(v) >= 0 && doc.setRowHeight(list, Number(v))) },
      { label: "Hide", action: () => doc.setHidden("r", list, true) },
      { label: "Unhide", action: () => doc.setHidden("r", Array.from({ length: rows + 2 }, (_, i) => range.r1 - 1 + i).filter((i) => i >= 0), false) },
    );
  }
  if (on === "col") {
    const list = Array.from({ length: Math.min(cols, 1000) }, (_, i) => range.c1 + i);
    items.push(
      { label: "Column width…", action: () => void askName("Column width (pixels)", String(doc.sheet.colWidths.get(range.c1) ?? doc.sheet.defaultColWidth), "OK").then((v) => v && Number(v) >= 0 && doc.setColWidth(list, Number(v))) },
      { label: "Fit to contents", action: () => list.forEach((c) => doc.setColWidth([c], grid.fitWidth(c))) },
      { label: "Hide", action: () => doc.setHidden("c", list, true) },
      { label: "Unhide", action: () => doc.setHidden("c", Array.from({ length: cols + 2 }, (_, i) => range.c1 - 1 + i).filter((i) => i >= 0), false) },
    );
  }
  if (on === "cell") items.push({ label: "Merge and centre", action: () => doc.toggleMerge(grid.expanded(range)) });
  // A menu needs something to sit under: a tiny element where the mouse is.
  const at = h("div", { class: "menu-anchor" });
  at.style.left = `${x}px`;
  at.style.top = `${y}px`;
  document.body.append(at);
  menu(at, items);
  setTimeout(() => at.remove(), 1000);
}

/** The filter list for a column: every value in it, ticked if shown; sorting too. */
function filterMenu(col: number, x: number, y: number) {
  const filter = doc.sheet.filter;
  if (!filter) return;
  const offset = col - filter.range.c1;
  const values = new Set<string>();
  for (let r = filter.range.r1 + 1; r <= filter.range.r2; r++) values.add(doc.shown(r, col));
  const all = [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const shown = filter.shown.get(offset) ?? new Set(all);
  const at = h("div", { class: "menu-anchor" });
  at.style.left = `${x}px`;
  at.style.top = `${y}px`;
  document.body.append(at);
  popover(at, (close) => {
    const search = h("input", { type: "search", class: "tb-input wide", placeholder: "Search", "aria-label": "Search the values" }) as HTMLInputElement;
    const boxes = all.map((v) => {
      const box = h("input", { type: "checkbox" }) as HTMLInputElement;
      box.checked = shown.has(v);
      return { v, box, row: h("label", { class: "filter-item" }, box, v === "" ? "(Blanks)" : v) };
    });
    const allBox = h("input", { type: "checkbox" }) as HTMLInputElement;
    allBox.checked = boxes.every((b) => b.box.checked);
    allBox.addEventListener("change", () => boxes.forEach((b) => (b.row.hidden ? null : (b.box.checked = allBox.checked))));
    search.addEventListener("input", () => boxes.forEach((b) => (b.row.hidden = !b.v.toLowerCase().includes(search.value.toLowerCase()))));
    const ok = h("button", { type: "button", class: "btn primary small" }, "OK");
    ok.addEventListener("click", () => {
      close();
      const chosen = new Set(boxes.filter((b) => b.box.checked).map((b) => b.v));
      doc.setFilter(offset, chosen.size === all.length ? null : chosen);
    });
    const sorter = (descending: boolean) => () => {
      close();
      doc.sort(filter.range, col, descending, true);
      if (doc.sheet.filter) doc.setFilter(offset, doc.sheet.filter.shown.get(offset) ?? null);
    };
    return h(
      "div",
      { class: "filter-menu" },
      h("button", { type: "button", class: "menu-item", onclick: sorter(false) }, "Sort A to Z"),
      h("button", { type: "button", class: "menu-item", onclick: sorter(true) }, "Sort Z to A"),
      h("button", { type: "button", class: "menu-item", onclick: () => (close(), doc.setFilter(offset, null)) }, "Clear this filter"),
      search,
      h("div", { class: "filter-list" }, h("label", { class: "filter-item" }, allBox, "(Select all)"), ...boxes.map((b) => b.row)),
      h("div", { class: "dialog-buttons" }, h("button", { type: "button", class: "btn small", onclick: () => close() }, "Cancel"), ok),
    );
  });
  setTimeout(() => at.remove(), 1000);
}

/** Ctrl+F finds; Ctrl+H also replaces. Searches what the cells show, and their formulas. */
function findReplace(replacing: boolean) {
  const anchor = editor.bar;
  popover(anchor, (close) => {
    const what = h("input", { type: "text", class: "tb-input wide", placeholder: "Find", "aria-label": "Find" }) as HTMLInputElement;
    const withText = h("input", { type: "text", class: "tb-input wide", placeholder: "Replace with", "aria-label": "Replace with", hidden: !replacing }) as HTMLInputElement;
    const result = h("div", { class: "find-result" });
    const matches = () => {
      const q = what.value.toLowerCase();
      const found: [number, number][] = [];
      if (!q) return found;
      for (const k of doc.sheet.cells.keys()) {
        const r = Math.floor(k / MAX_COLS);
        const c = k % MAX_COLS;
        if (doc.shown(r, c).toLowerCase().includes(q) || doc.inputText(r, c).toLowerCase().includes(q)) found.push([r, c]);
      }
      return found.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    };
    const next = () => {
      const found = matches();
      if (!found.length) return void (result.textContent = "Not found.");
      const { r, c } = doc.sel;
      const after = found.find(([fr, fc]) => fr > r || (fr === r && fc > c)) ?? found[0];
      go(after[0], after[1]);
      result.textContent = `${found.findIndex((f) => f === after) + 1} of ${found.length}`;
    };
    const replaceAll = () => {
      const q = what.value;
      if (!q) return;
      const found = matches();
      const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
      doc.change(() => {
        for (const [r, c] of found) {
          const cell = doc.cell(r, c)!;
          if (cell.f !== undefined) doc.sheet.cells.set(key(r, c), { ...cell, f: cell.f.replace(re, withText.value) });
          else if (typeof cell.v === "string") doc.sheet.cells.set(key(r, c), { ...cell, v: cell.v.replace(re, withText.value) });
        }
      });
      result.textContent = `Replaced in ${found.length} ${found.length === 1 ? "cell" : "cells"}.`;
    };
    what.addEventListener("keydown", (e) => {
      if (e.key === "Enter") next();
    });
    setTimeout(() => what.focus());
    return h(
      "div",
      { class: "find-box" },
      what,
      withText,
      h(
        "div",
        { class: "dialog-buttons" },
        h("button", { type: "button", class: "btn small", onclick: () => (withText.hidden = !withText.hidden) }, "Replace…"),
        h("button", { type: "button", class: "btn small", onclick: replaceAll, hidden: false }, "Replace all"),
        h("button", { type: "button", class: "btn primary small", onclick: next }, "Find next"),
        h("button", { type: "button", class: "btn subtle small", onclick: () => close() }, "Close"),
      ),
      result,
    );
  });
}

// ---- files ----------------------------------------------------------------------------------------------

/** Offers to save changes first; false means the user cancelled. */
async function saveFirst(): Promise<boolean> {
  if (!doc?.dirty) return true;
  const answer = await askSave(doc.name);
  if (answer === "cancel") return false;
  if (answer === "save") return save();
  return true;
}

let books = 1;

async function newBook() {
  if (!(await saveFirst())) return;
  mount(new Doc(newWorkbook(), `Book${++books}.xlsx`));
}

async function openFile() {
  if (!(await saveFirst())) return;
  const file = await pickFile();
  if (file) await load(file);
}

async function load(file: OpenedFile) {
  try {
    if (/\.csv$/i.test(file.name)) {
      const book = newWorkbook();
      book.sheets = [readCsv(file.bytes, file.name.replace(/\.csv$/i, "").slice(0, 31), book)];
      mount(new Doc(book, file.name, file.path));
    } else {
      mount(new Doc(await readXlsx(file.bytes), file.name, file.path));
    }
    showStatus(`Opened ${file.name}`);
  } catch (error) {
    console.error(error);
    toast(`Couldn't open ${file.name}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** Saves; resolves to false if cancelled or it failed. */
async function save(as = false): Promise<boolean> {
  if (editor.active) editor.finish();
  let path = doc.path;
  if (!path || as) {
    path = await pickSavePath(doc.name.replace(/\.(xlsm|csv)$/i, (m) => m) || "Book1.xlsx");
    if (!path) return false;
  }
  try {
    let bytes: Uint8Array;
    if (/\.csv$/i.test(path)) {
      bytes = new TextEncoder().encode(writeCsv(doc.book, doc.active, doc.calc));
      if (doc.book.sheets.length > 1) toast("A CSV file holds one sheet: only this one was saved, without formulas or formats.");
    } else {
      if (!/\.xls[xm]$/i.test(path)) path += ".xlsx";
      bytes = await writeXlsx(doc.book, doc.calc, doc.active);
    }
    await writePath(path, bytes);
    doc.path = path;
    doc.name = baseName(path);
    doc.dirty = false;
    showTitle();
    showStatus(`Saved ${doc.name}`);
    return true;
  } catch (error) {
    console.error(error);
    toast(`Couldn't save: ${error instanceof Error ? error.message : String(error)}`);
    return false;
  }
}

async function openPaths(paths: string[]) {
  const path = paths.find((p) => /\.(xlsx|xlsm|csv)$/i.test(p));
  if (!path || !(await saveFirst())) return;
  await load(await readPath(path));
}

mount(new Doc(newWorkbook()));
onCloseRequested(saveFirst);
onFilesOpened((paths) => void openPaths(paths));
void startupFiles().then((paths) => {
  if (paths.length) void openPaths(paths);
});
void offerUpdate(app, { beforeInstall: saveFirst });

// For the tests, which drive the app from outside.
Object.assign(window, {
  __ducksheets: {
    get doc() {
      return doc;
    },
    get grid() {
      return grid;
    },
    save: (path: string) => {
      doc.path = path;
      return save();
    },
    load: (path: string) => readPath(path).then(load),
  },
});
