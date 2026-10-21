// Format Cells (Ctrl+1), as in Excel: the number format by category, with
// decimals and a sample, or a code of your own; alignment; font; and fill.
// Only what is changed in the dialog is applied, so cells formatted in
// different ways keep what wasn't touched.

import { selRange, type Doc } from "./doc";
import { h } from "./dom";
import type { Grid } from "./grid";
import { formatValue, locale } from "./numfmt";
import { numberFormats } from "./toolbar";
import { modal } from "./ui";
import { lookOf, type StyleChange } from "./xlsx/styles";

const FONTS = ["Calibri", "Aptos", "Arial", "Cambria", "Consolas", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana"];

/** A category's format code with this many decimals (and thousands separators, for numbers). */
export function codeFor(category: string, decimals: number, thousands: boolean): string | null {
  const d = decimals > 0 ? `.${"0".repeat(decimals)}` : "";
  const kr = locale.decimal === ",";
  switch (category) {
    case "Number":
      return `${thousands ? "#,##0" : "0"}${d}`;
    case "Currency":
      return kr ? `#,##0${d} "kr"` : `$#,##0${d}`;
    case "Percentage":
      return `0${d}%`;
    case "Scientific":
      return `0${d}E+00`;
    default:
      return null;
  }
}

/** The category and decimals a code belongs to, to open the dialog where the cell is. */
function categoryOf(code: string): { category: string; decimals: number; thousands: boolean } {
  const decimals = (/\.(0+)/.exec(code)?.[1].length ?? 0) as number;
  const thousands = code.includes("#,##0");
  for (const [name, c] of numberFormats()) if (c === code) return { category: name, decimals: codeFor(name, decimals, thousands) ? decimals : 2, thousands };
  for (const name of ["Percentage", "Scientific", "Currency", "Number"])
    if (codeFor(name, decimals, thousands) === code) return { category: name, decimals, thousands };
  return { category: "Custom", decimals: 2, thousands: true };
}

export function formatCells(doc: Doc, grid: Grid) {
  const range = grid.expanded(selRange(doc.sel));
  const cell = doc.cell(doc.sel.r, doc.sel.c);
  const look = lookOf(doc.book.styles, cell?.s);
  const value = doc.value(doc.sel.r, doc.sel.c);
  const change: StyleChange = {};

  // ---- Number
  const start = categoryOf(look.numFmt);
  const categories = [...numberFormats().map(([name]) => name), "Custom"];
  const category = h("select", { class: "tb-select wide", size: 8, "aria-label": "Category" }, ...categories.map((name) => h("option", { value: name }, name))) as HTMLSelectElement;
  category.value = start.category;
  const decimals = h("input", { type: "number", min: 0, max: 15, class: "tb-input", value: String(start.decimals), "aria-label": "Decimal places" }) as HTMLInputElement;
  const thousands = h("input", { type: "checkbox", checked: start.thousands }) as HTMLInputElement;
  const code = h("input", { type: "text", class: "tb-input wide mono", value: look.numFmt, spellcheck: "false", "aria-label": "Format code" }) as HTMLInputElement;
  const sample = h("div", { class: "format-sample" });
  const decimalsRow = h("label", { class: "format-row" }, "Decimal places ", decimals);
  const thousandsRow = h("label", { class: "format-row" }, thousands, " Use thousands separator (1 000)");
  const showSample = () => {
    sample.textContent = value === null ? "(the selected cell is empty)" : formatValue(value, code.value || "General", doc.book.date1904).text || " ";
  };
  const fromCategory = () => {
    const made = codeFor(category.value, Math.max(0, Math.min(15, Number(decimals.value) || 0)), thousands.checked);
    decimalsRow.hidden = made === null;
    thousandsRow.hidden = category.value !== "Number";
    if (made) code.value = made;
    else if (category.value !== "Custom") code.value = numberFormats().find(([name]) => name === category.value)?.[1] ?? code.value;
    change.numFmt = code.value;
    showSample();
  };
  category.addEventListener("change", fromCategory);
  decimals.addEventListener("input", fromCategory);
  thousands.addEventListener("change", fromCategory);
  code.addEventListener("input", () => {
    category.value = "Custom";
    decimalsRow.hidden = thousandsRow.hidden = true;
    change.numFmt = code.value || "General";
    showSample();
  });
  decimalsRow.hidden = codeFor(start.category, 0, false) === null;
  thousandsRow.hidden = start.category !== "Number";
  showSample();
  const numberTab = h("div", { class: "format-tab" }, h("div", { class: "format-cols" }, category, h("div", { class: "format-side" }, h("p", {}, "Sample"), sample, decimalsRow, thousandsRow, h("p", {}, "Format code"), code)));

  // ---- Alignment
  const select = (label: string, options: [string, string][], current: string) => {
    const el = h("select", { class: "tb-select wide", "aria-label": label }, ...options.map(([v, text]) => h("option", { value: v }, text))) as HTMLSelectElement;
    el.value = current;
    return el;
  };
  const hAlign = select("Horizontal", [["general", "General"], ["left", "Left"], ["center", "Centre"], ["right", "Right"]], look.h ?? "general");
  const vAlign = select("Vertical", [["top", "Top"], ["center", "Centre"], ["bottom", "Bottom"]], look.v ?? "bottom");
  const wrap = h("input", { type: "checkbox", checked: !!look.wrap }) as HTMLInputElement;
  const indent = h("input", { type: "number", min: 0, max: 15, class: "tb-input", value: String(look.indent ?? 0), "aria-label": "Indent" }) as HTMLInputElement;
  hAlign.addEventListener("change", () => (change.h = hAlign.value === "general" ? null : hAlign.value));
  vAlign.addEventListener("change", () => (change.v = vAlign.value === "bottom" ? null : vAlign.value));
  wrap.addEventListener("change", () => (change.wrap = wrap.checked));
  indent.addEventListener("input", () => (change.indent = Math.max(0, Math.min(15, Number(indent.value) || 0))));
  const alignTab = h(
    "div",
    { class: "format-tab" },
    h("p", {}, "Horizontal"),
    hAlign,
    h("p", {}, "Vertical"),
    vAlign,
    h("label", { class: "format-row" }, "Indent ", indent),
    h("label", { class: "format-row" }, wrap, " Wrap text"),
  );

  // ---- Font
  const fontName = h("input", { type: "text", class: "tb-input wide", value: look.font.name, list: "format-fonts", "aria-label": "Font" }) as HTMLInputElement;
  const fontList = h("datalist", { id: "format-fonts" }, ...FONTS.map((f) => h("option", { value: f })));
  const fontSize = h("input", { type: "number", min: 1, max: 409, class: "tb-input", value: String(look.font.size), "aria-label": "Size" }) as HTMLInputElement;
  const check = (label: string, on: boolean | undefined, set: (v: boolean) => void) => {
    const box = h("input", { type: "checkbox", checked: !!on }) as HTMLInputElement;
    box.addEventListener("change", () => set(box.checked));
    return h("label", { class: "format-row" }, box, ` ${label}`);
  };
  const color = h("input", { type: "color", value: (look.font.color ?? "#000000").toLowerCase(), "aria-label": "Colour" }) as HTMLInputElement;
  fontName.addEventListener("input", () => fontName.value.trim() && (change.fontName = fontName.value.trim()));
  fontSize.addEventListener("input", () => Number(fontSize.value) > 0 && (change.fontSize = Number(fontSize.value)));
  color.addEventListener("input", () => (change.color = color.value.toUpperCase()));
  const fontTab = h(
    "div",
    { class: "format-tab" },
    h("p", {}, "Font"),
    fontName,
    fontList,
    h("label", { class: "format-row" }, "Size ", fontSize),
    check("Bold", look.font.bold, (v) => (change.bold = v)),
    check("Italic", look.font.italic, (v) => (change.italic = v)),
    check("Underline", look.font.underline, (v) => (change.underline = v)),
    check("Strikethrough", look.font.strike, (v) => (change.strike = v)),
    h("label", { class: "format-row" }, "Colour ", color),
  );

  // ---- Fill
  const fill = h("input", { type: "color", value: (look.fill ?? "#ffffff").toLowerCase(), "aria-label": "Fill colour" }) as HTMLInputElement;
  const noFill = h("input", { type: "checkbox", checked: !look.fill }) as HTMLInputElement;
  fill.addEventListener("input", () => {
    noFill.checked = false;
    change.fill = fill.value.toUpperCase();
  });
  noFill.addEventListener("change", () => (change.fill = noFill.checked ? null : fill.value.toUpperCase()));
  const fillTab = h("div", { class: "format-tab" }, h("label", { class: "format-row" }, "Background colour ", fill), h("label", { class: "format-row" }, noFill, " No colour"));

  // ---- Tabs and buttons
  const tabs: [string, HTMLElement][] = [
    ["Number", numberTab],
    ["Alignment", alignTab],
    ["Font", fontTab],
    ["Fill", fillTab],
  ];
  const tabBar = h("div", { class: "format-tabs", role: "tablist" });
  const show = (i: number) => {
    tabs.forEach(([, panel], j) => (panel.hidden = i !== j));
    Array.from(tabBar.children).forEach((b, j) => b.classList.toggle("active", i === j));
  };
  tabs.forEach(([name], i) => tabBar.append(h("button", { type: "button", class: "format-tab-button", role: "tab", onclick: () => show(i) }, name)));
  show(0);
  const ok = h("button", { type: "button", class: "btn primary" }, "OK");
  const cancel = h("button", { type: "button", class: "btn" }, "Cancel");
  const close = modal(h("div", { class: "dialog format-dialog", role: "dialog", "aria-label": "Format cells" }, h("h2", {}, "Format cells"), tabBar, ...tabs.map(([, panel]) => panel), h("div", { class: "dialog-buttons" }, cancel, ok)), (event) => {
    if (event.key === "Escape") close();
    if (event.key === "Enter" && !(event.target instanceof HTMLSelectElement)) apply();
  });
  const apply = () => {
    close();
    if (Object.keys(change).length) doc.format(range, change);
  };
  ok.addEventListener("click", apply);
  cancel.addEventListener("click", close);
  category.focus();
}
