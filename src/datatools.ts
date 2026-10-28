// Data tools, as on Excel's Data tab: Remove duplicates and Text to columns.

import { selRange, type Doc, type SummaryFn } from "./doc";
import { h } from "./dom";
import { address, colName, parseAddress, type FilterOp, type FilterRule, type Range } from "./model";
import { modal, toast, type MenuItem } from "./ui";

/** The selection, or (for a single cell) the block of data around it, as Excel does. */
function block(doc: Doc): Range {
  const r = selRange(doc.sel);
  return r.r1 === r.r2 && r.c1 === r.c2 ? doc.region(r.r1, r.c1) : r;
}

function dialog(title: string, body: HTMLElement[], onOk: () => void, ok = "OK") {
  const yes = h("button", { type: "button", class: "btn primary" }, ok);
  const no = h("button", { type: "button", class: "btn" }, "Cancel");
  const close = modal(h("div", { class: "dialog", role: "dialog", "aria-label": title }, h("h2", {}, title), ...body, h("div", { class: "dialog-buttons" }, no, yes)), (event) => {
    if (event.key === "Escape") close();
    if (event.key === "Enter") done();
  });
  const done = () => {
    close();
    onOk();
  };
  yes.addEventListener("click", done);
  no.addEventListener("click", close);
  (body.find((el) => el.querySelector("input, select"))?.querySelector("input, select") as HTMLElement | null)?.focus();
}

function removeDuplicates(doc: Doc) {
  const range = block(doc);
  const header = h("input", { type: "checkbox", checked: doc.hasHeader(range) }) as HTMLInputElement;
  const columns = Array.from({ length: Math.min(range.c2 - range.c1 + 1, 50) }, (_, i) => range.c1 + i);
  const name = (c: number) => (header.checked && doc.value(range.r1, c) !== null ? String(doc.value(range.r1, c)) : `Column ${colName(c)}`);
  const boxes = columns.map((c) => h("input", { type: "checkbox", checked: true, "data-col": c }) as HTMLInputElement);
  const labels = columns.map((c, i) => h("label", { class: "format-row" }, boxes[i], ` ${name(c)}`));
  header.addEventListener("change", () => columns.forEach((c, i) => (labels[i].lastChild!.textContent = ` ${name(c)}`)));
  dialog(
    "Remove duplicates",
    [h("p", {}, "Rows whose values repeat in the ticked columns are removed; the first of each stays."), h("label", { class: "format-row" }, header, " My data has headers"), h("div", { class: "column-list" }, ...labels)],
    () => {
      const cols = columns.filter((_, i) => boxes[i].checked);
      if (!cols.length) return toast("Tick at least one column.");
      const top = header.checked ? range.r1 + 1 : range.r1;
      const removed = doc.removeDuplicates(range, cols, header.checked);
      const left = range.r2 - top + 1 - removed;
      toast(removed ? `Removed ${removed} duplicate ${removed === 1 ? "row" : "rows"}; ${left} unique ${left === 1 ? "row remains" : "rows remain"}.` : "No duplicates found.");
    },
  );
}

function textToColumns(doc: Doc) {
  const range = selRange(doc.sel);
  const options: [string, string][] = [
    [",", "Comma"],
    [";", "Semicolon"],
    ["\t", "Tab"],
    [" ", "Space"],
    ["other", "Other:"],
  ];
  const radios = options.map(([value], i) => h("input", { type: "radio", name: "split-at", value, checked: i === 0 }) as HTMLInputElement);
  const other = h("input", { type: "text", class: "tb-input", maxlength: 3, "aria-label": "Other separator", style: "width: 60px" }) as HTMLInputElement;
  other.addEventListener("input", () => (radios[radios.length - 1].checked = true));
  dialog(
    "Text to columns",
    [h("p", {}, `Splits the text in column ${colName(range.c1)} into the columns to its right, at:`), ...options.map(([, label], i) => h("label", { class: "format-row" }, radios[i], ` ${label}`, i === options.length - 1 ? other : null))],
    () => {
      const chosen = radios.find((r) => r.checked)!.value;
      doc.textToColumns({ ...range, c2: range.c1 }, chosen === "other" ? other.value : chosen);
    },
  );
}

function summarize(doc: Doc) {
  const range = block(doc);
  if (range.r2 <= range.r1) return toast("Select data with a header row first.");
  const columns = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i);
  const pick = (label: string, chosen: number) => {
    const select = h("select", { class: "tb-input", "aria-label": label }, ...columns.map((c) => h("option", { value: c, selected: c === chosen }, String(doc.value(range.r1, c) ?? `Column ${colName(c)}`)))) as HTMLSelectElement;
    return [select, h("label", { class: "format-row" }, `${label} `, select)] as const;
  };
  const [by, byRow] = pick("Group by", range.c1);
  const [of, ofRow] = pick("Values from", columns.find((c) => typeof doc.value(range.r1 + 1, c) === "number") ?? range.c2);
  const fns: [SummaryFn, string][] = [["sum", "Sum"], ["count", "Count"], ["average", "Average"], ["max", "Largest"], ["min", "Smallest"]];
  const fn = h("select", { class: "tb-input", "aria-label": "Work out" }, ...fns.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
  dialog(
    "Summarize",
    [h("p", {}, "Makes a new sheet with each value of one column once, and the sum (or count, average, …) of another for it. It updates when the data changes, and Excel opens it as a PivotTable."), byRow, h("label", { class: "format-row" }, "Work out ", fn), ofRow],
    () => doc.summarize(range, Number(by.value), Number(of.value), fn.value as SummaryFn),
    "Summarize",
  );
}

/** The table styles to choose from, for the data around the selection. */
export function tableMenu(doc: Doc): MenuItem[] {
  return TABLE_STYLES.map(([label, header, band, style]) => ({ label, swatch: [header, band], action: () => doc.formatAsTable(block(doc), header, band, style) }));
}

/** Excel's table colours: the header's, the light one for every other row, and its own style of that colour. */
const TABLE_STYLES: [string, string, string, string][] = [
  ["Blue", "#4472C4", "#D9E1F2", "TableStyleMedium2"],
  ["Orange", "#ED7D31", "#FCE4D6", "TableStyleMedium3"],
  ["Green", "#70AD47", "#E2EFDA", "TableStyleMedium7"],
  ["Grey", "#7F7F7F", "#EDEDED", "TableStyleMedium4"],
  ["Gold", "#BF8F00", "#FFF2CC", "TableStyleMedium5"],
];

/** Goal Seek: asks which formula should give what, by changing which cell. */
function goalSeek(doc: Doc) {
  const here = address(doc.sel.r, doc.sel.c);
  const set = h("input", { type: "text", class: "tb-input wide", value: here, "aria-label": "Set cell" }) as HTMLInputElement;
  const to = h("input", { type: "text", class: "tb-input wide", "aria-label": "To value" }) as HTMLInputElement;
  const by = h("input", { type: "text", class: "tb-input wide", "aria-label": "By changing cell" }) as HTMLInputElement;
  dialog(
    "Goal seek",
    [h("p", {}, "Set cell (a formula)"), set, h("p", {}, "To value"), to, h("p", {}, "By changing cell (a number)"), by],
    () => {
      const target = parseAddress(set.value.replace(/\$/g, ""));
      const input = parseAddress(by.value.replace(/\$/g, ""));
      const goal = Number(to.value.replace(",", "."));
      if (!target || !input || !to.value.trim() || Number.isNaN(goal)) return toast("Give two cells and a number.");
      if (doc.cell(target.r, target.c)?.f === undefined) return toast(`${set.value} needs a formula.`);
      if (doc.cell(input.r, input.c)?.f !== undefined) return toast(`${by.value} must hold a number, not a formula.`);
      const found = doc.goalSeek(target, goal, input);
      toast(found === null ? "Goal Seek didn't find an answer." : `Found it: ${by.value.toUpperCase()} = ${Math.round(found * 1e6) / 1e6}.`);
    },
  );
}

/** Sort by up to three columns in turn, as Excel's Sort dialog. */
function sortDialog(doc: Doc) {
  const range = block(doc);
  const header = h("input", { type: "checkbox", checked: doc.hasHeader(range) }) as HTMLInputElement;
  const columns = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i);
  const name = (c: number) => (header.checked ? String(doc.value(range.r1, c) ?? `Column ${colName(c)}`) : `Column ${colName(c)}`);
  const levels = ["Sort by", "Then by", "Then by"].map((label, i) => {
    const col = h("select", { class: "tb-input", "aria-label": `${label} column` }) as HTMLSelectElement;
    const order = h("select", { class: "tb-input", "aria-label": `${label} order` }, h("option", { value: "asc" }, "A to Z, smallest first"), h("option", { value: "desc" }, "Z to A, largest first")) as HTMLSelectElement;
    const fill = () => {
      const was = col.value;
      col.replaceChildren(...(i ? [h("option", { value: "" }, "(none)")] : []), ...columns.map((c) => h("option", { value: c }, name(c))));
      col.value = was || (i ? "" : String(doc.sel.c >= range.c1 && doc.sel.c <= range.c2 ? doc.sel.c : range.c1));
    };
    fill();
    return { col, order, fill, row: h("label", { class: "format-row" }, `${label} `, col, " ", order) };
  });
  header.addEventListener("change", () => levels.forEach((l) => l.fill()));
  dialog(
    "Sort",
    [h("label", { class: "format-row" }, header, " My data has headers"), ...levels.map((l) => l.row)],
    () => doc.sort(range, levels.filter((l) => l.col.value !== "").map((l) => ({ col: Number(l.col.value), descending: l.order.value === "desc" })), false, header.checked),
  );
}

/** What the custom filter dialog offers, and the condition each makes from the value typed. */
const FILTER_CHOICES: [string, string, (v: string) => FilterRule][] = [
  ["equal", "equals", (v) => one("equal", v)],
  ["notEqual", "does not equal", (v) => one("notEqual", v)],
  ["greaterThan", "is greater than", (v) => one("greaterThan", v)],
  ["greaterThanOrEqual", "is greater than or equal to", (v) => one("greaterThanOrEqual", v)],
  ["lessThan", "is less than", (v) => one("lessThan", v)],
  ["lessThanOrEqual", "is less than or equal to", (v) => one("lessThanOrEqual", v)],
  ["begins", "begins with", (v) => one("equal", `${v}*`)],
  ["ends", "ends with", (v) => one("equal", `*${v}`)],
  ["contains", "contains", (v) => one("equal", `*${v}*`)],
  ["notContains", "does not contain", (v) => one("notEqual", `*${v}*`)],
  ["top", "top (number of items)", (v) => ({ kind: "top", top: true, percent: false, count: Number(v) })],
  ["bottom", "bottom (number of items)", (v) => ({ kind: "top", top: false, percent: false, count: Number(v) })],
  ["topPercent", "top %", (v) => ({ kind: "top", top: true, percent: true, count: Number(v) })],
];
const one = (op: FilterOp, val: string): FilterRule => ({ kind: "custom", and: false, conditions: [{ op, val }] });

/** Excel's "Custom filter": show the rows whose value in column `col` of the filter matches a condition. */
export function customFilter(doc: Doc, col: number) {
  const filter = doc.sheet.filter;
  if (!filter) return;
  const offset = col - filter.range.c1;
  const how = h("select", { class: "tb-input", "aria-label": "Show rows where the value" }, ...FILTER_CHOICES.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
  const value = h("input", { type: "text", class: "tb-input wide", "aria-label": "Value" }) as HTMLInputElement;
  dialog("Custom filter", [h("p", {}, `Show rows where ${String(doc.value(filter.range.r1, col) ?? `column ${colName(col)}`)}`), how, value], () => {
    const rule = FILTER_CHOICES.find(([v]) => v === how.value)![2](value.value.trim());
    if (rule.kind === "top" && !(rule.count > 0)) return toast("Give how many, as a number.");
    doc.setCustomFilter(offset, rule);
  });
}

export function dataMenu(doc: Doc): MenuItem[] {
  return [
    { label: "Sort…", action: () => sortDialog(doc) },
    { label: "Format as table", items: tableMenu(doc) },
    { label: "Summarize (PivotTable)…", action: () => summarize(doc) },
    { label: "Goal seek…", action: () => goalSeek(doc) },
    {
      label: "Flash Fill (Ctrl+E)",
      action: () => {
        const filled = doc.flashFill(doc.sel.r, doc.sel.c);
        toast(filled === null ? "Flash Fill: type a result or two next to the data first, then Ctrl+E." : `Flash Fill filled ${filled} ${filled === 1 ? "cell" : "cells"}.`);
      },
    },
    { label: "Remove duplicates…", action: () => removeDuplicates(doc) },
    { label: "Text to columns…", action: () => textToColumns(doc) },
  ];
}
