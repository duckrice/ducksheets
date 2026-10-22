// Data tools, as on Excel's Data tab: Remove duplicates and Text to columns.

import { selRange, type Doc } from "./doc";
import { h } from "./dom";
import { address, colName, parseAddress, type Range } from "./model";
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

/** Excel's table colours: the header's, and the light one for every other row. */
const TABLE_STYLES: [string, string, string][] = [
  ["Blue", "#4472C4", "#D9E1F2"],
  ["Orange", "#ED7D31", "#FCE4D6"],
  ["Green", "#70AD47", "#E2EFDA"],
  ["Grey", "#7F7F7F", "#EDEDED"],
  ["Gold", "#BF8F00", "#FFF2CC"],
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

export function dataMenu(doc: Doc): MenuItem[] {
  return [
    {
      label: "Format as table",
      items: TABLE_STYLES.map(([label, header, band]) => ({ label, action: () => doc.formatAsTable(block(doc), header, band) })),
    },
    { label: "Goal seek…", action: () => goalSeek(doc) },
    { label: "Remove duplicates…", action: () => removeDuplicates(doc) },
    { label: "Text to columns…", action: () => textToColumns(doc) },
  ];
}
