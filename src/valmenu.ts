// The Data validation dialog, as in Excel: what the selected cells may hold
// (any value, a value from a list, or a whole or decimal number within limits).

import type { CellOp } from "./conditional";
import { ruleValue } from "./cfmenu";
import { selRange, type Doc } from "./doc";
import { h } from "./dom";
import { modal } from "./ui";
import { validationAt, type Validation } from "./validation";

const OPS: [CellOp, string][] = [
  ["between", "between"],
  ["notBetween", "not between"],
  ["equal", "equal to"],
  ["notEqual", "not equal to"],
  ["greaterThan", "greater than"],
  ["lessThan", "less than"],
  ["greaterThanOrEqual", "greater than or equal to"],
  ["lessThanOrEqual", "less than or equal to"],
];

/** A value as the dialog shows it again: formulas with "=", quoted text without its quotes. */
const shown = (f: string) => (/^-?[\d.]+$/.test(f) ? f : f.startsWith('"') ? f.slice(1, -1).replace(/""/g, '"') : `=${f}`);

export function validationDialog(doc: Doc) {
  const range = selRange(doc.sel);
  const current = validationAt(doc.sheet, doc.sel.r, doc.sel.c);
  const allow = h(
    "select",
    { class: "tb-select wide", "aria-label": "Allow" },
    h("option", { value: "any" }, "Any value"),
    h("option", { value: "list" }, "List"),
    h("option", { value: "whole" }, "Whole number"),
    h("option", { value: "decimal" }, "Decimal"),
  ) as HTMLSelectElement;
  const source = h("input", { type: "text", class: "tb-input wide", placeholder: "Yes, No, Maybe   or   =$D$1:$D$5", "aria-label": "Source" }) as HTMLInputElement;
  const op = h("select", { class: "tb-select wide", "aria-label": "Data" }, ...OPS.map(([v, label]) => h("option", { value: v }, label))) as HTMLSelectElement;
  const min = h("input", { type: "text", class: "tb-input wide", "aria-label": "Minimum" }) as HTMLInputElement;
  const max = h("input", { type: "text", class: "tb-input wide", "aria-label": "Maximum" }) as HTMLInputElement;
  const listPart = h("div", {}, h("p", {}, "Source: the choices separated by commas, or the cells that hold them"), source);
  const minLabel = h("p", {}, "Minimum");
  const maxLabel = h("p", {}, "Maximum");
  const numberPart = h("div", {}, h("p", {}, "Data"), op, minLabel, min, maxLabel, max);
  const error = h("p", { class: "dialog-error", role: "alert" });

  if (current?.t === "list") {
    allow.value = "list";
    source.value = current.f.startsWith('"') ? current.f.slice(1, -1) : `=${current.f}`;
  } else if (current && current.t !== "other") {
    allow.value = current.t;
    op.value = current.op;
    min.value = shown(current.f[0] ?? "");
    max.value = shown(current.f[1] ?? "");
  }
  const show = () => {
    listPart.hidden = allow.value !== "list";
    numberPart.hidden = allow.value !== "whole" && allow.value !== "decimal";
    const two = op.value === "between" || op.value === "notBetween";
    minLabel.textContent = two ? "Minimum" : "Value";
    maxLabel.hidden = max.hidden = !two;
  };
  allow.addEventListener("change", show);
  op.addEventListener("change", show);
  show();

  const ok = h("button", { type: "button", class: "btn primary" }, "OK");
  const cancel = h("button", { type: "button", class: "btn" }, "Cancel");
  const clear = h("button", { type: "button", class: "btn subtle" }, "Clear all");
  const close = modal(
    h("div", { class: "dialog", role: "dialog", "aria-label": "Data validation" }, h("h2", {}, "Data validation"), h("p", {}, "Allow"), allow, listPart, numberPart, error, h("div", { class: "dialog-buttons" }, clear, h("span", { class: "spacer" }), cancel, ok)),
    (event) => {
      if (event.key === "Escape") close();
      if (event.key === "Enter") finish();
    },
  );
  const finish = () => {
    let rule: Validation | null = null;
    if (allow.value === "list") {
      const text = source.value.trim();
      if (!text) return void (error.textContent = "Give the choices, or the cells that hold them.");
      rule = { ranges: [], t: "list", f: text.startsWith("=") ? text.slice(1) : `"${text.split(",").map((s) => s.trim()).filter(Boolean).join(",")}"` };
    } else if (allow.value === "whole" || allow.value === "decimal") {
      const two = op.value === "between" || op.value === "notBetween";
      const values = (two ? [min, max] : [min]).map((i) => i.value.trim());
      if (values.some((v) => !v)) return void (error.textContent = two ? "Give a minimum and a maximum." : "Give a value.");
      rule = { ranges: [], t: allow.value, op: op.value as CellOp, f: values.map(ruleValue) };
    }
    close();
    doc.setValidation(range, rule);
  };
  ok.addEventListener("click", finish);
  cancel.addEventListener("click", close);
  clear.addEventListener("click", () => {
    close();
    doc.setValidation(range, null);
  });
  allow.focus();
}
