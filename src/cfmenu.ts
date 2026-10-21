// The Conditional formatting menu, laid out as Excel's: highlight rules,
// top and bottom rules, data bars, colour scales, a formula of your own, and
// clearing rules. Rules apply to the selected cells.

import { PRESETS, type CellOp, type CfRule } from "./conditional";
import { selRange, type Doc } from "./doc";
import { h } from "./dom";
import { address } from "./model";
import { modal, type MenuItem } from "./ui";

const BARS: [string, string][] = [
  ["Blue", "#638EC6"],
  ["Green", "#63C384"],
  ["Red", "#FF555A"],
  ["Orange", "#FFB628"],
  ["Light blue", "#008AEF"],
  ["Purple", "#D6007B"],
];

const SCALES: [string, string[]][] = [
  ["Green – yellow – red", ["#63BE7B", "#FFEB84", "#F8696B"]],
  ["Red – yellow – green", ["#F8696B", "#FFEB84", "#63BE7B"]],
  ["Green – white – red", ["#63BE7B", "#FCFCFF", "#F8696B"]],
  ["Red – white – green", ["#F8696B", "#FCFCFF", "#63BE7B"]],
  ["Blue – white – red", ["#5A8AC6", "#FCFCFF", "#F8696B"]],
  ["White – green", ["#FCFCFF", "#63BE7B"]],
];

/** What was typed as a rule's value, as it goes in the file: a number, a formula (typed with =), or quoted text. */
export function ruleValue(text: string): string {
  const t = text.trim();
  if (t.startsWith("=")) return t.slice(1);
  if (t !== "" && !Number.isNaN(Number(t.replace(",", ".")))) return String(Number(t.replace(",", ".")));
  return `"${t.replace(/"/g, '""')}"`;
}

/** Asks for a rule's values and its format; resolves to the values and the format's index, or null. */
function ask(title: string, prompt: string, fields: string[], defaults: string[]): Promise<{ values: string[]; preset: number } | null> {
  return new Promise((resolve) => {
    const inputs = fields.map((label, i) => h("input", { type: "text", class: "tb-input wide", value: defaults[i] ?? "", "aria-label": label }) as HTMLInputElement);
    const look = h("select", { class: "tb-select wide", "aria-label": "Format" }, ...PRESETS.map((p, i) => h("option", { value: String(i) }, p.name))) as HTMLSelectElement;
    const ok = h("button", { type: "button", class: "btn primary" }, "OK");
    const cancel = h("button", { type: "button", class: "btn" }, "Cancel");
    const rows = inputs.flatMap((input, i) => (i ? [h("p", {}, fields[i]), input] : [input]));
    const close = modal(
      h("div", { class: "dialog cf-dialog", role: "dialog", "aria-label": title }, h("h2", {}, title), h("p", {}, prompt), ...rows, h("p", {}, "with"), look, h("div", { class: "dialog-buttons" }, cancel, ok)),
      (event) => {
        if (event.key === "Escape") done(null);
        if (event.key === "Enter") finish();
      },
    );
    const done = (answer: { values: string[]; preset: number } | null) => {
      close();
      resolve(answer);
    };
    const finish = () => {
      if (inputs.some((i) => !i.value.trim())) return inputs.find((i) => !i.value.trim())!.focus();
      done({ values: inputs.map((i) => i.value), preset: Number(look.value) });
    };
    ok.addEventListener("click", finish);
    cancel.addEventListener("click", () => done(null));
    inputs[0]?.focus();
    inputs[0]?.select();
  });
}

export function conditionalMenu(doc: Doc): MenuItem[] {
  const range = () => selRange(doc.sel);
  /** A middle value of the selection, offered as a starting point, as Excel does. */
  const middle = () => {
    const r = range();
    const nums: number[] = [];
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++)
      for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 200); col++) {
        const v = doc.value(row, col);
        if (typeof v === "number") nums.push(v);
      }
    nums.sort((a, b) => a - b);
    return nums.length ? String(nums[Math.floor(nums.length / 2)]) : "";
  };
  const add = (rule: CfRule) => doc.addRule(range(), rule);
  const cell = (op: CellOp, title: string, prompt: string, two = false) => async () => {
    const answer = await ask(title, prompt, two ? ["From", "and"] : ["Value"], two ? [middle(), middle()] : [middle()]);
    if (answer) add({ t: "cell", op, f: answer.values.map(ruleValue), dxf: doc.dxf(PRESETS[answer.preset].style) });
  };
  const simple = (title: string, prompt: string, make: (dxf: number) => CfRule) => async () => {
    const answer = await ask(title, prompt, [], []);
    if (answer) add(make(doc.dxf(PRESETS[answer.preset].style)));
  };
  const top = (bottom: boolean, percent: boolean) => async () => {
    const answer = await ask(`${bottom ? "Bottom" : "Top"} ${percent ? "10%" : "10 items"}`, `Format cells that rank in the ${bottom ? "bottom" : "top"}${percent ? " (percent)" : ""}:`, ["Count"], ["10"]);
    const n = answer && Math.max(1, Math.round(Number(answer.values[0])));
    if (answer && n) add({ t: "top", n, bottom, percent, dxf: doc.dxf(PRESETS[answer.preset].style) });
  };
  return [
    {
      label: "Highlight cells rules",
      items: [
        { label: "Greater than…", action: cell("greaterThan", "Greater than", "Format cells that are greater than:") },
        { label: "Less than…", action: cell("lessThan", "Less than", "Format cells that are less than:") },
        { label: "Between…", action: cell("between", "Between", "Format cells that are between:", true) },
        { label: "Equal to…", action: cell("equal", "Equal to", "Format cells that are equal to:") },
        {
          label: "Text that contains…",
          action: async () => {
            const answer = await ask("Text that contains", "Format cells that contain the text:", ["Text"], [""]);
            if (answer) add({ t: "text", text: answer.values[0], dxf: doc.dxf(PRESETS[answer.preset].style) });
          },
        },
        { label: "Duplicate values…", action: simple("Duplicate values", "Format the cells whose value appears more than once.", (dxf) => ({ t: "duplicate", dxf })) },
        { label: "Unique values…", action: simple("Unique values", "Format the cells whose value appears only once.", (dxf) => ({ t: "unique", dxf })) },
      ],
    },
    {
      label: "Top/bottom rules",
      items: [
        { label: "Top 10 items…", action: top(false, false) },
        { label: "Top 10%…", action: top(false, true) },
        { label: "Bottom 10 items…", action: top(true, false) },
        { label: "Bottom 10%…", action: top(true, true) },
        { label: "Above average…", action: simple("Above average", "Format the cells above the average of the selection.", (dxf) => ({ t: "average", below: false, dxf })) },
        { label: "Below average…", action: simple("Below average", "Format the cells below the average of the selection.", (dxf) => ({ t: "average", below: true, dxf })) },
      ],
    },
    { label: "Data bars", items: BARS.map(([label, color]) => ({ label, action: () => add({ t: "bar", color }) })) },
    { label: "Colour scales", items: SCALES.map(([label, colors]) => ({ label, action: () => add({ t: "scale", colors }) })) },
    {
      label: "Use a formula…",
      action: async () => {
        const r = range();
        const at = address(r.r1, r.c1);
        const answer = await ask("Use a formula", `Format cells where this formula is true (written for the first cell, ${at}; it moves along like a copied formula):`, ["Formula"], [`=${at}>0`]);
        if (answer) add({ t: "formula", f: answer.values[0].trim().replace(/^=/, ""), dxf: doc.dxf(PRESETS[answer.preset].style) });
      },
    },
    {
      label: "Clear rules",
      items: [
        { label: "From the selected cells", action: () => doc.clearRules(range()) },
        { label: "From the entire sheet", action: () => doc.clearRules(null) },
      ],
    },
  ];
}
