// Paste special (Ctrl+Alt+V), as in Excel: paste only the values, the
// formulas or the formats of what was copied, optionally turned on its side.

import { h } from "./dom";
import { modal } from "./ui";

export type PasteWhat = "all" | "values" | "formulas" | "formats";

export function pasteSpecial(): Promise<{ what: PasteWhat; transpose: boolean } | null> {
  return new Promise((resolve) => {
    const options: [PasteWhat, string][] = [
      ["all", "All"],
      ["values", "Values (what the formulas show, without the formulas)"],
      ["formulas", "Formulas"],
      ["formats", "Formats"],
    ];
    const radios = options.map(([value, label], i) => {
      const radio = h("input", { type: "radio", name: "paste-what", value, checked: i === 1 }) as HTMLInputElement;
      return h("label", { class: "format-row" }, radio, ` ${label}`);
    });
    const transpose = h("input", { type: "checkbox" }) as HTMLInputElement;
    const ok = h("button", { type: "button", class: "btn primary" }, "OK");
    const cancel = h("button", { type: "button", class: "btn" }, "Cancel");
    const box = h("div", { class: "dialog", role: "dialog", "aria-label": "Paste special" }, h("h2", {}, "Paste special"), ...radios, h("label", { class: "format-row" }, transpose, " Transpose (rows become columns)"), h("div", { class: "dialog-buttons" }, cancel, ok));
    const close = modal(box, (event) => {
      if (event.key === "Escape") done(null);
      if (event.key === "Enter") finish();
    });
    const done = (answer: { what: PasteWhat; transpose: boolean } | null) => {
      close();
      resolve(answer);
    };
    const finish = () => done({ what: (box.querySelector('input[name="paste-what"]:checked') as HTMLInputElement).value as PasteWhat, transpose: transpose.checked });
    ok.addEventListener("click", finish);
    cancel.addEventListener("click", () => done(null));
    (radios[1].querySelector("input") as HTMLInputElement).focus();
  });
}
