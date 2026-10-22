// The toolbar: files, undo, the font, colours, alignment, number formats,
// borders, and the data tools (sum, sort, filter, freeze). Everything acts on
// the selected cells, as in Excel's Home tab.

import { selRange, type Doc } from "./doc";
import type { Grid } from "./grid";
import { h } from "./dom";
import { icon, type IconName } from "./icons";
import { MAX_COLS, MAX_ROWS, type Range } from "./model";
import { locale } from "./numfmt";
import { menu, popover } from "./ui";
import { conditionalMenu } from "./cfmenu";
import { dataMenu, tableMenu } from "./datatools";
import { validationDialog } from "./valmenu";
import { lookOf, type Border, type BorderSide, type StyleChange } from "./xlsx/styles";

export interface ToolbarActions {
  newBook(): void;
  open(): void;
  save(): void;
  autoSum(): void;
  find(): void;
  chart(): void;
  picture(): void;
  print(): void;
  /** Format painter: starts copying formats, or (when on) stops. */
  painter(): void;
  painting(): boolean;
}

const FONTS = ["Calibri", "Aptos", "Arial", "Cambria", "Consolas", "Georgia", "Segoe UI", "Times New Roman", "Verdana"];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
const PALETTE = [
  ["#000000", "#FFFFFF", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
  ["#7F7F7F", "#F2F2F2", "#D0CECE", "#D6DCE4", "#D9E2F3", "#FBE5D5", "#EDEDED", "#FFF2CC", "#DEEBF6", "#E2EFD9"],
  ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"],
];

/** The number formats offered by name, as in Excel's list. */
export function numberFormats(): [string, string][] {
  const kr = locale.decimal === ",";
  return [
    ["General", "General"],
    ["Number", "#,##0.00"],
    ["Currency", kr ? '#,##0.00 "kr"' : "$#,##0.00"],
    ["Accounting", kr ? '_-* #,##0.00 "kr"_-;-* #,##0.00 "kr"_-;_-* "-"?? "kr"_-;_-@_-' : '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'],
    ["Short date", locale.dateCode],
    ["Long date", "dddd d. mmmm yyyy"],
    ["Time", "hh:mm:ss"],
    ["Percentage", "0.00%"],
    ["Fraction", "# ?/?"],
    ["Scientific", "0.00E+00"],
    ["Text", "@"],
  ];
}

/** One more or one fewer decimal in a format code (the first number in each section). */
export function changeDecimals(code: string, delta: number): string {
  if (code === "General" || code === "@") return delta > 0 ? "0.0" : "0";
  return code
    .split(";")
    .map((section) => {
      // Only outside quoted text.
      let done = false;
      return section
        .split(/("[^"]*")/)
        .map((part) =>
          part.startsWith('"') || done
            ? part
            : part.replace(/0(?:\.(0+))?/, (_m, decimals: string = "") => {
                done = true;
                const n = Math.max(0, decimals.length + delta);
                return n ? `0.${"0".repeat(n)}` : "0";
              }),
        )
        .join("");
    })
    .join(";");
}

/** Every cell of the range (up to the data, for whole rows and columns) gets the border its place in the range calls for. */
function borderChange(kind: string, range: Range): (s: number | undefined, r: number, c: number) => StyleChange {
  const thin: BorderSide = { style: "thin" };
  const thick: BorderSide = { style: "medium" };
  return (_s, r, c) => {
    const edges = { top: r === range.r1, bottom: r === range.r2, left: c === range.c1, right: c === range.c2 };
    let border: Border = {};
    if (kind === "all") border = { top: thin, bottom: thin, left: thin, right: thin };
    else if (kind === "outside")
      border = { ...(edges.top ? { top: thin } : {}), ...(edges.bottom ? { bottom: thin } : {}), ...(edges.left ? { left: thin } : {}), ...(edges.right ? { right: thin } : {}) };
    else if (kind === "thick") border = { ...(edges.top ? { top: thick } : {}), ...(edges.bottom ? { bottom: thick } : {}), ...(edges.left ? { left: thick } : {}), ...(edges.right ? { right: thick } : {}) };
    else if (kind === "bottom") border = edges.bottom ? { bottom: thin } : {};
    else if (kind === "top") border = edges.top ? { top: thin } : {};
    else if (kind === "double") border = edges.bottom ? { bottom: { style: "double" } } : {};
    return { border };
  };
}

export function createToolbar(doc: Doc, grid: Grid, actions: ToolbarActions) {
  const updaters: (() => void)[] = [];
  const range = () => grid.expanded(selRange(doc.sel));
  const active = () => lookOf(doc.book.styles, doc.cell(doc.sel.r, doc.sel.c)?.s);
  const apply = (change: StyleChange | ((s: number | undefined, r: number, c: number) => StyleChange)) => {
    doc.format(range(), change);
    grid.el.focus();
  };

  const button = (name: IconName, title: string, onclick: () => void, isActive?: () => boolean) => {
    const el = h("button", { type: "button", class: "tb", title, "aria-label": title.replace(/ \(.*\)$/, ""), html: icon(name) });
    el.addEventListener("mousedown", (e) => e.preventDefault());
    el.addEventListener("click", onclick);
    if (isActive) updaters.push(() => el.classList.toggle("active", isActive()));
    return el;
  };
  const textButton = (label: string, title: string, onclick: () => void) => {
    const el = h("button", { type: "button", class: "tb text", title, "aria-label": title }, label);
    el.addEventListener("mousedown", (e) => e.preventDefault());
    el.addEventListener("click", onclick);
    return el;
  };

  // Font and size.
  const font = h("select", { class: "tb-select font-select", title: "Font", "aria-label": "Font" }, ...FONTS.map((f) => h("option", { value: f }, f))) as HTMLSelectElement;
  font.addEventListener("change", () => apply({ fontName: font.value }));
  const size = h("select", { class: "tb-select size-select", title: "Font size (Alt+scroll wheel)", "aria-label": "Font size" }, ...SIZES.map((s) => h("option", { value: String(s) }, String(s)))) as HTMLSelectElement;
  size.addEventListener("change", () => apply({ fontSize: Number(size.value) }));
  updaters.push(() => {
    const f = active().font;
    if (!FONTS.includes(f.name) && !Array.from(font.options).some((o) => o.value === f.name)) font.append(h("option", { value: f.name }, f.name));
    font.value = f.name;
    if (!Array.from(size.options).some((o) => o.value === String(f.size))) size.append(h("option", { value: String(f.size) }, String(f.size)));
    size.value = String(f.size);
  });

  const palette = (anchor: HTMLElement, choose: (color: string | null) => void, none: string) =>
    popover(anchor, (close) =>
      h(
        "div",
        { class: "palette-grid" },
        ...PALETTE.flat().map((color) =>
          h("button", {
            type: "button",
            class: "swatch",
            title: color,
            "aria-label": color,
            style: `--swatch:${color}`,
            onmousedown: (e: Event) => e.preventDefault(),
            onclick: () => {
              close();
              choose(color);
            },
          }),
        ),
        h("button", { type: "button", class: "btn subtle small wide", onclick: () => (close(), choose(null)) }, none),
      ),
    );
  /**
   * A colour button as Excel's: clicking it gives the selection the last colour chosen (shown as a
   * bar under the icon), the arrow beside it opens the palette.
   */
  const colorButton = (name: IconName, title: string, first: string, give: (color: string | null) => void, none: string) => {
    let last: string | null = first;
    const bar = h("span", { class: "color-bar" });
    const show = () => bar.style.setProperty("--bar", last ?? "transparent");
    show();
    const main = h("button", { type: "button", class: "tb color-main", title: `${title}`, "aria-label": title, html: icon(name) }, bar);
    const more = h("button", { type: "button", class: "tb color-more", title: `More colours for ${title.toLowerCase()}`, "aria-label": `Choose the ${title.toLowerCase()}`, html: icon("caret") });
    for (const el of [main, more]) el.addEventListener("mousedown", (e) => e.preventDefault());
    main.addEventListener("click", () => give(last));
    more.addEventListener("click", () =>
      palette(more, (color) => {
        last = color;
        show();
        give(color);
      }, none),
    );
    return h("span", { class: "color-split" }, main, more);
  };
  const textColor = colorButton("textColor", "Text colour", "#C00000", (color) => apply({ color }), "Automatic");
  const fillColor = colorButton("fillColor", "Fill colour", "#FFFF00", (fill) => apply({ fill }), "No fill");

  const conditional = button("conditional", "Conditional formatting", () => menu(conditional, conditionalMenu(doc)));
  const dataTools = button("table", "Data tools: remove duplicates, text to columns, summarize, goal seek", () => menu(dataTools, dataMenu(doc)));
  // Format as table, out in the open: a labelled button with the styles to pick from.
  const tableButton = h("button", { type: "button", class: "tb labelled", title: "Format as table: a coloured header, striped rows and filter buttons for the selected data", html: icon("formatTable") }, "Table");
  tableButton.addEventListener("mousedown", (e) => e.preventDefault());
  tableButton.addEventListener("click", () => menu(tableButton, tableMenu(doc)));

  // Number formats.
  const formats = numberFormats();
  const numFmt = h("select", { class: "tb-select format-select", title: "Number format", "aria-label": "Number format" }, ...formats.map(([name]) => h("option", { value: name }, name)), h("option", { value: "custom", disabled: true }, "Custom")) as HTMLSelectElement;
  numFmt.addEventListener("change", () => apply({ numFmt: formats.find(([name]) => name === numFmt.value)![1] }));
  updaters.push(() => {
    const code = active().numFmt;
    numFmt.value = formats.find(([, c]) => c === code)?.[0] ?? (code.includes("%") ? "Percentage" : "custom");
  });

  const borders = button("borders", "Borders", () =>
    menu(borders, [
      { label: "All borders", action: () => apply(borderChange("all", range())) },
      { label: "Outside borders", action: () => apply(borderChange("outside", range())) },
      { label: "Thick outside borders", action: () => apply(borderChange("thick", range())) },
      { label: "Bottom border", action: () => apply(borderChange("bottom", range())) },
      { label: "Top border", action: () => apply(borderChange("top", range())) },
      { label: "Double bottom border", action: () => apply(borderChange("double", range())) },
      { label: "No border", action: () => apply({ border: {} }) },
    ]),
  );

  const freeze = button("freeze", "Freeze panes", () =>
    menu(freeze, [
      { label: "Freeze panes (above and left of the selected cell)", action: () => doc.freeze(doc.sel.r, doc.sel.c) },
      { label: "Freeze top row", action: () => doc.freeze(1, 0) },
      { label: "Freeze first column", action: () => doc.freeze(0, 1) },
      { label: "Unfreeze panes", action: () => doc.freeze(0, 0) },
    ]),
  );

  const sortBy = (descending: boolean) => {
    const sel = selRange(doc.sel);
    const area = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? doc.region(sel.r1, sel.c1) : sel;
    if (area.r2 >= MAX_ROWS - 1 || area.c2 >= MAX_COLS - 1) return;
    doc.sort(area, doc.sel.c, descending);
  };

  const el = h(
    "div",
    { class: "toolbar", role: "toolbar", "aria-label": "Home" },
    h(
      "div",
      { class: "tb-group" },
      button("newDoc", "New workbook (Ctrl+N)", actions.newBook),
      button("open", "Open (Ctrl+O)", actions.open),
      button("save", "Save (Ctrl+S)", actions.save),
      button("print", "Print, or save as PDF (Ctrl+P)", actions.print),
    ),
    h("div", { class: "tb-group" }, button("undo", "Undo (Ctrl+Z)", () => doc.undo()), button("redo", "Redo (Ctrl+Y)", () => doc.redo())),
    h(
      "div",
      { class: "tb-group" },
      font,
      size,
      button("bold", "Bold (Ctrl+B)", () => apply({ bold: !active().font.bold }), () => !!active().font.bold),
      button("italic", "Italic (Ctrl+I)", () => apply({ italic: !active().font.italic }), () => !!active().font.italic),
      button("underline", "Underline (Ctrl+U)", () => apply({ underline: !active().font.underline }), () => !!active().font.underline),
      button("strike", "Strikethrough (Ctrl+5)", () => apply({ strike: !active().font.strike }), () => !!active().font.strike),
      textColor,
      fillColor,
      button("painter", "Format painter: copy the formats, then click or drag over the cells to give them", actions.painter, actions.painting),
    ),
    h(
      "div",
      { class: "tb-group" },
      button("alignLeft", "Align left", () => apply({ h: active().h === "left" ? null : "left" }), () => active().h === "left"),
      button("alignCenter", "Centre", () => apply({ h: active().h === "center" ? null : "center" }), () => active().h === "center"),
      button("alignRight", "Align right", () => apply({ h: active().h === "right" ? null : "right" }), () => active().h === "right"),
      button("alignTop", "Top", () => apply({ v: "top" }), () => active().v === "top"),
      button("alignMiddle", "Middle", () => apply({ v: "center" }), () => active().v === "center"),
      button("alignBottom", "Bottom", () => apply({ v: null }), () => !active().v || active().v === "bottom"),
      button("wrap", "Wrap text", () => apply({ wrap: !active().wrap }), () => !!active().wrap),
      button("merge", "Merge and centre", () => doc.toggleMerge(range()), () => doc.sheet.merges.some((m) => m.r1 === doc.sel.r && m.c1 === doc.sel.c)),
    ),
    h(
      "div",
      { class: "tb-group" },
      numFmt,
      textButton(locale.decimal === "," ? "kr" : "$", "Currency (Ctrl+Shift+4)", () => apply({ numFmt: formats[2][1] })),
      textButton("%", "Percentage (Ctrl+Shift+5)", () => apply({ numFmt: "0%" })),
      textButton(locale.decimal === "," ? "000" : ",", "Thousands separator", () => apply({ numFmt: "#,##0" })),
      button("moreDecimals", "More decimals", () => apply((s) => ({ numFmt: changeDecimals(lookOf(doc.book.styles, s).numFmt, 1) }))),
      button("fewerDecimals", "Fewer decimals", () => apply((s) => ({ numFmt: changeDecimals(lookOf(doc.book.styles, s).numFmt, -1) }))),
      borders,
    ),
    h(
      "div",
      { class: "tb-group" },
      button("sum", "AutoSum (Alt+=)", actions.autoSum),
      button("sortAsc", "Sort A to Z (smallest first)", () => sortBy(false)),
      button("sortDesc", "Sort Z to A (largest first)", () => sortBy(true)),
      button("filter", "Filter (Ctrl+Shift+L)", () => doc.toggleFilter(), () => !!doc.sheet.filter),
      tableButton,
      conditional,
      dataTools,
      button("validation", "Data validation: a list to choose from, or limits on numbers", () => validationDialog(doc)),
      freeze,
      button("chart", "Insert a chart of the selected data", actions.chart),
      button("picture", "Insert a picture (or paste one with Ctrl+V)", actions.picture),
      button("find", "Find and replace (Ctrl+F)", actions.find),
    ),
  );

  return {
    el,
    update() {
      for (const u of updaters) u();
    },
  };
}

/** The numbers above (or else to the left of) a cell, for AutoSum: from the cell up to the first gap or text. */
export function sumRange(doc: Doc, r: number, c: number): Range | null {
  const isNumber = (rr: number, cc: number) => rr >= 0 && cc >= 0 && typeof doc.value(rr, cc) === "number";
  if (isNumber(r - 1, c)) {
    let top = r - 1;
    while (isNumber(top - 1, c)) top--;
    return { r1: top, c1: c, r2: r - 1, c2: c };
  }
  if (isNumber(r, c - 1)) {
    let left = c - 1;
    while (isNumber(r, left - 1)) left--;
    return { r1: r, c1: left, r2: r, c2: c - 1 };
  }
  return null;
}

