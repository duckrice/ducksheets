// Cell formats as Excel keeps them (styles.xml): lists of fonts, fills and
// borders, and cell formats (xfs) that point into them. A cell's style is an
// index into the xfs. Formats read from a file are written back exactly as
// they were; changes add new entries rather than altering shared ones.

import { builtinFormat } from "../numfmt";

export interface Font {
  name: string;
  /** In points. */
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** #rrggbb */
  color?: string;
}

export interface BorderSide {
  style: string;
  color?: string;
}

export interface Border {
  left?: BorderSide;
  right?: BorderSide;
  top?: BorderSide;
  bottom?: BorderSide;
}

export interface Xf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  h?: string;
  v?: string;
  wrap?: boolean;
  indent?: number;
}

export interface Styles {
  /** Custom number formats by id (164 and up); built-in ones aren't listed. */
  numFmts: Map<number, string>;
  fonts: Font[];
  /** The fill colour of each fill (#rrggbb), or null for none or a pattern. */
  fills: (string | null)[];
  borders: Border[];
  xfs: Xf[];
  /** The formats conditional formatting rules use, as XML (<dxf>…</dxf>), by dxfId. */
  dxfs: string[];
  /** The XML each entry was read from, written back unchanged if it wasn't changed. */
  raw: { fonts: (string | null)[]; fills: (string | null)[]; borders: (string | null)[]; xfs: (string | null)[] };
  /** The rest of styles.xml (cell styles, table styles, …), kept as it was. */
  source: string | null;
  /** Theme colours, in Excel's theme index order. */
  theme: string[];
}

/** A fresh workbook's styles: Calibri 11, as in Excel. */
export function defaultStyles(): Styles {
  return {
    numFmts: new Map(),
    fonts: [{ name: "Calibri", size: 11 }],
    fills: [null, null],
    borders: [{}],
    xfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
    dxfs: [],
    raw: { fonts: [null], fills: ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders: [null], xfs: [null] },
    source: null,
    theme: DEFAULT_THEME,
  };
}

const DEFAULT_THEME = ["#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#0563C1", "#954F72"];

/** Excel's old 64-colour palette, for colours given by number. */
const INDEXED = [
  "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF", "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
  "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080", "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
  "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF", "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
  "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696", "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
];

/** Applies Excel's tint (-1 darker … 1 lighter) to a colour, in HSL as Excel does. */
function tinted(hex: string, tint: number): string {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const out = s === 0 ? [l, l, l] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  return `#${out.map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/** A <color> element (rgb, theme or indexed, with an optional tint) as #rrggbb. */
export function colorOf(el: Element | null | undefined, theme: string[]): string | undefined {
  if (!el) return undefined;
  let hex: string | undefined;
  const rgb = el.getAttribute("rgb");
  const themeIndex = el.getAttribute("theme");
  const indexed = el.getAttribute("indexed");
  if (rgb) hex = `#${rgb.slice(-6).toUpperCase()}`;
  else if (themeIndex !== null) hex = theme[Number(themeIndex)];
  else if (indexed !== null) {
    const i = Number(indexed);
    // 64 and 65 are the system's text and background colours.
    hex = i === 64 ? undefined : i === 65 ? "#FFFFFF" : INDEXED[i];
  }
  const tint = Number(el.getAttribute("tint") ?? 0);
  return hex && tint ? tinted(hex, tint) : hex;
}

// ---- reading -------------------------------------------------------------------------------

const child = (el: Element, name: string) => Array.from(el.children).find((c) => c.localName === name) ?? null;
const children = (el: Element | null, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const serializer = () => new XMLSerializer();

/** The theme colours from theme1.xml, in the order cells refer to them (light 1 and dark 1 swapped, as Excel does). */
export function readTheme(xml: string | null): string[] {
  if (!xml) return DEFAULT_THEME;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const scheme = doc.getElementsByTagNameNS("*", "clrScheme")[0];
  if (!scheme) return DEFAULT_THEME;
  const value = (name: string) => {
    const el = Array.from(scheme.children).find((c) => c.localName === name);
    const color = el?.firstElementChild;
    const hex = color?.getAttribute("val") ?? color?.getAttribute("lastClr");
    return hex ? `#${hex.toUpperCase()}` : "#000000";
  };
  return ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].map(value);
}

export function readStyles(xml: string | null, theme: string[]): Styles {
  if (!xml) return { ...defaultStyles(), theme };
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  const out: Styles = { numFmts: new Map(), fonts: [], fills: [], borders: [], xfs: [], dxfs: children(child(root, "dxfs"), "dxf").map((d) => serializer().serializeToString(d).replace(/sxmlns(:w+)?="[^"]*"/g, "")), raw: { fonts: [], fills: [], borders: [], xfs: [] }, source: xml, theme };
  for (const nf of children(child(root, "numFmts"), "numFmt")) out.numFmts.set(Number(nf.getAttribute("numFmtId")), nf.getAttribute("formatCode") ?? "General");
  const on = (el: Element, name: string) => {
    const e = child(el, name);
    return !!e && e.getAttribute("val") !== "0" && e.getAttribute("val") !== "false" && e.getAttribute("val") !== "none";
  };
  for (const f of children(child(root, "fonts"), "font")) {
    out.fonts.push({
      name: child(f, "name")?.getAttribute("val") ?? "Calibri",
      size: Number(child(f, "sz")?.getAttribute("val") ?? 11),
      ...(on(f, "b") ? { bold: true } : {}),
      ...(on(f, "i") ? { italic: true } : {}),
      ...(on(f, "u") ? { underline: true } : {}),
      ...(on(f, "strike") ? { strike: true } : {}),
      ...(colorOf(child(f, "color"), theme) ? { color: colorOf(child(f, "color"), theme) } : {}),
    });
    out.raw.fonts.push(serializer().serializeToString(f));
  }
  for (const fill of children(child(root, "fills"), "fill")) {
    const pattern = child(fill, "patternFill");
    const solid = pattern?.getAttribute("patternType") === "solid";
    out.fills.push(solid ? (colorOf(child(pattern!, "fgColor"), theme) ?? null) : null);
    out.raw.fills.push(serializer().serializeToString(fill));
  }
  for (const b of children(child(root, "borders"), "border")) {
    const border: Border = {};
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const el = child(b, side);
      const style = el?.getAttribute("style");
      if (el && style) border[side] = { style, ...(colorOf(child(el, "color"), theme) ? { color: colorOf(child(el, "color"), theme) } : {}) };
    }
    out.borders.push(border);
    out.raw.borders.push(serializer().serializeToString(b));
  }
  for (const xf of children(child(root, "cellXfs"), "xf")) {
    const align = child(xf, "alignment");
    out.xfs.push({
      numFmtId: Number(xf.getAttribute("numFmtId") ?? 0),
      fontId: Number(xf.getAttribute("fontId") ?? 0),
      fillId: Number(xf.getAttribute("fillId") ?? 0),
      borderId: Number(xf.getAttribute("borderId") ?? 0),
      ...(align?.getAttribute("horizontal") ? { h: align.getAttribute("horizontal")! } : {}),
      ...(align?.getAttribute("vertical") ? { v: align.getAttribute("vertical")! } : {}),
      ...(align && (align.getAttribute("wrapText") === "1" || align.getAttribute("wrapText") === "true") ? { wrap: true } : {}),
      ...(align?.getAttribute("indent") ? { indent: Number(align.getAttribute("indent")) } : {}),
    });
    out.raw.xfs.push(serializer().serializeToString(xf));
  }
  if (!out.fonts.length) out.fonts.push({ name: "Calibri", size: 11 }), out.raw.fonts.push(null);
  if (!out.fills.length) out.fills.push(null, null), out.raw.fills.push('<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>');
  if (!out.borders.length) out.borders.push({}), out.raw.borders.push(null);
  if (!out.xfs.length) out.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }), out.raw.xfs.push(null);
  return out;
}

// ---- how a style looks -------------------------------------------------------------------------

export interface Look {
  font: Font;
  fill: string | null;
  border: Border;
  numFmt: string;
  h?: string;
  v?: string;
  wrap?: boolean;
  indent?: number;
}

const looks = new WeakMap<Styles, Map<number, Look>>();

export function lookOf(styles: Styles, s = 0): Look {
  let cache = looks.get(styles);
  if (!cache) looks.set(styles, (cache = new Map()));
  let look = cache.get(s);
  if (!look) {
    const xf = styles.xfs[s] ?? styles.xfs[0];
    look = {
      font: styles.fonts[xf.fontId] ?? styles.fonts[0],
      fill: styles.fills[xf.fillId] ?? null,
      border: styles.borders[xf.borderId] ?? {},
      numFmt: styles.numFmts.get(xf.numFmtId) ?? builtinFormat(xf.numFmtId),
      ...(xf.h ? { h: xf.h } : {}),
      ...(xf.v ? { v: xf.v } : {}),
      ...(xf.wrap ? { wrap: true } : {}),
      ...(xf.indent ? { indent: xf.indent } : {}),
    };
    cache.set(s, look);
  }
  return look;
}

// ---- changing a style ---------------------------------------------------------------------------

export interface StyleChange {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontName?: string;
  fontSize?: number;
  color?: string | null;
  fill?: string | null;
  numFmt?: string;
  h?: string | null;
  v?: string | null;
  wrap?: boolean;
  indent?: number;
  border?: Border;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function findOrAdd<T>(list: T[], raw: (string | null)[], item: T): number {
  const i = list.findIndex((x) => same(x, item));
  if (i >= 0) return i;
  list.push(item);
  raw.push(null);
  return list.length - 1;
}

/** The style index for style `s` with `change` applied; adds what's new to the styles. */
export function withStyle(styles: Styles, s: number | undefined, change: StyleChange): number {
  const xf = { ...(styles.xfs[s ?? 0] ?? styles.xfs[0]) };
  const fontKeys = ["bold", "italic", "underline", "strike", "fontName", "fontSize", "color"] as const;
  if (fontKeys.some((k) => k in change)) {
    const font: Font = { ...styles.fonts[xf.fontId] };
    for (const k of ["bold", "italic", "underline", "strike"] as const) {
      if (!(k in change)) continue;
      if (change[k]) font[k] = true;
      else delete font[k];
    }
    if (change.fontName) font.name = change.fontName;
    if (change.fontSize) font.size = change.fontSize;
    if ("color" in change) {
      if (change.color) font.color = change.color;
      else delete font.color;
    }
    xf.fontId = findOrAdd(styles.fonts, styles.raw.fonts, font);
  }
  if ("fill" in change) {
    xf.fillId = change.fill ? findOrAdd(styles.fills, styles.raw.fills, change.fill) : 0;
  }
  if (change.border) xf.borderId = findOrAdd(styles.borders, styles.raw.borders, change.border);
  if (change.numFmt !== undefined) {
    const builtin = Array.from({ length: 50 }, (_, i) => i).find((i) => builtinFormat(i) === change.numFmt && (i !== 0 || change.numFmt === "General"));
    if (builtin !== undefined && builtin !== 14 && builtin !== 22) xf.numFmtId = builtin;
    else {
      let id = [...styles.numFmts].find(([, code]) => code === change.numFmt)?.[0];
      if (id === undefined) {
        id = Math.max(163, ...styles.numFmts.keys()) + 1;
        styles.numFmts.set(id, change.numFmt);
      }
      xf.numFmtId = id;
    }
  }
  for (const k of ["h", "v"] as const) {
    if (!(k in change)) continue;
    if (change[k]) xf[k] = change[k]!;
    else delete xf[k];
  }
  if ("wrap" in change) {
    if (change.wrap) xf.wrap = true;
    else delete xf.wrap;
  }
  if ("indent" in change) {
    if (change.indent) xf.indent = change.indent;
    else delete xf.indent;
  }
  return findOrAdd(styles.xfs, styles.raw.xfs, xf);
}

// ---- writing --------------------------------------------------------------------------------------

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const rgb = (hex: string) => `FF${hex.slice(1).toUpperCase()}`;

function fontXml(f: Font): string {
  return `<font>${f.bold ? "<b/>" : ""}${f.italic ? "<i/>" : ""}${f.strike ? "<strike/>" : ""}${f.underline ? "<u/>" : ""}<sz val="${f.size}"/>${f.color ? `<color rgb="${rgb(f.color)}"/>` : '<color theme="1"/>'}<name val="${esc(f.name)}"/><family val="2"/></font>`;
}

function borderXml(b: Border): string {
  const side = (name: string, s?: BorderSide) => (s ? `<${name} style="${s.style}">${s.color ? `<color rgb="${rgb(s.color)}"/>` : '<color indexed="64"/>'}</${name}>` : `<${name}/>`);
  return `<border>${side("left", b.left)}${side("right", b.right)}${side("top", b.top)}${side("bottom", b.bottom)}<diagonal/></border>`;
}

function xfXml(x: Xf): string {
  const align = x.h || x.v || x.wrap || x.indent ? `<alignment${x.h ? ` horizontal="${x.h}"` : ""}${x.v ? ` vertical="${x.v}"` : ""}${x.wrap ? ' wrapText="1"' : ""}${x.indent ? ` indent="${x.indent}"` : ""}/>` : "";
  const applies = `${x.numFmtId ? ' applyNumberFormat="1"' : ""}${x.fontId ? ' applyFont="1"' : ""}${x.fillId ? ' applyFill="1"' : ""}${x.borderId ? ' applyBorder="1"' : ""}${align ? ' applyAlignment="1"' : ""}`;
  return `<xf numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0"${applies}>${align}</xf>`.replace("></xf>", "/>");
}

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** styles.xml for the styles: the lists rebuilt, everything else as it was read. */
export function writeStyles(styles: Styles): string {
  const list = (name: string, items: string[], extra = "") => `<${name} count="${items.length}"${extra}>${items.join("")}</${name}>`;
  const numFmts = [...styles.numFmts].map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
  const fonts = styles.fonts.map((f, i) => styles.raw.fonts[i] ?? fontXml(f));
  const fills = styles.fills.map((f, i) => styles.raw.fills[i] ?? `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(f ?? "#FFFFFF")}"/><bgColor indexed="64"/></patternFill></fill>`);
  const borders = styles.borders.map((b, i) => styles.raw.borders[i] ?? borderXml(b));
  const xfs = styles.xfs.map((x, i) => styles.raw.xfs[i] ?? xfXml(x));
  const parts: Record<string, string> = {
    numFmts: numFmts.length ? list("numFmts", numFmts) : "",
    fonts: list("fonts", fonts),
    fills: list("fills", fills),
    borders: list("borders", borders),
    cellXfs: list("cellXfs", xfs),
    dxfs: list("dxfs", styles.dxfs),
  };
  if (!styles.source) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${MAIN_NS}">${parts.numFmts}${parts.fonts}${parts.fills}${parts.borders}<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>${parts.cellXfs}<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${parts.dxfs}<tableStyles count="0"/></styleSheet>`;
  }
  // Replace each list in the original, in place; add numFmts first, and dxfs after the cell styles, if they weren't there.
  let xml = styles.source;
  for (const name of ["numFmts", "fonts", "fills", "borders", "cellXfs", "dxfs"]) {
    const re = new RegExp(`<(?:\\w+:)?${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</(?:\\w+:)?${name}>)`);
    if (re.test(xml)) xml = xml.replace(re, () => parts[name]);
    else if (name === "numFmts" && parts.numFmts) xml = xml.replace(/(<(?:\w+:)?styleSheet\b[^>]*>)/, (m) => m + parts.numFmts);
    else if (name === "dxfs" && styles.dxfs.length) xml = xml.replace(/<(?:\w+:)?(?:tableStyles|colors|extLst)\b|<\/(?:\w+:)?styleSheet>/, (m) => parts.dxfs + m);
  }
  return xml;
}
