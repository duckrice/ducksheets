// Charts in Excel files (DrawingML). Reading: the column, bar, line and pie
// charts of a sheet's drawing become charts DuckSheets can show and change;
// anything else in the drawing (pictures, shapes, other kinds of chart) is
// kept as it was. Writing: each chart as Excel writes one, its series linked
// to the cells, with the values cached for programs that don't calculate.

import type JSZip from "jszip";
import { quoteSheet } from "../formula/parse";
import { parseRange, type Chart, type ChartKind, type Range, type Sheet } from "../model";
import { readRels } from "./read";

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU = 9525; // per pixel

const xml = (text: string) => new DOMParser().parseFromString(text, "application/xml");
const kids = (el: Element | null | undefined, name: string) => (el ? Array.from(el.children).filter((c) => c.localName === name) : []);
const kid = (el: Element | null | undefined, name: string) => kids(el, name)[0] ?? null;
const deep = (el: Element | null | undefined, name: string) => (el ? (Array.from(el.getElementsByTagNameNS("*", name))[0] ?? null) : null);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A chart reference like 'My sheet'!$B$2:$B$5, if it's on `sheetName`. */
function refRange(f: string | null | undefined, sheetName: string): Range | null {
  if (!f) return null;
  const at = f.lastIndexOf("!");
  if (at < 0) return null;
  const sheet = f.slice(0, at).replace(/^'|'$/g, "").replace(/''/g, "'");
  if (sheet.toLowerCase() !== sheetName.toLowerCase()) return null;
  return parseRange(f.slice(at + 1).replace(/\$/g, ""));
}

/** A chart part as a DuckSheets chart, if it's one kind of chart with its series in columns of one block. */
export function parseChart(text: string, sheetName: string): Omit<Chart, "id" | "at"> | null {
  const doc = xml(text);
  const plot = deep(doc.documentElement, "plotArea");
  const types = Array.from(plot?.children ?? []).filter((el) => /Chart$/.test(el.localName));
  if (types.length !== 1) return null;
  const type = types[0];
  let kind: ChartKind;
  if (type.localName === "barChart") kind = kid(type, "barDir")?.getAttribute("val") === "bar" ? "bar" : "column";
  else if (type.localName === "lineChart") kind = "line";
  else if (type.localName === "pieChart") kind = "pie";
  else if (type.localName === "doughnutChart") kind = "doughnut";
  else if (type.localName === "scatterChart") kind = "scatter";
  else if (type.localName === "areaChart" && ["standard", null].includes(kid(type, "grouping")?.getAttribute("val") ?? null)) kind = "area";
  else return null;
  if (type.localName === "barChart" && !["clustered", null].includes(kid(type, "grouping")?.getAttribute("val") ?? null)) return null;
  const series = kids(type, "ser");
  if (!series.length) return null;
  const f = (el: Element | null, part: string) => deep(kid(el, part), "f")?.textContent ?? null;
  const vals = series.map((s) => refRange(f(s, "val") ?? f(s, "yVal"), sheetName));
  if (vals.some((v) => !v || v.c1 !== v.c2)) return null;
  const v = vals as Range[];
  const r1 = v[0].r1;
  const r2 = v[0].r2;
  // Series in neighbouring columns, over the same rows.
  if (v.some((x, i) => x.r1 !== r1 || x.r2 !== r2 || x.c1 !== v[0].c1 + i)) return null;
  const cat = refRange(f(series[0], "cat") ?? f(series[0], "xVal"), sheetName);
  if (cat && (cat.c1 !== cat.c2 || cat.c1 !== v[0].c1 - 1 || cat.r1 !== r1 || cat.r2 !== r2)) return null;
  const names = series.map((s) => refRange(f(s, "tx"), sheetName));
  const headed = names.every((n, i) => n && n.r1 === r1 - 1 && n.c1 === v[i].c1);
  const title = Array.from(deep(doc.documentElement, "title")?.getElementsByTagNameNS(A, "t") ?? [])
    .map((t) => t.textContent ?? "")
    .join("");
  return { kind, title, data: { r1: headed ? r1 - 1 : r1, c1: cat ? cat.c1 : v[0].c1, r2, c2: v[v.length - 1].c1 } };
}

/** Pixel positions of rows or columns, for turning anchors into a place and size. */
function edges(sizes: Map<number, number>, hidden: Set<number>, def: number) {
  const size = (i: number) => (hidden.has(i) ? 0 : (sizes.get(i) ?? def));
  return {
    size,
    /** From cell `from` (plus `off` pixels) to cell `to` (plus `toOff`), in pixels. */
    span(from: number, off: number, to: number, toOff: number) {
      let px = -off;
      for (let i = from; i < to; i++) px += size(i);
      return px + toOff;
    },
    /** The cell `px` pixels on from cell `from` (plus `off`), and how far into it. */
    reach(from: number, off: number, px: number): [number, number] {
      let i = from;
      let left = px + off;
      while (left >= size(i) && i < 1048575) {
        left -= size(i);
        i++;
      }
      return [i, Math.round(left)];
    },
  };
}

/** Reads the charts of a sheet's drawing; what it can't show stays in the drawing's XML, kept for writing. */
export async function readCharts(zip: JSZip, sheet: Sheet): Promise<void> {
  if (!sheet.origin) return;
  const ws = xml(sheet.origin.xml).documentElement;
  const drawingEl = kid(ws, "drawing");
  if (!drawingEl) return;
  const rels = await readRels(zip, sheet.origin.path);
  const rel = rels.get(drawingEl.getAttributeNS(R, "id") ?? "");
  const text = rel && (await zip.file(rel.path)?.async("string"));
  if (!rel || !text) return;
  const doc = xml(text);
  const drawingRels = await readRels(zip, rel.path);
  const cols = edges(sheet.colWidths, sheet.hiddenCols, sheet.defaultColWidth);
  const rows = edges(sheet.rowHeights, sheet.hiddenRows, sheet.defaultRowHeight);
  const owned: string[] = [];
  const point = (el: Element | null) => ({
    c: Number(kid(el, "col")?.textContent ?? 0),
    cOff: Math.round(Number(kid(el, "colOff")?.textContent ?? 0) / EMU),
    r: Number(kid(el, "row")?.textContent ?? 0),
    rOff: Math.round(Number(kid(el, "rowOff")?.textContent ?? 0) / EMU),
  });
  const place = (anchor: Element) => {
    const from = point(kid(anchor, "from"));
    const to = point(kid(anchor, "to"));
    return { r: from.r, c: from.c, dx: from.cOff, dy: from.rOff, w: Math.max(20, cols.span(from.c, from.cOff, to.c, to.cOff)), h: Math.max(20, rows.span(from.r, from.rOff, to.r, to.rOff)) };
  };
  for (const anchor of kids(doc.documentElement, "twoCellAnchor")) {
    // Pictures: shown, moved and saved by DuckSheets as charts are.
    const blip = kid(anchor, "pic") && deep(anchor, "blip");
    const media = blip && drawingRels.get(blip.getAttributeNS(R, "embed") ?? "");
    if (media) {
      const bytes = await zip.file(media.path)?.async("base64");
      const ext = media.path.split(".").pop()?.toLowerCase() ?? "png";
      if (bytes && ["png", "jpg", "jpeg", "gif", "bmp"].includes(ext)) {
        const type = ext === "jpg" ? "jpeg" : ext;
        sheet.charts.push({ id: Math.random().toString(36).slice(2, 9), kind: "picture", src: `data:image/${type};base64,${bytes}`, data: { r1: 0, c1: 0, r2: 0, c2: 0 }, title: "", at: place(anchor) });
        owned.push(media.path);
        doc.documentElement.removeChild(anchor);
        continue;
      }
    }
    const chartEl = deep(anchor, "chart");
    const chartRel = chartEl && drawingRels.get(chartEl.getAttributeNS(R, "id") ?? "");
    const chartText = chartRel && (await zip.file(chartRel.path)?.async("string"));
    if (!chartRel || !chartText) continue;
    const parsed = parseChart(chartText, sheet.name);
    if (!parsed) continue;
    const at = place(anchor);
    sheet.charts.push({ ...parsed, id: Math.random().toString(36).slice(2, 9), at: { ...at, w: Math.max(60, at.w), h: Math.max(40, at.h) } });
    owned.push(chartRel.path);
    doc.documentElement.removeChild(anchor);
  }
  sheet.origin.drawing = { path: rel.path, xml: new XMLSerializer().serializeToString(doc), owned };
}

// ---- writing ----------------------------------------------------------------------------------------

export interface ChartContent {
  categories: string[];
  series: { name: string; values: (number | null)[] }[];
  /** Whether the first row names the series, and the first column the categories. */
  headed: boolean;
  labelled: boolean;
  format: string;
  /** The first column as numbers: a scatter chart's x values. */
  xs: (number | null)[];
}

const COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9E480E", "636363", "997300"];

const abs = (sheet: string, r: Range) => `${quoteSheet(sheet)}!$${colLetters(r.c1)}$${r.r1 + 1}${r.r1 === r.r2 && r.c1 === r.c2 ? "" : `:$${colLetters(r.c2)}$${r.r2 + 1}`}`;
function colLetters(c: number) {
  let name = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** A chart part, as Excel writes one. */
export function chartXml(chart: Chart, sheetName: string, content: ChartContent): string {
  const { data } = chart;
  const top = content.headed ? data.r1 + 1 : data.r1;
  const left = content.labelled ? data.c1 + 1 : data.c1;
  const points = (values: (string | number | null)[]) => values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`)).join("");
  const series = content.series.map((s, i) => {
    const c = left + i;
    const tx = content.headed
      ? `<c:tx><c:strRef><c:f>${esc(abs(sheetName, { r1: data.r1, c1: c, r2: data.r1, c2: c }))}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
      : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
    const color = COLORS[i % COLORS.length];
    const fill = chart.kind === "line" ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:marker>` : chart.kind === "pie" || chart.kind === "doughnut" ? "" : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${chart.kind === "area" ? "" : '<c:invertIfNegative val="0"/>'}`;
    const slices = chart.kind === "pie" || chart.kind === "doughnut" ? content.categories.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${COLORS[k % COLORS.length]}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("") : "";
    const cat = content.labelled
      ? `<c:cat><c:strRef><c:f>${esc(abs(sheetName, { r1: top, c1: data.c1, r2: data.r2, c2: data.c1 }))}</c:f><c:strCache><c:ptCount val="${content.categories.length}"/>${points(content.categories)}</c:strCache></c:strRef></c:cat>`
      : "";
    const val = `<c:val><c:numRef><c:f>${esc(abs(sheetName, { r1: top, c1: c, r2: data.r2, c2: c }))}</c:f><c:numCache><c:formatCode>${esc(content.format)}</c:formatCode><c:ptCount val="${s.values.length}"/>${points(s.values)}</c:numCache></c:numRef></c:val>`;
    if (chart.kind === "scatter") {
      const xRange = esc(abs(sheetName, { r1: top, c1: data.c1, r2: data.r2, c2: data.c1 }));
      const xVal = `<c:xVal><c:numRef><c:f>${xRange}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${content.xs.length}"/>${points(content.xs)}</c:numCache></c:numRef></c:xVal>`;
      const marker = `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:marker>`;
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${marker}${xVal}${val.replace(/^<c:val>/, "<c:yVal>").replace(/<\/c:val>$/, "</c:yVal>")}<c:smooth val="0"/></c:ser>`;
    }
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${fill}${slices}${cat}${val}${chart.kind === "line" ? '<c:smooth val="0"/>' : ""}</c:ser>`;
  });
  const axes = (catPos: string, valPos: string) =>
    `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="${esc(content.format)}" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const ids = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  let plot: string;
  if (chart.kind === "pie") plot = `<c:pieChart><c:varyColors val="1"/>${series[0] ?? ""}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (chart.kind === "doughnut") plot = `<c:doughnutChart><c:varyColors val="1"/>${series[0] ?? ""}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
  else if (chart.kind === "scatter") {
    const valAx = (id: number, cross: number, pos: string, grid: boolean) =>
      `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : ""}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
    plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series.join("")}${ids}</c:scatterChart>${valAx(500000001, 500000002, "b", false)}${valAx(500000002, 500000001, "l", true)}`;
  } else if (chart.kind === "area") plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.join("")}${ids}</c:areaChart>${axes("b", "l")}`;
  else if (chart.kind === "line") plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.join("")}<c:marker val="1"/>${ids}</c:lineChart>${axes("b", "l")}`;
  else plot = `<c:barChart><c:barDir val="${chart.kind === "bar" ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.join("")}<c:gapWidth val="150"/>${ids}</c:barChart>${axes(chart.kind === "bar" ? "l" : "b", chart.kind === "bar" ? "b" : "l")}`;
  const title = chart.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = chart.kind === "pie" || chart.kind === "doughnut" || content.series.length > 1 ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
}

/** The anchor for a picture: as a chart's, with the picture (linked by `relId`) in place of the chart. */
export function pictureAnchorXml(chart: Chart, sheet: Sheet, relId: string, shapeId: number): string {
  const pic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${shapeId}" name="Picture ${shapeId}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${R}" r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${chart.at.w * EMU}" cy="${chart.at.h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
  return anchorXml(chart, sheet, relId, shapeId).replace(/<xdr:graphicFrame[\s\S]*<\/xdr:graphicFrame>/, pic);
}

/** The anchor that places a chart in the drawing, from its cell to the cell its far corner reaches. */
export function anchorXml(chart: Chart, sheet: Sheet, relId: string, shapeId: number): string {
  const cols = edges(sheet.colWidths, sheet.hiddenCols, sheet.defaultColWidth);
  const rows = edges(sheet.rowHeights, sheet.hiddenRows, sheet.defaultRowHeight);
  const [toC, toCOff] = cols.reach(chart.at.c, chart.at.dx, chart.at.w);
  const [toR, toROff] = rows.reach(chart.at.r, chart.at.dy, chart.at.h);
  const point = (tag: string, c: number, cOff: number, r: number, rOff: number) => `<xdr:${tag}><xdr:col>${c}</xdr:col><xdr:colOff>${cOff * EMU}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${rOff * EMU}</xdr:rowOff></xdr:${tag}>`;
  return `<xdr:twoCellAnchor editAs="oneCell">${point("from", chart.at.c, chart.at.dx, chart.at.r, chart.at.dy)}${point("to", toC, toCOff, toR, toROff)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="Chart ${shapeId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

export const EMPTY_DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}"></xdr:wsDr>`;
