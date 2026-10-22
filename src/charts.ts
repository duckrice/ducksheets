// Charts: column, bar, line and pie charts of a block of data, floating over
// the sheet. The first row of the block names the series and the first
// column the categories, as Excel reads them. Click a chart to move it (drag),
// resize it (the corner), or change it (right-click); Delete removes it.

import { selRange, type Doc } from "./doc";
import { svg, h } from "./dom";
import type { Grid } from "./grid";
import { MAX_COLS, MAX_ROWS, type Chart, type ChartKind } from "./model";
import { chartData } from "./chartdata";
import { formatValue } from "./numfmt";
import { askName, menu, toast } from "./ui";

export const CHART_COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E", "#636363", "#997300"];

/** Round steps for an axis: 0, 2 000, 4 000 … rather than 0, 1 873 … */
export function niceScale(min: number, max: number, ticks = 5): { lo: number; hi: number; step: number } {
  if (min === max) {
    min = Math.min(0, min);
    max = max === 0 ? 1 : Math.max(0, max);
  }
  const lo0 = Math.min(0, min);
  const hi0 = Math.max(0, max);
  const raw = (hi0 - lo0) / ticks;
  const magnitude = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * magnitude).find((s) => s >= raw) ?? 10 * magnitude;
  return { lo: Math.floor(lo0 / step) * step, hi: Math.ceil(hi0 / step) * step, step };
}

const text = (x: number, y: number, content: string, attrs: Record<string, string | number> = {}) => {
  const el = svg("text", { x, y, "font-size": 11, fill: "#404040", ...attrs });
  el.textContent = content;
  return el;
};

/** The chart drawn as an SVG, w × h. */
export function drawChart(doc: Doc, chart: Chart, w: number, h: number): SVGSVGElement {
  const root = svg("svg", { width: w, height: h, viewBox: `0 0 ${w} ${h}`, class: "chart-svg" });
  root.append(svg("rect", { x: 0, y: 0, width: w, height: h, fill: "#ffffff" }));
  const { categories, series, format } = chartData(doc, chart, doc.active);
  const shown = (n: number) => formatValue(n, format === "General" || format === "@" ? "General" : format, doc.book.date1904).text;
  let top = 12;
  if (chart.title) {
    root.append(text(w / 2, 26, chart.title, { "font-size": 15, "text-anchor": "middle", fill: "#262626" }));
    top = 42;
  }
  if (!series.length || !categories.length) {
    root.append(text(w / 2, h / 2, "No numbers to show", { "text-anchor": "middle" }));
    return root;
  }
  const round = chart.kind === "pie" || chart.kind === "doughnut";
  const legendItems = round ? categories : series.length > 1 ? series.map((s) => s.name) : [];
  const legendH = legendItems.length ? 22 : 0;
  // The legend along the bottom, as in Excel.
  if (legendItems.length) {
    let x = 0;
    const items = legendItems.map((name, i) => {
      const g = svg("g");
      g.append(svg("rect", { x, y: 0, width: 9, height: 9, fill: CHART_COLORS[i % CHART_COLORS.length] }), text(x + 13, 9, name));
      x += 26 + name.length * 6.2;
      return g;
    });
    const legend = svg("g", { transform: `translate(${Math.max(8, (w - x) / 2)},${h - 16})` });
    legend.append(...items);
    root.append(legend);
  }
  const bottom = h - 10 - legendH;

  if (round) {
    const values = series[0].values.map((v) => Math.max(0, v ?? 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const r = Math.max(10, Math.min(w - 20, bottom - top) / 2 - 4);
    const cx = w / 2;
    const cy = (top + bottom) / 2;
    let angle = -Math.PI / 2;
    values.forEach((v, i) => {
      const a = (v / total) * Math.PI * 2;
      const end = angle + a;
      const large = a > Math.PI ? 1 : 0;
      const p = (t: number) => `${cx + r * Math.cos(t)},${cy + r * Math.sin(t)}`;
      const d = values.filter((x) => x > 0).length === 1 && v > 0 ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` : `M${cx},${cy}L${p(angle)}A${r},${r} 0 ${large} 1 ${p(end)}Z`;
      root.append(svg("path", { d, fill: CHART_COLORS[i % CHART_COLORS.length], stroke: "#ffffff", "stroke-width": 1.5 }));
      if (v / total > 0.04) {
        const mid = angle + a / 2;
        const at = chart.kind === "doughnut" ? 0.75 : 0.65;
        root.append(text(cx + r * at * Math.cos(mid), cy + r * at * Math.sin(mid) + 4, `${Math.round((v / total) * 100)}%`, { "text-anchor": "middle", fill: "#ffffff", "font-weight": "bold" }));
      }
      angle = end;
    });
    // A doughnut is a pie with its middle cut out.
    if (chart.kind === "doughnut") root.append(svg("circle", { cx, cy, r: r * 0.5, fill: "#ffffff" }));
    return root;
  }

  const all = series.flatMap((s) => s.values).filter((v): v is number => v !== null);
  const scale = niceScale(Math.min(...all), Math.max(...all));
  const horizontal = chart.kind === "bar";
  const labelW = Math.max(...[scale.lo, scale.hi].map((n) => shown(n).length)) * 6.5 + 10;
  const catW = horizontal ? Math.min(w * 0.3, Math.max(...categories.map((c) => c.length)) * 6.2 + 12) : 0;
  const plot = { x: horizontal ? catW + 8 : labelW + 6, y: top, w: 0, h: 0 };
  plot.w = w - plot.x - 14;
  plot.h = bottom - top - (horizontal ? 18 : 20);
  if (plot.w < 20 || plot.h < 20) return root;
  const pos = (v: number) => (v - scale.lo) / (scale.hi - scale.lo);

  // Value gridlines and labels.
  for (let v = scale.lo; v <= scale.hi + scale.step / 2; v += scale.step) {
    const t = pos(v);
    if (horizontal) {
      const x = plot.x + t * plot.w;
      root.append(svg("line", { x1: x, y1: plot.y, x2: x, y2: plot.y + plot.h, stroke: "#d9d9d9" }), text(x, plot.y + plot.h + 14, shown(v), { "text-anchor": "middle" }));
    } else {
      const y = plot.y + plot.h - t * plot.h;
      root.append(svg("line", { x1: plot.x, y1: y, x2: plot.x + plot.w, y2: y, stroke: "#d9d9d9" }), text(plot.x - 6, y + 4, shown(v), { "text-anchor": "end" }));
    }
  }
  const zero = pos(Math.max(scale.lo, Math.min(0, scale.hi)));
  const n = categories.length;
  const band = (horizontal ? plot.h : plot.w) / n;
  // Category labels, thinned out when they'd overlap.
  const every = Math.max(1, Math.ceil((horizontal ? 14 : Math.max(...categories.map((c) => c.length)) * 6.2 + 6) / band));
  categories.forEach((name, i) => {
    if (i % every) return;
    const mid = i * band + band / 2;
    if (horizontal) root.append(text(plot.x - 6, plot.y + mid + 4, name, { "text-anchor": "end" }));
    else root.append(text(plot.x + mid, plot.y + plot.h + 15, name, { "text-anchor": "middle" }));
  });

  if (chart.kind === "area") {
    // Filled down to zero, the first series at the back.
    const base = plot.y + plot.h - zero * plot.h;
    series.forEach((s, k) => {
      const points = s.values.map((v, i) => [plot.x + i * band + band / 2, plot.y + plot.h - pos(v ?? 0) * plot.h]);
      const d = `M${points[0][0]},${base}${points.map((p) => `L${p[0]},${p[1]}`).join("")}L${points[points.length - 1][0]},${base}Z`;
      root.append(svg("path", { d, fill: CHART_COLORS[k % CHART_COLORS.length], "fill-opacity": 0.75, stroke: CHART_COLORS[k % CHART_COLORS.length] }));
    });
  } else if (chart.kind === "line") {
    series.forEach((s, k) => {
      const color = CHART_COLORS[k % CHART_COLORS.length];
      const points = s.values.map((v, i) => (v === null ? null : [plot.x + i * band + band / 2, plot.y + plot.h - pos(v) * plot.h]));
      let d = "";
      points.forEach((p, i) => {
        if (p) d += `${d && points[i - 1] ? "L" : "M"}${p[0]},${p[1]}`;
      });
      root.append(svg("path", { d, fill: "none", stroke: color, "stroke-width": 2.25, "stroke-linejoin": "round" }));
      for (const p of points) if (p) root.append(svg("circle", { cx: p[0], cy: p[1], r: 3, fill: color }));
    });
  } else {
    const gap = band * 0.25;
    const thick = (band - gap * 2) / series.length;
    series.forEach((s, k) =>
      s.values.forEach((v, i) => {
        if (v === null) return;
        const a = Math.min(pos(v), zero);
        const b = Math.max(pos(v), zero);
        const along = i * band + gap + k * thick;
        root.append(
          horizontal
            ? svg("rect", { x: plot.x + a * plot.w, y: plot.y + along, width: Math.max(0.5, (b - a) * plot.w), height: Math.max(1, thick - 1), fill: CHART_COLORS[k % CHART_COLORS.length] })
            : svg("rect", { x: plot.x + along, y: plot.y + plot.h - b * plot.h, width: Math.max(1, thick - 1), height: Math.max(0.5, (b - a) * plot.h), fill: CHART_COLORS[k % CHART_COLORS.length] }),
        );
      }),
    );
  }
  // The axis line at zero.
  if (horizontal) root.append(svg("line", { x1: plot.x + zero * plot.w, y1: plot.y, x2: plot.x + zero * plot.w, y2: plot.y + plot.h, stroke: "#bfbfbf" }));
  else root.append(svg("line", { x1: plot.x, y1: plot.y + plot.h - zero * plot.h, x2: plot.x + plot.w, y2: plot.y + plot.h - zero * plot.h, stroke: "#bfbfbf" }));
  return root;
}

// ---- on the sheet ---------------------------------------------------------------------------------

let chosen: string | null = null;

function layer(grid: Grid): HTMLElement {
  let el = grid.el.querySelector<HTMLElement>(".charts");
  if (!el) {
    el = h("div", { class: "charts" });
    grid.el.append(el);
  }
  return el;
}

/** Draws the sheet's charts where they sit, at the current scroll and zoom. */
export function renderCharts(doc: Doc, grid: Grid) {
  const el = layer(grid);
  const z = grid.zoom;
  // Charts go under the column and row headers, not over them.
  el.style.clipPath = `inset(${grid.origin.y}px 0 0 ${grid.origin.x}px)`;
  el.replaceChildren(
    ...doc.sheet.charts.map((chart) => {
      const cell = grid.cellBox(chart.at.r, chart.at.c);
      const box = h("div", { class: `chart${chosen === chart.id ? " chosen" : ""}`, tabindex: "0", "aria-label": `${chart.kind} chart${chart.title ? `: ${chart.title}` : ""}` });
      box.style.left = `${cell.x + chart.at.dx * z}px`;
      box.style.top = `${cell.y + chart.at.dy * z}px`;
      box.style.width = `${chart.at.w * z}px`;
      box.style.height = `${chart.at.h * z}px`;
      const drawing = drawChart(doc, chart, chart.at.w, chart.at.h);
      drawing.setAttribute("width", String(chart.at.w * z));
      drawing.setAttribute("height", String(chart.at.h * z));
      const handle = h("span", { class: "chart-handle", title: "Drag to resize" });
      box.append(drawing, handle);
      box.addEventListener("pointerdown", (e) => grab(e, doc, grid, chart, e.target === handle));
      box.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        chosen = chart.id;
        chartMenu(doc, chart, box);
      });
      box.addEventListener("dblclick", () => chartMenu(doc, chart, box));
      box.addEventListener("keydown", (e) => {
        if (e.key === "Delete" || e.key === "Backspace") {
          e.preventDefault();
          e.stopPropagation();
          removeChart(doc, chart.id);
        }
      });
      return box;
    }),
  );
}

function update(doc: Doc, id: string, change: (c: Chart) => Chart) {
  doc.change(() => (doc.sheet.charts = doc.sheet.charts.map((c) => (c.id === id ? change(c) : c))));
}

function removeChart(doc: Doc, id: string) {
  chosen = null;
  doc.change(() => (doc.sheet.charts = doc.sheet.charts.filter((c) => c.id !== id)));
}

/** Moving (or, by the corner, resizing) a chart; it lands in whichever cell its corner is over. */
function grab(e: PointerEvent, doc: Doc, grid: Grid, chart: Chart, resize: boolean) {
  if (e.button !== 0) return;
  e.preventDefault();
  e.stopPropagation();
  chosen = chart.id;
  const box = e.currentTarget as HTMLElement;
  box.focus({ preventScroll: true });
  box.classList.add("chosen");
  box.setPointerCapture(e.pointerId);
  const z = grid.zoom;
  const start = { x: e.clientX, y: e.clientY, left: box.offsetLeft, top: box.offsetTop, w: box.offsetWidth, h: box.offsetHeight };
  let moved = false;
  const onMove = (ev: PointerEvent) => {
    const dx = ev.clientX - start.x;
    const dy = ev.clientY - start.y;
    if (!moved && Math.hypot(dx, dy) < 3) return;
    moved = true;
    if (resize) {
      box.style.width = `${Math.max(120, start.w + dx)}px`;
      box.style.height = `${Math.max(90, start.h + dy)}px`;
    } else {
      box.style.left = `${start.left + dx}px`;
      box.style.top = `${start.top + dy}px`;
    }
  };
  const onUp = () => {
    box.removeEventListener("pointermove", onMove);
    box.removeEventListener("pointerup", onUp);
    if (!moved) return;
    if (resize) {
      update(doc, chart.id, (c) => ({ ...c, at: { ...c.at, w: Math.round(box.offsetWidth / z), h: Math.round(box.offsetHeight / z) } }));
      return;
    }
    // Which cell the new corner is over, and how far into it.
    const anchor = grid.cellUnder(box.offsetLeft, box.offsetTop);
    const cell = grid.cellBox(anchor.r, anchor.c);
    update(doc, chart.id, (c) => ({ ...c, at: { ...c.at, r: anchor.r, c: anchor.c, dx: Math.max(0, Math.round((box.offsetLeft - cell.x) / z)), dy: Math.max(0, Math.round((box.offsetTop - cell.y) / z)) } }));
  };
  box.addEventListener("pointermove", onMove);
  box.addEventListener("pointerup", onUp);
}

function chartMenu(doc: Doc, chart: Chart, anchor: HTMLElement) {
  const kind = (k: ChartKind, label: string) => ({ label: `${chart.kind === k ? "✓ " : ""}${label}`, action: () => update(doc, chart.id, (c) => ({ ...c, kind: k })) });
  menu(anchor, [
    { label: "Chart type", items: [kind("column", "Column"), kind("bar", "Bar"), kind("line", "Line"), kind("area", "Area"), kind("pie", "Pie"), kind("doughnut", "Doughnut")] },
    {
      label: "Title…",
      action: () =>
        void askName("Chart title", chart.title, "OK").then((title) => {
          if (title !== null) update(doc, chart.id, (c) => ({ ...c, title }));
        }),
    },
    { label: "Use the selected cells as its data", action: () => update(doc, chart.id, (c) => ({ ...c, data: selRange(doc.sel) })) },
    { label: "Delete chart", danger: true, action: () => removeChart(doc, chart.id) },
  ]);
}

/** A chart of the selected cells (or the block of data around the cursor), next to the data. */
export function insertChart(doc: Doc, grid: Grid) {
  const sel = selRange(doc.sel);
  const data = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? doc.region(sel.r1, sel.c1) : sel;
  if (data.r2 >= MAX_ROWS - 1 || data.c2 >= MAX_COLS - 1 || (data.r1 === data.r2 && data.c1 === data.c2)) {
    toast("Select the data to chart first: a block of numbers, with names in its first row and column.");
    return;
  }
  const chart: Chart = {
    id: Math.random().toString(36).slice(2, 9),
    kind: "column",
    data,
    title: "",
    at: { r: data.r1, c: data.c2 + 2, dx: 0, dy: 0, w: 480, h: 290 },
  };
  if (!chartData(doc, chart, doc.active).series.some((s) => s.values.some((v) => v !== null))) {
    toast("There are no numbers in the selected cells to chart.");
    return;
  }
  chosen = chart.id;
  doc.change(() => doc.sheet.charts.push(chart));
  grid.reveal(chart.at.r, chart.at.c + 4);
}
