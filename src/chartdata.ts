// What a chart shows, read from its block of data: the first row names the
// series and the first column the categories, when they hold text, as Excel
// reads them. Used to draw charts and to write them into files.

import type { Cell, Chart, Scalar } from "./model";
import { lookOf, type Styles } from "./xlsx/styles";

export interface Series {
  name: string;
  values: (number | null)[];
}

/** Where a chart's numbers come from: the open document, or the workbook being saved. */
export interface ChartSource {
  value(r: number, c: number, sheet: number): Scalar;
  shown(r: number, c: number, sheet: number): string;
  cell(r: number, c: number, sheet: number): Cell | undefined;
  book: { styles: Styles };
}

/** What a chart shows: category names along the axis, and series of numbers. */
export function chartData(doc: ChartSource, chart: Chart, sheet: number): { categories: string[]; series: Series[]; format: string; headed: boolean; labelled: boolean } {
  const { data } = chart;
  const value = (r: number, c: number) => doc.value(r, c, sheet);
  // A first row of text (or a blank corner over text) names the series.
  const headed = data.r2 > data.r1 && [...Array(data.c2 - data.c1 + 1).keys()].some((j) => typeof value(data.r1, data.c1 + j) === "string") && typeof value(data.r1 + 1, data.c2) === "number";
  const firstColLabels = data.c2 > data.c1 && [...Array(data.r2 - data.r1 + 1).keys()].slice(headed ? 1 : 0).some((i) => typeof value(data.r1 + i, data.c1) === "string");
  const top = headed ? data.r1 + 1 : data.r1;
  const left = firstColLabels ? data.c1 + 1 : data.c1;
  const categories: string[] = [];
  for (let r = top; r <= data.r2; r++) categories.push(firstColLabels ? doc.shown(r, data.c1, sheet) : String(r - top + 1));
  const series: Series[] = [];
  for (let c = left; c <= data.c2; c++) {
    const values: (number | null)[] = [];
    for (let r = top; r <= data.r2; r++) {
      const v = value(r, c);
      values.push(typeof v === "number" ? v : null);
    }
    series.push({ name: headed ? doc.shown(data.r1, c, sheet) : `Series ${c - left + 1}`, values });
  }
  const format = lookOf(doc.book.styles, doc.cell(top, left, sheet)?.s).numFmt;
  return { categories, series, format, headed, labelled: firstColLabels };
}

