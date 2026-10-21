// Comma- (or semicolon-, or tab-) separated text: what Excel reads and writes
// as .csv. In countries that write decimals with a comma, Excel separates
// with semicolons, so DuckSheets does too.

import { key, newSheet, type Sheet, type Workbook } from "./model";
import { formatValue, locale, parseInput } from "./numfmt";
import { lookOf, withStyle } from "./xlsx/styles";
import type { Calculator } from "./calc";

/** The separator a CSV text uses: whichever of ; , and tab is most common outside quotes in its first line. */
export function separatorOf(text: string): string {
  const line = text.split(/\r?\n/, 1)[0].replace(/"[^"]*"/g, "");
  const counts = [";", ",", "\t"].map((s) => [s, line.split(s).length - 1] as const);
  const best = counts.sort((a, b) => b[1] - a[1])[0];
  return best[1] > 0 ? best[0] : locale.decimal === "," ? ";" : ",";
}

export function parseCsv(text: string, separator = separatorOf(text)): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
    } else if (ch === '"' && field === "") quoted = true;
    else if (ch === separator) {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** A CSV file as one sheet: numbers, dates and percentages read as Excel reads them. */
export function readCsv(bytes: Uint8Array, name: string, book: Workbook): Sheet {
  let text = new TextDecoder("utf-8").decode(bytes);
  // Not UTF-8? Then it's most likely Windows' own encoding.
  if (text.includes("�")) text = new TextDecoder("windows-1252").decode(bytes);
  text = text.replace(/^﻿/, "");
  const sheet = newSheet(name);
  parseCsv(text).forEach((fields, r) =>
    fields.forEach((field, c) => {
      if (field === "") return;
      const typed = parseInput(field, book.date1904);
      sheet.cells.set(key(r, c), { v: typed.value, ...(typed.format ? { s: withStyle(book.styles, 0, { numFmt: typed.format }) } : {}) });
    }),
  );
  return sheet;
}

/** A sheet as CSV text, values as they're shown, with a byte-order mark so Excel reads æøå right. */
export function writeCsv(book: Workbook, index: number, calc: Calculator): string {
  const sheet = book.sheets[index];
  const separator = locale.decimal === "," ? ";" : ",";
  let rows = 0;
  let cols = 0;
  for (const k of sheet.cells.keys()) {
    rows = Math.max(rows, Math.floor(k / 16384) + 1);
    cols = Math.max(cols, (k % 16384) + 1);
  }
  const lines: string[] = [];
  for (let r = 0; r < rows; r++) {
    const fields: string[] = [];
    for (let c = 0; c < cols; c++) {
      const cell = sheet.cells.get(key(r, c));
      const value = calc.value(index, r, c);
      const text = formatValue(value, lookOf(book.styles, cell?.s).numFmt, book.date1904).text;
      fields.push(/["\n\r]/.test(text) || text.includes(separator) ? `"${text.replace(/"/g, '""')}"` : text);
    }
    lines.push(fields.join(separator));
  }
  return `﻿${lines.join("\r\n")}\r\n`;
}
