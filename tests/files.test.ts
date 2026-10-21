import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { Calculator } from "../src/calc";
import { parseCsv, readCsv, separatorOf, writeCsv } from "../src/csv";
import { CellError, key, newSheet, type Workbook } from "../src/model";
import { readSheet, readXlsx } from "../src/xlsx/read";
import { defaultStyles, lookOf, withStyle } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Budget"), newSheet("Rates 2026")], styles: defaultStyles(), date1904: false, names: new Map() };
  const [a, b] = book.sheets;
  const bold = withStyle(book.styles, 0, { bold: true, fill: "#FFFF00" });
  const money = withStyle(book.styles, 0, { numFmt: '#,##0.00 "kr"', border: { bottom: { style: "thin" } } });
  a.cells.set(key(0, 0), { v: "Item", s: bold });
  a.cells.set(key(0, 1), { v: "Cost", s: bold });
  a.cells.set(key(1, 0), { v: "Rent" });
  a.cells.set(key(1, 1), { v: 9500, s: money });
  a.cells.set(key(2, 0), { v: "Food & drink <weekly>" });
  a.cells.set(key(2, 1), { v: 1200.5, s: money });
  a.cells.set(key(3, 0), { v: true });
  a.cells.set(key(3, 1), { f: "SUM(B2:B3)*'Rates 2026'!A1", s: money });
  a.cells.set(key(4, 1), { f: "XLOOKUP(\"Rent\",A2:A3,B2:B3)" });
  a.cells.set(key(5, 1), { v: new CellError("#N/A") });
  a.colWidths.set(0, 150);
  a.rowHeights.set(0, 30);
  a.merges.push({ r1: 6, c1: 0, r2: 6, c2: 1 });
  a.freeze = { rows: 1, cols: 0 };
  a.filter = { range: { r1: 0, c1: 0, r2: 2, c2: 1 }, shown: new Map([[0, new Set(["Rent"])]]) };
  a.hiddenRows.add(2);
  b.cells.set(key(0, 0), { v: 1.25 });
  return book;
}

describe("Excel files", () => {
  it("write and read back what's in them", async () => {
    const book = sample();
    const bytes = await writeXlsx(book, new Calculator(book));
    const back = await readXlsx(bytes);
    const [a, b] = back.sheets;
    expect(back.sheets.map((s) => s.name)).toEqual(["Budget", "Rates 2026"]);
    expect(a.cells.get(key(2, 0))?.v).toBe("Food & drink <weekly>");
    expect(a.cells.get(key(1, 1))?.v).toBe(9500);
    expect(a.cells.get(key(3, 0))?.v).toBe(true);
    expect((a.cells.get(key(5, 1))?.v as CellError).code).toBe("#N/A");
    expect(a.cells.get(key(3, 1))?.f).toBe("SUM(B2:B3)*'Rates 2026'!A1");
    expect(a.cells.get(key(4, 1))?.f).toBe('XLOOKUP("Rent",A2:A3,B2:B3)');
    expect(new Calculator(back).value(0, 3, 1)).toBeCloseTo(13375.625);
    expect(lookOf(back.styles, a.cells.get(key(0, 0))?.s).font.bold).toBe(true);
    expect(lookOf(back.styles, a.cells.get(key(0, 0))?.s).fill).toBe("#FFFF00");
    expect(lookOf(back.styles, a.cells.get(key(1, 1))?.s).numFmt).toBe('#,##0.00 "kr"');
    expect(lookOf(back.styles, a.cells.get(key(1, 1))?.s).border.bottom?.style).toBe("thin");
    expect(a.colWidths.get(0)).toBe(150);
    expect(a.rowHeights.get(0)).toBe(30);
    expect(a.merges).toEqual([{ r1: 6, c1: 0, r2: 6, c2: 1 }]);
    expect(a.freeze).toEqual({ rows: 1, cols: 0 });
    expect(a.filter?.range).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 });
    expect([...a.filter!.shown.get(0)!]).toEqual(["Rent"]);
    expect(a.hiddenRows.has(2)).toBe(true);
    expect(b.cells.get(key(0, 0))?.v).toBe(1.25);
  });

  it("write new functions the way Excel expects, with values for other programs", async () => {
    const book = sample();
    const zip = await JSZip.loadAsync(await writeXlsx(book, new Calculator(book)));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheet).toContain("<f>_xlfn.XLOOKUP(");
    expect(sheet).toMatch(/<f>SUM\(B2:B3\)\*'Rates 2026'!A1<\/f><v>13375.625<\/v>/);
    const workbook = await zip.file("xl/workbook.xml")!.async("string");
    expect(workbook).toContain('fullCalcOnLoad="1"');
  });

  it("keep the parts DuckSheets doesn't edit, like charts", async () => {
    const book = sample();
    const first = await JSZip.loadAsync(await writeXlsx(book, new Calculator(book)));
    // As if Excel had put a chart on the first sheet.
    first.file("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'/>");
    const sheetXml = (await first.file("xl/worksheets/sheet1.xml")!.async("string")).replace("</worksheet>", '<drawing r:id="rId9"/></worksheet>');
    first.file("xl/worksheets/sheet1.xml", sheetXml);
    const reopened = await readXlsx(await first.generateAsync({ type: "uint8array" }));
    reopened.sheets[0].cells.set(key(10, 0), { v: "added" });
    const again = await JSZip.loadAsync(await writeXlsx(reopened, new Calculator(reopened)));
    expect(again.file("xl/drawings/drawing1.xml")).toBeTruthy();
    const sheet = await again.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheet).toContain('<drawing r:id="rId9"/>');
    expect(sheet.indexOf("<mergeCells")).toBeLessThan(sheet.indexOf("<drawing"));
    expect(sheet).toContain('<c r="A11" t="s">');
    expect(await again.file("xl/sharedStrings.xml")!.async("string")).toContain("<t>added</t>");
  });

  it("write charts as Excel charts, and read them back", async () => {
    const book = sample();
    book.sheets[0].charts.push({ id: "c1", kind: "column", title: "Costs", data: { r1: 0, c1: 0, r2: 2, c2: 1 }, at: { r: 1, c: 4, dx: 5, dy: 3, w: 400, h: 240 } });
    book.sheets[1].charts.push({ id: "c2", kind: "pie", title: "", data: { r1: 0, c1: 0, r2: 0, c2: 0 }, at: { r: 0, c: 2, dx: 0, dy: 0, w: 300, h: 200 } });
    const bytes = await writeXlsx(book, new Calculator(book));
    const zip = await JSZip.loadAsync(bytes);
    const chart = await zip.file("xl/charts/chart1.xml")!.async("string");
    expect(chart).toContain("<c:barChart><c:barDir val=\"col\"/>");
    expect(chart).toContain("<c:f>Budget!$B$2:$B$3</c:f>");
    expect(chart).toContain("<c:v>9500</c:v>");
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("drawingml.chart+xml");
    const back = await readXlsx(bytes);
    const [c] = back.sheets[0].charts;
    expect({ kind: c.kind, title: c.title, data: c.data }).toEqual({ kind: "column", title: "Costs", data: { r1: 0, c1: 0, r2: 2, c2: 1 } });
    expect([c.at.r, c.at.c, c.at.dx, c.at.dy]).toEqual([1, 4, 5, 3]);
    expect(Math.abs(c.at.w - 400)).toBeLessThan(2);
    // Saving again keeps one drawing per sheet with one chart each, not two.
    const again = await JSZip.loadAsync(await writeXlsx(back, new Calculator(back)));
    const sheetXml = await again.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheetXml.match(/<drawing /g)).toHaveLength(1);
    const drawing = await again.file("xl/drawings/drawing1.xml")!.async("string");
    expect(drawing.match(/twoCellAnchor /g)).toHaveLength(1);
  });

  it("read Excel's shared formulas, filled down a column", () => {
    const xml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
      <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>
    </sheetData></worksheet>`;
    const sheet = readSheet("S", xml, []);
    expect([0, 1, 2].map((r) => sheet.cells.get(key(r, 1))?.f)).toEqual(["A1*2", "A2*2", "A3*2"]);
  });

  it("say clearly when a file isn't a workbook", async () => {
    await expect(readXlsx(new TextEncoder().encode("not a zip"))).rejects.toThrow(/isn't an Excel workbook/);
  });
});

describe("CSV", () => {
  it("finds the separator and reads quotes", () => {
    expect(separatorOf("a;b;c\n1;2;3")).toBe(";");
    expect(separatorOf('a,"b;c",d')).toBe(",");
    expect(parseCsv('a,"b ""c"", d"\r\n1,2', ",")).toEqual([["a", 'b "c", d'], ["1", "2"]]);
  });

  it("reads numbers and dates, and writes what's shown", () => {
    const book: Workbook = { sheets: [], styles: defaultStyles(), date1904: false, names: new Map() };
    const sheet = readCsv(new TextEncoder().encode("Name;Amount\nRent;9500\nDay;2026-10-20\n"), "data", book);
    book.sheets.push(sheet);
    expect(sheet.cells.get(key(1, 1))?.v).toBe(9500);
    expect(lookOf(book.styles, sheet.cells.get(key(2, 1))?.s).numFmt).toBe("yyyy-mm-dd");
    const csv = writeCsv(book, 0, new Calculator(book));
    expect(csv.startsWith("﻿")).toBe(true);
    expect(csv).toContain("2026-10-20");
  });
});
