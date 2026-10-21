// Checks DuckSheets' files against the real Excel, where it's installed:
// Excel opens a file DuckSheets wrote and reports what it finds, then saves
// it, and DuckSheets reads Excel's own file back. Run with EXCEL_CHECK=1.

import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { CellError, key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles, lookOf, withStyle } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

const run = process.env.EXCEL_CHECK ? describe : describe.skip;

run("with the real Excel", () => {
  it("opens DuckSheets' file as written, and DuckSheets reads Excel's", async () => {
    const book: Workbook = { sheets: [newSheet("Budget"), newSheet("Rates")], styles: defaultStyles(), date1904: false, names: new Map() };
    const [a, b] = book.sheets;
    const bold = withStyle(book.styles, 0, { bold: true, fill: "#FFFF00" });
    const money = withStyle(book.styles, 0, { numFmt: '#,##0.00 "kr"' });
    a.cells.set(key(0, 0), { v: "Item", s: bold });
    a.cells.set(key(0, 1), { v: "Cost", s: bold });
    a.cells.set(key(1, 0), { v: "Rent" });
    a.cells.set(key(1, 1), { v: 9500, s: money });
    a.cells.set(key(2, 0), { v: "Food" });
    a.cells.set(key(2, 1), { v: 1200.5, s: money });
    a.cells.set(key(3, 1), { f: "SUM(B2:B3)*Rates!A1", s: money });
    a.cells.set(key(4, 1), { f: 'XLOOKUP("Food",A2:A3,B2:B3)' });
    a.cells.set(key(5, 1), { f: 'COUNTIF(A2:A3,"F*")' });
    a.merges.push({ r1: 6, c1: 0, r2: 6, c2: 1 });
    a.cells.set(key(6, 0), { v: "Merged" });
    a.filter = { range: { r1: 0, c1: 0, r2: 2, c2: 1 }, shown: new Map() };
    a.colWidths.set(0, 150);
    a.colWidths.set(1, 120);
    a.freeze = { rows: 1, cols: 0 };
    b.cells.set(key(0, 0), { v: 1.25 });
    a.charts.push({ id: "c1", kind: "column", title: "Costs", data: { r1: 0, c1: 0, r2: 2, c2: 1 }, at: { r: 1, c: 4, dx: 0, dy: 0, w: 400, h: 240 } });

    const dir = mkdtempSync(join(tmpdir(), "ducksheets-excel-"));
    const ours = join(dir, "ours.xlsx");
    const theirs = join(dir, "theirs.xlsx");
    writeFileSync(ours, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${ours}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{
    sheets = @($wb.Worksheets | ForEach-Object { $_.Name })
    total = $s.Range('B4').Value2; totalFormula = $s.Range('B4').Formula
    lookup = $s.Range('B5').Value2; count = $s.Range('B6').Value2
    shown = $s.Range('B2').Text; bold = $s.Range('A1').Font.Bold; fill = $s.Range('A1').Interior.Color
    merged = $s.Range('A7').MergeCells; filter = $s.AutoFilterMode; width = $s.Columns.Item(1).ColumnWidth
    charts = $s.ChartObjects().Count; chartType = $s.ChartObjects(1).Chart.ChartType; chartTitle = $s.ChartObjects(1).Chart.ChartTitle.Text
    seriesName = $s.ChartObjects(1).Chart.SeriesCollection(1).Name; seriesFormula = $s.ChartObjects(1).Chart.SeriesCollection(1).Formula
  }
  $wb.SaveAs('${theirs}', 51); $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }`;
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-Command", script], { encoding: "utf8", timeout: 120000 }));
    expect(report.sheets).toEqual(["Budget", "Rates"]);
    expect(report.total).toBeCloseTo(13375.625);
    expect(report.totalFormula).toBe("=SUM(B2:B3)*Rates!A1");
    expect(report.lookup).toBe(1200.5);
    expect(report.count).toBe(1);
    expect(report.shown).toMatch(/9.500,00 kr|9,500.00 kr/);
    expect(report.bold).toBe(true);
    expect(report.fill).toBe(0x00ffff); // Excel's colours are BGR: yellow
    expect(report.merged).toBe(true);
    expect(report.filter).toBe(true);
    // 150 pixels is 20.71 characters of Calibri 11 (7 pixels each, plus 5 of padding).
    expect(report.width).toBeCloseTo(20.71, 1);
    expect(report.charts).toBe(1);
    expect(report.chartType).toBe(51); // a clustered column chart
    expect(report.chartTitle).toBe("Costs");
    expect(report.seriesName).toBe("Cost");
    expect(report.seriesFormula).toBe("=SERIES(Budget!$B$1,Budget!$A$2:$A$3,Budget!$B$2:$B$3,1)");

    const back = await readXlsx(new Uint8Array(readFileSync(theirs)));
    const s = back.sheets[0];
    expect(s.cells.get(key(3, 1))?.f).toBe("SUM(B2:B3)*Rates!A1");
    expect(s.cells.get(key(4, 1))?.f).toBe('XLOOKUP("Food",A2:A3,B2:B3)');
    expect(new Calculator(back).value(0, 3, 1)).toBeCloseTo(13375.625);
    expect(lookOf(back.styles, s.cells.get(key(0, 0))?.s).font.bold).toBe(true);
    expect(s.merges).toEqual([{ r1: 6, c1: 0, r2: 6, c2: 1 }]);
    expect(s.colWidths.get(0)).toBe(150);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    // Excel's own chart, read back as one DuckSheets can change.
    expect(s.charts.map((c) => [c.kind, c.title, c.data])).toEqual([["column", "Costs", { r1: 0, c1: 0, r2: 2, c2: 1 }]]);
    expect(s.cells.get(key(5, 1))?.v ?? new Calculator(back).value(0, 5, 1)).not.toBeInstanceOf(CellError);
  }, 180000);
});
