import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { key, newSheet, type Chart, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

const EDITED: Omit<Chart, "id" | "at" | "data">[] = [
  { kind: "column", title: "", xTitle: "Month", yTitle: "Rain (mm)", colors: ["#C00000", "#70AD47"] },
  { kind: "bar", title: "Bars", xTitle: "Rain (mm)", yTitle: "Month", colors: ["#7030A0", "#ED7D31"] },
  { kind: "line", title: "Lines", colors: ["#000000", "#FFC000"] },
  { kind: "scatter", title: "Dots", xTitle: "Oslo", yTitle: "Bergen", colors: ["#00B050"] },
  { kind: "pie", title: "Pie", colors: ["#002060", "#FF0000", "#A5A5A5"] },
];

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Sales")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  [["Month", "Oslo", "Bergen"], ["Jan", 10, 7], ["Feb", 14, 9], ["Mar", 12, 15]].forEach((row, r) => row.forEach((v, c) => s.cells.set(key(r, c), { v })));
  EDITED.forEach((chart, i) => s.charts.push({ ...chart, id: `c${i}`, data: { r1: 0, c1: 0, r2: 3, c2: chart.kind === "pie" || chart.kind === "scatter" ? 1 : 2 }, at: { r: 5 + i * 16, c: 1, dx: 0, dy: 0, w: 400, h: 240 } }));
  return book;
}

describe("edited charts", () => {
  it("keep their axis titles and colours in the .xlsx file", async () => {
    const book = sample();
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    expect(again.sheets[0].charts.map(({ kind, title, xTitle, yTitle, colors }) => ({ kind, title, xTitle, yTitle, colors }))).toEqual(EDITED.map((c) => ({ xTitle: undefined, yTitle: undefined, ...c })));
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("edited charts in the real Excel", () => {
  it("show the same titles and colours", async () => {
    const book = sample();
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-chartedit-"));
    const file = join(dir, "charts.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
function title($axis) { if ($axis.HasTitle) { $axis.AxisTitle.Text } else { "" } }
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $col = $s.ChartObjects(1).Chart
  $bar = $s.ChartObjects(2).Chart
  $pie = $s.ChartObjects(5).Chart
  $out = [ordered]@{
    columnTitle = $col.HasTitle
    columnAxes = @((title $col.Axes(1)), (title $col.Axes(2)))
    columnFill = $col.SeriesCollection(1).Format.Fill.ForeColor.RGB
    barAxes = @((title $bar.Axes(1)), (title $bar.Axes(2)))
    barFill = $bar.SeriesCollection(1).Format.Fill.ForeColor.RGB
    lineColour = $s.ChartObjects(3).Chart.SeriesCollection(2).Format.Line.ForeColor.RGB
    pieSlice = $pie.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB
  }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // Excel's RGB numbers are red + green * 256 + blue * 65536.
    const rgb = (hex: string) => parseInt(hex.slice(1, 3), 16) + parseInt(hex.slice(3, 5), 16) * 256 + parseInt(hex.slice(5, 7), 16) * 65536;
    expect(report).toEqual({
      columnTitle: false,
      columnAxes: ["Month", "Rain (mm)"],
      columnFill: rgb("#C00000"),
      // A bar chart's category axis is the vertical one.
      barAxes: ["Month", "Rain (mm)"],
      barFill: rgb("#7030A0"),
      lineColour: rgb("#FFC000"),
      pieSlice: rgb("#FF0000"),
    });
  }, 60_000);
});
