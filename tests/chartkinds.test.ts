import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { key, newSheet, type ChartKind, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

function sample(kinds: ChartKind[]): Workbook {
  const book: Workbook = { sheets: [newSheet("Sales")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  [["Month", "Oslo", "Bergen"], ["Jan", 10, 7], ["Feb", 14, 9], ["Mar", 12, 15]].forEach((row, r) => row.forEach((v, c) => s.cells.set(key(r, c), { v })));
  kinds.forEach((kind, i) => s.charts.push({ id: `c${i}`, kind, title: kind, data: { r1: 0, c1: 0, r2: 3, c2: kind === "doughnut" ? 1 : 2 }, at: { r: 5 + i * 16, c: 1, dx: 0, dy: 0, w: 400, h: 240 } }));
  return book;
}

describe("area and doughnut charts", () => {
  it("are kept in the .xlsx file as what they are", async () => {
    const book = sample(["area", "doughnut"]);
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    expect(again.sheets[0].charts.map((c) => c.kind)).toEqual(["area", "doughnut"]);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("area and doughnut charts in the real Excel", () => {
  it("are Excel's own charts of those kinds", async () => {
    const book = sample(["area", "doughnut"]);
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-charts-"));
    const file = join(dir, "charts.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{ count = $s.ChartObjects().Count; first = $s.ChartObjects(1).Chart.ChartType; second = $s.ChartObjects(2).Chart.ChartType; series = $s.ChartObjects(1).Chart.SeriesCollection().Count }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // xlArea is 1, xlDoughnut -4120.
    expect(report).toEqual({ count: 2, first: 1, second: -4120, series: 2 });
  }, 60_000);
});
