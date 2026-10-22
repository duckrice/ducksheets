import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

const PNG_1PX = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Pictures")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  s.cells.set(key(0, 0), { v: 1 });
  s.charts.push({ id: "p1", kind: "picture", src: PNG_1PX, data: { r1: 0, c1: 0, r2: 0, c2: 0 }, title: "", at: { r: 2, c: 1, dx: 5, dy: 3, w: 160, h: 90 } });
  return book;
}

describe("pictures on sheets", () => {
  it("are kept in the .xlsx file, where they were and at their size", async () => {
    const book = sample();
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    const [picture] = again.sheets[0].charts;
    expect(picture).toMatchObject({ kind: "picture", src: PNG_1PX, at: { r: 2, c: 1, dx: 5, dy: 3, w: 160, h: 90 } });
    // Saving again doesn't pile up copies of the picture.
    const third = await readXlsx(await writeXlsx(again, new Calculator(again)));
    expect(third.sheets[0].charts.length).toBe(1);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("pictures in the real Excel", () => {
  it("are Excel's own pictures", async () => {
    const book = sample();
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-pics-"));
    const file = join(dir, "pics.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $p = $s.Shapes.Item(1)
  $out = [ordered]@{ count = $s.Shapes.Count; type = $p.Type; cell = $p.TopLeftCell.Address(0, 0); w = [math]::Round($p.Width); h = [math]::Round($p.Height) }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // msoPicture is 13; 160 × 90 pixels are 120 × 67.5 points.
    expect(report).toMatchObject({ count: 1, type: 13, cell: "B3", w: 120 });
    // Placed by its corner cells, as Excel does: within a few points, from rounding row heights.
    expect(Math.abs(report.h - 67.5)).toBeLessThan(4);
  }, 60_000);
});
