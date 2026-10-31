import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Doc, newWorkbook } from "../src/doc";
import { writeXlsx } from "../src/xlsx/write";

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("a summary in the real Excel", () => {
  it("works out the same sums, counts and averages", async () => {
    const doc = new Doc(newWorkbook());
    doc.renameSheet(0, "My data");
    [["Fruit", "Kg"], ["pear", "2"], ["Apple", "3"], ["pear", "4"], ["apple", "1"], ["Plum", "5"]].forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
    const ours: unknown[] = [];
    for (const fn of ["sum", "count", "average"] as const) {
      doc.active = 0;
      doc.summarize({ r1: 0, c1: 0, r2: 5, c2: 1 }, 0, 1, fn);
      ours.push([1, 2, 3, 4].map((r) => doc.value(r, 1)));
    }
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-summary-"));
    const file = join(dir, "summary.xlsx");
    writeFileSync(file, await writeXlsx(doc.book, doc.calc));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
# Excel can be busy a moment after opening (refreshing the PivotTables): ask again.
function Retry($block) { for ($i = 0; $i -lt 40; $i++) { try { return & $block } catch { Start-Sleep -Milliseconds 500 } }; throw 'Excel stayed busy' }
try {
  $wb = Retry { $xl.Workbooks.Open('${file}') }
  Retry { $wb.Application.CalculateFull() } | Out-Null
  $out = Retry { @(2, 3, 4) | ForEach-Object { $s = $wb.Worksheets.Item($_); ,@(2, 3, 4, 5 | ForEach-Object { $s.Cells.Item($_, 2).Value2 }) } }
  $wb.Close($false)
  ConvertTo-Json -Compress -InputObject $out
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const theirs = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // Each summary goes right after the data, so the last one made comes first.
    expect(theirs).toEqual(ours.reverse());
  }, 60_000);
});
