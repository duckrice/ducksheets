import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Doc, newWorkbook } from "../src/doc";
import { readXlsx } from "../src/xlsx/read";
import { writeXlsx } from "../src/xlsx/write";

/** A workbook with data on "Kari's data" and summaries of it. */
function book() {
  const doc = new Doc(newWorkbook());
  doc.renameSheet(0, "Kari's data");
  [["Fruit", "Kg", "Shop"], ["pear", "2", "A"], ["Apple", "3", "B"], ["pear", "4", "A"], ["apple", "1", "B"], ["Plum", "5", "A"]].forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
  doc.summarize({ r1: 0, c1: 0, r2: 5, c2: 2 }, 0, 1, "sum");
  doc.active = 0;
  doc.summarize({ r1: 0, c1: 0, r2: 5, c2: 2 }, 2, 1, "average");
  return doc;
}

describe("summaries saved as PivotTables", () => {
  it("write a cache and a table for each, and values (not formulas) in their cells", async () => {
    const doc = book();
    const zip = await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc));
    const files = Object.keys(zip.files);
    expect(files.filter((f) => /pivot(Tables|Cache)\/pivot\w+\d\.xml$/.test(f)).length).toBe(4);
    const workbook = await zip.file("xl/workbook.xml")!.async("string");
    expect(workbook.match(/<pivotCache /g)?.length).toBe(2);
    const caches = await Promise.all(files.filter((f) => f.includes("pivotCacheDefinition")).map((f) => zip.file(f)!.async("string")));
    expect(caches.join("")).toContain(`<worksheetSource ref="A1:B6" sheet="Kari's data"/>`);
    const sheets = await Promise.all(files.filter((f) => /worksheets\/sheet\d\.xml$/.test(f)).map((f) => zip.file(f)!.async("string")));
    expect(sheets.filter((s) => s.includes("<f>")).length).toBe(0);
  });

  it("still in DuckSheets are formulas, so they follow the data", () => {
    const doc = book();
    doc.active = 0;
    doc.setInput(1, 1, "20");
    doc.active = 2;
    expect([doc.sheet.name, doc.value(2, 1), doc.inputText(2, 1).startsWith("=")]).toEqual(["Summary", 24, true]);
  });

  it("a summary changed by hand is saved as plain cells", async () => {
    const doc = book();
    doc.active = 2;
    doc.setInput(2, 1, "7");
    const zip = await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc));
    expect((await zip.file("xl/workbook.xml")!.async("string")).match(/<pivotCache /g)?.length).toBe(1);
  });

  it("a file read back and saved again keeps them", async () => {
    const doc = book();
    const again = await readXlsx(await writeXlsx(doc.book, doc.calc));
    const zip = await JSZip.loadAsync(await writeXlsx(again, new Doc(again).calc));
    expect((await zip.file("xl/workbook.xml")!.async("string")).match(/<pivotCache /g)?.length).toBe(2);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("summaries in the real Excel", () => {
  it("are PivotTables with the same numbers, also after saving twice", async () => {
    const doc = book();
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-pivot-"));
    const once = join(dir, "once.xlsx");
    const twice = join(dir, "twice.xlsx");
    writeFileSync(once, await writeXlsx(doc.book, doc.calc));
    const again = await readXlsx(await writeXlsx(doc.book, doc.calc));
    writeFileSync(twice, await writeXlsx(again, new Doc(again).calc));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
# Excel refreshes the PivotTables as it opens the file, and turns calls away until it's done.
function v($x) { if ($x -is [double]) { "$([math]::Round($x, 3))" } else { "$x" } }
function Retry($block) { for ($i = 0; ; $i++) { try { return (& $block) } catch { if ($i -gt 60) { throw }; Start-Sleep -Milliseconds 500 } } }
try {
  $out = foreach ($file in @('${once}', '${twice}')) {
    $wb = $xl.Workbooks.Open($file)
    foreach ($n in @(2, 3)) {
      Retry {
        $s = $wb.Worksheets.Item($n)
        $pt = $s.PivotTables(1)
        "$($s.PivotTables().Count)|$($pt.DataFields(1).Function)|" + ((1..5 | ForEach-Object { $s.Cells.Item($_, 1).Text + '=' + (v $s.Cells.Item($_, 2).Value2) }) -join ';')
      }
    }
    $wb.Close($false)
  }
  ConvertTo-Json -Compress -InputObject @($out)
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const theirs = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // xlSum is -4157, xlAverage -4106. The summary made last is the first sheet after the data.
    const expected = ["1|-4106|Shop=Average of Kg;A=3.667;B=2;Total=3;=", "1|-4157|Fruit=Sum of Kg;Apple=4;pear=6;Plum=5;Total=15"];
    expect(theirs).toEqual([...expected, ...expected]);
  }, 120_000);
});
