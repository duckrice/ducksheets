import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Doc, newWorkbook } from "../src/doc";
import { readXlsx } from "../src/xlsx/read";
import { writeXlsx } from "../src/xlsx/write";

function sample() {
  const doc = new Doc(newWorkbook());
  [["Name", "Mark", ""], ["Kari", "5", "x"], ["Ola", "4", "y"], ["Per", "6", "z"]].forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
  doc.formatAsTable({ r1: 0, c1: 0, r2: 3, c2: 2 }, "#4472C4", "#D9E1F2", "TableStyleMedium2");
  return doc;
}

describe("Excel tables", () => {
  it("are made by Format as table, with every header named", () => {
    const doc = sample();
    expect(doc.sheet.tables).toMatchObject([{ name: "Table1", range: { r1: 0, c1: 0, r2: 3, c2: 2 }, style: "TableStyleMedium2" }]);
    expect(doc.value(0, 2)).toBe("Column3");
  });

  it("are saved as table parts and read back, with their filter", async () => {
    const doc = sample();
    const bytes = await writeXlsx(doc.book, doc.calc);
    const again = await readXlsx(bytes);
    expect(again.sheets[0].tables).toMatchObject([{ name: "Table1", range: { r1: 0, c1: 0, r2: 3, c2: 2 }, style: "TableStyleMedium2" }]);
    expect(again.sheets[0].filter?.range).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 });
  });

  it("can be referred to by name and column in formulas", () => {
    const doc = sample();
    doc.setInput(5, 0, "=SUM(Table1[Mark])");
    doc.setInput(5, 1, "=ROWS(Table1[#All])");
    doc.setInput(5, 2, "=COUNTA(Table1[[#Headers],[Name]:[Mark]])");
    doc.setInput(1, 2, "=[@Mark]*2");
    expect([doc.value(5, 0), doc.value(5, 1), doc.value(5, 2), doc.value(1, 2)]).toEqual([15, 4, 2, 10]);
    expect(doc.cell(5, 0)?.f).toBe("SUM(Table1[Mark])");
    doc.setInput(5, 0, "=SUM(Table1[Nope])");
    expect(doc.value(5, 0)).toEqual(expect.objectContaining({ code: "#REF!" }));
  });

  it("rename a column in formulas when its header is changed", () => {
    const doc = sample();
    doc.setInput(5, 0, "=SUM(Table1[Mark])");
    doc.setInput(1, 2, "=Table1[[#This Row],[Mark]]+[@Mark]");
    doc.setInput(0, 1, "Score");
    expect(doc.cell(5, 0)?.f).toBe("SUM(Table1[Score])");
    expect(doc.cell(1, 2)?.f).toBe("Table1[[#This Row],[Score]]+[@Score]");
    expect(doc.value(5, 0)).toBe(15);
  });

  it("grow and move with inserted rows", () => {
    const doc = sample();
    doc.insertDelete("r", 2, 1);
    expect(doc.sheet.tables[0].range).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
    doc.insertDelete("r", 0, 2);
    expect(doc.sheet.tables[0].range).toEqual({ r1: 2, c1: 0, r2: 6, c2: 2 });
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("tables in the real Excel", () => {
  it("are Excel's own tables, and survive a round trip through DuckSheets", async () => {
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-tables-"));
    const ours = join(dir, "ours.xlsx");
    const theirs = join(dir, "theirs.xlsx");
    const back = join(dir, "back.xlsx");
    const doc = sample();
    doc.setInput(1, 1, "=1+4");
    writeFileSync(ours, await writeXlsx(doc.book, doc.calc));
    const run = (script: string) => {
      const file = join(dir, `s${Math.random().toString(36).slice(2)}.ps1`);
      writeFileSync(file, `$ErrorActionPreference = 'Stop'\n$xl = New-Object -ComObject Excel.Application\n$xl.Visible = $false; $xl.DisplayAlerts = $false\ntry {\n${script}\n} finally { $xl.Quit() }`, "utf8");
      return JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", file], { encoding: "utf8" }));
    };
    const describeTables = (path: string) => `
  $wb = $xl.Workbooks.Open('${path}')
  $s = $wb.Worksheets.Item(1)
  $out = @($s.ListObjects | ForEach-Object { [ordered]@{ name = $_.Name; ref = $_.Range.Address(0, 0); style = $_.TableStyle.Name; headers = @($_.HeaderRowRange.Value2); filter = $_.ShowAutoFilter } })
  $wb.Close($false)
  ConvertTo-Json -Compress -Depth 4 -InputObject $out`;
    expect(run(describeTables(ours))).toEqual([{ name: "Table1", ref: "A1:C4", style: "TableStyleMedium2", headers: ["Name", "Mark", "Column3"], filter: true }]);
    // A table made in Excel, with a totals row, read and saved by DuckSheets after a change.
    run(`
  $wb = $xl.Workbooks.Add()
  $s = $wb.Worksheets.Item(1)
  $s.Range("A1:B1").Value2 = @("Fruit", "Kg")
  $s.Range("A2:B3").Value2 = [object[,]]::new(2, 2)
  $s.Range("A2").Value2 = "pear"; $s.Range("B2").Value2 = 2; $s.Range("A3").Value2 = "plum"; $s.Range("B3").Value2 = 3
  $t = $s.ListObjects.Add(1, $s.Range("A1:B3"), $null, 1)
  $t.Name = "Fruits"; $t.TableStyle = "TableStyleLight9"; $t.ShowTotals = $true
  $wb.SaveAs('${theirs}', 51)
  $wb.Close($false)
  '{}'`);
    const book = await readXlsx(new Uint8Array(readFileSync(theirs)));
    const doc2 = new Doc(book);
    doc2.setInput(0, 1, "Weight");
    writeFileSync(back, await writeXlsx(doc2.book, doc2.calc));
    expect(doc2.value(3, 1)).toBe(5);
    expect(doc2.cell(3, 1)?.f).toBe("SUBTOTAL(109,Fruits[Weight])");
    expect(run(describeTables(back))).toEqual([{ name: "Fruits", ref: "A1:B4", style: "TableStyleLight9", headers: ["Fruit", "Weight"], filter: true }]);
  }, 120_000);
});

describe("table styles on screen", () => {
  it("colour the header and every other row in the style's colour", async () => {
    const { tableLook } = await import("../src/structured");
    const sheet = { tables: [{ name: "T", range: { r1: 0, c1: 0, r2: 4, c2: 1 }, style: "TableStyleMedium2" }] };
    expect(tableLook(sheet, 0, 0)).toEqual({ fill: "#4472C4", color: "#FFFFFF", bold: true });
    expect([tableLook(sheet, 1, 0)?.fill, tableLook(sheet, 2, 0)?.fill]).toEqual(["#DAE3F3", null]);
    expect(tableLook(sheet, 5, 0)).toBeNull();
    expect(tableLook({ tables: [{ ...sheet.tables[0], style: "TableStyleLight1" }] }, 0, 0)).toEqual({ fill: null, bold: true });
  });
});
