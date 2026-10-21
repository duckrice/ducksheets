import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { CellError, key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

/** Names in A1:A6, cities in B, marks in C. */
function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Data")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  const rows: [string, string, number][] = [
    ["Kari", "Oslo", 78],
    ["Ola", "Bergen", 45],
    ["Per", "Oslo", 90],
    ["Anne", "Tromsø", 12],
    ["Lise", "Bergen", 66],
    ["Nils", "Oslo", 45],
  ];
  rows.forEach((row, r) => row.forEach((v, c) => s.cells.set(key(r, c), { v })));
  return book;
}

/** Puts a formula in E1 and reads the block it fills (rows × cols from E1). */
function spill(formula: string, rows: number, cols = 1, book = sample()) {
  book.sheets[0].cells.set(key(0, 4), { f: formula });
  const calc = new Calculator(book);
  return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => calc.value(0, i, 4 + j)));
}

describe("dynamic arrays", () => {
  it("FILTER keeps the rows that meet a condition, and spills them", () => {
    expect(spill('FILTER(A1:A6,B1:B6="Oslo")', 4)).toEqual([["Kari"], ["Per"], ["Nils"], [null]]);
    expect(spill("FILTER(A1:C6,(C1:C6>50)*(B1:B6<>\"Oslo\"))", 2, 3)).toEqual([["Lise", "Bergen", 66], [null, null, null]]);
    expect(spill('FILTER(A1:A6,C1:C6>100,"none")', 1)).toEqual([["none"]]);
  });

  it("SORT, SORTBY, UNIQUE, SEQUENCE and TRANSPOSE", () => {
    expect(spill("SORT(C1:C6,1,-1)", 6).flat()).toEqual([90, 78, 66, 45, 45, 12]);
    expect(spill("SORTBY(A1:A6,B1:B6,1,C1:C6,-1)", 6).flat()).toEqual(["Lise", "Ola", "Per", "Kari", "Nils", "Anne"]);
    expect(spill("UNIQUE(B1:B6)", 4).flat()).toEqual(["Oslo", "Bergen", "Tromsø", null]);
    expect(spill("UNIQUE(B1:B6,FALSE,TRUE)", 1).flat()).toEqual(["Tromsø"]);
    expect(spill("SEQUENCE(2,3,10,5)", 2, 3)).toEqual([
      [10, 15, 20],
      [25, 30, 35],
    ]);
    expect(spill("TRANSPOSE(A1:A3)", 1, 3)).toEqual([["Kari", "Ola", "Per"]]);
  });

  it("arrays work inside other functions, and conditions inside SUMPRODUCT", () => {
    expect(spill('SUM(FILTER(C1:C6,B1:B6="Oslo"))', 1)).toEqual([[213]]);
    expect(spill('SUMPRODUCT((B1:B6="Bergen")*C1:C6)', 1)).toEqual([[111]]);
    expect(spill("COUNTA(UNIQUE(B1:B6))", 1)).toEqual([[3]]);
  });

  it("other formulas read the spilled cells", () => {
    const book = sample();
    book.sheets[0].cells.set(key(0, 4), { f: "SORT(C1:C6)" });
    book.sheets[0].cells.set(key(0, 6), { f: "E2+E6" });
    expect(new Calculator(book).value(0, 0, 6)).toBe(45 + 90);
  });

  it("shows #SPILL! when something is in the way", () => {
    const book = sample();
    book.sheets[0].cells.set(key(2, 4), { v: "blocking" });
    expect(spill("SORT(C1:C6)", 1, 1, book)[0][0]).toEqual(new CellError("#SPILL!"));
  });

  it("is saved as a dynamic array and read back without the saved copies of its values", async () => {
    const book = sample();
    book.sheets[0].cells.set(key(0, 4), { f: 'FILTER(A1:A6,B1:B6="Oslo")' });
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    const s = again.sheets[0];
    expect(s.cells.get(key(0, 4))?.f).toBe('FILTER(A1:A6,B1:B6="Oslo")');
    expect(s.cells.has(key(1, 4))).toBe(false);
    expect(new Calculator(again).value(0, 2, 4)).toBe("Nils");
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("dynamic arrays in the real Excel", () => {
  it("are Excel's own dynamic arrays, with the same values", async () => {
    const book = sample();
    book.sheets[0].cells.set(key(0, 4), { f: 'FILTER(A1:A6,B1:B6="Oslo")' });
    book.sheets[0].cells.set(key(0, 6), { f: "SORT(UNIQUE(B1:B6))" });
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-da-"));
    const file = join(dir, "da.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{ spills = $s.Range('E1').HasSpill; area = $s.Range('E1').SpillingToRange.Address(0, 0); third = $s.Range('E3').Text; formula = $s.Range('E1').Formula2; cities = (@($s.Range('G1:G3').Value2) -join ',') }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script);
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    expect(report).toMatchObject({ spills: true, area: "E1:E3", third: "Nils", cities: "Bergen,Oslo,Tromsø" });
    expect(report.formula).toMatch(/^=FILTER\(A1:A6[,;]B1:B6="Oslo"\)$/);
  }, 60_000);
});
