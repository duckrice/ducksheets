import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { conditionalEffects, dxfXml, readDxf, withoutRules, type CfStyle } from "../src/conditional";
import { ruleValue } from "../src/cfmenu";
import { key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

const RED: CfStyle = { fill: "#FFC7CE", color: "#9C0006" };
const GREEN: CfStyle = { fill: "#C6EFCE", bold: true };

/** Marks in A1:A6, names in B1:B6, and a format list with red and green. */
function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Marks")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  [12, 45, 78, 45, 90, 3].forEach((v, r) => s.cells.set(key(r, 0), { v }));
  ["Kari", "Ola", "Per", "kari", "Anne", "Lise"].forEach((v, r) => s.cells.set(key(r, 1), { v }));
  book.styles.dxfs.push(dxfXml(RED), dxfXml(GREEN));
  return book;
}

const col = (c: number) => ({ r1: 0, c1: c, r2: 5, c2: c });
const look = (book: Workbook) => book.styles.dxfs.map((x) => readDxf(x, book.styles.theme));
const effects = (book: Workbook) => conditionalEffects(book.sheets[0], 0, new Calculator(book), look(book));
const filled = (book: Workbook, c: number) => [0, 1, 2, 3, 4, 5].map((r) => effects(book).get(key(r, c))?.fill ?? "-");

describe("conditional formatting", () => {
  it("reads back the formats it writes", () => {
    expect(readDxf(dxfXml(RED), [])).toEqual(RED);
    expect(readDxf(dxfXml(GREEN), [])).toEqual(GREEN);
  });

  it("colours cells by comparing them with a value, the first matching rule winning", () => {
    const book = sample();
    book.sheets[0].conditional = [
      { ranges: [col(0)], rules: [{ t: "cell", op: "greaterThan", f: ["50"], dxf: 0 }] },
      { ranges: [col(0)], rules: [{ t: "cell", op: "between", f: ["40", "80"], dxf: 1 }] },
    ];
    expect(filled(book, 0)).toEqual(["-", "#C6EFCE", "#FFC7CE", "#C6EFCE", "#FFC7CE", "-"]);
    expect(effects(book).get(key(1, 0))).toMatchObject({ bold: true });
  });

  it("finds text, duplicates (ignoring case), the top items and those above average", () => {
    const book = sample();
    const s = book.sheets[0];
    s.conditional = [{ ranges: [col(1)], rules: [{ t: "text", text: "ar", dxf: 0 }] }];
    expect(filled(book, 1)).toEqual(["#FFC7CE", "-", "-", "#FFC7CE", "-", "-"]);
    s.conditional = [{ ranges: [col(1)], rules: [{ t: "duplicate", dxf: 0 }] }];
    expect(filled(book, 1)).toEqual(["#FFC7CE", "-", "-", "#FFC7CE", "-", "-"]);
    s.conditional = [{ ranges: [col(0)], rules: [{ t: "top", n: 2, bottom: false, percent: false, dxf: 0 }] }];
    expect(filled(book, 0)).toEqual(["-", "-", "#FFC7CE", "-", "#FFC7CE", "-"]);
    s.conditional = [{ ranges: [col(0)], rules: [{ t: "average", below: true, dxf: 0 }] }];
    expect(filled(book, 0)).toEqual(["#FFC7CE", "#FFC7CE", "-", "#FFC7CE", "-", "#FFC7CE"]);
  });

  it("uses a formula written for the first cell, moved along for the others", () => {
    const book = sample();
    book.sheets[0].conditional = [{ ranges: [col(1)], rules: [{ t: "formula", f: "A1>=45", dxf: 0 }] }];
    expect(filled(book, 1)).toEqual(["-", "#FFC7CE", "#FFC7CE", "#FFC7CE", "#FFC7CE", "-"]);
  });

  it("shades colour scales and sizes data bars from the lowest to the highest", () => {
    const book = sample();
    book.sheets[0].conditional = [{ ranges: [col(0)], rules: [{ t: "scale", colors: ["#FFFFFF", "#000000"] }, { t: "bar", color: "#638EC6" }] }];
    const e = effects(book);
    expect(e.get(key(5, 0))?.fill).toBe("#ffffff");
    expect(e.get(key(4, 0))?.fill).toBe("#000000");
    expect(e.get(key(4, 0))?.bar?.frac).toBe(1);
    expect(e.get(key(5, 0))?.bar?.frac).toBeCloseTo(0.1);
  });

  it("takes rules off part of a block, keeping the rest", () => {
    const blocks = withoutRules([{ ranges: [{ r1: 0, c1: 0, r2: 9, c2: 0 }], rules: [] }], { r1: 3, c1: 0, r2: 4, c2: 0 });
    expect(blocks[0].ranges).toEqual([
      { r1: 0, c1: 0, r2: 2, c2: 0 },
      { r1: 5, c1: 0, r2: 9, c2: 0 },
    ]);
  });

  it("reads typed values as Excel does: numbers, formulas and text", () => {
    expect(ruleValue("50")).toBe("50");
    expect(ruleValue("2,5")).toBe("2.5");
    expect(ruleValue("=$C$1")).toBe("$C$1");
    expect(ruleValue('Say "hi"')).toBe('"Say ""hi"""');
  });

  it("keeps rules and their formats in the .xlsx file, and rules it doesn't show unchanged", async () => {
    const book = sample();
    const icons = '<cfRule type="iconSet" priority="9"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
    book.sheets[0].conditional = [
      { ranges: [col(0), { r1: 0, c1: 3, r2: 2, c2: 3 }], rules: [{ t: "cell", op: "greaterThan", f: ["50"], dxf: 0 }, { t: "other", xml: icons }] },
      { ranges: [col(1)], rules: [{ t: "text", text: "ar", dxf: 1 }] },
    ];
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    expect(again.sheets[0].conditional).toEqual([
      { ranges: [col(0), { r1: 0, c1: 3, r2: 2, c2: 3 }], rules: [{ t: "cell", op: "greaterThan", f: ["50"], dxf: 0 }, { t: "other", xml: expect.stringContaining('type="iconSet" priority="2"') }] },
      { ranges: [col(1)], rules: [{ t: "text", text: "ar", dxf: 1 }] },
    ]);
    expect(look(again)).toEqual([RED, GREEN]);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("conditional formatting in the real Excel", () => {
  it("shows DuckSheets' rules as Excel's own", async () => {
    const book = sample();
    book.sheets[0].conditional = [
      { ranges: [col(0)], rules: [{ t: "cell", op: "greaterThan", f: ["50"], dxf: 0 }] },
      { ranges: [col(1)], rules: [{ t: "duplicate", dxf: 1 }] },
      { ranges: [{ r1: 0, c1: 2, r2: 5, c2: 2 }], rules: [{ t: "bar", color: "#638EC6" }] },
    ];
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-cf-"));
    const file = join(dir, "cf.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{
    rules = $s.Cells.FormatConditions.Count
    big = $s.Range('A3').DisplayFormat.Interior.Color; small = $s.Range('A1').DisplayFormat.Interior.Color
    dup = $s.Range('B1').DisplayFormat.Interior.Color; once = $s.Range('B2').DisplayFormat.Interior.Color
    bold = $s.Range('B1').DisplayFormat.Font.Bold
  }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script);
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // Excel's colours are BGR numbers: #FFC7CE is 0xCEC7FF.
    expect(report).toMatchObject({ rules: 3, big: 0xcec7ff, small: 0xffffff, dup: 0xceefc6, once: 0xffffff, bold: true });
    void readFileSync;
  }, 60_000);
});
