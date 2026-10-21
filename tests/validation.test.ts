import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { Doc } from "../src/doc";
import { key, newSheet, type Workbook } from "../src/model";
import { allowed, describe as describeRule, listItems, validationAt, type Validation } from "../src/validation";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Form"), newSheet("Lists")], styles: defaultStyles(), date1904: false, names: new Map() };
  ["Oslo", "Bergen", "", "Trondheim"].forEach((v, r) => v && book.sheets[1].cells.set(key(r, 0), { v }));
  return book;
}

const list = (f: string): Validation => ({ ranges: [{ r1: 0, c1: 0, r2: 9, c2: 0 }], t: "list", f });

describe("data validation", () => {
  it("lists typed choices, or the values in cells (on another sheet too), skipping blanks", () => {
    const calc = new Calculator(sample());
    expect(listItems(list('"Yes,No, Maybe"') as Extract<Validation, { t: "list" }>, 0, calc)).toEqual(["Yes", "No", "Maybe"]);
    expect(listItems(list("Lists!$A$1:$A$4") as Extract<Validation, { t: "list" }>, 0, calc)).toEqual(["Oslo", "Bergen", "Trondheim"]);
  });

  it("allows only what the rule allows; an empty cell always", () => {
    const calc = new Calculator(sample());
    const cities = list("Lists!$A$1:$A$4");
    expect(allowed(cities, "bergen", 0, calc)).toBe(true);
    expect(allowed(cities, "Stavanger", 0, calc)).toBe(false);
    expect(allowed(cities, null, 0, calc)).toBe(true);
    const age: Validation = { ranges: [], t: "whole", op: "between", f: ["0", "120"] };
    expect([allowed(age, 30, 0, calc), allowed(age, 30.5, 0, calc), allowed(age, 130, 0, calc), allowed(age, "thirty", 0, calc)]).toEqual([true, false, false, false]);
    const price: Validation = { ranges: [], t: "decimal", op: "greaterThan", f: ["0"] };
    expect([allowed(price, 0.5, 0, calc), allowed(price, 0, 0, calc)]).toEqual([true, false]);
    expect(describeRule(age)).toBe("a whole number between 0 and 120");
  });

  it("replaces the rule for the cells it's set on, and Undo brings the old one back", () => {
    const doc = new Doc(sample());
    doc.setValidation({ r1: 0, c1: 0, r2: 9, c2: 0 }, list('"A,B"'));
    doc.setValidation({ r1: 2, c1: 0, r2: 3, c2: 0 }, { ranges: [], t: "whole", op: "greaterThan", f: ["0"] });
    expect(validationAt(doc.sheet, 0, 0)?.t).toBe("list");
    expect(validationAt(doc.sheet, 2, 0)?.t).toBe("whole");
    doc.undo();
    expect(validationAt(doc.sheet, 2, 0)?.t).toBe("list");
  });

  it("keeps rules in the .xlsx file, and ones it doesn't check unchanged", async () => {
    const book = sample();
    const dates = '<dataValidation type="date" operator="greaterThan" allowBlank="1" showErrorMessage="1" sqref="C1"><formula1>45000</formula1></dataValidation>';
    book.sheets[0].validations = [
      list("Lists!$A$1:$A$4"),
      { ranges: [{ r1: 0, c1: 1, r2: 4, c2: 1 }], t: "whole", op: "between", f: ["1", "6"] },
      { ranges: [{ r1: 0, c1: 2, r2: 0, c2: 2 }], t: "other", xml: dates },
    ];
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    expect(again.sheets[0].validations.map((v) => v.t)).toEqual(["list", "whole", "other"]);
    expect(again.sheets[0].validations[0]).toEqual(list("Lists!$A$1:$A$4"));
    expect(again.sheets[0].validations[1]).toMatchObject({ op: "between", f: ["1", "6"] });
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("data validation in the real Excel", () => {
  it("is Excel's own validation", async () => {
    const book = sample();
    book.sheets[0].validations = [list('"Yes,No"'), { ranges: [{ r1: 0, c1: 1, r2: 4, c2: 1 }], t: "whole", op: "between", f: ["1", "6"] }];
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-dv-"));
    const file = join(dir, "dv.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{ listType = $s.Range('A3').Validation.Type; list = $s.Range('A3').Validation.Formula1; numType = $s.Range('B2').Validation.Type; min = $s.Range('B2').Validation.Formula1; max = $s.Range('B2').Validation.Formula2 }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script);
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    // xlValidateList is 3, xlValidateWholeNumber 1; Excel shows the list with the separator of its language (; in Norwegian).
    expect(report).toMatchObject({ listType: 3, numType: 1, min: "1", max: "6" });
    expect(report.list).toMatch(/^Yes[,;]No$/);
  }, 60_000);
});
