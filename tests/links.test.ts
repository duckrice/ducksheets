import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { Doc, newWorkbook } from "../src/doc";
import { key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Links"), newSheet("Data")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  s.cells.set(key(0, 0), { v: "Duckware" });
  s.links.set(key(0, 0), "https://andris.no/duckware?x=1&y=2");
  s.cells.set(key(1, 0), { v: "Mail" });
  s.links.set(key(1, 0), "mailto:kari@example.com");
  s.cells.set(key(2, 0), { v: "The data" });
  s.links.set(key(2, 0), "#Data!B5");
  s.cells.set(key(3, 0), { f: 'HYPERLINK("https://example.com","Example")' });
  return book;
}

describe("links", () => {
  it("are kept in the .xlsx file: addresses and places in the workbook", async () => {
    const book = sample();
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    expect([...again.sheets[0].links]).toEqual([...book.sheets[0].links]);
  });

  it("HYPERLINK shows its name and points at its address", () => {
    const calc = new Calculator(sample());
    expect(calc.value(0, 3, 0)).toBe("Example");
    expect(calc.hyperlink(0, 3, 0)).toBe("https://example.com");
    expect(calc.hyperlink(0, 0, 0)).toBe(null);
  });

  it("move with their cells, and setting one fills an empty cell with the address", () => {
    const doc = new Doc(newWorkbook());
    doc.setLink(4, 1, "https://example.com");
    expect(doc.value(4, 1)).toBe("https://example.com");
    doc.insertDelete("r", 0, 1);
    expect(doc.sheet.links.get(key(5, 1))).toBe("https://example.com");
    doc.setLink(5, 1, "");
    expect(doc.sheet.links.size).toBe(0);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("links in the real Excel", () => {
  it("are Excel's own hyperlinks", async () => {
    const book = sample();
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-links-"));
    const file = join(dir, "links.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{ count = $s.Hyperlinks.Count; web = $s.Range('A1').Hyperlinks.Item(1).Address; mail = $s.Range('A2').Hyperlinks.Item(1).Address; place = $s.Range('A3').Hyperlinks.Item(1).SubAddress }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    expect(report).toEqual({ count: 3, web: "https://andris.no/duckware?x=1&y=2", mail: "mailto:kari@example.com", place: "Data!B5" });
  }, 60_000);
});
