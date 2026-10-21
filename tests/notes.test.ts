import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { Doc, newWorkbook } from "../src/doc";
import { key, newSheet, type Workbook } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

function sample(): Workbook {
  const book: Workbook = { sheets: [newSheet("Notes")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = book.sheets[0];
  s.cells.set(key(0, 0), { v: "Rent" });
  s.notes.set(key(0, 0), { text: "Paid on the 1st", author: "Kari" });
  s.notes.set(key(4, 2), { text: "Check this <later> & again", author: "" });
  return book;
}

describe("notes", () => {
  it("are kept in the .xlsx file, with their authors", async () => {
    const book = sample();
    const bytes = await writeXlsx(book, new Calculator(book));
    const again = await readXlsx(bytes);
    expect([...again.sheets[0].notes]).toEqual([...book.sheets[0].notes]);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file("xl/worksheets/sheet1.xml")!.async("string")).toMatch(/<legacyDrawing r:id="rId\d+"\/>/);
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain('Extension="vml"');
  });

  it("go when the last one is taken off", async () => {
    const book = sample();
    const again = await readXlsx(await writeXlsx(book, new Calculator(book)));
    again.sheets[0].notes.clear();
    const zip = await JSZip.loadAsync(await writeXlsx(again, new Calculator(again)));
    expect(zip.file(/comments/).length).toBe(0);
    expect(await zip.file("xl/worksheets/sheet1.xml")!.async("string")).not.toContain("legacyDrawing");
  });

  it("move with their cells when rows are inserted, and can be undone", () => {
    const doc = new Doc(newWorkbook());
    doc.setNote(3, 1, "Hello", "");
    doc.insertDelete("r", 0, 2);
    expect(doc.sheet.notes.get(key(5, 1))?.text).toBe("Hello");
    doc.undo();
    expect(doc.sheet.notes.get(key(3, 1))?.text).toBe("Hello");
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("notes in the real Excel", () => {
  it("are Excel's own notes", async () => {
    const book = sample();
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-notes-"));
    const file = join(dir, "notes.xlsx");
    writeFileSync(file, await writeXlsx(book, new Calculator(book)));
    const script = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $s = $wb.Worksheets.Item(1)
  $out = [ordered]@{ count = $s.Comments.Count; first = $s.Range('A1').Comment.Text(); author = $s.Range('A1').Comment.Author; second = $s.Range('C5').Comment.Text() }
  $wb.Close($false)
  $out | ConvertTo-Json -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    expect(report).toMatchObject({ count: 2, author: "Kari", second: "Check this <later> & again" });
    expect(report.first.replace(/\r/g, "")).toBe("Kari:\nPaid on the 1st");
  }, 60_000);
});
