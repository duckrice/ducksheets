import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Doc, newWorkbook } from "../src/doc";
import type { FilterRule } from "../src/model";
import { readXlsx } from "../src/xlsx/read";
import { writeXlsx } from "../src/xlsx/write";

const DATA = [
  ["Name", "Team", "Score"],
  ["Kari", "Blue", "40"],
  ["Ola", "Red", "75"],
  ["Anne", "Blue", "90"],
  ["Per", "Red", "75"],
  ["Lise", "Green", "12"],
  ["Nils", "Blue", ""],
];

function sheet() {
  const doc = new Doc(newWorkbook());
  DATA.forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
  return doc;
}
const all = { r1: 0, c1: 0, r2: DATA.length - 1, c2: 2 };
const names = (doc: Doc) => Array.from({ length: DATA.length - 1 }, (_, i) => i + 1).map((r) => doc.value(r, 0));
const visible = (doc: Doc) => names(doc).filter((_, i) => !doc.sheet.hiddenRows.has(i + 1));

const RULES: [string, number, FilterRule, string[]][] = [
  ["greater than", 2, { kind: "custom", and: false, conditions: [{ op: "greaterThan", val: "50" }] }, ["Ola", "Anne", "Per"]],
  ["between (and)", 2, { kind: "custom", and: true, conditions: [{ op: "greaterThanOrEqual", val: "20" }, { op: "lessThan", val: "80" }] }, ["Kari", "Ola", "Per"]],
  ["contains", 0, { kind: "custom", and: false, conditions: [{ op: "equal", val: "*i*" }] }, ["Kari", "Lise", "Nils"]],
  ["does not contain", 0, { kind: "custom", and: false, conditions: [{ op: "notEqual", val: "*i*" }] }, ["Ola", "Anne", "Per"]],
  ["begins with or ends with", 0, { kind: "custom", and: false, conditions: [{ op: "equal", val: "a*" }, { op: "equal", val: "*s" }] }, ["Anne", "Nils"]],
  ["top 2 (ties count)", 2, { kind: "top", top: true, percent: false, count: 2 }, ["Ola", "Anne", "Per"]],
  ["bottom 1", 2, { kind: "top", top: false, percent: false, count: 1 }, ["Lise"]],
];

describe("sorting by several columns", () => {
  it("sorts by team, then by score largest first", () => {
    const doc = sheet();
    doc.sort(all, [{ col: 1, descending: false }, { col: 2, descending: true }], false, true);
    expect(names(doc)).toEqual(["Anne", "Kari", "Nils", "Lise", "Ola", "Per"]);
  });
});

describe("custom filters", () => {
  for (const [label, col, rule, shown] of RULES)
    it(label, () => {
      const doc = sheet();
      doc.toggleFilter();
      doc.setCustomFilter(col, rule);
      expect(visible(doc)).toEqual(shown);
    });

  it("a list filter on the column replaces the custom one, and both undo", () => {
    const doc = sheet();
    doc.toggleFilter();
    doc.setCustomFilter(2, RULES[0][2]);
    doc.setFilter(2, null);
    expect(visible(doc).length).toBe(6);
    doc.undo();
    expect(visible(doc)).toEqual(["Ola", "Anne", "Per"]);
  });

  it("are saved as Excel's own and read back", async () => {
    const doc = sheet();
    doc.toggleFilter();
    doc.setCustomFilter(2, RULES[1][2]);
    doc.setCustomFilter(0, RULES[5][2]);
    const again = await readXlsx(await writeXlsx(doc.book, doc.calc));
    expect([...again.sheets[0].filter!.custom!]).toEqual([
      [0, RULES[5][2]],
      [2, RULES[1][2]],
    ]);
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("custom filters in the real Excel", () => {
  it("hide the same rows", async () => {
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-filter-"));
    const files = RULES.map(([, col, rule], i) => {
      const doc = sheet();
      doc.toggleFilter();
      doc.setCustomFilter(col, rule);
      return [join(dir, `f${i}.xlsx`), doc] as const;
    });
    for (const [file, doc] of files) writeFileSync(file, await writeXlsx(doc.book, doc.calc));
    const script = `
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $out = @(${files.map(([f]) => `'${f}'`).join(",")}) | ForEach-Object {
    $wb = $xl.Workbooks.Open($_)
    $s = $wb.Worksheets.Item(1)
    $s.AutoFilter.ApplyFilter()
    $names = @(2..${DATA.length} | Where-Object { -not $s.Rows.Item($_).Hidden } | ForEach-Object { $s.Cells.Item($_, 1).Value2 })
    $wb.Close($false)
    ,$names
  }
  ConvertTo-Json -Compress -InputObject @($out)
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const theirs = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    expect(theirs).toEqual(RULES.map(([, , , shown]) => shown));
  }, 120_000);
});
