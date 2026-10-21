import { describe, expect, it } from "vitest";
import { Doc, newWorkbook, seriesOf } from "../src/doc";
import { key } from "../src/model";
import { dateToSerial, locale } from "../src/numfmt";
import { lookOf } from "../src/xlsx/styles";

const doc = () => new Doc(newWorkbook());
const put = (d: Doc, cells: Record<string, string>) => {
  for (const [at, text] of Object.entries(cells)) {
    const m = /^([A-Z])(\d+)$/.exec(at)!;
    d.setInput(Number(m[2]) - 1, m[1].charCodeAt(0) - 65, text);
  }
};
const v = (d: Doc, at: string) => {
  const m = /^([A-Z])(\d+)$/.exec(at)!;
  return d.value(Number(m[2]) - 1, m[1].charCodeAt(0) - 65);
};
const f = (d: Doc, at: string) => {
  const m = /^([A-Z])(\d+)$/.exec(at)!;
  return d.cell(Number(m[2]) - 1, m[1].charCodeAt(0) - 65)?.f;
};

describe("editing a sheet", () => {
  it("reads typing as Excel does, and formats percentages and dates", () => {
    const d = doc();
    put(d, { A1: "10%", A2: "2026-10-20", A3: "=a1*2", A4: `1${locale.decimal}5` });
    expect(v(d, "A1")).toBe(0.1);
    expect(d.shown(0, 0)).toBe("10%");
    expect(d.shown(1, 0)).toBe("2026-10-20");
    expect(f(d, "A3")).toBe("A1*2");
    expect(v(d, "A3")).toBe(0.2);
    expect(v(d, "A4")).toBe(1.5);
    expect(d.inputText(0, 0)).toBe("10%");
  });

  it("undoes and redoes", () => {
    const d = doc();
    put(d, { A1: "1" });
    put(d, { A1: "2" });
    d.undo();
    expect(v(d, "A1")).toBe(1);
    d.redo();
    expect(v(d, "A1")).toBe(2);
  });

  it("keeps formulas right when rows and columns are inserted or deleted", () => {
    const d = doc();
    put(d, { A1: "1", A2: "2", A3: "3", B1: "=SUM(A1:A3)", C1: "=A3*2" });
    d.insertDelete("r", 1, 2);
    expect(f(d, "B1")).toBe("SUM(A1:A5)");
    expect(v(d, "A5")).toBe(3);
    expect(v(d, "B1")).toBe(6);
    d.insertDelete("r", 4, -1);
    expect(f(d, "C1")).toBe("#REF!*2");
    d.insertDelete("c", 0, 1);
    expect(d.cell(0, 2)?.f).toBe("SUM(B1:B4)");
  });

  it("sorts by a column, keeping the heading row", () => {
    const d = doc();
    put(d, { A1: "Name", B1: "Score", A2: "Cid", B2: "3", A3: "Ada", B3: "9", A4: "Bo", B4: "5" });
    d.sort({ r1: 0, c1: 0, r2: 3, c2: 1 }, 1, true);
    expect([v(d, "A1"), v(d, "A2"), v(d, "A3"), v(d, "A4")]).toEqual(["Name", "Ada", "Bo", "Cid"]);
  });

  it("fills series and copies formulas with the fill handle", () => {
    const d = doc();
    put(d, { A1: "1", A2: "3", B1: "Week 1", C1: "=A1*10", D1: "January" });
    d.fill({ r1: 0, c1: 0, r2: 1, c2: 0 }, { r1: 0, c1: 0, r2: 4, c2: 0 });
    expect([v(d, "A3"), v(d, "A4"), v(d, "A5")]).toEqual([5, 7, 9]);
    d.fill({ r1: 0, c1: 1, r2: 0, c2: 3 }, { r1: 0, c1: 1, r2: 2, c2: 3 });
    expect(v(d, "B3")).toBe("Week 3");
    expect(f(d, "C3")).toBe("A3*10");
    expect(v(d, "C3")).toBe(50);
    expect(v(d, "D2")).toBe("February");
    expect(seriesOf([5])).toBeNull();
    expect(seriesOf([dateToSerial(2026, 1, 31)], true)!(1)).toBe(dateToSerial(2026, 2, 1));
  });

  it("copies and pastes formulas so their references move", () => {
    const d = doc();
    put(d, { A1: "2", B1: "=A1*$A$1" });
    const clip = d.copy({ r1: 0, c1: 1, r2: 0, c2: 1 });
    d.paste(clip, { r: 3, c: 2 });
    expect(d.cell(3, 2)?.f).toBe("B4*$A$1");
    d.pasteText("1\t2\r\n3\t4", { r: 10, c: 0 });
    expect(d.value(11, 1)).toBe(4);
  });

  it("merges, formats and filters", () => {
    const d = doc();
    put(d, { A1: "Title", B1: "x", A3: "Fruit", A4: "apple", A5: "pear", A6: "apple" });
    d.toggleMerge({ r1: 0, c1: 0, r2: 0, c2: 1 });
    expect(d.sheet.merges).toHaveLength(1);
    expect(d.cell(0, 1)).toBeUndefined();
    d.format({ r1: 2, c1: 0, r2: 2, c2: 0 }, { bold: true });
    expect(lookOf(d.book.styles, d.cell(2, 0)?.s).font.bold).toBe(true);
    d.sel = { r: 2, c: 0, ar: 2, ac: 0 };
    d.toggleFilter();
    expect(d.sheet.filter?.range).toEqual({ r1: 2, c1: 0, r2: 5, c2: 0 });
    d.setFilter(0, new Set(["apple"]));
    expect([...d.sheet.hiddenRows]).toEqual([4]);
  });

  it("renames sheets and keeps formulas pointing at them", () => {
    const d = doc();
    d.addSheet();
    put(d, { A1: "5" });
    d.active = 0;
    put(d, { A1: "=Sheet2!A1*2" });
    expect(d.renameSheet(1, "Data")).toBeNull();
    expect(f(d, "A1")).toBe("Data!A1*2");
    expect(v(d, "A1")).toBe(10);
    expect(d.renameSheet(1, "a/b")).toMatch(/can't contain/);
    d.deleteSheet(1);
    expect(f(d, "A1")).toBe("#REF!*2");
  });

  it("finds the block of data around a cell", () => {
    const d = doc();
    put(d, { B2: "1", C2: "2", B3: "3", C4: "4" });
    expect(d.region(1, 1)).toEqual({ r1: 1, c1: 1, r2: 3, c2: 2 });
    expect(d.sheet.cells.has(key(1, 1))).toBe(true);
  });
});
