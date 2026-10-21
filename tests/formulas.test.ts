import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { afterInsertDelete, parse, print, shifted } from "../src/formula/parse";
import { CellError, key, newSheet, parseAddress, type Scalar, type Workbook } from "../src/model";
import { dateToSerial, formatValue, locale, parseInput } from "../src/numfmt";
import { defaultStyles } from "../src/xlsx/styles";

/** A workbook from {"A1": value or "=formula"} per sheet. */
function book(...sheets: Record<string, Scalar>[]): Workbook {
  const wb: Workbook = { sheets: [], styles: defaultStyles(), date1904: false, names: new Map() };
  sheets.forEach((cells, i) => {
    const sheet = newSheet(`Sheet${i + 1}`);
    for (const [at, v] of Object.entries(cells)) {
      const { r, c } = parseAddress(at)!;
      sheet.cells.set(key(r, c), typeof v === "string" && v.startsWith("=") ? { f: v.slice(1) } : { v });
    }
    wb.sheets.push(sheet);
  });
  return wb;
}

/** The value of a formula on a sheet with the given cells. */
function calc(formula: string, cells: Record<string, Scalar> = {}): Scalar {
  const wb = book({ ...cells, Z100: `=${formula}` });
  return new Calculator(wb).value(0, 99, 25);
}

const code = (v: Scalar) => (v instanceof CellError ? v.code : v);

describe("reading and writing formulas", () => {
  it("follows Excel's order of operations", () => {
    expect(calc("1+2*3")).toBe(7);
    expect(calc("-2^2")).toBe(4);
    expect(calc("2^3^2")).toBe(64);
    expect(calc("(1+2)*3")).toBe(9);
    expect(calc("50%*4")).toBe(2);
    expect(calc('"a"&1+1')).toBe("a2");
    expect(calc("1=1")).toBe(true);
    expect(calc('"abc"="ABC"')).toBe(true);
  });

  it("writes formulas back as they were", () => {
    for (const f of ["SUM(A1:B2)", "$A$1+Sheet2!B3", "'My sheet'!A1*2", "IF(A1>0,\"yes\",\"no\")", "SUM(A:A)", "SUM(2:3)", "-A1%"]) expect(print(parse(f))).toBe(f);
    expect(print(parse("SUM(A1;B2)"))).toBe("SUM(A1,B2)");
    expect(print(parse("_xlfn.XLOOKUP(A1,B:B,C:C)"))).toBe("XLOOKUP(A1,B:B,C:C)");
    expect(print(parse("XLOOKUP(A1,B:B,C:C)"), true)).toBe("_xlfn.XLOOKUP(A1,B:B,C:C)");
  });

  it("moves relative references when copied, not $ ones", () => {
    expect(print(shifted(parse("A1+$B$1+$C2+D$3"), 2, 1))).toBe("B3+$B$1+$C4+E$3");
    expect(print(shifted(parse("A1"), -1, 0))).toBe("#REF!");
  });

  it("keeps references right when rows are inserted or deleted", () => {
    const after = (f: string, at: number, count: number) => print(afterInsertDelete(parse(f), "Sheet1", "Sheet1", "r", at, count));
    expect(after("SUM(A1:A10)", 4, 2)).toBe("SUM(A1:A12)");
    expect(after("A5+A2", 2, 1)).toBe("A6+A2");
    expect(after("A5", 4, -1)).toBe("#REF!");
    expect(after("SUM(A1:A10)", 2, -3)).toBe("SUM(A1:A7)");
    expect(print(afterInsertDelete(parse("Sheet2!A5"), "Sheet1", "Sheet1", "r", 0, 1))).toBe("Sheet2!A5");
  });

  it("reports what it can't read", () => {
    expect(() => parse("SUM(A1")).toThrow();
    expect(code(calc("NOSUCHFUNCTION(1)"))).toBe("#NAME?");
    expect(code(calc("1/0"))).toBe("#DIV/0!");
  });
});

describe("functions", () => {
  const data = { A1: 3, A2: 5, A3: "x", A4: 10, A5: null, B1: "apple", B2: "banana", B3: "cherry", B4: "apple", C1: 1, C2: 2, C3: 3, C4: 4 };

  it("adds up and counts as Excel does", () => {
    expect(calc("SUM(A1:A5)", data)).toBe(18);
    expect(calc("AVERAGE(A1:A4)", data)).toBe(6);
    expect(calc("COUNT(A1:A5)", data)).toBe(3);
    expect(calc("COUNTA(A1:A5)", data)).toBe(4);
    expect(calc('SUM("3",TRUE)')).toBe(4);
    expect(calc("MAX(A1:A4)", data)).toBe(10);
    expect(calc("MEDIAN(1,2,3,4)")).toBe(2.5);
    expect(calc("ROUND(2.675,2)")).toBe(2.68);
    expect(calc("ROUND(1234.5,-2)")).toBe(1200);
    expect(calc("MOD(-3,2)")).toBe(1);
  });

  it("does conditional sums and counts", () => {
    expect(calc('COUNTIF(B1:B4,"apple")', data)).toBe(2);
    expect(calc('COUNTIF(A1:A4,">4")', data)).toBe(2);
    expect(calc('SUMIF(B1:B4,"apple",C1:C4)', data)).toBe(5);
    expect(calc('SUMIFS(C1:C4,B1:B4,"a*",C1:C4,">1")', data)).toBe(4);
    expect(calc('COUNTIF(B1:B4,"<>apple")', data)).toBe(2);
    expect(calc('AVERAGEIF(B1:B4,"apple",C1:C4)', data)).toBe(2.5);
  });

  it("looks things up", () => {
    expect(calc("VLOOKUP(\"banana\",B1:C4,2,FALSE)", data)).toBe(2);
    expect(code(calc("VLOOKUP(\"kiwi\",B1:C4,2,FALSE)", data))).toBe("#N/A");
    expect(calc("XLOOKUP(\"cherry\",B1:B4,C1:C4)", data)).toBe(3);
    expect(calc("XLOOKUP(\"kiwi\",B1:B4,C1:C4,\"none\")", data)).toBe("none");
    expect(calc("INDEX(C1:C4,MATCH(\"cherry\",B1:B4,0))", data)).toBe(3);
    expect(calc("VLOOKUP(3.5,C1:C4,1)", data)).toBe(3);
    expect(calc("MATCH(2.5,C1:C4)", data)).toBe(2);
  });

  it("works with text", () => {
    expect(calc('CONCAT("a","b",1)')).toBe("ab1");
    expect(calc('TEXTJOIN(", ",TRUE,B1:B3)', data)).toBe("apple, banana, cherry");
    expect(calc('LEFT("Duck",2)&RIGHT("Duck",1)&MID("Duck",2,2)')).toBe("Dukuc");
    expect(calc('UPPER("æøå")')).toBe("ÆØÅ");
    expect(calc('PROPER("hello world")')).toBe("Hello World");
    expect(calc('SUBSTITUTE("a-b-c","-","+")')).toBe("a+b+c");
    expect(calc('FIND("c","abc")')).toBe(3);
    expect(calc('SEARCH("B*","abc")')).toBe(2);
    expect(calc('TRIM("  a   b ")')).toBe("a b");
  });

  it("works with dates", () => {
    expect(calc("DATE(2026,10,20)")).toBe(dateToSerial(2026, 10, 20));
    expect(calc("YEAR(DATE(2026,10,20))")).toBe(2026);
    expect(calc("DATE(2026,13,1)")).toBe(dateToSerial(2027, 1, 1));
    expect(calc("WEEKDAY(DATE(2026,10,20),2)")).toBe(2);
    expect(calc("EOMONTH(DATE(2026,2,10),0)")).toBe(dateToSerial(2026, 2, 28));
    expect(calc('DATEDIF(DATE(2000,5,1),DATE(2026,4,30),"Y")')).toBe(25);
    expect(calc("ISOWEEKNUM(DATE(2026,1,1))")).toBe(1);
    expect(calc("NETWORKDAYS(DATE(2026,10,19),DATE(2026,10,25))")).toBe(5);
    expect(calc("DATE(1900,3,1)")).toBe(61);
  });

  it("does logic and catches errors", () => {
    expect(calc("IF(A1>2,\"big\",\"small\")", data)).toBe("big");
    expect(calc("IFERROR(1/0,\"none\")")).toBe("none");
    expect(calc("AND(TRUE,1)")).toBe(true);
    expect(calc('IFS(A1>5,"a",A1>2,"b")', data)).toBe("b");
    expect(calc('SWITCH(2,1,"one",2,"two","other")')).toBe("two");
    expect(code(calc("A3+1", data))).toBe("#VALUE!");
  });

  it("does money", () => {
    expect(Math.round((calc("PMT(0.05/12,360,200000)") as number) * 100) / 100).toBe(-1073.64);
    expect(Math.round(calc("FV(0.05,10,-100)") as number)).toBe(1258);
  });
});

describe("whole workbooks", () => {
  it("follows formulas across sheets and named ranges", () => {
    const wb = book({ A1: 2, B1: "=Sheet2!A1*A1" }, { A1: 21 });
    wb.names.set("RATE", "Sheet1!$A$1");
    wb.sheets[0].cells.set(key(0, 2), { f: "B1*RATE" });
    const c = new Calculator(wb);
    expect(c.value(0, 0, 1)).toBe(42);
    expect(c.value(0, 0, 2)).toBe(84);
  });

  it("finds circular references instead of hanging", () => {
    const wb = book({ A1: "=B1+1", B1: "=A1+1" });
    const c = new Calculator(wb);
    c.all();
    expect(c.circular).toMatch(/Sheet1!(A1|B1)/);
  });

  it("works out long chains", () => {
    const cells: Record<string, Scalar> = { A1: 1 };
    for (let i = 2; i <= 20000; i++) cells[`A${i}`] = `=A${i - 1}+1`;
    const c = new Calculator(book(cells));
    c.all();
    expect(c.value(0, 19999, 0)).toBe(20000);
  });
});

describe("number formats", () => {
  const show = (v: Scalar, f: string) => formatValue(v, f).text;
  const d = locale.decimal;
  const g = locale.group;

  it("shows numbers", () => {
    expect(show(1234.567, "#,##0.00")).toBe(`1${g}234${d}57`);
    expect(show(0.256, "0.0%")).toBe(`25${d}6%`);
    expect(show(-5, "0;(0)")).toBe("(5)");
    expect(show(0, '0;-0;"zero"')).toBe("zero");
    expect(show(1234567, "#,##0,")).toBe(`1${g}235`);
    expect(show(12345, "0.00E+00")).toBe(`1${d}23E+04`);
    expect(show(5, '0 "kr"')).toBe("5 kr");
    expect(show(3.1, "0.##")).toBe(`3${d}1`);
    expect(show(42, "000")).toBe("042");
    expect(show(1 / 3, "General")).toBe(`0${d}3333333333`);
  });

  it("shows dates and times", () => {
    const serial = dateToSerial(2026, 10, 20) + 0.5;
    expect(show(serial, "yyyy-mm-dd")).toBe("2026-10-20");
    expect(show(serial, "dd.mm.yy hh:mm")).toBe("20.10.26 12:00");
    expect(show(serial, "h:mm AM/PM")).toBe("12:00 PM");
    expect(show(1.5, "[h]:mm")).toBe("36:00");
  });

  it("reads what's typed", () => {
    expect(parseInput(`1${d}5`).value).toBe(1.5);
    expect(parseInput("10%")).toEqual({ value: 0.1, format: "0%" });
    expect(parseInput("2026-10-20")).toEqual({ value: dateToSerial(2026, 10, 20), format: "yyyy-mm-dd" });
    expect(parseInput("12:30").value).toBeCloseTo(0.5208333);
    expect(parseInput("TRUE").value).toBe(true);
    expect(parseInput("'123").value).toBe("123");
    expect(parseInput("hello").value).toBe("hello");
  });
});
