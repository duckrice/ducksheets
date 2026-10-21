import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Calculator } from "../src/calc";
import { CellError, key, newSheet, type Scalar, type Workbook } from "../src/model";
import { defaultStyles } from "../src/xlsx/styles";
import { writeXlsx } from "../src/xlsx/write";

/** Data in A1:C6 for the formulas to use. */
function book(formulas: string[]): Workbook {
  const b: Workbook = { sheets: [newSheet("F")], styles: defaultStyles(), date1904: false, names: new Map() };
  const s = b.sheets[0];
  [
    [2, 1, -1000],
    [4, 3, 300],
    [6, 7, 400],
    [8, 8, 500],
    [10, 12, 45000],
    [12, 11, 45400],
  ].forEach((row, r) => row.forEach((v, c) => s.cells.set(key(r, c), { v })));
  formulas.forEach((f, i) => s.cells.set(key(i, 4), { f }));
  return b;
}

const FORMULAS = [
  "SIN(PI()/6)",
  "ATAN2(1,1)",
  "DEGREES(PI())",
  "GCD(24,36,60)",
  "LCM(4,6,10)",
  "MROUND(17,5)",
  "COMBIN(10,3)",
  "PERMUT(10,3)",
  "DEC2BIN(10)",
  "DEC2BIN(-3)",
  "DEC2HEX(255,4)",
  "HEX2DEC(\"FF\")",
  "BIN2DEC(\"1111111111\")",
  "HEX2BIN(\"1F\")",
  "BASE(255,16)",
  "DECIMAL(\"ff\",16)",
  "BITAND(12,10)",
  "BITXOR(12,10)",
  "BITLSHIFT(1,4)",
  "SLOPE(B1:B6,A1:A6)",
  "INTERCEPT(B1:B6,A1:A6)",
  "FORECAST.LINEAR(14,B1:B6,A1:A6)",
  "CORREL(A1:A6,B1:B6)",
  "RSQ(B1:B6,A1:A6)",
  "COVARIANCE.S(A1:A6,B1:B6)",
  "GEOMEAN(A1:A6)",
  "HARMEAN(A1:A6)",
  "DEVSQ(A1:A6)",
  "AVEDEV(A1:A6)",
  "SKEW(B1:B6)",
  "KURT(B1:B6)",
  "STANDARDIZE(5,4,2)",
  "NORM.S.DIST(1,TRUE)",
  "NORM.DIST(12,10,2,TRUE)",
  "NORM.INV(0.9,10,2)",
  "NORM.S.INV(0.975)",
  "BINOM.DIST(3,10,0.5,FALSE)",
  "BINOM.DIST(3,10,0.5,TRUE)",
  "POISSON.DIST(2,3,TRUE)",
  "EXPON.DIST(1,2,TRUE)",
  "RANK.AVG(8,B1:B6)",
  "PERCENTRANK.INC(A1:A6,7)",
  "IPMT(0.05/12,1,36,8000)",
  "IPMT(0.05/12,3,36,8000,0,1)",
  "PPMT(0.05/12,10,36,8000)",
  "NPER(0.05/12,-200,8000)",
  "RATE(36,-250,8000)",
  "IRR(C1:C4)",
  "XNPV(0.08,C1:C4,C3:C6)",
  "SLN(10000,1000,5)",
  'TEXTBEFORE("user@example.com","@")',
  'TEXTAFTER("a-b-c","-",-1)',
  "UNICHAR(9731)",
  'UNICODE("Å")',
  "FIXED(1234.567,1)",
  "LOOKUP(7,A1:A6,B1:B6)",
  "ADDRESS(3,28,2)",
  "SUM(TAKE(A1:A6,2))",
  "SUM(DROP(A1:A6,4))",
  "SUM(CHOOSECOLS(A1:B6,2))",
  "ROWS(VSTACK(A1:A2,B1:B3))",
  "TYPE(A1)",
  "YEARFRAC(45000,45400)",
  "DAYS360(45000,45400)",
  "SUBTOTAL(9,A1:A6)",
  "SUBTOTAL(101,A1:A6)",
  "LET(x,A2*2,y,3,x+y)",
  'COMPLEX(3,4)',
  'IMABS("3+4i")',
  'IMSUM("3+4i","1-2i")',
  'IMPRODUCT("3+4i","1-2i")',
  'IMDIV("3+4i","1-2i")',
  'IMSUB("3+4j","1+j")',
  'IMREAL("2-5i")',
  'IMAGINARY("2-5i")',
  'IMCONJUGATE("2-5i")',
  'IMARGUMENT("0+1i")',
];

const ours = (formulas: string[]) => {
  const b = book(formulas);
  const calc = new Calculator(b);
  return formulas.map((_, i) => calc.value(0, i, 4));
};

describe("more of Excel's functions", () => {
  it("work out as expected", () => {
    const at = (f: string) => ours([f])[0];
    expect(at("DEC2BIN(10)")).toBe("1010");
    expect(at("DEC2BIN(-3)")).toBe("1111111101");
    expect(at("HEX2DEC(\"FF\")")).toBe(255);
    expect(at("GCD(24,36,60)")).toBe(12);
    expect(at("LET(x,A2*2,y,3,x+y)")).toBe(11);
    expect(at('IMPRODUCT("3+4i","1-2i")')).toBe("11-2i");
    expect(at('TEXTBEFORE("user@example.com","@")')).toBe("user");
    expect(at("SUBTOTAL(9,A1:A6)")).toBe(42);
    expect(at("NORM.S.INV(0.975)") as number).toBeCloseTo(1.959964, 5);
    expect(at("NOSUCHTHING(1)")).toEqual(new CellError("#NAME?"));
  });
});

const excel = process.env.EXCEL_CHECK ? describe : describe.skip;

excel("more functions, against the real Excel", () => {
  it("give the same results as Excel", async () => {
    const b = book(FORMULAS);
    const dir = mkdtempSync(join(tmpdir(), "ducksheets-fn-"));
    const file = join(dir, "fn.xlsx");
    writeFileSync(file, await writeXlsx(b, new Calculator(b)));
    const script = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open('${file}')
  $xl.CalculateFull()
  $s = $wb.Worksheets.Item(1)
  $out = @(1..${FORMULAS.length} | ForEach-Object { $v = $s.Cells.Item($_, 5).Value2; if ($v -is [int]) { "E:" + $s.Cells.Item($_, 5).Text } else { $v } })
  $wb.Close($false)
  ConvertTo-Json -InputObject $out -Compress
} finally { $xl.Quit() }`;
    const scriptFile = join(dir, "check.ps1");
    writeFileSync(scriptFile, script, "utf8");
    const theirs: (string | number | boolean)[] = JSON.parse(execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptFile], { encoding: "utf8" }));
    const mine = ours(FORMULAS);
    const wrong = FORMULAS.map((f, i) => [f, mine[i], theirs[i]] as [string, Scalar, unknown]).filter(([, a, b]) => {
      if (a instanceof CellError) return b !== `E:${a.code}`;
      if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) > 1e-6 * Math.max(1, Math.abs(b));
      return a !== b;
    });
    expect(wrong).toEqual([]);
  }, 120_000);
});
