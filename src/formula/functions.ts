// Excel's functions, the ones people use day to day: sums and statistics,
// the conditional ...IF(S) family, logic, lookups, text, dates and a few
// financial ones. Each follows Excel's rules for what it accepts and returns.

import { CellError, type Scalar } from "../model";
import { dateToSerial, formatValue, parseInput, serialToDate } from "../numfmt";
import { ArrayValue, FUNCTIONS, RangeValue, cellsOf, compare, err, evaluate, isErr, scalar, toBool, toGrid, toNumber, toText, type EvalContext, type Fn, type Value } from "./eval";
import { parse, type Node } from "./parse";

// ---- helpers ------------------------------------------------------------------------------

const ev = (node: Node | undefined, ctx: EvalContext): Value => (node ? evaluate(node, ctx) : null);
const one = (node: Node | undefined, ctx: EvalContext): Scalar => scalar(ev(node, ctx), ctx);
const missing = (node: Node | undefined) => !node || (node.t === "name" && node.name === "");

function num(node: Node | undefined, ctx: EvalContext, fallback?: number): number | CellError {
  if (missing(node) && fallback !== undefined) return fallback;
  return toNumber(one(node, ctx));
}

/** Numbers for SUM, AVERAGE and the like: from ranges only real numbers, typed-in values converted. */
function numbers(args: Node[], ctx: EvalContext): number[] | CellError {
  const out: number[] = [];
  for (const arg of args) {
    const v = ev(arg, ctx);
    if (v instanceof RangeValue || v instanceof ArrayValue) {
      for (const c of cellsOf(v, ctx)) {
        if (isErr(c)) return c;
        if (typeof c === "number") out.push(c);
      }
    } else {
      if (v === null) continue;
      const n = toNumber(v);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

/** All values of the arguments, ranges flattened. */
function values(args: Node[], ctx: EvalContext): Scalar[] {
  const out: Scalar[] = [];
  for (const arg of args) {
    const v = ev(arg, ctx);
    if (v instanceof RangeValue || v instanceof ArrayValue) out.push(...cellsOf(v, ctx));
    else out.push(v);
  }
  return out;
}

function range(node: Node | undefined, ctx: EvalContext): RangeValue | CellError {
  const v = ev(node, ctx);
  if (v instanceof RangeValue) return v;
  return isErr(v) ? v : err("#VALUE!");
}

const cellAt = (r: RangeValue, i: number, j: number, ctx: EvalContext) => ctx.cell(r.sheet, r.r1 + i, r.c1 + j);

/** Wildcards * and ? (with ~ to escape them), as in COUNTIF and VLOOKUP. */
function wildcard(pattern: string): RegExp {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (ch === "*") re += "[\\s\\S]*";
    else if (ch === "?") re += "[\\s\\S]";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "i");
}

/** A COUNTIF criterion (5, ">5", "<>x", "a*") as a test. */
export function criterion(crit: Scalar): (v: Scalar) => boolean {
  if (typeof crit === "number") return (v) => (typeof v === "number" ? v === crit : typeof v === "string" && parseInput(v).value === crit);
  if (typeof crit === "boolean") return (v) => v === crit;
  if (crit === null || isErr(crit)) return (v) => v === crit;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(crit)!;
  const op = m[1] ?? "=";
  const text = m[2];
  const typed = parseInput(text).value;
  if (typeof typed === "number" && text !== "") {
    return (v) => {
      const n = typeof v === "number" ? v : typeof v === "string" && op === "=" ? parseInput(v).value : null;
      if (typeof n !== "number") return op === "<>";
      return { "=": n === typed, "<>": n !== typed, "<": n < typed, ">": n > typed, "<=": n <= typed, ">=": n >= typed }[op as "="];
    };
  }
  if (op === "=" || op === "<>") {
    // "=" alone matches empty cells, "<>" alone everything that isn't empty.
    if (text === "") return op === "=" ? (v) => v === null || v === "" : (v) => v !== null && v !== "";
    const re = wildcard(text);
    return op === "=" ? (v) => re.test(toText(v)) && typeof v !== "number" : (v) => !re.test(toText(v)) || typeof v === "number";
  }
  return (v) => {
    if (typeof v !== "string") return false;
    const c = compare(v, text);
    return { "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0 }[op as "<"];
  };
}

/** For the ...IFS functions: the positions in the ranges where every criterion holds. */
function matches(pairs: Node[], ctx: EvalContext): { rows: number; cols: number; ok: (i: number, j: number) => boolean } | CellError {
  const ranges: RangeValue[] = [];
  const tests: ((v: Scalar) => boolean)[] = [];
  for (let k = 0; k + 1 < pairs.length; k += 2) {
    const r = range(pairs[k], ctx);
    if (isErr(r)) return r;
    ranges.push(r);
    tests.push(criterion(one(pairs[k + 1], ctx)));
  }
  if (!ranges.length) return err("#VALUE!");
  const { rows, cols } = ranges[0];
  if (ranges.some((r) => r.rows !== rows || r.cols !== cols)) return err("#VALUE!");
  return { rows, cols, ok: (i, j) => ranges.every((r, k) => tests[k](cellAt(r, i, j, ctx))) };
}

const fns: Record<string, Fn> = {};
const def = (names: string, fn: Fn) => {
  for (const name of names.split(" ")) fns[name] = fn;
};
/** A function of plain numbers: arguments converted, errors passed through. */
const math = (names: string, f: (...n: number[]) => number | CellError, arity?: number) =>
  def(names, (args, ctx) => {
    const ns: number[] = [];
    for (let i = 0; i < (arity ?? args.length); i++) {
      const n = num(args[i], ctx);
      if (isErr(n)) return n;
      ns.push(n);
    }
    const r = f(...ns);
    return typeof r === "number" && !Number.isFinite(r) ? err("#NUM!") : r;
  });
const stat = (names: string, f: (ns: number[]) => number | CellError) =>
  def(names, (args, ctx) => {
    const ns = numbers(args, ctx);
    return isErr(ns) ? ns : f(ns);
  });

const sum = (ns: number[]) => ns.reduce((a, b) => a + b, 0);
const mean = (ns: number[]) => (ns.length ? sum(ns) / ns.length : err("#DIV/0!"));
const variance = (ns: number[], sample: boolean) => {
  if (ns.length < (sample ? 2 : 1)) return err("#DIV/0!");
  const m = sum(ns) / ns.length;
  return sum(ns.map((n) => (n - m) ** 2)) / (ns.length - (sample ? 1 : 0));
};
const roundTo = (n: number, digits: number, mode: "round" | "up" | "down") => {
  const f = 10 ** digits;
  const x = Math.abs(n) * f;
  // A little slack so 2.675 rounds as Excel does (to 2.68).
  const r = mode === "round" ? Math.round(x + 1e-9) : mode === "up" ? Math.ceil(x - 1e-9) : Math.floor(x + 1e-9);
  return (Math.sign(n) * r) / f;
};

// ---- maths and statistics -------------------------------------------------------------------

stat("SUM", sum);
stat("PRODUCT", (ns) => ns.reduce((a, b) => a * b, 1));
stat("AVERAGE", mean);
stat("MIN", (ns) => (ns.length ? Math.min(...ns) : 0));
stat("MAX", (ns) => (ns.length ? Math.max(...ns) : 0));
stat("MEDIAN", (ns) => {
  if (!ns.length) return err("#NUM!");
  const s = [...ns].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
});
stat("MODE MODE.SNGL", (ns) => {
  const counts = new Map<number, number>();
  for (const n of ns) counts.set(n, (counts.get(n) ?? 0) + 1);
  let best: number | null = null;
  for (const [n, c] of counts) if (c > 1 && (best === null || c > counts.get(best)!)) best = n;
  return best ?? err("#N/A");
});
stat("STDEV STDEV.S", (ns) => {
  const v = variance(ns, true);
  return isErr(v) ? v : Math.sqrt(v);
});
stat("STDEVP STDEV.P", (ns) => {
  const v = variance(ns, false);
  return isErr(v) ? v : Math.sqrt(v);
});
stat("VAR VAR.S", (ns) => variance(ns, true));
stat("VARP VAR.P", (ns) => variance(ns, false));
stat("SUMSQ", (ns) => sum(ns.map((n) => n * n)));
def("COUNT", (args, ctx) => values(args, ctx).filter((v) => typeof v === "number").length);
def("COUNTA", (args, ctx) => values(args, ctx).filter((v) => v !== null).length);
def("COUNTBLANK", (args, ctx) => values(args, ctx).filter((v) => v === null || v === "").length);
const nth = (largest: boolean): Fn => (args, ctx) => {
  const ns = numbers([args[0]], ctx);
  const k = num(args[1], ctx);
  if (isErr(ns)) return ns;
  if (isErr(k)) return k;
  const s = [...ns].sort((a, b) => (largest ? b - a : a - b));
  return k >= 1 && k <= s.length ? s[Math.ceil(k) - 1] : err("#NUM!");
};
def("LARGE", nth(true));
def("SMALL", nth(false));
def("RANK RANK.EQ", (args, ctx) => {
  const n = num(args[0], ctx);
  const ns = numbers([args[1]], ctx);
  const ascending = num(args[2], ctx, 0);
  if (isErr(n)) return n;
  if (isErr(ns)) return ns;
  if (isErr(ascending)) return ascending;
  if (!ns.includes(n)) return err("#N/A");
  return 1 + ns.filter((x) => (ascending ? x < n : x > n)).length;
});
def("PERCENTILE PERCENTILE.INC", (args, ctx) => {
  const ns = numbers([args[0]], ctx);
  const k = num(args[1], ctx);
  if (isErr(ns)) return ns;
  if (isErr(k)) return k;
  if (!ns.length || k < 0 || k > 1) return err("#NUM!");
  const s = [...ns].sort((a, b) => a - b);
  const pos = k * (s.length - 1);
  const lo = Math.floor(pos);
  return s[lo] + (s[Math.min(lo + 1, s.length - 1)] - s[lo]) * (pos - lo);
});
def("QUARTILE QUARTILE.INC", (args, ctx) => {
  const q = num(args[1], ctx);
  if (isErr(q)) return q;
  if (q < 0 || q > 4) return err("#NUM!");
  return fns.PERCENTILE([args[0], { t: "num", v: Math.floor(q) / 4 }], ctx);
});
def("SUMPRODUCT", (args, ctx) => {
  // Arrays too, so conditions work as in Excel: SUMPRODUCT((A1:A9="Oslo")*B1:B9).
  const grids = args.map((a) => grid(a, ctx));
  if (!grids.length || grids.some((g) => g.length !== grids[0].length || g[0].length !== grids[0][0].length)) return err("#VALUE!");
  let total = 0;
  for (let i = 0; i < grids[0].length; i++)
    for (let j = 0; j < grids[0][0].length; j++) {
      let p = 1;
      for (const g of grids) {
        const v = g[i][j];
        if (isErr(v)) return v;
        p *= typeof v === "number" ? v : 0;
      }
      total += p;
    }
  return total;
});

math("ABS", Math.abs, 1);
math("SQRT", (n) => (n < 0 ? err("#NUM!") : Math.sqrt(n)), 1);
math("EXP", Math.exp, 1);
math("LN", (n) => (n <= 0 ? err("#NUM!") : Math.log(n)), 1);
math("LOG10", (n) => (n <= 0 ? err("#NUM!") : Math.log10(n)), 1);
def("LOG", (args, ctx) => {
  const n = num(args[0], ctx);
  const base = num(args[1], ctx, 10);
  if (isErr(n)) return n;
  if (isErr(base)) return base;
  return n <= 0 || base <= 0 || base === 1 ? err("#NUM!") : Math.log(n) / Math.log(base);
});
math("POWER", (a, b) => a ** b, 2);
math("MOD", (a, b) => (b === 0 ? err("#DIV/0!") : a - b * Math.floor(a / b)), 2);
math("QUOTIENT", (a, b) => (b === 0 ? err("#DIV/0!") : Math.trunc(a / b)), 2);
math("INT", Math.floor, 1);
math("SIGN", Math.sign, 1);
math("EVEN", (n) => Math.sign(n) * Math.ceil(Math.abs(n) / 2) * 2, 1);
math("ODD", (n) => {
  const up = Math.ceil(Math.abs(n));
  return Math.sign(n || 1) * (up % 2 ? up : up + 1);
}, 1);
math("FACT", (n) => {
  if (n < 0) return err("#NUM!");
  let f = 1;
  for (let i = 2; i <= Math.floor(n); i++) f *= i;
  return f;
}, 1);
def("PI", () => Math.PI);
def("RAND", () => Math.random());
math("RANDBETWEEN", (a, b) => Math.ceil(a) + Math.floor(Math.random() * (Math.floor(b) - Math.ceil(a) + 1)), 2);
const rounding = (mode: "round" | "up" | "down"): Fn => (args, ctx) => {
  const n = num(args[0], ctx);
  const d = num(args[1], ctx, 0);
  if (isErr(n)) return n;
  if (isErr(d)) return d;
  return roundTo(n, Math.trunc(d), mode);
};
def("ROUND", rounding("round"));
def("ROUNDUP", rounding("up"));
def("ROUNDDOWN TRUNC", rounding("down"));
const toMultiple = (up: boolean): Fn => (args, ctx) => {
  const n = num(args[0], ctx);
  const m = num(args[1], ctx, 1);
  if (isErr(n)) return n;
  if (isErr(m)) return m;
  if (m === 0) return 0;
  const k = n / m;
  return (up ? Math.ceil(k - 1e-12) : Math.floor(k + 1e-12)) * m;
};
def("CEILING CEILING.MATH", toMultiple(true));
def("FLOOR FLOOR.MATH", toMultiple(false));

// ---- conditional sums and counts -------------------------------------------------------------

def("COUNTIF", (args, ctx) => {
  const r = range(args[0], ctx);
  if (isErr(r)) return r;
  const test = criterion(one(args[1], ctx));
  let n = 0;
  for (const v of cellsOf(r, ctx)) if (test(v)) n++;
  return n;
});
def("COUNTIFS", (args, ctx) => {
  const m = matches(args, ctx);
  if (isErr(m)) return m;
  let n = 0;
  for (let i = 0; i < m.rows; i++) for (let j = 0; j < m.cols; j++) if (m.ok(i, j)) n++;
  return n;
});
/** SUMIF and AVERAGEIF: the criterion's range, and the range to add up (the same if not given). */
const ifAggregate = (average: boolean): Fn => (args, ctx) => {
  const r = range(args[0], ctx);
  if (isErr(r)) return r;
  const target = missing(args[2]) ? r : range(args[2], ctx);
  if (isErr(target)) return target;
  const test = criterion(one(args[1], ctx));
  const ns: number[] = [];
  for (let i = 0; i < r.rows; i++)
    for (let j = 0; j < r.cols; j++) {
      if (!test(cellAt(r, i, j, ctx))) continue;
      const v = cellAt(target, i, j, ctx);
      if (isErr(v)) return v;
      if (typeof v === "number") ns.push(v);
    }
  return average ? mean(ns) : sum(ns);
};
def("SUMIF", ifAggregate(false));
def("AVERAGEIF", ifAggregate(true));
/** SUMIFS and friends: the range first, then pairs of range and criterion. */
const ifsAggregate = (f: (ns: number[]) => number | CellError): Fn => (args, ctx) => {
  const target = range(args[0], ctx);
  if (isErr(target)) return target;
  const m = matches(args.slice(1), ctx);
  if (isErr(m)) return m;
  if (m.rows !== target.rows || m.cols !== target.cols) return err("#VALUE!");
  const ns: number[] = [];
  for (let i = 0; i < m.rows; i++)
    for (let j = 0; j < m.cols; j++) {
      if (!m.ok(i, j)) continue;
      const v = cellAt(target, i, j, ctx);
      if (isErr(v)) return v;
      if (typeof v === "number") ns.push(v);
    }
  return f(ns);
};
def("SUMIFS", ifsAggregate(sum));
def("AVERAGEIFS", ifsAggregate(mean));
def("MAXIFS", ifsAggregate((ns) => (ns.length ? Math.max(...ns) : 0)));
def("MINIFS", ifsAggregate((ns) => (ns.length ? Math.min(...ns) : 0)));

// ---- logic --------------------------------------------------------------------------------------

def("IF", (args, ctx) => {
  const test = toBool(one(args[0], ctx));
  if (isErr(test)) return test;
  if (test) return missing(args[1]) ? 0 : ev(args[1], ctx);
  return args.length < 3 ? false : missing(args[2]) ? 0 : ev(args[2], ctx);
});
def("IFS", (args, ctx) => {
  for (let i = 0; i + 1 < args.length; i += 2) {
    const test = toBool(one(args[i], ctx));
    if (isErr(test)) return test;
    if (test) return ev(args[i + 1], ctx);
  }
  return err("#N/A");
});
def("IFERROR", (args, ctx) => {
  const v = one(args[0], ctx);
  return isErr(v) ? one(args[1], ctx) : v;
});
def("IFNA", (args, ctx) => {
  const v = one(args[0], ctx);
  return isErr(v) && v.code === "#N/A" ? one(args[1], ctx) : v;
});
const logic = (f: (bs: boolean[]) => boolean): Fn => (args, ctx) => {
  const bs: boolean[] = [];
  for (const v of values(args, ctx)) {
    if (v === null || typeof v === "string") continue;
    const b = toBool(v);
    if (isErr(b)) return b;
    bs.push(b);
  }
  return bs.length ? f(bs) : err("#VALUE!");
};
def("AND", logic((bs) => bs.every(Boolean)));
def("OR", logic((bs) => bs.some(Boolean)));
def("XOR", logic((bs) => bs.filter(Boolean).length % 2 === 1));
def("NOT", (args, ctx) => {
  const b = toBool(one(args[0], ctx));
  return isErr(b) ? b : !b;
});
def("TRUE", () => true);
def("FALSE", () => false);
def("SWITCH", (args, ctx) => {
  const value = one(args[0], ctx);
  if (isErr(value)) return value;
  let i = 1;
  for (; i + 1 < args.length; i += 2) if (compare(value, one(args[i], ctx)) === 0) return ev(args[i + 1], ctx);
  return i < args.length ? ev(args[i], ctx) : err("#N/A");
});

// ---- lookups --------------------------------------------------------------------------------------

/** Where a value is in a list: exact (with wildcards for text), or the largest not above it in a sorted list. */
function find(list: Scalar[], value: Scalar, mode: "exact" | "below" | "above"): number {
  if (mode === "exact") {
    const re = typeof value === "string" && /[*?~]/.test(value) ? wildcard(value) : null;
    return list.findIndex((v) => (re ? typeof v === "string" && re.test(v) : v !== null && compare(v, value) === 0));
  }
  // Sorted ascending (below) or descending (above): a binary search, as Excel does.
  let lo = 0;
  let hi = list.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = list[mid];
    const c = v === null ? (mode === "below" ? -1 : 1) : compare(v, value);
    if (c === 0) found = mid;
    if (mode === "below" ? c <= 0 : c >= 0) {
      if (c !== 0 || found === -1) found = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return found;
}

const column = (r: RangeValue, j: number, ctx: EvalContext) => Array.from({ length: r.rows }, (_, i) => cellAt(r, i, j, ctx));
const row = (r: RangeValue, i: number, ctx: EvalContext) => Array.from({ length: r.cols }, (_, j) => cellAt(r, i, j, ctx));

const vhlookup = (vertical: boolean): Fn => (args, ctx) => {
  const value = one(args[0], ctx);
  const table = range(args[1], ctx);
  const index = num(args[2], ctx);
  const approx = missing(args[3]) ? true : toBool(one(args[3], ctx));
  if (isErr(value)) return value;
  if (isErr(table)) return table;
  if (isErr(index)) return index;
  if (isErr(approx)) return approx;
  const n = Math.trunc(index);
  if (n < 1) return err("#VALUE!");
  if (n > (vertical ? table.cols : table.rows)) return err("#REF!");
  const keys = vertical ? column(table, 0, ctx) : row(table, 0, ctx);
  const at = find(keys, value, approx ? "below" : "exact");
  if (at < 0) return err("#N/A");
  return vertical ? cellAt(table, at, n - 1, ctx) : cellAt(table, n - 1, at, ctx);
};
def("VLOOKUP", vhlookup(true));
def("HLOOKUP", vhlookup(false));
def("MATCH", (args, ctx) => {
  const value = one(args[0], ctx);
  const r = range(args[1], ctx);
  const type = num(args[2], ctx, 1);
  if (isErr(value)) return value;
  if (isErr(r)) return r;
  if (isErr(type)) return type;
  const list = r.cols === 1 ? column(r, 0, ctx) : row(r, 0, ctx);
  const at = find(list, value, type === 0 ? "exact" : type > 0 ? "below" : "above");
  return at < 0 ? err("#N/A") : at + 1;
});
def("XMATCH", (args, ctx) => {
  const value = one(args[0], ctx);
  const r = range(args[1], ctx);
  if (isErr(value)) return value;
  if (isErr(r)) return r;
  const list = r.cols === 1 ? column(r, 0, ctx) : row(r, 0, ctx);
  const at = find(list, value, "exact");
  return at < 0 ? err("#N/A") : at + 1;
});
def("XLOOKUP", (args, ctx) => {
  const value = one(args[0], ctx);
  const keys = range(args[1], ctx);
  const results = range(args[2], ctx);
  const mode = num(args[4], ctx, 0);
  const search = num(args[5], ctx, 1);
  if (isErr(value)) return value;
  if (isErr(keys)) return keys;
  if (isErr(results)) return results;
  if (isErr(mode)) return mode;
  if (isErr(search)) return search;
  const vertical = keys.cols === 1;
  let list = vertical ? column(keys, 0, ctx) : row(keys, 0, ctx);
  const reversed = search < 0;
  if (reversed) list = [...list].reverse();
  let at: number;
  if (mode === 0 || mode === 2) at = find(list, mode === 2 ? value : typeof value === "string" ? value.replace(/[*?~]/g, "~$&") : value, "exact");
  else {
    // -1: the next smaller, 1: the next larger, when there is no exact match.
    at = find(list, value, "exact");
    if (at < 0) {
      let best = -1;
      list.forEach((v, i) => {
        if (v === null || typeof v !== typeof value) return;
        const c = compare(v, value);
        if ((mode < 0 && c < 0 && (best < 0 || compare(v, list[best]) > 0)) || (mode > 0 && c > 0 && (best < 0 || compare(v, list[best]) < 0))) best = i;
      });
      at = best;
    }
  }
  if (at >= 0 && reversed) at = list.length - 1 - at;
  if (at < 0) return missing(args[3]) ? err("#N/A") : ev(args[3], ctx);
  return vertical ? cellAt(results, at, 0, ctx) : cellAt(results, 0, at, ctx);
});
def("INDEX", (args, ctx) => {
  const r = range(args[0], ctx);
  if (isErr(r)) return r;
  let i = num(args[1], ctx, 0);
  let j = num(args[2], ctx, 0);
  if (isErr(i)) return i;
  if (isErr(j)) return j;
  // INDEX(a single row, n) counts along the row.
  if (r.rows === 1 && missing(args[2]) && args.length < 3) [i, j] = [1, i];
  if (i < 0 || j < 0 || i > r.rows || j > r.cols) return err("#REF!");
  if (i === 0 && j === 0) return r;
  if (i === 0) return new RangeValue(r.sheet, r.r1, r.c1 + j - 1, r.r2, r.c1 + j - 1);
  if (j === 0) return r.cols === 1 ? new RangeValue(r.sheet, r.r1 + i - 1, r.c1, r.r1 + i - 1, r.c1) : new RangeValue(r.sheet, r.r1 + i - 1, r.c1, r.r1 + i - 1, r.c2);
  return new RangeValue(r.sheet, r.r1 + i - 1, r.c1 + j - 1, r.r1 + i - 1, r.c1 + j - 1);
});
def("CHOOSE", (args, ctx) => {
  const i = num(args[0], ctx);
  if (isErr(i)) return i;
  const k = Math.trunc(i);
  return k >= 1 && k < args.length ? ev(args[k], ctx) : err("#VALUE!");
});
def("ROW", (args, ctx) => {
  if (missing(args[0])) return ctx.row + 1;
  const r = range(args[0], ctx);
  return isErr(r) ? r : r.r1 + 1;
});
def("COLUMN", (args, ctx) => {
  if (missing(args[0])) return ctx.col + 1;
  const r = range(args[0], ctx);
  return isErr(r) ? r : r.c1 + 1;
});
def("ROWS", (args, ctx) => {
  const v = ev(args[0], ctx);
  return v instanceof RangeValue || v instanceof ArrayValue ? v.rows : isErr(v) ? v : 1;
});
def("COLUMNS", (args, ctx) => {
  const v = ev(args[0], ctx);
  return v instanceof RangeValue || v instanceof ArrayValue ? v.cols : isErr(v) ? v : 1;
});
def("OFFSET", (args, ctx) => {
  const r = range(args[0], ctx);
  if (isErr(r)) return r;
  const dr = num(args[1], ctx, 0);
  const dc = num(args[2], ctx, 0);
  const h = num(args[3], ctx, r.rows);
  const w = num(args[4], ctx, r.cols);
  for (const n of [dr, dc, h, w]) if (isErr(n)) return n;
  const [a, b, hh, ww] = [dr, dc, h, w].map((n) => Math.trunc(n as number));
  if (r.r1 + a < 0 || r.c1 + b < 0 || hh < 1 || ww < 1) return err("#REF!");
  return new RangeValue(r.sheet, r.r1 + a, r.c1 + b, r.r1 + a + hh - 1, r.c1 + b + ww - 1);
});
def("INDIRECT", (args, ctx) => {
  const text = toText(one(args[0], ctx));
  try {
    const node = parse(text);
    if (node.t !== "ref" && node.t !== "range") return err("#REF!");
    return evaluate(node, ctx);
  } catch {
    return err("#REF!");
  }
});

// ---- text ---------------------------------------------------------------------------------------

const text = (names: string, f: (...args: Scalar[]) => Scalar, arity?: number) =>
  def(names, (args, ctx) => {
    const vs = (arity === undefined ? args : args.slice(0, arity)).map((a) => one(a, ctx));
    const bad = vs.find(isErr);
    return bad ?? f(...vs);
  });
const join = (vs: Scalar[]) => {
  const bad = vs.find(isErr);
  return bad ?? vs.map(toText).join("");
};
def("CONCAT", (args, ctx) => join(values(args, ctx)));
def("CONCATENATE", (args, ctx) => join(args.map((a) => one(a, ctx))));
def("TEXTJOIN", (args, ctx) => {
  const delimiter = toText(one(args[0], ctx));
  const skip = toBool(one(args[1], ctx));
  if (isErr(skip)) return skip;
  const vs = values(args.slice(2), ctx);
  const bad = vs.find(isErr);
  if (bad) return bad;
  return vs.filter((v) => !skip || (v !== null && v !== "")).map(toText).join(delimiter);
});
const count = (v: Scalar, fallback: number) => (v === null ? fallback : toNumber(v));
text("LEFT", (s, n) => {
  const k = count(n ?? null, 1);
  return isErr(k) ? k : k < 0 ? err("#VALUE!") : toText(s).slice(0, Math.trunc(k));
});
text("RIGHT", (s, n) => {
  const k = count(n ?? null, 1);
  if (isErr(k)) return k;
  if (k < 0) return err("#VALUE!");
  const str = toText(s);
  return Math.trunc(k) === 0 ? "" : str.slice(-Math.trunc(k));
});
text("MID", (s, start, n) => {
  const a = toNumber(start);
  const k = toNumber(n);
  if (isErr(a)) return a;
  if (isErr(k)) return k;
  if (a < 1 || k < 0) return err("#VALUE!");
  return toText(s).substr(Math.trunc(a) - 1, Math.trunc(k));
});
text("LEN", (s) => toText(s).length, 1);
text("UPPER", (s) => toText(s).toUpperCase(), 1);
text("LOWER", (s) => toText(s).toLowerCase(), 1);
text("PROPER", (s) => toText(s).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a, b) => a + b.toUpperCase()), 1);
text("TRIM", (s) => toText(s).replace(/ +/g, " ").trim(), 1);
text("CLEAN", (s) => toText(s).replace(/[\x00-\x1f]/g, ""), 1);
text("REPT", (s, n) => {
  const k = toNumber(n);
  return isErr(k) ? k : k < 0 ? err("#VALUE!") : toText(s).repeat(Math.trunc(k));
});
text("EXACT", (a, b) => toText(a) === toText(b));
text("CHAR", (n) => {
  const k = toNumber(n);
  return isErr(k) ? k : k < 1 || k > 255 ? err("#VALUE!") : String.fromCharCode(Math.trunc(k));
}, 1);
text("CODE", (s) => {
  const str = toText(s);
  return str ? str.charCodeAt(0) : err("#VALUE!");
}, 1);
text("T", (v) => (typeof v === "string" ? v : ""), 1);
text("N", (v) => (typeof v === "number" ? v : typeof v === "boolean" ? Number(v) : 0), 1);
text("SUBSTITUTE", (s, from, to, nth) => {
  const str = toText(s);
  const a = toText(from);
  const b = toText(to);
  if (!a) return str;
  if (nth === undefined || nth === null) return str.split(a).join(b);
  const k = toNumber(nth);
  if (isErr(k)) return k;
  let at = -1;
  for (let i = 0; i < k; i++) {
    at = str.indexOf(a, at + 1);
    if (at < 0) return str;
  }
  return str.slice(0, at) + b + str.slice(at + a.length);
});
text("REPLACE", (s, start, n, to) => {
  const a = toNumber(start);
  const k = toNumber(n);
  if (isErr(a)) return a;
  if (isErr(k)) return k;
  const str = toText(s);
  return str.slice(0, a - 1) + toText(to) + str.slice(a - 1 + k);
});
const findText = (caseless: boolean) => (needle: Scalar, hay: Scalar, start?: Scalar) => {
  const from = start === undefined || start === null ? 1 : toNumber(start);
  if (isErr(from)) return from;
  const h = toText(hay);
  const n = toText(needle);
  if (from < 1 || from > h.length + 1) return err("#VALUE!");
  const rest = h.slice(from - 1);
  // SEARCH ignores case and knows wildcards; FIND does neither.
  const at = caseless ? rest.search(new RegExp(wildcard(n).source.slice(1, -1), "i")) : rest.indexOf(n);
  return at < 0 ? err("#VALUE!") : at + from;
};
text("FIND", findText(false));
text("SEARCH", findText(true));
text("TEXT", (v, fmt) => {
  const code = toText(fmt);
  if (typeof v === "string") {
    const typed = parseInput(v).value;
    return formatValue(typeof typed === "number" ? typed : v, code).text;
  }
  return formatValue(v, code).text;
}, 2);
text("VALUE NUMBERVALUE", (s) => {
  if (typeof s === "number") return s;
  const typed = parseInput(toText(s));
  return typeof typed.value === "number" ? typed.value : err("#VALUE!");
}, 1);

// ---- information --------------------------------------------------------------------------------

text("ISBLANK", (v) => v === null, 1);
text("ISNUMBER", (v) => typeof v === "number", 1);
text("ISTEXT", (v) => typeof v === "string", 1);
text("ISNONTEXT", (v) => typeof v !== "string", 1);
text("ISLOGICAL", (v) => typeof v === "boolean", 1);
def("ISERROR", (args, ctx) => isErr(one(args[0], ctx)));
def("ISERR", (args, ctx) => {
  const v = one(args[0], ctx);
  return isErr(v) && v.code !== "#N/A";
});
def("ISNA", (args, ctx) => {
  const v = one(args[0], ctx);
  return isErr(v) && v.code === "#N/A";
});
const parity = (even: boolean) => (v: Scalar) => {
  const n = toNumber(v);
  return isErr(n) ? n : (Math.trunc(n) % 2 === 0) === even;
};
text("ISEVEN", parity(true), 1);
text("ISODD", parity(false), 1);
def("NA", () => err("#N/A"));

// ---- dates and times ------------------------------------------------------------------------------

const today = (ctx: EvalContext) => {
  const now = new Date();
  return dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate(), ctx.date1904);
};
def("TODAY", (_args, ctx) => today(ctx));
def("NOW", (_args, ctx) => {
  const now = new Date();
  return today(ctx) + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
});
def("DATE", (args, ctx) => {
  const [y, m, d] = [0, 1, 2].map((i) => num(args[i], ctx));
  for (const n of [y, m, d]) if (isErr(n)) return n;
  let year = Math.trunc(y as number);
  if (year < 1900) year += 1900;
  // Months and days past the end roll over, as in Excel: DATE(2026,13,1) is January 2027.
  const date = new Date(Date.UTC(year, Math.trunc(m as number) - 1, 1));
  return dateToSerial(date.getUTCFullYear(), date.getUTCMonth() + 1, 1, ctx.date1904) + Math.trunc(d as number) - 1;
});
math("TIME", (h, m, s) => ((h * 3600 + m * 60 + s) / 86400) % 1, 3);
const part = (f: (d: Date) => number): Fn => (args, ctx) => {
  const n = num(args[0], ctx);
  if (isErr(n)) return n;
  return n < 0 ? err("#NUM!") : f(serialToDate(n, ctx.date1904));
};
def("YEAR", part((d) => d.getUTCFullYear()));
def("MONTH", part((d) => d.getUTCMonth() + 1));
def("DAY", part((d) => d.getUTCDate()));
const secondsOf = (n: number) => Math.round((n % 1) * 86400);
math("HOUR", (n) => Math.floor(secondsOf(n) / 3600) % 24, 1);
math("MINUTE", (n) => Math.floor(secondsOf(n) / 60) % 60, 1);
math("SECOND", (n) => secondsOf(n) % 60, 1);
def("WEEKDAY", (args, ctx) => {
  const n = num(args[0], ctx);
  const type = num(args[1], ctx, 1);
  if (isErr(n)) return n;
  if (isErr(type)) return type;
  const day = serialToDate(n, ctx.date1904).getUTCDay(); // 0 = Sunday
  if (type === 2) return day === 0 ? 7 : day;
  if (type === 3) return day === 0 ? 6 : day - 1;
  return day + 1;
});
const isoWeek = (d: Date) => {
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  t.setUTCDate(t.getUTCDate() + 4 - (t.getUTCDay() || 7));
  const yearStart = Date.UTC(t.getUTCFullYear(), 0, 1);
  return Math.ceil(((t.getTime() - yearStart) / 86400000 + 1) / 7);
};
def("ISOWEEKNUM", part(isoWeek));
def("WEEKNUM", (args, ctx) => {
  const n = num(args[0], ctx);
  const type = num(args[1], ctx, 1);
  if (isErr(n)) return n;
  if (isErr(type)) return type;
  const d = serialToDate(n, ctx.date1904);
  if (type === 21) return isoWeek(d);
  const start = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const offset = type === 2 ? (start.getUTCDay() + 6) % 7 : start.getUTCDay();
  return Math.floor(((d.getTime() - start.getTime()) / 86400000 + offset) / 7) + 1;
});
const addMonths = (end: boolean): Fn => (args, ctx) => {
  const n = num(args[0], ctx);
  const months = num(args[1], ctx);
  if (isErr(n)) return n;
  if (isErr(months)) return months;
  const d = serialToDate(n, ctx.date1904);
  const target = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + Math.trunc(months) + (end ? 1 : 0), end ? 0 : 1));
  const day = end ? target.getUTCDate() : Math.min(d.getUTCDate(), new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate());
  return dateToSerial(target.getUTCFullYear(), target.getUTCMonth() + 1, day, ctx.date1904);
};
def("EDATE", addMonths(false));
def("EOMONTH", addMonths(true));
math("DAYS", (end, start) => Math.trunc(end) - Math.trunc(start), 2);
def("DATEDIF", (args, ctx) => {
  const a = num(args[0], ctx);
  const b = num(args[1], ctx);
  const unit = toText(one(args[2], ctx)).toUpperCase();
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (a > b) return err("#NUM!");
  const x = serialToDate(a, ctx.date1904);
  const y = serialToDate(b, ctx.date1904);
  let months = (y.getUTCFullYear() - x.getUTCFullYear()) * 12 + y.getUTCMonth() - x.getUTCMonth();
  if (y.getUTCDate() < x.getUTCDate()) months--;
  if (unit === "D") return Math.trunc(b) - Math.trunc(a);
  if (unit === "M") return months;
  if (unit === "Y") return Math.floor(months / 12);
  if (unit === "YM") return months % 12;
  if (unit === "MD") return (y.getUTCDate() - x.getUTCDate() + 31) % 31;
  return err("#NUM!");
});
const workdays = (args: Node[], ctx: EvalContext, from: number) => {
  const holidays = new Set<number>();
  if (!missing(args[from])) for (const v of values([args[from]], ctx)) if (typeof v === "number") holidays.add(Math.trunc(v));
  return (serial: number) => {
    const day = serialToDate(serial, ctx.date1904).getUTCDay();
    return day !== 0 && day !== 6 && !holidays.has(serial);
  };
};
def("NETWORKDAYS", (args, ctx) => {
  const a = num(args[0], ctx);
  const b = num(args[1], ctx);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const works = workdays(args, ctx, 2);
  const [lo, hi] = a <= b ? [Math.trunc(a), Math.trunc(b)] : [Math.trunc(b), Math.trunc(a)];
  let n = 0;
  for (let s = lo; s <= hi; s++) if (works(s)) n++;
  return a <= b ? n : -n;
});
def("WORKDAY", (args, ctx) => {
  const a = num(args[0], ctx);
  const days = num(args[1], ctx);
  if (isErr(a)) return a;
  if (isErr(days)) return days;
  const works = workdays(args, ctx, 2);
  let s = Math.trunc(a);
  const step = days < 0 ? -1 : 1;
  for (let left = Math.abs(Math.trunc(days)); left > 0; ) {
    s += step;
    if (works(s)) left--;
  }
  return s;
});
text("DATEVALUE", (s) => {
  const typed = parseInput(toText(s));
  return typeof typed.value === "number" && typed.format ? Math.floor(typed.value) : err("#VALUE!");
}, 1);
text("TIMEVALUE", (s) => {
  const typed = parseInput(toText(s));
  return typeof typed.value === "number" && typed.format ? typed.value % 1 : err("#VALUE!");
}, 1);

// ---- money ------------------------------------------------------------------------------------------

/** The loan or saving functions share one equation; type 1 means paying at the start of each period. */
def("PMT", (args, ctx) => {
  const [rate, n, pv, fv, type] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx), num(args[3], ctx, 0), num(args[4], ctx, 0)];
  for (const x of [rate, n, pv, fv, type]) if (isErr(x)) return x;
  const [r, k, p, f, t] = [rate, n, pv, fv, type] as number[];
  if (r === 0) return -(p + f) / k;
  const g = (1 + r) ** k;
  return (-(p * g + f) * r) / ((1 + r * t) * (g - 1));
});
def("FV", (args, ctx) => {
  const [rate, n, pmt, pv, type] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx), num(args[3], ctx, 0), num(args[4], ctx, 0)];
  for (const x of [rate, n, pmt, pv, type]) if (isErr(x)) return x;
  const [r, k, m, p, t] = [rate, n, pmt, pv, type] as number[];
  if (r === 0) return -(p + m * k);
  const g = (1 + r) ** k;
  return -(p * g + (m * (1 + r * t) * (g - 1)) / r);
});
def("PV", (args, ctx) => {
  const [rate, n, pmt, fv, type] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx), num(args[3], ctx, 0), num(args[4], ctx, 0)];
  for (const x of [rate, n, pmt, fv, type]) if (isErr(x)) return x;
  const [r, k, m, f, t] = [rate, n, pmt, fv, type] as number[];
  if (r === 0) return -(f + m * k);
  const g = (1 + r) ** k;
  return -(f + (m * (1 + r * t) * (g - 1)) / r) / g;
});
def("NPV", (args, ctx) => {
  const rate = num(args[0], ctx);
  if (isErr(rate)) return rate;
  const flows = numbers(args.slice(1), ctx);
  if (isErr(flows)) return flows;
  return flows.reduce((total, v, i) => total + v / (1 + rate) ** (i + 1), 0);
});

// ---- dynamic arrays: results that spill into the cells below and to the right -----------------

/** An argument as rows of values, with comparisons like A1:A9>5 worked out for every cell. */
const grid = (node: Node | undefined, ctx: EvalContext): Scalar[][] => toGrid(ev(node, { ...ctx, arrays: true }), ctx);
const truthy = (v: Scalar) => {
  const b = toBool(v);
  return b === true;
};

def("FILTER", (args, ctx) => {
  const data = grid(args[0], ctx);
  const keep = grid(args[1], ctx);
  const errIn = keep.flat().find(isErr);
  if (errIn) return errIn;
  let rows: Scalar[][];
  if (keep.length === data.length && keep[0].length === 1) rows = data.filter((_, i) => truthy(keep[i][0]));
  else if (keep.length === 1 && keep[0].length === data[0].length) rows = data.map((row) => row.filter((_, j) => truthy(keep[0][j]))).filter((row) => row.length);
  else return err("#VALUE!");
  if (!rows.length || !rows[0].length) return args.length > 2 ? ev(args[2], ctx) : err("#CALC!");
  return new ArrayValue(rows);
});

/** Rows (or columns) sorted by one or more keys, each up (1) or down (-1); the original order kept where keys are equal. */
function sortRows(rows: Scalar[][], keys: { values: Scalar[]; order: number }[]): Scalar[][] {
  const index = rows.map((_, i) => i);
  index.sort((a, b) => {
    for (const k of keys) {
      const c = compare(k.values[a], k.values[b]);
      if (c) return c * k.order;
    }
    return a - b;
  });
  return index.map((i) => rows[i]);
}

const transpose = (g: Scalar[][]) => (g.length ? g[0].map((_, j) => g.map((row) => row[j])) : g);

def("SORT", (args, ctx) => {
  let data = grid(args[0], ctx);
  const byCol = !missing(args[3]) && toBool(one(args[3], ctx)) === true;
  if (byCol) data = transpose(data);
  const at = num(args[1], ctx, 1);
  const order = num(args[2], ctx, 1);
  if (isErr(at)) return at;
  if (isErr(order)) return order;
  if (at < 1 || at > data[0].length || (order !== 1 && order !== -1)) return err("#VALUE!");
  const sorted = sortRows(data, [{ values: data.map((row) => row[at - 1]), order }]);
  return new ArrayValue(byCol ? transpose(sorted) : sorted);
});

def("SORTBY", (args, ctx) => {
  const data = grid(args[0], ctx);
  const keys: { values: Scalar[]; order: number }[] = [];
  for (let i = 1; i < args.length; i += 2) {
    const by = grid(args[i], ctx);
    const order = num(args[i + 1], ctx, 1);
    if (isErr(order)) return order;
    const values = by.length === data.length && by[0].length === 1 ? by.map((row) => row[0]) : null;
    if (!values || (order !== 1 && order !== -1)) return err("#VALUE!");
    keys.push({ values, order });
  }
  return new ArrayValue(sortRows(data, keys));
});

def("UNIQUE", (args, ctx) => {
  let data = grid(args[0], ctx);
  const byCol = !missing(args[1]) && toBool(one(args[1], ctx)) === true;
  const once = !missing(args[2]) && toBool(one(args[2], ctx)) === true;
  if (byCol) data = transpose(data);
  const id = (row: Scalar[]) => JSON.stringify(row.map((v) => (typeof v === "string" ? v.toLowerCase() : isErr(v) ? v.code : v)));
  const counts = new Map<string, number>();
  for (const row of data) counts.set(id(row), (counts.get(id(row)) ?? 0) + 1);
  const seen = new Set<string>();
  const rows = data.filter((row) => {
    const k = id(row);
    if (seen.has(k) || (once && counts.get(k)! > 1)) return false;
    seen.add(k);
    return true;
  });
  if (!rows.length) return err("#CALC!");
  return new ArrayValue(byCol ? transpose(rows) : rows);
});

def("SEQUENCE", (args, ctx) => {
  const rows = num(args[0], ctx);
  const cols = num(args[1], ctx, 1);
  const start = num(args[2], ctx, 1);
  const step = num(args[3], ctx, 1);
  for (const v of [rows, cols, start, step]) if (isErr(v)) return v;
  const [r, c] = [Math.trunc(rows as number), Math.trunc(cols as number)];
  if (r < 1 || c < 1) return err("#CALC!");
  if (r * c > 1_000_000) return err("#NUM!");
  return new ArrayValue(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => (start as number) + (i * c + j) * (step as number))));
});

/** HYPERLINK shows its name (or the address); clicking the cell opens the address. */
def("HYPERLINK", (args, ctx) => {
  const link = one(args[0], ctx);
  if (isErr(link)) return link;
  return missing(args[1]) ? toText(link) : one(args[1], ctx);
});

def("TRANSPOSE", (args, ctx) => new ArrayValue(transpose(grid(args[0], ctx))));

// ---- maths and trigonometry -------------------------------------------------------------------

const finite = (n: number) => (Number.isFinite(n) ? n : err("#NUM!"));
math("SIN", (x) => Math.sin(x), 1);
math("COS", (x) => Math.cos(x), 1);
math("TAN", (x) => Math.tan(x), 1);
math("ASIN", (x) => (Math.abs(x) > 1 ? err("#NUM!") : Math.asin(x)), 1);
math("ACOS", (x) => (Math.abs(x) > 1 ? err("#NUM!") : Math.acos(x)), 1);
math("ATAN", (x) => Math.atan(x), 1);
math("ATAN2", (x, y) => (x === 0 && y === 0 ? err("#DIV/0!") : Math.atan2(y, x)), 2);
math("SINH", (x) => finite(Math.sinh(x)), 1);
math("COSH", (x) => finite(Math.cosh(x)), 1);
math("TANH", (x) => Math.tanh(x), 1);
math("DEGREES", (x) => (x * 180) / Math.PI, 1);
math("RADIANS", (x) => (x * Math.PI) / 180, 1);
math("SQRTPI", (x) => (x < 0 ? err("#NUM!") : Math.sqrt(x * Math.PI)), 1);
math("MROUND", (x, m) => (m === 0 ? 0 : x * m < 0 ? err("#NUM!") : Math.round(x / m) * m), 2);
const gcd = (a: number, b: number): number => (b ? gcd(b, a % b) : a);
def("GCD", (args, ctx) => {
  const ns = numbers(args, ctx);
  if (isErr(ns)) return ns;
  if (ns.some((n) => n < 0)) return err("#NUM!");
  return ns.map(Math.trunc).reduce(gcd, 0);
});
def("LCM", (args, ctx) => {
  const ns = numbers(args, ctx);
  if (isErr(ns)) return ns;
  if (ns.some((n) => n < 0)) return err("#NUM!");
  return ns.map(Math.trunc).reduce((a, b) => (a && b ? (a / gcd(a, b)) * b : 0), 1);
});
const factorial = (n: number) => {
  let f = 1;
  for (let i = 2; i <= n; i++) f *= i;
  return f;
};
math("COMBIN", (n, k) => (n < 0 || k < 0 || k > n ? err("#NUM!") : Math.round(factorial(Math.trunc(n)) / (factorial(Math.trunc(k)) * factorial(Math.trunc(n) - Math.trunc(k))))), 2);
math("PERMUT", (n, k) => (n < 0 || k < 0 || k > n ? err("#NUM!") : Math.round(factorial(Math.trunc(n)) / factorial(Math.trunc(n) - Math.trunc(k)))), 2);

// ---- number bases, for programmers ----------------------------------------------------------------

/** A number in another base, with at least `places` digits; negative numbers as Excel writes them (two's complement in 10 digits). */
function toBase(n: number, base: number, bits: number, places?: number): string | CellError {
  const limit = 2 ** (bits - 1);
  if (!Number.isInteger(n) || n < -limit || n >= limit) return err("#NUM!");
  const text = (n < 0 ? 2 ** bits + n : n).toString(base).toUpperCase();
  if (places === undefined || n < 0) return text;
  return places < text.length ? err("#NUM!") : text.padStart(places, "0");
}
function fromBase(text: string, base: number, bits: number): number | CellError {
  if (!text || text.length > 10 || !new RegExp(`^[0-${Math.min(base, 10) - 1}${base > 10 ? "A-Fa-f" : ""}]+$`).test(text)) return err("#NUM!");
  const n = parseInt(text, base);
  // Ten digits with the top bit set are negative, as in Excel.
  return text.length === 10 && n >= 2 ** (bits - 1) ? n - 2 ** bits : n;
}
const BASES: [string, number, number][] = [
  ["BIN", 2, 10],
  ["OCT", 8, 30],
  ["HEX", 16, 40],
];
for (const [from, fromRadix, fromBits] of BASES) {
  def(`${from}2DEC`, (args, ctx) => fromBase(toText(one(args[0], ctx)).trim(), fromRadix, fromBits));
  def(`DEC2${from}`, (args, ctx) => {
    const n = num(args[0], ctx);
    const places = missing(args[1]) ? undefined : num(args[1], ctx);
    if (isErr(n)) return n;
    if (isErr(places as number | CellError)) return places as CellError;
    return toBase(Math.trunc(n), fromRadix, fromBits, places as number | undefined);
  });
  for (const [to, toRadix, toBits] of BASES) {
    if (to === from) continue;
    def(`${from}2${to}`, (args, ctx) => {
      const n = fromBase(toText(one(args[0], ctx)).trim(), fromRadix, fromBits);
      const places = missing(args[1]) ? undefined : num(args[1], ctx);
      if (isErr(n)) return n;
      if (isErr(places as number | CellError)) return places as CellError;
      return toBase(n, toRadix, toBits, places as number | undefined);
    });
  }
}
def("BASE", (args, ctx) => {
  const n = num(args[0], ctx);
  const radix = num(args[1], ctx);
  const length = num(args[2], ctx, 0);
  for (const v of [n, radix, length]) if (isErr(v)) return v;
  if ((n as number) < 0 || (radix as number) < 2 || (radix as number) > 36) return err("#NUM!");
  return Math.trunc(n as number).toString(radix as number).toUpperCase().padStart(length as number, "0");
});
def("DECIMAL", (args, ctx) => {
  const text = toText(one(args[0], ctx)).trim();
  const radix = num(args[1], ctx);
  if (isErr(radix)) return radix;
  const n = parseInt(text, radix);
  return radix < 2 || radix > 36 || Number.isNaN(n) || n.toString(radix).toUpperCase() !== text.replace(/^0+(?=.)/, "").toUpperCase() ? err("#NUM!") : n;
});
const bits = (f: (a: bigint, b: bigint) => bigint) => (args: Node[], ctx: EvalContext) => {
  const a = num(args[0], ctx);
  const b = num(args[1], ctx);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (a < 0 || !Number.isInteger(a) || a >= 2 ** 48) return err("#NUM!");
  return Number(f(BigInt(a), BigInt(Math.trunc(b))));
};
def("BITAND", bits((a, b) => a & b));
def("BITOR", bits((a, b) => a | b));
def("BITXOR", bits((a, b) => a ^ b));
def("BITLSHIFT", bits((a, s) => (s >= 0n ? a << s : a >> -s)));
def("BITRSHIFT", bits((a, s) => (s >= 0n ? a >> s : a << -s)));

// ---- statistics ---------------------------------------------------------------------------------

/** Two ranges' numbers, in pairs where both are numbers, as Excel's regression functions take them. */
function pairs(args: Node[], ctx: EvalContext): [number[], number[]] | CellError {
  const ys = values([args[0]], ctx);
  const xs = values([args[1]], ctx);
  if (ys.length !== xs.length) return err("#N/A");
  const y: number[] = [];
  const x: number[] = [];
  ys.forEach((v, i) => {
    if (typeof v === "number" && typeof xs[i] === "number") {
      y.push(v);
      x.push(xs[i] as number);
    }
  });
  return [y, x];
}
const avg = (ns: number[]) => ns.reduce((s, n) => s + n, 0) / ns.length;
/** Slope and intercept of the least-squares line through the points. */
function line(args: Node[], ctx: EvalContext): { slope: number; intercept: number; y: number[]; x: number[] } | CellError {
  const p = pairs(args, ctx);
  if (isErr(p)) return p;
  const [y, x] = p;
  if (y.length < 2) return err("#DIV/0!");
  const mx = avg(x);
  const my = avg(y);
  const sxx = x.reduce((s, v) => s + (v - mx) ** 2, 0);
  if (!sxx) return err("#DIV/0!");
  const slope = x.reduce((s, v, i) => s + (v - mx) * (y[i] - my), 0) / sxx;
  return { slope, intercept: my - slope * mx, y, x };
}
def("SLOPE", (args, ctx) => {
  const l = line(args, ctx);
  return isErr(l) ? l : l.slope;
});
def("INTERCEPT", (args, ctx) => {
  const l = line(args, ctx);
  return isErr(l) ? l : l.intercept;
});
def("FORECAST FORECAST.LINEAR", (args, ctx) => {
  const x = num(args[0], ctx);
  if (isErr(x)) return x;
  const l = line([args[1], args[2]], ctx);
  return isErr(l) ? l : l.intercept + l.slope * x;
});
def("CORREL PEARSON", (args, ctx) => {
  const p = pairs(args, ctx);
  if (isErr(p)) return p;
  const [y, x] = p;
  const my = avg(y);
  const mx = avg(x);
  const sxy = x.reduce((s, v, i) => s + (v - mx) * (y[i] - my), 0);
  const sxx = x.reduce((s, v) => s + (v - mx) ** 2, 0);
  const syy = y.reduce((s, v) => s + (v - my) ** 2, 0);
  return sxx && syy ? sxy / Math.sqrt(sxx * syy) : err("#DIV/0!");
});
def("RSQ", (args, ctx) => {
  const r = fns.CORREL(args, ctx);
  return typeof r === "number" ? r * r : r;
});
const covariance = (sample: boolean) => (args: Node[], ctx: EvalContext) => {
  const p = pairs(args, ctx);
  if (isErr(p)) return p;
  const [y, x] = p;
  if (y.length < (sample ? 2 : 1)) return err("#DIV/0!");
  const my = avg(y);
  const mx = avg(x);
  return x.reduce((s, v, i) => s + (v - mx) * (y[i] - my), 0) / (y.length - (sample ? 1 : 0));
};
def("COVARIANCE.S", covariance(true));
def("COVARIANCE.P COVAR", covariance(false));
stat("GEOMEAN", (ns) => !ns.length ? err("#DIV/0!") : (ns.some((n) => n <= 0) ? err("#NUM!") : Math.exp(avg(ns.map(Math.log)))));
stat("HARMEAN", (ns) => !ns.length ? err("#DIV/0!") : (ns.some((n) => n <= 0) ? err("#NUM!") : ns.length / ns.reduce((s, n) => s + 1 / n, 0)));
stat("DEVSQ", (ns) => !ns.length ? err("#DIV/0!") : ns.reduce((s, n) => s + (n - avg(ns)) ** 2, 0));
stat("AVEDEV", (ns) => !ns.length ? err("#DIV/0!") : avg(ns.map((n) => Math.abs(n - avg(ns)))));
const sd = (ns: number[]) => Math.sqrt(ns.reduce((s, n) => s + (n - avg(ns)) ** 2, 0) / (ns.length - 1));
stat("SKEW", (ns) => !ns.length ? err("#DIV/0!") : (ns.length < 3 ? err("#DIV/0!") : (ns.length / ((ns.length - 1) * (ns.length - 2))) * ns.reduce((s, n) => s + ((n - avg(ns)) / sd(ns)) ** 3, 0)));
stat("KURT", (ns) => {
  const n = ns.length;
  if (n < 4) return err("#DIV/0!");
  const s4 = ns.reduce((s, v) => s + ((v - avg(ns)) / sd(ns)) ** 4, 0);
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
});
/** Text and TRUE/FALSE counted too (text as 0, TRUE as 1), as the …A functions do. */
const statA = (names: string, f: (ns: number[]) => number) =>
  def(names, (args, ctx) => {
    const ns = values(args, ctx)
      .filter((v) => v !== null)
      .map((v) => (typeof v === "number" ? v : v === true ? 1 : 0));
    return ns.length ? f(ns) : 0;
  });
statA("AVERAGEA", avg);
statA("MAXA", (ns) => Math.max(...ns));
statA("MINA", (ns) => Math.min(...ns));
math("STANDARDIZE", (x, m, s) => (s <= 0 ? err("#NUM!") : (x - m) / s), 3);

/** The standard normal distribution: its cumulative probability (an approximation good to 1e-7). */
function normCdf(z: number): number {
  const t = 1 / (1 + 0.2316419 * Math.abs(z));
  const d = 0.3989422804014327 * Math.exp((-z * z) / 2);
  const p = d * t * (0.31938153 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))));
  return z >= 0 ? 1 - p : p;
}
/** Its inverse (Acklam's algorithm). */
function normInv(p: number): number {
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const lo = 0.02425;
  if (p < lo) {
    const q = Math.sqrt(-2 * Math.log(p));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  if (p > 1 - lo) return -normInv(1 - p);
  const q = p - 0.5;
  const r = q * q;
  return ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}
def("NORM.S.DIST", (args, ctx) => {
  const z = num(args[0], ctx);
  if (isErr(z)) return z;
  const cumulative = toBool(one(args[1], ctx));
  return cumulative === false ? Math.exp((-z * z) / 2) / Math.sqrt(2 * Math.PI) : normCdf(z);
});
def("NORM.DIST NORMDIST", (args, ctx) => {
  const [x, m, s] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx)];
  for (const v of [x, m, s]) if (isErr(v)) return v;
  if ((s as number) <= 0) return err("#NUM!");
  const z = ((x as number) - (m as number)) / (s as number);
  return toBool(one(args[3], ctx)) === false ? Math.exp((-z * z) / 2) / ((s as number) * Math.sqrt(2 * Math.PI)) : normCdf(z);
});
math("NORM.S.INV NORMSINV", (p) => (p <= 0 || p >= 1 ? err("#NUM!") : normInv(p)), 1);
math("NORM.INV NORMINV", (p, m, s) => (p <= 0 || p >= 1 || s <= 0 ? err("#NUM!") : m + s * normInv(p)), 3);
def("BINOM.DIST BINOMDIST", (args, ctx) => {
  const [k, n, p] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx)];
  for (const v of [k, n, p]) if (isErr(v)) return v;
  const [kk, nn, pp] = [Math.trunc(k as number), Math.trunc(n as number), p as number];
  if (kk < 0 || kk > nn || pp < 0 || pp > 1) return err("#NUM!");
  const pmf = (i: number) => (factorial(nn) / (factorial(i) * factorial(nn - i))) * pp ** i * (1 - pp) ** (nn - i);
  if (toBool(one(args[3], ctx)) !== true) return pmf(kk);
  let total = 0;
  for (let i = 0; i <= kk; i++) total += pmf(i);
  return total;
});
def("POISSON.DIST POISSON", (args, ctx) => {
  const [x, m] = [num(args[0], ctx), num(args[1], ctx)];
  if (isErr(x)) return x;
  if (isErr(m)) return m;
  const k = Math.trunc(x);
  if (k < 0 || m < 0) return err("#NUM!");
  const pmf = (i: number) => (Math.exp(-m) * m ** i) / factorial(i);
  if (toBool(one(args[2], ctx)) !== true) return pmf(k);
  let total = 0;
  for (let i = 0; i <= k; i++) total += pmf(i);
  return total;
});
def("EXPON.DIST EXPONDIST", (args, ctx) => {
  const [x, l] = [num(args[0], ctx), num(args[1], ctx)];
  if (isErr(x)) return x;
  if (isErr(l)) return l;
  if (x < 0 || l <= 0) return err("#NUM!");
  return toBool(one(args[2], ctx)) === true ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x);
});
def("RANK.AVG", (args, ctx) => {
  const x = num(args[0], ctx);
  const ns = numbers([args[1]], ctx);
  if (isErr(x)) return x;
  if (isErr(ns)) return ns;
  const up = !missing(args[2]) && toNumber(one(args[2], ctx)) !== 0;
  const same = ns.filter((n) => n === x).length;
  if (!same) return err("#N/A");
  const before = ns.filter((n) => (up ? n < x : n > x)).length;
  return before + (same + 1) / 2;
});
def("PERCENTRANK PERCENTRANK.INC", (args, ctx) => {
  const ns = numbers([args[0]], ctx);
  const x = num(args[1], ctx);
  const digits = num(args[2], ctx, 3);
  if (isErr(ns)) return ns;
  if (isErr(x)) return x;
  if (isErr(digits)) return digits;
  const sorted = [...ns].sort((a, b) => a - b);
  if (!sorted.length || x < sorted[0] || x > sorted[sorted.length - 1]) return err("#N/A");
  let i = sorted.findIndex((n) => n >= x);
  const rank = sorted[i] === x ? sorted.indexOf(x) : i - 1 + (x - sorted[i - 1]) / (sorted[i] - sorted[i - 1]);
  i = 10 ** digits;
  return Math.floor((rank / (sorted.length - 1)) * i) / i;
});

// ---- money ---------------------------------------------------------------------------------

/** The payment for a loan or saving: rate per period, periods, present and future value, paid at the start (1) or end (0). */
const payment = (rate: number, n: number, pv: number, fv: number, type: number) =>
  rate === 0 ? -(pv + fv) / n : (-(pv * (1 + rate) ** n + fv) * rate) / ((1 + rate * type) * ((1 + rate) ** n - 1));
/** The balance after n periods, in Excel's signs (FV). */
const balanceAfter = (rate: number, n: number, pmt: number, pv: number, type: number) =>
  rate === 0 ? -(pv + pmt * n) : -(pv * (1 + rate) ** n + (pmt * (1 + rate * type) * ((1 + rate) ** n - 1)) / rate);
/** The interest in one period's payment: the balance at its start times the rate (none in the first, when paying in advance). */
function interestPart(rate: number, per: number, n: number, pv: number, fv: number, type: number): number {
  const pmt = payment(rate, n, pv, fv, type);
  const interest = balanceAfter(rate, per - 1, pmt, pv, type) * rate;
  if (!type) return interest;
  return per === 1 ? 0 : interest / (1 + rate);
}
const loanArgs = (args: Node[], ctx: EvalContext, from: number) => {
  const out = [num(args[from], ctx), num(args[from + 1], ctx), num(args[from + 2], ctx), num(args[from + 3], ctx, 0), num(args[from + 4], ctx, 0)];
  const bad = out.find(isErr);
  return bad ? (bad as CellError) : (out as number[]);
};
def("IPMT", (args, ctx) => {
  const rate = num(args[0], ctx);
  const l = loanArgs(args, ctx, 1);
  if (isErr(rate)) return rate;
  if (isErr(l)) return l;
  const [per, n, pv, fv, type] = l;
  if (per < 1 || per > n) return err("#NUM!");
  return interestPart(rate, per, n, pv, fv, type);
});
def("PPMT", (args, ctx) => {
  const rate = num(args[0], ctx);
  const l = loanArgs(args, ctx, 1);
  if (isErr(rate)) return rate;
  if (isErr(l)) return l;
  const [per, n, pv, fv, type] = l;
  if (per < 1 || per > n) return err("#NUM!");
  return payment(rate, n, pv, fv, type) - interestPart(rate, per, n, pv, fv, type);
});
def("NPER", (args, ctx) => {
  const [rate, pmt, pv, fv, type] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx), num(args[3], ctx, 0), num(args[4], ctx, 0)];
  for (const v of [rate, pmt, pv, fv, type]) if (isErr(v)) return v;
  const [r, p, v0, v1, t] = [rate, pmt, pv, fv, type] as number[];
  if (r === 0) return p === 0 ? err("#NUM!") : -(v0 + v1) / p;
  const a = p * (1 + r * t) - v1 * r;
  const b = v0 * r + p * (1 + r * t);
  return a / b <= 0 ? err("#NUM!") : Math.log(a / b) / Math.log(1 + r);
});
/** Finds where f is 0 by Newton's method, from a guess. */
function solve(f: (x: number) => number, guess: number): number | CellError {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    const dy = (f(x + 1e-7) - y) / 1e-7;
    if (!Number.isFinite(dy) || dy === 0) break;
    const next = x - y / dy;
    if (Math.abs(next - x) < 1e-10) return next;
    x = next;
  }
  return Math.abs(f(x)) < 1e-6 ? x : err("#NUM!");
}
def("RATE", (args, ctx) => {
  const [n, pmt, pv, fv, type, guess] = [num(args[0], ctx), num(args[1], ctx), num(args[2], ctx), num(args[3], ctx, 0), num(args[4], ctx, 0), num(args[5], ctx, 0.1)];
  for (const v of [n, pmt, pv, fv, type, guess]) if (isErr(v)) return v;
  const [N, P, V0, V1, T] = [n, pmt, pv, fv, type] as number[];
  return solve((r) => (Math.abs(r) < 1e-12 ? V0 + P * N + V1 : V0 * (1 + r) ** N + P * (1 + r * T) * (((1 + r) ** N - 1) / r) + V1), guess as number);
});
def("IRR", (args, ctx) => {
  const flows = numbers([args[0]], ctx);
  const guess = num(args[1], ctx, 0.1);
  if (isErr(flows)) return flows;
  if (isErr(guess)) return guess;
  if (!flows.some((f) => f > 0) || !flows.some((f) => f < 0)) return err("#NUM!");
  return solve((r) => flows.reduce((s, f, i) => s + f / (1 + r) ** i, 0), guess);
});
def("XNPV", (args, ctx) => {
  const rate = num(args[0], ctx);
  const flows = numbers([args[1]], ctx);
  const dates = numbers([args[2]], ctx);
  if (isErr(rate)) return rate;
  if (isErr(flows)) return flows;
  if (isErr(dates)) return dates;
  if (flows.length !== dates.length) return err("#NUM!");
  return flows.reduce((s, f, i) => s + f / (1 + rate) ** ((dates[i] - dates[0]) / 365), 0);
});
def("XIRR", (args, ctx) => {
  const flows = numbers([args[0]], ctx);
  const dates = numbers([args[1]], ctx);
  const guess = num(args[2], ctx, 0.1);
  if (isErr(flows)) return flows;
  if (isErr(dates)) return dates;
  if (isErr(guess)) return guess;
  if (flows.length !== dates.length || !flows.some((f) => f > 0) || !flows.some((f) => f < 0)) return err("#NUM!");
  return solve((r) => flows.reduce((s, f, i) => s + f / (1 + r) ** ((dates[i] - dates[0]) / 365), 0), guess);
});
math("SLN", (cost, salvage, life) => (life === 0 ? err("#DIV/0!") : (cost - salvage) / life), 3);

// ---- text ------------------------------------------------------------------------------------

const around = (after: boolean): Fn => (args, ctx) => {
  const text = toText(one(args[0], ctx));
  const delim = toText(one(args[1], ctx));
  const n = num(args[2], ctx, 1);
  if (isErr(n)) return n;
  if (!n) return err("#VALUE!");
  const ignoreCase = !missing(args[3]) && toNumber(one(args[3], ctx)) === 1;
  const hay = ignoreCase ? text.toLowerCase() : text;
  const needle = ignoreCase ? delim.toLowerCase() : delim;
  let at = -1;
  if (n > 0) for (let i = 0, from = 0; i < n; i++, from = at + needle.length) at = hay.indexOf(needle, from);
  else for (let i = 0, from = hay.length; i < -n; i++, from = at - 1) at = from < 0 ? -1 : hay.lastIndexOf(needle, from);
  if (at < 0) return missing(args[5]) ? err("#N/A") : one(args[5], ctx);
  return after ? text.slice(at + delim.length) : text.slice(0, at);
};
def("TEXTBEFORE", around(false));
def("TEXTAFTER", around(true));
def("TEXTSPLIT", (args, ctx) => {
  const text = toText(one(args[0], ctx));
  const cols = missing(args[1]) ? null : toText(one(args[1], ctx));
  const rows = missing(args[2]) ? null : toText(one(args[2], ctx));
  const lines = rows ? text.split(rows) : [text];
  const grid2 = lines.map((line) => (cols ? line.split(cols) : [line]));
  const width = Math.max(...grid2.map((r) => r.length));
  return new ArrayValue(grid2.map((r) => Array.from({ length: width }, (_, j) => r[j] ?? err("#N/A"))));
});
def("UNICHAR", (args, ctx) => {
  const n = num(args[0], ctx);
  return isErr(n) ? n : n < 1 || n > 0x10ffff ? err("#VALUE!") : String.fromCodePoint(Math.trunc(n));
});
def("UNICODE", (args, ctx) => {
  const t = toText(one(args[0], ctx));
  return t ? t.codePointAt(0)! : err("#VALUE!");
});
def("FIXED", (args, ctx) => {
  const n = num(args[0], ctx);
  const d = num(args[1], ctx, 2);
  if (isErr(n)) return n;
  if (isErr(d)) return d;
  const noCommas = !missing(args[2]) && toBool(one(args[2], ctx)) === true;
  const rounded = Math.round(n * 10 ** d) / 10 ** d;
  return formatValue(rounded, `${noCommas ? "0" : "#,##0"}${d > 0 ? `.${"0".repeat(d)}` : ""}`).text;
});

// ---- lookups, arrays and the rest ---------------------------------------------------------------

def("LOOKUP", (args, ctx) => {
  const x = one(args[0], ctx);
  const look = grid(args[1], ctx);
  const result = missing(args[2]) ? null : grid(args[2], ctx);
  const vector = look.length === 1 || look[0].length === 1 ? look.flat() : look.length >= look[0].length ? look.map((r) => r[0]) : look[0];
  let at = -1;
  vector.forEach((v, i) => {
    if (v !== null && compare(v, x) <= 0) at = i;
  });
  if (at < 0) return err("#N/A");
  if (result) return result.flat()[at] ?? err("#N/A");
  if (look.length === 1 || look[0].length === 1) return vector[at];
  return look.length >= look[0].length ? look[at][look[0].length - 1] : look[look.length - 1][at];
});
def("ADDRESS", (args, ctx) => {
  const r = num(args[0], ctx);
  const c = num(args[1], ctx);
  const abs = num(args[2], ctx, 1);
  for (const v of [r, c, abs]) if (isErr(v)) return v;
  if ((r as number) < 1 || (c as number) < 1) return err("#VALUE!");
  let col = "";
  for (let n = c as number; n > 0; n = Math.floor((n - 1) / 26)) col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
  const rowAbs = abs === 1 || abs === 2;
  const colAbs = abs === 1 || abs === 3;
  const sheet = missing(args[4]) ? "" : `${toText(one(args[4], ctx))}!`;
  return `${sheet}${colAbs ? "$" : ""}${col}${rowAbs ? "$" : ""}${r}`;
});
const pick = (rowsWanted: boolean): Fn => (args, ctx) => {
  const data = grid(args[0], ctx);
  const size = rowsWanted ? data.length : data[0].length;
  const out: Scalar[][] = [];
  for (const a of args.slice(1)) {
    const n = num(a, ctx);
    if (isErr(n)) return n;
    const i = n < 0 ? size + n : n - 1;
    if (!n || i < 0 || i >= size) return err("#VALUE!");
    out.push(rowsWanted ? data[i] : data.map((row) => row[i]));
  }
  return new ArrayValue(rowsWanted ? out : transpose(out));
};
def("CHOOSEROWS", pick(true));
def("CHOOSECOLS", pick(false));
const slice = (take: boolean): Fn => (args, ctx) => {
  const data = grid(args[0], ctx);
  const r = missing(args[1]) ? (take ? data.length : 0) : num(args[1], ctx);
  const c = missing(args[2]) ? (take ? data[0].length : 0) : num(args[2], ctx);
  if (isErr(r)) return r;
  if (isErr(c)) return c;
  const cut = <T,>(list: T[], n: number) => (take ? (n >= 0 ? list.slice(0, n) : list.slice(n)) : n >= 0 ? list.slice(n) : list.slice(0, n));
  const rows = cut(data, r).map((row) => cut(row, c));
  return rows.length && rows[0].length ? new ArrayValue(rows) : err("#CALC!");
};
def("TAKE", slice(true));
def("DROP", slice(false));
def("VSTACK", (args, ctx) => {
  const parts = args.map((a) => grid(a, ctx));
  const width = Math.max(...parts.map((p) => p[0].length));
  return new ArrayValue(parts.flatMap((p) => p.map((row) => Array.from({ length: width }, (_, j) => (j < row.length ? row[j] : err("#N/A"))))));
});
def("HSTACK", (args, ctx) => {
  const parts = args.map((a) => grid(a, ctx));
  const height = Math.max(...parts.map((p) => p.length));
  return new ArrayValue(Array.from({ length: height }, (_, i) => parts.flatMap((p) => (i < p.length ? p[i] : p[0].map(() => err("#N/A"))))));
});
def("TYPE", (args, ctx) => {
  const v = ev(args[0], ctx);
  if (v instanceof ArrayValue || (v instanceof RangeValue && v.rows * v.cols > 1)) return 64;
  const s = scalar(v, ctx);
  return typeof s === "number" || s === null ? 1 : typeof s === "string" ? 2 : typeof s === "boolean" ? 4 : 16;
});
def("YEARFRAC", (args, ctx) => {
  const a = num(args[0], ctx);
  const b = num(args[1], ctx);
  const basis = num(args[2], ctx, 0);
  for (const v of [a, b, basis]) if (isErr(v)) return v;
  const [start, end] = [Math.min(a as number, b as number), Math.max(a as number, b as number)];
  const d1 = serialToDate(start, ctx.date1904);
  const d2 = serialToDate(end, ctx.date1904);
  switch (basis) {
    case 1:
      return (end - start) / 365.25;
    case 2:
      return (end - start) / 360;
    case 3:
      return (end - start) / 365;
    default: {
      // 30/360 (US), Excel's default.
      let [day1, day2] = [d1.getUTCDate(), d2.getUTCDate()];
      if (day1 === 31) day1 = 30;
      if (day2 === 31 && day1 === 30) day2 = 30;
      return ((d2.getUTCFullYear() - d1.getUTCFullYear()) * 360 + (d2.getUTCMonth() - d1.getUTCMonth()) * 30 + (day2 - day1)) / 360;
    }
  }
});
def("DAYS360", (args, ctx) => {
  const a = num(args[0], ctx);
  const b = num(args[1], ctx);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const d1 = serialToDate(a, ctx.date1904);
  const d2 = serialToDate(b, ctx.date1904);
  let [day1, day2] = [d1.getUTCDate(), d2.getUTCDate()];
  if (day1 === 31) day1 = 30;
  if (day2 === 31 && day1 >= 30) day2 = 30;
  return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 360 + (d2.getUTCMonth() - d1.getUTCMonth()) * 30 + (day2 - day1);
});

/** SUBTOTAL: a summary that leaves out other SUBTOTALs (and, for 101–111, hidden rows), as under filters. */
const SUBTOTALS: Record<number, string> = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP" };
def("SUBTOTAL", (args, ctx) => {
  const which = num(args[0], ctx);
  if (isErr(which)) return which;
  const name = SUBTOTALS[which > 100 ? which - 100 : which];
  if (!name) return err("#VALUE!");
  return fns[name](args.slice(1), ctx);
});

/** LET: names for values, used in the formula that follows. */
def("LET", (args, ctx) => {
  if (args.length < 3 || args.length % 2 === 0) return err("#VALUE!");
  const local = new Map<string, Value>();
  let inner: EvalContext = ctx;
  for (let i = 0; i < args.length - 1; i += 2) {
    const nameNode = args[i];
    if (nameNode.t !== "name" || !nameNode.name) return err("#VALUE!");
    local.set(nameNode.name.toUpperCase(), ev(args[i + 1], inner));
    const outer = inner;
    inner = { ...outer, local: new Map([...(outer.local ?? []), ...local]) };
  }
  return ev(args[args.length - 1], inner);
});

// ---- complex numbers, as text like "3+4i" (for electronics) ---------------------------------------

type Complex = { re: number; im: number; unit: "i" | "j" };
function readComplex(v: Scalar): Complex | CellError {
  if (typeof v === "number") return { re: v, im: 0, unit: "i" };
  const t = toText(v).replace(/\s+/g, "");
  if (!t) return { re: 0, im: 0, unit: "i" };
  const m = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?(?:([+-])((?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?([ij]))?$/.exec(t);
  const imOnly = /^([+-]?)((?:\d+\.?\d*|\.\d+)?)([ij])$/.exec(t);
  if (imOnly) return { re: 0, im: Number(`${imOnly[1]}${imOnly[2] || "1"}`), unit: imOnly[3] as "i" };
  if (!m || (!m[1] && !m[4])) return err("#NUM!");
  return { re: Number(m[1] ?? 0), im: m[4] ? Number(`${m[2]}${m[3] ?? "1"}`) : 0, unit: (m[4] as "i") ?? "i" };
}
function writeComplex(c: Complex): string {
  const round = (n: number) => Number(n.toPrecision(15));
  const re = round(c.re);
  const im = round(c.im);
  if (!im) return String(re);
  const imText = im === 1 ? "" : im === -1 ? "-" : String(im);
  if (!re) return `${imText}${c.unit}`;
  return `${re}${im > 0 ? "+" : ""}${imText}${c.unit}`;
}
const complexArgs = (args: Node[], ctx: EvalContext) => {
  const out: Complex[] = [];
  for (const v of values(args, ctx)) {
    if (v === null) continue;
    const c = readComplex(v);
    if (isErr(c)) return c;
    out.push(c);
  }
  return out;
};
def("COMPLEX", (args, ctx) => {
  const re = num(args[0], ctx);
  const im = num(args[1], ctx);
  if (isErr(re)) return re;
  if (isErr(im)) return im;
  const unit = missing(args[2]) ? "i" : toText(one(args[2], ctx));
  return unit !== "i" && unit !== "j" ? err("#VALUE!") : writeComplex({ re, im, unit });
});
const complexOne = (names: string, f: (c: Complex) => Scalar | Complex) =>
  def(names, (args, ctx) => {
    const c = readComplex(one(args[0], ctx));
    if (isErr(c)) return c;
    const out = f(c);
    return out !== null && typeof out === "object" && "re" in out ? writeComplex(out) : (out as Scalar);
  });
complexOne("IMREAL", (c) => c.re);
complexOne("IMAGINARY", (c) => c.im);
complexOne("IMABS", (c) => Math.hypot(c.re, c.im));
complexOne("IMARGUMENT", (c) => (c.re === 0 && c.im === 0 ? err("#DIV/0!") : Math.atan2(c.im, c.re)));
complexOne("IMCONJUGATE", (c) => ({ ...c, im: -c.im }));
complexOne("IMSQRT", (c) => {
  const r = Math.sqrt(Math.hypot(c.re, c.im));
  const a = Math.atan2(c.im, c.re) / 2;
  return { re: r * Math.cos(a), im: r * Math.sin(a), unit: c.unit };
});
complexOne("IMEXP", (c) => ({ re: Math.exp(c.re) * Math.cos(c.im), im: Math.exp(c.re) * Math.sin(c.im), unit: c.unit }));
complexOne("IMLN", (c) => (c.re === 0 && c.im === 0 ? err("#NUM!") : { re: Math.log(Math.hypot(c.re, c.im)), im: Math.atan2(c.im, c.re), unit: c.unit }));
def("IMSUM", (args, ctx) => {
  const cs = complexArgs(args, ctx);
  return isErr(cs) ? cs : writeComplex(cs.reduce((s, c) => ({ re: s.re + c.re, im: s.im + c.im, unit: c.unit }), { re: 0, im: 0, unit: "i" } as Complex));
});
def("IMPRODUCT", (args, ctx) => {
  const cs = complexArgs(args, ctx);
  return isErr(cs) ? cs : writeComplex(cs.reduce((s, c) => ({ re: s.re * c.re - s.im * c.im, im: s.re * c.im + s.im * c.re, unit: c.unit }), { re: 1, im: 0, unit: "i" } as Complex));
});
def("IMSUB", (args, ctx) => {
  const a = readComplex(one(args[0], ctx));
  const b = readComplex(one(args[1], ctx));
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return writeComplex({ re: a.re - b.re, im: a.im - b.im, unit: a.unit });
});
def("IMDIV", (args, ctx) => {
  const a = readComplex(one(args[0], ctx));
  const b = readComplex(one(args[1], ctx));
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const d = b.re * b.re + b.im * b.im;
  if (!d) return err("#NUM!");
  return writeComplex({ re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d, unit: a.unit });
});
def("IMPOWER", (args, ctx) => {
  const c = readComplex(one(args[0], ctx));
  const p = num(args[1], ctx);
  if (isErr(c)) return c;
  if (isErr(p)) return p;
  const r = Math.hypot(c.re, c.im) ** p;
  const a = Math.atan2(c.im, c.re) * p;
  return writeComplex({ re: r * Math.cos(a), im: r * Math.sin(a), unit: c.unit });
});

Object.assign(FUNCTIONS, fns);

/** The names of all the functions, for suggestions while typing. */
export const FUNCTION_NAMES = Object.keys(fns).sort();
