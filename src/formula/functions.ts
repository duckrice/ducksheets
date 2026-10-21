// Excel's functions, the ones people use day to day: sums and statistics,
// the conditional ...IF(S) family, logic, lookups, text, dates and a few
// financial ones. Each follows Excel's rules for what it accepts and returns.

import { CellError, type Scalar } from "../model";
import { dateToSerial, formatValue, parseInput, serialToDate } from "../numfmt";
import { FUNCTIONS, RangeValue, cellsOf, compare, err, evaluate, isErr, scalar, toBool, toNumber, toText, type EvalContext, type Fn, type Value } from "./eval";
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
    if (v instanceof RangeValue) {
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
    if (v instanceof RangeValue) out.push(...cellsOf(v, ctx));
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
  const ranges = args.map((a) => range(a, ctx));
  const bad = ranges.find(isErr);
  if (bad) return bad as CellError;
  const rs = ranges as RangeValue[];
  if (rs.some((r) => r.rows !== rs[0].rows || r.cols !== rs[0].cols)) return err("#VALUE!");
  let total = 0;
  for (let i = 0; i < rs[0].rows; i++)
    for (let j = 0; j < rs[0].cols; j++) {
      let p = 1;
      for (const r of rs) {
        const v = cellAt(r, i, j, ctx);
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
  const r = range(args[0], ctx);
  return isErr(r) ? r : r.rows;
});
def("COLUMNS", (args, ctx) => {
  const r = range(args[0], ctx);
  return isErr(r) ? r : r.cols;
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

Object.assign(FUNCTIONS, fns);

/** The names of all the functions, for suggestions while typing. */
export const FUNCTION_NAMES = Object.keys(fns).sort();
