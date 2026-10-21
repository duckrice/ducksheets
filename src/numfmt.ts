// How values look: Excel's number format codes (#,##0.00, 0%, dd.mm.yyyy,
// "kr" …), in the computer's own decimal comma or point. And the reverse:
// what someone types into a cell, read as Excel reads it (10% is 0.1 shown as
// a percentage, 20.10.2026 is a date).

import { CellError, type Scalar } from "./model";

// ---- the computer's locale ---------------------------------------------------------------

function localeParts() {
  const parts = new Intl.NumberFormat(undefined).formatToParts(12345.6);
  const decimal = parts.find((p) => p.type === "decimal")?.value ?? ".";
  const group = parts.find((p) => p.type === "group")?.value ?? ",";
  const month = (style: "long" | "short") => Array.from({ length: 12 }, (_, i) => new Intl.DateTimeFormat(undefined, { month: style, timeZone: "UTC" }).format(Date.UTC(2021, i, 1)).replace(/\.$/, ""));
  const day = (style: "long" | "short") => Array.from({ length: 7 }, (_, i) => new Intl.DateTimeFormat(undefined, { weekday: style, timeZone: "UTC" }).format(Date.UTC(2021, 0, 3 + i)).replace(/\.$/, ""));
  // The short date as a format code: 20.10.2026 → dd.mm.yyyy, 10/20/2026 → mm/dd/yyyy.
  const dateCode = new Intl.DateTimeFormat(undefined, { year: "numeric", month: "2-digit", day: "2-digit", timeZone: "UTC" })
    .formatToParts(Date.UTC(2026, 9, 20))
    .map((p) => (p.type === "day" ? "dd" : p.type === "month" ? "mm" : p.type === "year" ? "yyyy" : p.value))
    .join("");
  return { decimal, group: group === " " || group === " " ? " " : group, months: month("long"), monthsShort: month("short"), days: day("long"), daysShort: day("short"), dateCode };
}

export const locale = localeParts();

// ---- dates ---------------------------------------------------------------------------------

const DAY = 86400000;

/** An Excel date number as a UTC date. Excel counts 29 February 1900, which didn't exist. */
export function serialToDate(serial: number, date1904 = false): Date {
  if (date1904) return new Date(Date.UTC(1904, 0, 1) + Math.round(serial * DAY));
  const days = serial >= 60 ? serial - 1 : serial;
  return new Date(Date.UTC(1899, 11, 31) + Math.round(days * DAY));
}

export function dateToSerial(y: number, m: number, d: number, date1904 = false): number {
  const ms = Date.UTC(y, m - 1, d);
  if (date1904) return Math.round((ms - Date.UTC(1904, 0, 1)) / DAY);
  const days = Math.round((ms - Date.UTC(1899, 11, 31)) / DAY);
  return days >= 60 ? days + 1 : days;
}

// ---- built-in formats ------------------------------------------------------------------------

const BUILTIN: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "hh:mm", 21: "hh:mm:ss",
  37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mm:ss.0", 48: "##0.0E+0", 49: "@",
};

/** The format code for a built-in number format id. 14 and 22 follow the computer's date order, as in Excel. */
export function builtinFormat(id: number): string {
  if (id === 14) return locale.dateCode;
  if (id === 22) return `${locale.dateCode} hh:mm`;
  return BUILTIN[id] ?? "General";
}

// ---- formatting ---------------------------------------------------------------------------------

export interface Shown {
  text: string;
  color?: string;
}

const COLORS: Record<string, string> = { black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#008000", magenta: "#ff00ff", red: "#ff0000", white: "#ffffff", yellow: "#ffff00" };

/** Splits a format code into its ;-separated sections, ignoring ; inside quotes and brackets. */
function sections(code: string): string[] {
  const out: string[] = [];
  let current = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      current += code.slice(i, end < 0 ? code.length : end + 1);
      i = end < 0 ? code.length : end;
    } else if (ch === "\\") {
      current += code.slice(i, i + 2);
      i++;
    } else if (ch === "[") {
      const end = code.indexOf("]", i);
      current += code.slice(i, end < 0 ? code.length : end + 1);
      i = end < 0 ? code.length : end;
    } else if (ch === ";") {
      out.push(current);
      current = "";
    } else current += ch;
  }
  out.push(current);
  return out;
}

const formatCache = new Map<string, string[]>();

/** Shows a value in a format code. */
export function formatValue(value: Scalar, code: string, date1904 = false): Shown {
  if (value === null || value === undefined) return { text: "" };
  if (value instanceof CellError) return { text: value.code };
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE" };
  let parts = formatCache.get(code);
  if (!parts) formatCache.set(code, (parts = sections(code)));
  if (typeof value === "string") {
    const text = parts.length >= 4 ? parts[3] : parts.find((p) => p.includes("@"));
    return text ? render(text, 0, date1904, value) : { text: value };
  }
  if (!Number.isFinite(value)) return { text: "#NUM!" };
  // Which section: conditions like [<0] if there are any, else positive;negative;zero.
  let section = parts[0];
  let abs = false;
  const condition = (p: string) => /\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]/.exec(p);
  if (parts.some((p) => condition(p))) {
    const hit = parts.find((p) => {
      const m = condition(p);
      if (!m) return false;
      const n = Number(m[2]);
      return { "<": value < n, ">": value > n, "=": value === n, "<=": value <= n, ">=": value >= n, "<>": value !== n }[m[1] as "<"];
    });
    section = hit ?? parts.find((p) => !condition(p) && !p.includes("@")) ?? parts[parts.length - 1];
  } else if (value < 0 && parts.length >= 2 && parts[1] !== "") {
    section = parts[1];
    abs = true;
  } else if (value === 0 && parts.length >= 3) section = parts[2];
  return render(section, abs ? Math.abs(value) : value, date1904);
}

function general(value: number): string {
  const abs = Math.abs(value);
  let text: string;
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) {
    text = value.toExponential(5).replace(/\.?0+e/, "e").replace("e", "E").replace(/E([+-])(\d)$/, "E$10$2");
  } else {
    text = String(Number(value.toPrecision(10)));
    if (/e/.test(text)) text = value.toFixed(10).replace(/\.?0+$/, "");
  }
  return text.replace(".", locale.decimal);
}

function render(section: string, value: number, date1904: boolean, text?: string): Shown {
  let color: string | undefined;
  let body = section.replace(/\[([^\]]*)\]/g, (whole, inner: string) => {
    const lower = inner.toLowerCase();
    if (COLORS[lower]) color = COLORS[lower];
    else if (/^color\d+$/.test(lower)) color = undefined;
    else if (inner.startsWith("$")) return `"${inner.slice(1).split("-")[0]}"`; // [$kr-414] shows "kr"
    else if (/^(h+|m+|s+)$/.test(lower)) return whole; // elapsed time, kept for the date code
    return "";
  });
  if (text !== undefined) return { text: literal(body).replace(/@/g, text), color };
  if (/^\s*general\s*$/i.test(body)) return { text: general(value), color };
  if (/^general/i.test(body.trim())) body = body.replace(/general/i, "");
  // Without quotes, is it a date or time?
  const bare = body.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  if (/[dyhs]|\[h|AM\/PM|A\/P/i.test(bare) || (/m/i.test(bare) && !/[0#?]/.test(bare))) return { text: dateText(body, value, date1904), color };
  return { text: numberText(body, value), color };
}

/** The literal parts of a section: quotes and escapes resolved, padding and fills dropped. */
function literal(body: string): string {
  return body.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1").replace(/_./g, " ").replace(/\*./g, "");
}

function numberText(body: string, value: number): string {
  // Tokens: digit placeholders, the decimal point, and everything else as text.
  const tokens: { k: "d" | "." | "t" | "e"; v: string }[] = [];
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      tokens.push({ k: "t", v: body.slice(i + 1, end < 0 ? body.length : end) });
      i = end < 0 ? body.length : end;
    } else if (ch === "\\") tokens.push({ k: "t", v: body[++i] ?? "" });
    else if (ch === "_") {
      tokens.push({ k: "t", v: " " });
      i++;
    } else if (ch === "*") i++;
    else if (ch === "0" || ch === "#" || ch === "?") tokens.push({ k: "d", v: ch });
    else if (ch === ".") tokens.push({ k: ".", v: ch });
    else if ((ch === "E" || ch === "e") && /^[+-]/.test(body[i + 1] ?? "")) {
      tokens.push({ k: "e", v: body.slice(i, i + 2) });
      i++;
    } else tokens.push({ k: "t", v: ch });
  }
  let n = value;
  const percent = tokens.filter((t) => t.k === "t" && t.v === "%").length;
  n *= 100 ** percent;
  // Commas right after the last digit scale by 1000; commas between digits group.
  const lastDigit = tokens.map((t) => t.k).lastIndexOf("d");
  let scale = 0;
  for (let i = lastDigit + 1; i < tokens.length && tokens[i].k === "t" && tokens[i].v === ","; i++) {
    scale++;
    tokens[i] = { k: "t", v: "" };
  }
  n /= 1000 ** scale;
  const grouped = tokens.some((t, i) => t.k === "t" && t.v === "," && i < lastDigit);
  for (const t of tokens) if (t.k === "t" && t.v === ",") t.v = "";

  const point = tokens.findIndex((t) => t.k === ".");
  const exp = tokens.findIndex((t) => t.k === "e");
  const intEnd = point >= 0 ? point : exp >= 0 ? exp : tokens.length;
  const decEnd = exp >= 0 ? exp : tokens.length;
  const decimals = point >= 0 ? tokens.slice(point + 1, decEnd).filter((t) => t.k === "d") : [];
  let exponent = 0;
  if (exp >= 0 && n !== 0) {
    const intDigits = tokens.slice(0, intEnd).filter((t) => t.k === "d").length || 1;
    exponent = Math.floor(Math.log10(Math.abs(n))) - (intDigits - 1);
    n /= 10 ** exponent;
  }
  const negative = n < 0 && !/^-/.test(body);
  n = Math.abs(n);
  const fixed = n.toFixed(decimals.length);
  let [intPart, decPart = ""] = fixed.split(".");
  if (intPart === "0") intPart = "";

  // Integer digits fill the placeholders from the right; extra digits go before the first.
  const intTokens = tokens.slice(0, intEnd);
  const out: string[] = [];
  let digits = intPart;
  const places = intTokens.filter((t) => t.k === "d").length;
  for (let i = intTokens.length - 1, seen = 0; i >= 0; i--) {
    const t = intTokens[i];
    if (t.k !== "d") {
      out.unshift(t.v);
      continue;
    }
    seen++;
    let ch = digits ? digits[digits.length - 1] : t.v === "0" ? "0" : t.v === "?" ? " " : "";
    digits = digits.slice(0, -1);
    if (seen === places && digits) {
      ch = digits + ch;
      digits = "";
    }
    out.unshift(ch);
  }
  let intText = out.join("");
  if (grouped) {
    // Group the digit run that holds the integer.
    intText = intText.replace(/\d+/, (run) => run.replace(/\B(?=(\d{3})+(?!\d))/g, locale.group));
  }
  let result = intText;
  if (point >= 0) {
    let dec = "";
    let d = 0;
    for (const t of tokens.slice(point + 1, decEnd)) {
      if (t.k === "d") {
        dec += decPart[d++] ?? "0";
      } else dec += t.v;
    }
    // # drops trailing zeros, ? turns them into spaces.
    const chars = dec.split("");
    for (let i = chars.length - 1, k = decimals.length - 1; i >= 0 && k >= 0; i--) {
      if (!/\d/.test(chars[i])) continue;
      if (chars[i] !== "0" || decimals[k].v === "0") break;
      chars[i] = decimals[k].v === "?" ? " " : "";
      k--;
    }
    dec = chars.join("");
    result += (dec || decimals.some((t) => t.v === "0") ? locale.decimal : "") + dec;
    if (!dec && !intText.trim()) result = "0";
  }
  if (exp >= 0) {
    const sign = exponent < 0 ? "-" : tokens[exp].v[1] === "+" ? "+" : "";
    const width = tokens.slice(exp + 1).filter((t) => t.k === "d").length;
    result += `E${sign}${String(Math.abs(exponent)).padStart(width, "0")}${tokens
      .slice(exp + 1)
      .filter((t) => t.k === "t")
      .map((t) => t.v)
      .join("")}`;
  }
  return (negative ? "-" : "") + result;
}

function dateText(body: string, value: number, date1904: boolean): string {
  const elapsed = /\[(h+|m+|s+)\]/i.exec(body);
  const date = serialToDate(value, date1904);
  const totalSeconds = Math.round(value * 86400);
  const secondsOfDay = Math.round((value - Math.floor(value)) * 86400) % 86400;
  const h24 = Math.floor(secondsOfDay / 3600);
  const minutes = Math.floor((secondsOfDay % 3600) / 60);
  const seconds = secondsOfDay % 60;
  const ampm = /AM\/PM|A\/P/i.test(body);
  const out: string[] = [];
  // m after h, or before s, is minutes.
  const tokens = body.match(/"[^"]*"|\\.|\[[hms]+\]|AM\/PM|A\/P|y+|m+|d+|h+|s+(\.0+)?|./gi) ?? [];
  const isTime = (i: number) => {
    for (let j = i - 1; j >= 0; j--) if (/^[hm]|\[h/i.test(tokens[j]) || /^s/i.test(tokens[j])) return /^h|\[h/i.test(tokens[j]);
    for (let j = i + 1; j < tokens.length; j++) if (/^[dhmsy]/i.test(tokens[j])) return /^s/i.test(tokens[j]);
    return false;
  };
  tokens.forEach((tok, i) => {
    const lower = tok.toLowerCase();
    if (tok.startsWith('"')) out.push(tok.slice(1, -1));
    else if (tok.startsWith("\\")) out.push(tok.slice(1));
    else if (elapsed && /^\[[hms]+\]$/.test(lower)) {
      const total = lower[1] === "h" ? Math.floor(totalSeconds / 3600) : lower[1] === "m" ? Math.floor(totalSeconds / 60) : totalSeconds;
      out.push(String(total).padStart(lower.length - 2, "0"));
    } else if (/^y+$/.test(lower)) out.push(lower.length <= 2 ? String(date.getUTCFullYear() % 100).padStart(2, "0") : String(date.getUTCFullYear()));
    else if (/^d+$/.test(lower)) {
      const day = date.getUTCDate();
      out.push(lower.length === 1 ? String(day) : lower.length === 2 ? String(day).padStart(2, "0") : lower.length === 3 ? locale.daysShort[date.getUTCDay()] : locale.days[date.getUTCDay()]);
    } else if (/^m+$/.test(lower)) {
      if (lower.length <= 2 && isTime(i)) out.push(lower.length === 1 ? String(minutes) : String(minutes).padStart(2, "0"));
      else {
        const month = date.getUTCMonth();
        out.push(lower.length === 1 ? String(month + 1) : lower.length === 2 ? String(month + 1).padStart(2, "0") : lower.length === 3 ? locale.monthsShort[month] : lower.length === 5 ? locale.months[month][0] : locale.months[month]);
      }
    } else if (/^h+$/.test(lower)) {
      const h = ampm ? h24 % 12 || 12 : h24;
      out.push(lower.length === 1 ? String(h) : String(h).padStart(2, "0"));
    } else if (/^s+/.test(lower)) {
      const fraction = /\.(0+)/.exec(lower);
      const s = lower.startsWith("ss") ? String(seconds).padStart(2, "0") : String(seconds);
      out.push(fraction ? `${s}${locale.decimal}${String(Math.round(((value * 86400) % 1) * 10 ** fraction[1].length)).padStart(fraction[1].length, "0")}` : s);
    } else if (lower === "am/pm") out.push(h24 < 12 ? "AM" : "PM");
    else if (lower === "a/p") out.push(h24 < 12 ? "A" : "P");
    else if (tok === "_" || tok === "*") out.push("");
    else out.push(tok);
  });
  return out.join("");
}

// ---- reading what's typed ---------------------------------------------------------------------

export interface Typed {
  value: Scalar;
  /** A format the value suggests, if the cell has none yet: typing 10% asks for a percentage. */
  format?: string;
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** What a typed text means: a number, percentage, date, time, TRUE/FALSE, or just text. */
export function parseInput(input: string, date1904 = false): Typed {
  const text = input.trim();
  if (input === "") return { value: null };
  if (input.startsWith("'")) return { value: input.slice(1) };
  const upper = text.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") return { value: upper === "TRUE" };

  // A number: the computer's decimal sign, optional grouping, %, a currency.
  const dec = escapeRe(locale.decimal);
  const group = locale.group === " " ? "[ \\u00a0\\u202f]" : escapeRe(locale.group);
  const numberRe = new RegExp(`^([-+])?(kr\\.?\\s?|\\$|€|£)?(\\d{1,3}(?:${group}\\d{3})+|\\d+)?(?:${dec}(\\d+))?(?:[eE]([-+]?\\d+))?\\s?(%|\\s?kr\\.?|\\s?€)?$`, "i");
  const m = numberRe.exec(text);
  if (m && (m[3] !== undefined || m[4] !== undefined)) {
    const digits = (m[3] ?? "0").replace(new RegExp(group, "g"), "");
    let n = Number(`${digits}.${m[4] ?? "0"}${m[5] ? `e${m[5]}` : ""}`);
    if (m[1] === "-") n = -n;
    const places = m[4]?.length ?? 0;
    const fixed = places ? `0.${"0".repeat(places)}` : "0";
    if (m[6] === "%") return { value: n / 100, format: places ? `0.${"0".repeat(places)}%` : "0%" };
    const currency = (m[2] ?? m[6] ?? "").trim();
    if (currency) {
      const amount = `#,##${fixed}`;
      return { value: n, format: /kr/i.test(currency) ? `${amount} "kr"` : currency === "€" ? `${amount} €` : `${currency}${amount}` };
    }
    if (m[3] && new RegExp(group).test(m[3])) return { value: n, format: `#,##${fixed}` };
    return { value: n };
  }

  // A date: in the computer's order (day first in Norway), or ISO 2026-10-20.
  const time = /(?:\s+|^)(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(text);
  const datePart = time ? text.slice(0, time.index).trim() : text;
  let date: [number, number, number] | null = null;
  let iso = false;
  let d: RegExpExecArray | null;
  if ((d = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart))) {
    date = [Number(d[1]), Number(d[2]), Number(d[3])];
    iso = true;
  } else if ((d = /^(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2}|\d{4}))?$/.exec(datePart))) {
    const dayFirst = !locale.dateCode.startsWith("m");
    const [a, b] = [Number(d[1]), Number(d[2])];
    let year = d[3] ? Number(d[3]) : new Date().getFullYear();
    if (year < 100) year += year < 30 ? 2000 : 1900;
    date = dayFirst ? [year, b, a] : [year, a, b];
  }
  if (date && date[1] >= 1 && date[1] <= 12 && date[2] >= 1 && date[2] <= 31) {
    const check = new Date(Date.UTC(date[0], date[1] - 1, date[2]));
    if (check.getUTCDate() === date[2]) {
      let serial = dateToSerial(date[0], date[1], date[2], date1904);
      let format = iso ? "yyyy-mm-dd" : locale.dateCode;
      if (time) {
        serial += timeFraction(time);
        format += time[3] ? " hh:mm:ss" : " hh:mm";
      }
      return { value: serial, format };
    }
  }
  if (time && !datePart) {
    const fraction = timeFraction(time);
    if (fraction >= 0) return { value: fraction, format: time[4] ? (time[3] ? "h:mm:ss AM/PM" : "h:mm AM/PM") : time[3] ? "hh:mm:ss" : "hh:mm" };
  }
  return { value: input };
}

function timeFraction(m: RegExpExecArray): number {
  let h = Number(m[1]);
  const pm = m[4]?.toLowerCase() === "pm";
  if (m[4]) h = (h % 12) + (pm ? 12 : 0);
  const min = Number(m[2]);
  const s = Number(m[3] ?? 0);
  if (h > 23 || min > 59 || s > 59) return -1;
  return (h * 3600 + min * 60 + s) / 86400;
}

/** Whether a format shows dates or times (for the date picker, and for sorting labels). */
export function isDateFormat(code: string): boolean {
  const bare = code.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
  return /[dyhs]/i.test(bare) || (/m/i.test(bare) && !/[0#?]/.test(bare));
}

/** How a value is shown in the formula bar while editing: numbers plainly, dates as dates. */
export function editText(value: Scalar, code: string, date1904 = false): string {
  if (value === null) return "";
  if (typeof value === "number") {
    if (isDateFormat(code)) {
      const whole = Math.floor(value);
      const hasTime = value !== whole;
      const datePart = whole ? formatValue(value, locale.dateCode, date1904).text : "";
      const timePart = hasTime ? formatValue(value, "hh:mm:ss", date1904).text : "";
      return [datePart, timePart].filter(Boolean).join(" ");
    }
    if (/%/.test(code.replace(/"[^"]*"/g, ""))) return `${general(value * 100)}%`;
    return general(value);
  }
  if (typeof value === "string") return value;
  return formatValue(value, "General").text;
}
