// Real mouse clicks on cells, then keys, in many situations: typing must always reach
// the cell clicked (a click while typing elsewhere finishes that typing first, as in Excel).
//   npm run tauri build -- --no-bundle && node scripts/clicks-test.mjs
import { spawn } from "node:child_process";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const child = spawn("src-tauri/target/release/ducksheets.exe", [], { env: { ...process.env, DUCKSHEETS_TEST_OFFSCREEN: "1", DUCKSHEETS_MULTI_INSTANCE: "1", WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: "--remote-debugging-port=9893", WEBVIEW2_USER_DATA_FOLDER: mkdtempSync(join(tmpdir(), "ds-clicks-")) }, stdio: "ignore" });
let t;
for (let i = 0; i < 60 && !t; i++) {
  await sleep(500);
  try {
    t = (await (await fetch("http://127.0.0.1:9893/json")).json()).find((x) => x.type === "page");
  } catch {}
}
const ws = new WebSocket(t.webSocketDebuggerUrl);
await new Promise((r) => (ws.onopen = r));
let id = 1;
const p = new Map();
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (p.has(m.id)) p.get(m.id)(m);
};
const send = (method, params = {}) => new Promise((r) => { const i = id++; p.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
const ev = async (expression) => (await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true })).result?.result?.value;
await sleep(2500);
const at = (r, c) => ev(`(() => { const g = window.__ducksheets.grid; const b = g.cellBox(${r}, ${c}); const rect = g.el.getBoundingClientRect(); return [rect.left + b.x + b.w / 2, rect.top + b.y + b.h / 2]; })()`);
const click = async (r, c, count = 1) => {
  const [x, y] = await at(r, c);
  await send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y, button: "none" });
  for (let i = 1; i <= count; i++) {
    await send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", buttons: 1, clickCount: i });
    await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", buttons: 0, clickCount: i });
  }
  await sleep(150);
};
const char = async (ch) => {
  await send("Input.dispatchKeyEvent", { type: "keyDown", key: ch, text: ch, unmodifiedText: ch, code: `Key${ch.toUpperCase()}`, windowsVirtualKeyCode: ch.toUpperCase().charCodeAt(0) });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key: ch, code: `Key${ch.toUpperCase()}`, windowsVirtualKeyCode: ch.toUpperCase().charCodeAt(0) });
};
const key = async (name, code) => {
  await send("Input.dispatchKeyEvent", { type: "rawKeyDown", key: name, code: name, windowsVirtualKeyCode: code });
  await send("Input.dispatchKeyEvent", { type: "keyUp", key: name, code: name, windowsVirtualKeyCode: code });
  await sleep(100);
};
const state = () => ev(`(() => { const d = window.__ducksheets.doc; const b = document.querySelector(".cell-editor"); return { sel: [d.sel.r, d.sel.c], editing: !b.hidden, box: b.value, focus: document.activeElement?.className || document.activeElement?.tagName }; })()`);
const value = (r, c) => ev(`window.__ducksheets.doc.value(${r}, ${c})`);

const results = [];
const test = async (name, fn) => {
  try {
    results.push([name, await fn()]);
  } catch (e) {
    results.push([name, `ERROR ${e}`]);
  }
};

await test("1 click a cell, type, Enter", async () => {
  await click(2, 1);
  await char("a");
  await key("Enter", 13);
  return { ok: (await value(2, 1)) === "a", ...(await state()) };
});
await test("2 click a cell, arrow key moves", async () => {
  await click(5, 1);
  await key("ArrowDown", 40);
  const s = await state();
  return { ok: s.sel[0] === 6 && s.sel[1] === 1, ...s };
});
await test("3 typing in one cell, click another, type", async () => {
  await click(8, 1);
  await char("q");
  await click(9, 2);
  const mid = await state();
  await char("w");
  await key("Enter", 13);
  return { ok: (await value(9, 2)) === "w" && (await value(8, 1)) === "q", mid, ...(await state()) };
});
await test("4 double-click a cell to edit, click another, type", async () => {
  await click(12, 1, 2);
  const mid = await state();
  await click(13, 3);
  await char("z");
  await key("Enter", 13);
  return { ok: (await value(13, 3)) === "z", mid, ...(await state()) };
});
await test("5 click the formula bar, then a cell, type", async () => {
  const bar = await ev(`(() => { const r = document.querySelector(".bar-input").getBoundingClientRect(); return [r.left + 20, r.top + r.height / 2]; })()`);
  await send("Input.dispatchMouseEvent", { type: "mousePressed", x: bar[0], y: bar[1], button: "left", buttons: 1, clickCount: 1 });
  await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: bar[0], y: bar[1], button: "left", buttons: 0, clickCount: 1 });
  await sleep(150);
  const mid = await state();
  await click(16, 1);
  await char("k");
  await key("Enter", 13);
  return { ok: (await value(16, 1)) === "k", mid, ...(await state()) };
});
await test("6 click the name box, then a cell, type", async () => {
  const box = await ev(`(() => { const r = document.querySelector(".name-box").getBoundingClientRect(); return [r.left + 10, r.top + r.height / 2]; })()`);
  await send("Input.dispatchMouseEvent", { type: "mousePressed", x: box[0], y: box[1], button: "left", buttons: 1, clickCount: 1 });
  await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: box[0], y: box[1], button: "left", buttons: 0, clickCount: 1 });
  await sleep(150);
  await click(18, 1);
  await char("m");
  await key("Enter", 13);
  return { ok: (await value(18, 1)) === "m", ...(await state()) };
});
await test("7 drag over cells, then type", async () => {
  const [x1, y1] = await at(20, 1);
  const [x2, y2] = await at(22, 2);
  await send("Input.dispatchMouseEvent", { type: "mousePressed", x: x1, y: y1, button: "left", buttons: 1, clickCount: 1 });
  for (let i = 1; i <= 5; i++) await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: x1 + ((x2 - x1) * i) / 5, y: y1 + ((y2 - y1) * i) / 5, button: "left", buttons: 1 });
  await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: x2, y: y2, button: "left", buttons: 0, clickCount: 1 });
  await sleep(150);
  await char("d");
  await key("Enter", 13);
  return { ok: (await value(20, 1)) === "d", ...(await state()) };
});
await test("8 click a sheet tab area then a cell", async () => {
  const tab = await ev(`(() => { const r = document.querySelector(".sheet-tab").getBoundingClientRect(); return [r.left + 10, r.top + r.height / 2]; })()`);
  await send("Input.dispatchMouseEvent", { type: "mousePressed", x: tab[0], y: tab[1], button: "left", buttons: 1, clickCount: 1 });
  await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: tab[0], y: tab[1], button: "left", buttons: 0, clickCount: 1 });
  await sleep(150);
  await click(24, 1);
  await char("t");
  await key("Enter", 13);
  return { ok: (await value(24, 1)) === "t", ...(await state()) };
});
await test("9 type a formula, click a cell (points), Enter, then click another and type", async () => {
  await click(26, 1);
  await char("=");
  await click(2, 1);
  await key("Enter", 13);
  await click(27, 2);
  await char("p");
  await key("Enter", 13);
  return { ok: (await value(27, 2)) === "p", formula: await ev(`window.__ducksheets.doc.cell(26, 1)?.f`), ...(await state()) };
});

for (const [name, r] of results) console.log(r?.ok ? "ok  " : "FAIL", name, JSON.stringify(r));
child.kill();
const failed = results.filter(([, r]) => !r?.ok).length;
console.log(failed ? `\n${failed} check(s) failed` : "\nAll checks passed");
process.exit(failed ? 1 : 0);
