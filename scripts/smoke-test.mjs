// Drives the built app over the DevTools protocol (WebView2 remote debugging):
// typing values and formulas, pointing at cells in a formula, the fill
// handle, formats, sorting, filtering, inserting rows, zoom, undo, and saving
// an .xlsx that opens again (and, with EXCEL_CHECK=1, that Excel reads).
//
//   npm run tauri build -- --no-bundle && npm run smoke
//
// The test never touches the clipboard, so it can't disturb what you copied.
import { execFileSync, spawn } from "node:child_process";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "ducksheets.exe");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const work = mkdtempSync(join(tmpdir(), "ducksheets-smoke-"));
const profile = join(work, "profile");
const shot = process.argv.find((a) => a.startsWith("--screenshot="))?.split("=")[1];

let failures = 0;
function check(name, ok, detail = "") {
  if (!ok) failures++;
  console.log(`${ok ? "ok  " : "FAIL"}  ${name}${detail && !ok ? `  (${detail})` : ""}`);
}

const requestClose = (pid) =>
  execFileSync("powershell", [
    "-NoProfile",
    "-Command",
    `Add-Type -Name W -Namespace U -MemberDefinition '[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);'; [void][U.W]::PostMessage((Get-Process -Id ${pid}).MainWindowHandle, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)`,
  ]);

async function session(port, run) {
  const child = spawn(exe, [], {
    env: {
      ...process.env,
      
      DUCKSHEETS_TEST_OFFSCREEN: "1",
      DUCKSHEETS_MULTI_INSTANCE: "1",
      WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`,
      WEBVIEW2_USER_DATA_FOLDER: profile,
    },
    stdio: "ignore",
  });
  let target;
  for (let i = 0; i < 60 && !target; i++) {
    await sleep(500);
    try {
      target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === "page");
    } catch {
      // Not listening yet.
    }
  }
  if (!target) {
    child.kill();
    throw new Error("The app did not start, or remote debugging is unavailable");
  }
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve) => (ws.onopen = resolve));
  let nextId = 1;
  const pending = new Map();
  const errors = [];
  const ipc = [];
  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    if (msg.id && pending.has(msg.id)) {
      pending.get(msg.id)(msg);
      pending.delete(msg.id);
    } else if (msg.method === "Network.requestWillBeSent" && msg.params.request.url.includes("ipc.localhost")) {
      ipc.push(msg.params.request);
    } else if (msg.method === "Runtime.exceptionThrown") {
      const details = msg.params.exceptionDetails;
      errors.push(details.exception?.description ?? details.text);
    } else if (msg.method === "Runtime.consoleAPICalled" && msg.params.type === "error") {
      errors.push(msg.params.args.map((arg) => arg.value ?? arg.description).join(" "));
    }
  };
  ws.onclose = () => {
    for (const resolve of pending.values()) resolve({ result: { result: {} } });
  };
  const send = (method, params = {}) =>
    new Promise((resolve) => {
      const id = nextId++;
      pending.set(id, resolve);
      ws.send(JSON.stringify({ id, method, params }));
    });

  const app = {
    errors,
    send,
    ipc,
    child,
    async evaluate(expression) {
      const { result } = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result?.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? "evaluate failed");
      return result?.result?.value;
    },
    mouse: (type, x, y, extra = {}) => send("Input.dispatchMouseEvent", { type, x, y, button: "left", buttons: type === "mouseReleased" ? 0 : 1, clickCount: 1, ...extra }),
    async click(x, y, modifiers = 0) {
      await app.mouse("mouseMoved", x, y, { buttons: 0, modifiers });
      await app.mouse("mousePressed", x, y, { modifiers });
      await app.mouse("mouseReleased", x, y, { modifiers });
    },
    async drag(x1, y1, x2, y2) {
      await app.mouse("mouseMoved", x1, y1, { buttons: 0 });
      await app.mouse("mousePressed", x1, y1);
      for (let i = 1; i <= 8; i++) await app.mouse("mouseMoved", x1 + ((x2 - x1) * i) / 8, y1 + ((y2 - y1) * i) / 8);
      await app.mouse("mouseReleased", x2, y2);
    },
    async key(name, code, modifiers = 0) {
      await send("Input.dispatchKeyEvent", { type: "keyDown", key: name, code: name, windowsVirtualKeyCode: code, modifiers });
      await send("Input.dispatchKeyEvent", { type: "keyUp", key: name, code: name, windowsVirtualKeyCode: code, modifiers });
    },
    /** Types one character at a time, as a person does, so typing shortcuts run. */
    async type(text) {
      for (const ch of text) {
        if (ch === "\n") await app.key("Enter", 13);
        else await send("Input.insertText", { text: ch });
      }
    },
    center: (selector, index = 0) =>
      app.evaluate(`(() => {
        const el = document.querySelectorAll(${JSON.stringify(selector)})[${index}];
        if (!el) return null;
        el.scrollIntoView({ block: "center" });
        const r = el.getBoundingClientRect();
        return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height];
      })()`),
    async clickOn(selector, index = 0, modifiers = 0) {
      const at = await app.center(selector, index);
      if (!at) throw new Error(`Nothing matches ${selector} [${index}]`);
      await app.click(at[0], at[1], modifiers);
    },
    count: (selector) => app.evaluate(`document.querySelectorAll(${JSON.stringify(selector)}).length`),
    async waitFor(expression, ms = 8000) {
      for (let t = 0; t < ms; t += 150) {
        if (await app.evaluate(expression)) return true;
        await sleep(150);
      }
      return false;
    },
    /** A scroll-wheel turn; `modifiers`: 1 is Alt, 2 is Ctrl. */
    wheel: (x, y, deltaY, modifiers = 0) => send("Input.dispatchMouseEvent", { type: "mouseWheel", x, y, deltaX: 0, deltaY, modifiers }),
    /** Saves a picture of the window. */
    async shot(file) {
      const { result } = await send("Page.captureScreenshot", { format: "png" });
      writeFileSync(file, Buffer.from(result.data, "base64"));
    },
    saved: () => app.waitFor(`document.querySelector(".status-text").textContent === "Saved"`),
  };
  await send("Runtime.enable");
  await send("Network.enable");
  try {
    await app.waitFor(`!!window.__ducksheets`, 15000);
    await sleep(400);
    await run(app);
  } finally {
    ws.close();
    if (child.exitCode === null) child.kill();
    await sleep(500);
  }
}


const CTRL = 2;
const SHIFT = 8;

/** Helpers for the sheet: where a cell is on screen, typing into cells, reading values. */
function sheetHelpers(app) {
  const at = async (r, c) =>
    app.evaluate(`(() => {
      const g = window.__ducksheets.grid;
      g.reveal(${r}, ${c});
      const b = g.cellBox(${r}, ${c});
      const rect = g.el.getBoundingClientRect();
      return [rect.left + b.x + b.w / 2, rect.top + b.y + b.h / 2];
    })()`);
  return {
    at,
    async select(r, c) {
      await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: ${r}, c: ${c}, ar: ${r}, ac: ${c} }; d.emit(); window.__ducksheets.grid.el.focus(); })()`);
    },
    /** Types into the selected cell as a person does: the first key starts editing, the rest go in the box. */
    async type(text, end = "Enter") {
      const [first, ...rest] = [...text];
      await app.send("Input.dispatchKeyEvent", { type: "keyDown", key: first, text: first, unmodifiedText: first });
      await app.send("Input.dispatchKeyEvent", { type: "keyUp", key: first });
      for (const ch of rest) await app.send("Input.insertText", { text: ch });
      if (end === "Enter") await app.key("Enter", 13);
      else if (end === "Tab") await app.key("Tab", 9);
    },
    value: (r, c) => app.evaluate(`(() => { const v = window.__ducksheets.doc.value(${r}, ${c}); return v && typeof v === "object" ? v.code : v; })()`),
    shown: (r, c) => app.evaluate(`window.__ducksheets.doc.shown(${r}, ${c})`),
    formula: (r, c) => app.evaluate(`window.__ducksheets.doc.cell(${r}, ${c})?.f ?? null`),
    sel: () => app.evaluate(`(() => { const s = window.__ducksheets.doc.sel; return [s.r, s.c, s.ar, s.ac]; })()`),
  };
}

const saved = join(work, "budget.xlsx");

await session(9801, async (app) => {
  const s = sheetHelpers(app);
  check("a new workbook opens, ready to type in", (await app.evaluate(`window.__ducksheets.doc.name`)) === "Book1.xlsx");

  // Typing moves on as in Excel: Tab to the right, Enter down (back to the first column typed in).
  await s.select(0, 0);
  await s.type("Item", "Tab");
  await s.type("Cost");
  check("Tab moves right, Enter moves down", JSON.stringify(await s.sel()) === JSON.stringify([1, 1, 1, 1]), JSON.stringify(await s.sel()));
  await s.select(1, 0);
  await s.type("Rent", "Tab");
  await s.type("9500");
  await s.select(2, 0);
  await s.type("Food", "Tab");
  await s.type("1200.5");
  check("numbers are read as numbers", (await s.value(1, 1)) === 9500 && (await s.value(2, 1)) === 1200.5, `${await s.value(1, 1)} ${await s.value(2, 1)}`);

  // A formula, by typing it.
  await s.select(3, 1);
  await s.type("=sum(b2:b3)");
  check("a typed formula works it out", (await s.value(3, 1)) === 10700.5);
  check("and is written as Excel writes it", (await s.formula(3, 1)) === "SUM(B2:B3)");

  // A formula by pointing: = then click a cell.
  await s.select(4, 1);
  await s.type("=", "none");
  const [x, y] = await s.at(1, 1);
  await app.click(x, y);
  check("clicking a cell while typing a formula puts its address in", (await app.evaluate(`document.querySelector(".cell-editor").value`)) === "=B2");
  check("and outlines it in colour", (await app.evaluate(`window.__ducksheets.grid.highlights.length`)) === 1);
  await app.send("Input.insertText", { text: "*2" });
  await app.key("Enter", 13);
  check("pointing makes a working formula", (await s.value(4, 1)) === 19000);

  // Percentages, dates and text are read as Excel reads them.
  await s.select(6, 0);
  await s.type("10%");
  await s.select(7, 0);
  await s.type("2026-10-20");
  check("10% is a percentage", (await s.value(6, 0)) === 0.1 && (await s.shown(6, 0)) === "10%");
  check("a date is a date", (await s.shown(7, 0)) === "2026-10-20" && typeof (await s.value(7, 0)) === "number");

  // The fill handle: 1, 2 dragged down continues 3, 4, 5.
  await s.select(9, 3);
  await s.type("1");
  await s.type("2");
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 9, c: 3, ar: 10, ac: 3 }; d.emit(); })()`);
  const handle = await app.evaluate(`(() => {
    const g = window.__ducksheets.grid;
    const b = g.cellBox(10, 3);
    const rect = g.el.getBoundingClientRect();
    return [rect.left + b.x + b.w - 1, rect.top + b.y + b.h - 1];
  })()`);
  const [, ty] = await s.at(13, 3);
  await app.drag(handle[0], handle[1], handle[0], ty);
  check("the fill handle continues a series", (await s.value(13, 3)) === 5, String(await s.value(13, 3)));

  // Formats from the keyboard and the toolbar.
  await s.select(0, 0);
  await app.key("b", 66, CTRL);
  check("Ctrl+B makes it bold", (await app.evaluate(`window.__ducksheets.doc.book.styles.fonts[window.__ducksheets.doc.book.styles.xfs[window.__ducksheets.doc.cell(0, 0).s].fontId].bold`)) === true);
  await s.select(3, 1);
  await app.evaluate(`(() => { const sel = document.querySelector(".format-select"); sel.value = "Currency"; sel.dispatchEvent(new Event("change")); })()`);
  check("the toolbar sets a currency format", /kr|\$/.test(await s.shown(3, 1)), await s.shown(3, 1));
  check("the status bar sums a selection", await (async () => {
    await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 1, c: 1, ar: 2, ac: 1 }; d.emit(); })()`);
    return /Sum/.test(await app.evaluate(`document.querySelector(".summary").textContent`));
  })());

  // Ctrl+arrow jumps to the edge of the data.
  await s.select(1, 0);
  await app.key("ArrowDown", 40, CTRL);
  check("Ctrl+Down jumps to the end of the data", JSON.stringify((await s.sel()).slice(0, 2)) === JSON.stringify([2, 0]), JSON.stringify(await s.sel()));

  // Sorting the list (selected, as the totals under it touch it), largest first.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 0, c: 1, ar: 2, ac: 0 }; d.emit(); })()`);
  await app.clickOn(`.tb[title^="Sort Z to A"]`);
  check("sorting keeps the heading and puts the largest first", (await s.value(0, 0)) === "Item" && (await s.value(1, 0)) === "Rent", `${await s.value(1, 0)}`);

  // Inserting a row moves the formulas' references.
  await s.select(2, 0);
  await app.key(" ", 32, SHIFT);
  await app.key("+", 187, CTRL);
  check("a row inserted above keeps the sum right", (await s.formula(4, 1)) === "SUM(B2:B4)" && (await s.value(4, 1)) === 10700.5, `${await s.formula(4, 1)}`);
  await app.key("z", 90, CTRL);
  check("and Ctrl+Z takes it away again", (await s.formula(3, 1)) === "SUM(B2:B3)");

  // Filtering from the heading's button.
  await s.select(0, 0);
  await app.key("L", 76, CTRL | SHIFT);
  check("Ctrl+Shift+L turns on filtering", (await app.evaluate(`!!window.__ducksheets.doc.sheet.filter`)) === true);
  const button = await app.evaluate(`(() => {
    const g = window.__ducksheets.grid;
    const b = g.cellBox(0, 0);
    const rect = g.el.getBoundingClientRect();
    return [rect.left + b.x + b.w - 9, rect.top + b.y + b.h - 9];
  })()`);
  await app.click(button[0], button[1]);
  check("its button lists the values", await app.waitFor(`document.querySelectorAll(".filter-item").length >= 3`, 3000));
  await app.evaluate(`(() => { const item = [...document.querySelectorAll(".filter-item")].find((l) => l.textContent === "Food"); item.querySelector("input").click(); [...document.querySelectorAll(".filter-menu .btn")].find((b) => b.textContent === "OK").click(); })()`);
  check("unticking a value hides its rows", (await app.evaluate(`[...window.__ducksheets.doc.sheet.hiddenRows]`)).length === 1);
  // Custom filter from the Cost button: less than 5000 shows Food, hides Rent.
  await app.evaluate(`window.__ducksheets.doc.setFilter(0, null)`);
  const costButton = await app.evaluate(`(() => {
    const g = window.__ducksheets.grid;
    const b = g.cellBox(0, 1);
    const rect = g.el.getBoundingClientRect();
    return [rect.left + b.x + b.w - 9, rect.top + b.y + b.h - 9];
  })()`);
  await app.click(costButton[0], costButton[1]);
  await app.waitFor(`!!document.querySelector(".filter-menu")`, 3000);
  await app.evaluate(`[...document.querySelectorAll(".filter-menu .menu-item")].find((b) => b.textContent === "Custom filter…").click()`);
  check("Custom filter… asks for a condition", await app.waitFor(`!!document.querySelector('[aria-label="Custom filter"] select')`, 3000));
  await app.evaluate(`(() => { const d = document.querySelector('[aria-label="Custom filter"]'); d.querySelector("select").value = "lessThan"; d.querySelector("input").value = "5000"; d.querySelector(".btn.primary").click(); })()`);
  const hidden = await app.evaluate(`[...window.__ducksheets.doc.sheet.hiddenRows].sort().join()`);
  check("and hides the rows that don't match", hidden.split(",").includes("1") && !hidden.split(",").includes("2"), hidden);
  await app.key("L", 76, CTRL | SHIFT);

  // Sort… from Data tools: by Item, A to Z.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 0, c: 0, ar: 2, ac: 1 }; d.emit(); })()`);
  await app.clickOn('.toolbar [title^="Data tools"]');
  await app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.textContent === "Sort…").click()`);
  check("Sort… opens with the headings to choose from", await app.waitFor(`[...document.querySelectorAll('[aria-label="Sort by column"] option')].map((o) => o.textContent).join() === "Item,Cost"`, 3000));
  await app.evaluate(`(() => { const d = document.querySelector('[role="dialog"][aria-label="Sort"]'); d.querySelector('[aria-label="Sort by column"]').value = "0"; d.querySelector(".btn.primary").click(); })()`);
  check("and sorts by that column", (await s.value(1, 0)) === "Food" && (await s.value(2, 0)) === "Rent", `${await s.value(1, 0)}`);
  await app.key("z", 90, CTRL);

  // Zoom with Ctrl+wheel.
  const [gx, gy] = await s.at(5, 5);
  await app.wheel(gx, gy, -120, CTRL);
  check("Ctrl+wheel zooms", (await app.evaluate(`window.__ducksheets.grid.zoom`)) > 1);

  // A chart of the list.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 0, c: 0, ar: 2, ac: 1 }; d.emit(); })()`);
  await app.clickOn(`.tb[title^="Insert a chart"]`);
  check("the chart button draws a chart of the selection", await app.waitFor(`document.querySelectorAll(".chart svg rect").length >= 3`, 3000));
  const chartBox = await app.center(".chart");
  await app.drag(chartBox[0], chartBox[1], chartBox[0] + 60, chartBox[1] + 40);
  check("a chart can be dragged to another place", await app.evaluate(`(() => { const at = window.__ducksheets.doc.sheet.charts[0].at; return at.c !== 3 || at.dx > 40; })()`), JSON.stringify(await app.evaluate(`window.__ducksheets.doc.sheet.charts[0].at`)));

  // The chart's menu (right-click): a colour for its series, and axis titles.
  const chartMenu = () => app.evaluate(`(() => { const b = document.querySelector(".chart"); const r = b.getBoundingClientRect(); b.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: r.left + 20, clientY: r.top + 20 })); })()`);
  const pick = (label) => app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.textContent.trim().startsWith(${JSON.stringify(label)})).click()`);
  await chartMenu();
  await app.waitFor(`!!document.querySelector(".menu")`, 2000);
  await pick("Colours");
  await pick("Cost");
  await app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.querySelector(".menu-swatch span")?.style.background === "rgb(192, 0, 0)").click()`);
  check("a series can get another colour", await app.waitFor(`[...document.querySelectorAll(".chart svg rect")].some((r) => r.getAttribute("fill") === "#C00000")`, 3000));
  await chartMenu();
  await app.waitFor(`!!document.querySelector(".menu")`, 2000);
  await pick("Axis titles");
  await app.evaluate(`(() => { const d = document.querySelector('[aria-label="Axis titles"]'); d.querySelector('[aria-label="Vertical axis title"]').value = "Kroner"; d.querySelector(".btn.primary").click(); })()`);
  check("and axis titles", await app.waitFor(`[...document.querySelectorAll(".chart svg text")].some((t) => t.textContent === "Kroner")`, 3000) && (await app.evaluate(`window.__ducksheets.doc.sheet.charts[0].xTitle === undefined`)));

  // Saving an .xlsx file, and opening it again.
  await app.evaluate(`window.__ducksheets.save(${JSON.stringify(saved)})`);
  check("it saves an Excel file", await app.waitFor(`document.querySelector(".status-text").textContent === "Saved budget.xlsx"`, 8000));
  const zip = await JSZip.loadAsync(readFileSync(saved));
  const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
  check("with the formulas and their values in it", /<f>SUM\(B2:B3\)<\/f><v>10700.5<\/v>/.test(sheetXml), sheetXml.slice(0, 300));
  await app.evaluate(`window.__ducksheets.load(${JSON.stringify(saved)})`);
  check("and opens it again as it was", (await app.waitFor(`window.__ducksheets.doc.name === "budget.xlsx"`)) && (await s.value(3, 1)) === 10700.5 && (await s.formula(4, 1)) === "B2*2");
  check("with its chart", (await app.evaluate(`window.__ducksheets.doc.sheet.charts.length`)) === 1 && (await app.waitFor(`!!document.querySelector(".chart svg")`, 3000)));
  if (shot) {
    await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 3, c: 1, ar: 3, ac: 1 }; d.emit(); window.__ducksheets.grid.setZoom(1); })()`);
    await app.shot(shot);
  }
  // As in Excel: click a cell, then just type, even after using the toolbar; the cell keeps its colours while typing.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.format({ r1: 15, c1: 3, r2: 15, c2: 3 }, { fill: "#ffff00", color: "#c00000" }); })()`);
  const [cx, cy] = await s.at(15, 3);
  await app.click(cx, cy);
  await app.clickOn('.toolbar [title^="Bold"]');
  await app.clickOn('.toolbar [title^="Bold"]');
  await app.send("Input.dispatchKeyEvent", { type: "keyDown", key: "7", text: "7", unmodifiedText: "7" });
  await app.send("Input.dispatchKeyEvent", { type: "keyUp", key: "7" });
  const typing = await app.evaluate(`(() => { const b = document.querySelector(".cell-editor"); const st = getComputedStyle(b); return [b.hidden, b.value, st.backgroundColor, st.color, document.activeElement === b].join("|"); })()`);
  check("clicking a cell and typing writes into it, even after clicking the toolbar", typing.startsWith("false|7|"), typing);
  check("in the cell's own colours", typing.includes("rgb(255, 255, 0)|rgb(192, 0, 0)|true"), typing);
  await app.key("Enter", 13);
  check("and Enter puts it in", (await s.value(15, 3)) === 7);

  // Conditional formatting from the toolbar: Highlight cells rules → Greater than…
  for (const [r, v] of [[30, "5"], [31, "50"], [32, "500"]]) {
    await s.select(r, 0);
    await s.type(v);
  }
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 30, c: 0, ar: 32, ac: 0 }; d.emit(); })()`);
  await app.clickOn('.toolbar [title="Conditional formatting"]');
  await app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.textContent.startsWith("Highlight cells rules")).click()`);
  await app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.textContent.startsWith("Greater than")).click()`);
  check("Greater than… asks for the value", await app.waitFor(`!!document.querySelector('.cf-dialog input')`, 3000));
  await app.evaluate(`(() => { const i = document.querySelector('.cf-dialog input'); i.value = "40"; [...document.querySelectorAll(".cf-dialog .btn")].find((b) => b.textContent === "OK").click(); })()`);
  const fills = await app.evaluate(`[30, 31, 32].map((r) => window.__ducksheets.grid.effects.get(r * 16384)?.fill ?? "-").join()`);
  check("and colours the cells above it", fills === "-,#FFC7CE,#FFC7CE", fills);

  // Data validation: a list shows a drop-down beside the cell; a value not in it isn't put in.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.sel = { r: 40, c: 1, ar: 40, ac: 1 }; d.setValidation({ r1: 40, c1: 1, r2: 45, c2: 1 }, { ranges: [], t: "list", f: '"Yes,No"' }); window.__ducksheets.grid.el.focus(); })()`);
  check("a cell with a list shows a drop-down button", await app.waitFor(`!document.querySelector(".list-pick").hidden`, 2000));
  await app.clickOn(".list-pick");
  await app.evaluate(`[...document.querySelectorAll(".menu-item")].find((b) => b.textContent === "No").click()`);
  check("choosing from it fills the cell", (await s.value(40, 1)) === "No");
  await s.select(41, 1);
  await s.type("Maybe");
  check("a value not in the list is refused, with a message", (await app.waitFor(`document.querySelector('[role="alertdialog"] h2')?.textContent.includes("doesn't match")`, 3000)) && (await s.value(41, 1)) === null);
  await app.evaluate(`[...document.querySelectorAll('[role="alertdialog"] .btn')].find((b) => b.textContent === "Cancel").click()`);

  // Format cells (Ctrl+1): a percentage with one decimal.
  await s.select(31, 0);
  await app.key("1", 49, 2);
  check("Ctrl+1 opens Format cells", await app.waitFor(`!!document.querySelector(".format-dialog")`, 3000));
  await app.evaluate(`(() => { const sel = document.querySelector('.format-dialog select[aria-label="Category"]'); sel.value = "Percentage"; sel.dispatchEvent(new Event("change")); const d = document.querySelector('.format-dialog input[aria-label="Decimal places"]'); d.value = "1"; d.dispatchEvent(new Event("input")); [...document.querySelectorAll(".format-dialog .btn")].find((b) => b.textContent === "OK").click(); })()`);
  check("and sets the number format", (await s.shown(31, 0)) === "5000.0%", await s.shown(31, 0));

  // AutoComplete: typing the start of a word from the column above finishes it.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.setInput(50, 0, "Bergen"); d.setInput(51, 0, "Oslo"); d.setInput(52, 0, "Trondheim"); })()`);
  await s.select(53, 0);
  await s.type("Os", "none");
  const completed = await app.evaluate(`(() => { const b = document.querySelector(".cell-editor"); return [b.value, b.selectionStart, b.selectionEnd].join(); })()`);
  check("typing the start of a word above finishes it, the rest selected", completed === "Oslo,2,4", completed);
  await app.key("Enter", 13);
  check("and Enter puts the whole word in", (await s.value(53, 0)) === "Oslo");

  // Format painter: copy bold from one cell, click another.
  await app.evaluate(`(() => { const d = window.__ducksheets.doc; d.format({ r1: 50, c1: 0, r2: 50, c2: 0 }, { bold: true }); d.sel = { r: 50, c: 0, ar: 50, ac: 0 }; d.emit(); })()`);
  await app.clickOn('.toolbar [title^="Format painter"]');
  const [px, py] = await s.at(52, 0);
  await app.click(px, py);
  check("the format painter gives the formats to the cell clicked", await app.waitFor(`window.__ducksheets.doc.book.styles.fonts[window.__ducksheets.doc.book.styles.xfs[window.__ducksheets.doc.cell(52, 0)?.s ?? 0].fontId]?.bold === true`, 3000));

  // Notes: Shift+F2 writes one; the mouse over the cell shows it.
  await s.select(60, 1);
  await app.key("F2", 113, 8);
  check("Shift+F2 asks for a note", await app.waitFor(`!!document.querySelector(".note-input")`, 3000));
  await app.evaluate(`(() => { document.querySelector(".note-input").value = "Remember the deposit"; [...document.querySelectorAll('[aria-label="Note"] .btn')].find((b) => b.textContent === "OK").click(); })()`);
  check("and puts it on the cell", (await app.evaluate(`window.__ducksheets.doc.sheet.notes.get(60 * 16384 + 1)?.text`)) === "Remember the deposit");
  const [nx, ny] = await s.at(60, 1);
  await app.mouse("mouseMoved", nx, ny, { buttons: 0 });
  check("the mouse over the cell shows the note", await app.waitFor(`(() => { const t = document.querySelector(".note-tip"); return !t.hidden && t.textContent.includes("Remember the deposit"); })()`, 2000));

  // Links: Ctrl+K puts one on a cell; a click on the cell follows it.
  await s.select(62, 1);
  await app.key("k", 75, 2);
  check("Ctrl+K asks for a link", await app.waitFor(`!!document.querySelector('[aria-label="Address"]')`, 3000));
  await app.evaluate(`(() => { document.querySelector('[aria-label="Address"]').value = "A100"; document.querySelector('[aria-label="Text to show"]').value = "Go down"; [...document.querySelectorAll('[aria-label="Link"] .btn')].find((b) => b.textContent === "OK").click(); })()`);
  check("and puts it on the cell, with its text", (await s.value(62, 1)) === "Go down" && (await app.evaluate(`window.__ducksheets.doc.sheet.links.get(62 * 16384 + 1)`)) === "#A100");
  await s.select(70, 3);
  const [lx, ly] = await s.at(62, 1);
  await app.click(lx, ly);
  check("clicking a link goes where it points", await app.waitFor(`window.__ducksheets.doc.sel.r === 99 && window.__ducksheets.doc.sel.c === 0`, 3000), JSON.stringify(await s.sel()));

  // Ctrl+` shows formulas.
  await app.send("Input.dispatchKeyEvent", { type: "keyDown", key: "`", code: "Backquote", windowsVirtualKeyCode: 192, modifiers: 2 });
  await app.send("Input.dispatchKeyEvent", { type: "keyUp", key: "`", code: "Backquote", windowsVirtualKeyCode: 192, modifiers: 2 });
  check("Ctrl+` shows the formulas", (await app.evaluate(`window.__ducksheets.grid.showFormulas`)) === true);
  await app.send("Input.dispatchKeyEvent", { type: "keyDown", key: "`", code: "Backquote", windowsVirtualKeyCode: 192, modifiers: 2 });
  await app.send("Input.dispatchKeyEvent", { type: "keyUp", key: "`", code: "Backquote", windowsVirtualKeyCode: 192, modifiers: 2 });

  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));

  // Closing with changes asks whether to save.
  await s.select(20, 0);
  await s.type("changed");
  requestClose(app.child.pid);
  check("closing with changes asks whether to save", await app.waitFor(`!!document.querySelector(".dialog h2")?.textContent.startsWith("Save the changes")`, 5000));
  await app.evaluate(`[...document.querySelectorAll(".dialog .btn")].find((b) => b.textContent === "Don't save").click()`);
  for (let i = 0; i < 40 && app.child.exitCode === null; i++) await sleep(250);
  check("and closes when told not to", app.child.exitCode !== null);
});

// Excel itself reads the saved file the same way, where it's installed.
if (process.env.EXCEL_CHECK) {
  const script = `$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try { $wb = $xl.Workbooks.Open('${saved}'); $s = $wb.Worksheets.Item(1); @{ total = $s.Range('B4').Value2; formula = $s.Range('B5').Formula } | ConvertTo-Json -Compress; $wb.Close($false) } finally { $xl.Quit() }`;
  const report = JSON.parse(execFileSync("powershell", ["-NoProfile", "-Command", script], { encoding: "utf8" }));
  check("Excel opens the saved file with the same results", report.total === 10700.5 && report.formula === "=B2*2", JSON.stringify(report));
}

console.log(failures ? `\n${failures} check(s) failed` : "\nAll checks passed");
process.exit(failures ? 1 : 0);
