// Links on cells (Excel's hyperlinks) in .xlsx files: web and e-mail
// addresses through the sheet's relationships, places in the workbook
// ("Sheet2!A1") written straight in the sheet.

import type JSZip from "jszip";
import { address, colOf, key, parseAddress, rowOf, type Sheet } from "../model";
import { relsPath } from "./read";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const LINK_TYPE = `${REL}/hyperlink`;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A place in the workbook is written "#Sheet2!A1"; anything else is an address outside it. */
export const isPlace = (link: string) => link.startsWith("#");

export async function readLinks(zip: JSZip, sheetPath: string, ws: Element): Promise<Map<number, string>> {
  const links = new Map<number, string>();
  const list = Array.from(ws.children).find((el) => el.localName === "hyperlinks");
  if (!list) return links;
  // Read as written: outside addresses are not paths in the file.
  const relsText = (await zip.file(relsPath(sheetPath))?.async("string")) ?? "";
  const raw = new Map([...relsText.matchAll(/<Relationship\b[^>]*>/g)].map((m) => [/Id="([^"]*)"/.exec(m[0])?.[1] ?? "", /Target="([^"]*)"/.exec(m[0])?.[1] ?? ""]));
  for (const el of Array.from(list.children)) {
    const ref = (el.getAttribute("ref") ?? "").split(":")[0];
    const at = parseAddress(ref);
    if (!at) continue;
    const id = el.getAttributeNS(REL, "id") ?? el.getAttribute("r:id");
    const location = el.getAttribute("location");
    const target = id && raw.get(id) ? raw.get(id)!.replace(/&amp;/g, "&") : location ? `#${location}` : null;
    if (target) links.set(key(at.r, at.c), target);
  }
  return links;
}

/** Writes the links' relationships (replacing the sheet's old link ones) and returns the sheet's <hyperlinks>, or null. */
export async function writeLinks(zip: JSZip, sheet: Sheet, sheetPath: string): Promise<string | null> {
  const file = relsPath(sheetPath);
  const text = await zip.file(file)?.async("string");
  if (!text && !sheet.links.size) return null;
  const doc = new DOMParser().parseFromString(text ?? `<Relationships xmlns="${PKG_REL}"/>`, "application/xml");
  const root = doc.documentElement;
  for (const el of Array.from(root.children)) if (el.getAttribute("Type") === LINK_TYPE) root.removeChild(el);
  let next = Math.max(0, ...Array.from(root.children).map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1;
  const items = [...sheet.links].sort(([a], [b]) => a - b).map(([k, link]) => {
    const ref = address(rowOf(k), colOf(k));
    if (isPlace(link)) return `<hyperlink ref="${ref}" location="${esc(link.slice(1))}" display="${esc(link.slice(1))}"/>`;
    const id = `rId${next++}`;
    const rel = doc.createElementNS(PKG_REL, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", LINK_TYPE);
    rel.setAttribute("Target", link);
    rel.setAttribute("TargetMode", "External");
    root.appendChild(rel);
    return `<hyperlink ref="${ref}" r:id="${id}"/>`;
  });
  zip.file(file, new XMLSerializer().serializeToString(doc));
  return items.length ? `<hyperlinks>${items.join("")}</hyperlinks>` : null;
}
