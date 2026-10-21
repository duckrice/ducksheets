// Cell notes (Excel's comments) in .xlsx files: the text in a comments part,
// and the yellow boxes Excel shows them in, in a VML drawing linked as the
// sheet's legacy drawing.

import type JSZip from "jszip";
import { address, colOf, key, parseAddress, rowOf, type Sheet } from "../model";
import { readRels, relsPath, resolvePath } from "./read";

export interface Note {
  text: string;
  author: string;
}

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const NOTE_TYPES = { comments: `${REL}/comments`, vml: `${REL}/vmlDrawing` };
const COMMENTS_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A sheet's notes, by cell. */
export async function readNotes(zip: JSZip, sheetPath: string): Promise<Map<number, Note>> {
  const notes = new Map<number, Note>();
  const rel = [...(await readRels(zip, sheetPath)).values()].find((r) => r.type === NOTE_TYPES.comments);
  const text = rel && (await zip.file(rel.path)?.async("string"));
  if (!text) return notes;
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const authors = Array.from(doc.getElementsByTagNameNS("*", "author")).map((a) => a.textContent ?? "");
  for (const c of Array.from(doc.getElementsByTagNameNS("*", "comment"))) {
    const at = parseAddress(c.getAttribute("ref") ?? "");
    if (!at) continue;
    const body = Array.from(c.getElementsByTagNameNS("*", "t"))
      .map((t) => t.textContent ?? "")
      .join("");
    const author = authors[Number(c.getAttribute("authorId") ?? 0)] ?? "";
    // Excel starts a note with its author's name in bold ("Kari:\n"); that's shown separately here.
    notes.set(key(at.r, at.c), { author, text: author && body.startsWith(`${author}:\n`) ? body.slice(author.length + 2) : body });
  }
  return notes;
}

function commentsXml(notes: Map<number, Note>): string {
  const authors = [...new Set([...notes.values()].map((n) => n.author))];
  const list = [...notes]
    .sort(([a], [b]) => a - b)
    .map(([k, n]) => {
      const lead = n.author ? `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${esc(n.author)}:\n</t></r>` : "";
      return `<comment ref="${address(rowOf(k), colOf(k))}" authorId="${authors.indexOf(n.author)}"><text>${lead}<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${esc(n.text)}</t></r></text></comment>`;
    });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${MAIN}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join("")}</authors><commentList>${list.join("")}</commentList></comments>`;
}

function vmlXml(notes: Map<number, Note>): string {
  const shapes = [...notes.keys()].map((k, i) => {
    const r = rowOf(k);
    const c = colOf(k);
    return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 2, ${c + 3}, 15, ${r + 3}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  });
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes.join("")}</xml>`;
}

/**
 * Writes a sheet's notes (or takes away the ones it had). Returns the ID of the sheet's link to
 * the VML drawing, for its <legacyDrawing>; null when that is to go; undefined to leave the sheet as it is.
 */
export async function writeNotes(zip: JSZip, sheet: Sheet, sheetPath: string, override: (path: string, type: string | null) => void): Promise<string | null | undefined> {
  const relsFile = relsPath(sheetPath);
  const relsText = await zip.file(relsFile)?.async("string");
  if (!relsText && !sheet.notes.size) return undefined;
  const doc = new DOMParser().parseFromString(relsText ?? `<Relationships xmlns="${PKG_REL}"/>`, "application/xml");
  const root = doc.documentElement;
  const rels = Array.from(root.children);
  const find = (type: string) => rels.find((el) => el.getAttribute("Type") === type);
  let comments = find(NOTE_TYPES.comments);
  let vml = find(NOTE_TYPES.vml);
  if (!sheet.notes.size) {
    // A VML drawing without notes holds something else (buttons, form controls): it stays.
    if (!comments) return undefined;
    for (const el of [comments, vml]) {
      if (!el) continue;
      const path = resolvePath(sheetPath, el.getAttribute("Target") ?? "");
      zip.remove(path);
      override(path, null);
      root.removeChild(el);
    }
    if (comments || vml) zip.file(relsFile, new XMLSerializer().serializeToString(doc));
    return null;
  }
  let nextId = Math.max(0, ...rels.map((el) => Number((el.getAttribute("Id") ?? "").replace(/\D/g, "")) || 0)) + 1;
  const add = (type: string, path: string) => {
    const el = doc.createElementNS(PKG_REL, "Relationship");
    el.setAttribute("Id", `rId${nextId++}`);
    el.setAttribute("Type", type);
    el.setAttribute("Target", `../${path.replace(/^xl\//, "")}`);
    root.appendChild(el);
    return el;
  };
  const free = (make: (n: number) => string) => {
    let n = 1;
    while (zip.file(make(n))) n++;
    return make(n);
  };
  comments ??= add(NOTE_TYPES.comments, free((n) => `xl/comments${n}.xml`));
  vml ??= add(NOTE_TYPES.vml, free((n) => `xl/drawings/vmlDrawing${n}.vml`));
  const commentsPath = resolvePath(sheetPath, comments.getAttribute("Target") ?? "");
  const vmlPath = resolvePath(sheetPath, vml.getAttribute("Target") ?? "");
  zip.file(commentsPath, commentsXml(sheet.notes));
  zip.file(vmlPath, vmlXml(sheet.notes));
  override(commentsPath, COMMENTS_CONTENT);
  zip.file(relsFile, new XMLSerializer().serializeToString(doc));
  return vml.getAttribute("Id");
}
