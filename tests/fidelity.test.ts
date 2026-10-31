// What a workbook made in real Excel showed DuckSheets losing or getting wrong, each put right:
// a file is made here, changed as Excel writes these things, read, saved again and checked.
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { Doc, newWorkbook } from "../src/doc";
import { formatValue } from "../src/numfmt";
import { readXlsx } from "../src/xlsx/read";
import { writeXlsx } from "../src/xlsx/write";

/** A small book saved by DuckSheets, its parts then edited as `edit` says, and read back. */
async function excelMade(edit: (part: (path: string, change: (xml: string) => string) => Promise<void>) => Promise<void>) {
  const doc = new Doc(newWorkbook());
  [["Item", "Cost"], ["Bread", "2"], ["Milk", "3"], ["Tea", "4"]].forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
  doc.addSheet();
  const zip = await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc));
  await edit(async (path, change) => zip.file(path, change(await zip.file(path)!.async("string"))));
  return readXlsx(await zip.generateAsync({ type: "uint8array" }));
}
const saved = async (book: Awaited<ReturnType<typeof readXlsx>>, part = "xl/worksheets/sheet1.xml") => {
  const doc = new Doc(book);
  return (await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc))).file(part)!.async("string");
};

describe("kept as Excel wrote it", () => {
  it("conditional formatting with a data bar's Excel 2010 part (all of it was lost before)", async () => {
    const book = await excelMade((part) =>
      part("xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(
          "</sheetData>",
          '</sheetData><conditionalFormatting sqref="B2:B4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{C3278926-6621-4CB1-B58A-28203D8EF981}</x14:id></ext></extLst></cfRule></conditionalFormatting><conditionalFormatting sqref="A2:A4"><cfRule type="expression" dxfId="0" priority="2"><formula>$A2="Milk"</formula></cfRule></conditionalFormatting>',
        ),
      ),
    );
    const xml = await saved(book);
    expect(xml.match(/<conditionalFormatting /g)?.length).toBe(2);
    expect(xml).toContain('<x14:id>{C3278926-6621-4CB1-B58A-28203D8EF981}</x14:id>');
  });

  it("grouped rows and columns, and a column's style", async () => {
    const book = await excelMade((part) =>
      part("xl/worksheets/sheet1.xml", (xml) =>
        xml
          .replace("<worksheet ", '<worksheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ')
          .replace('<row r="3"', '<row r="3" outlineLevel="1" x14ac:dyDescent="0.25"')
          .replace("<sheetData>", '<cols><col min="2" max="3" width="12" style="1" outlineLevel="1"/></cols><sheetData>'),
      ),
    );
    expect(book.sheets[0].rowAttrs?.get(2)).toBe(' outlineLevel="1"');
    let xml = await saved(book);
    expect(xml).toMatch(/<row r="3" outlineLevel="1">/);
    expect(xml).toMatch(/<col min="2" max="3" width="[\d.]+"( customWidth="1")? style="1" outlineLevel="1"\/>/);
    // A row inserted above moves the grouping along with its row.
    const doc = new Doc(book);
    doc.insertDelete("r", 0, 1);
    xml = await (await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc))).file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).toMatch(/<row r="4" outlineLevel="1">/);
  });

  it("a hidden sheet stays hidden, and can be shown and hidden here", async () => {
    const book = await excelMade((part) => part("xl/workbook.xml", (xml) => xml.replace(/(<sheet name="Sheet2" sheetId="2")/, '$1 state="hidden"')));
    expect(book.sheets[1].state).toBe("hidden");
    expect(await saved(book, "xl/workbook.xml")).toMatch(/<sheet name="Sheet2" sheetId="2" state="hidden"/);
    const doc = new Doc(book);
    doc.setSheetHidden(1, false);
    expect(doc.book.sheets[1].state).toBeUndefined();
    doc.setSheetHidden(0, true);
    expect([doc.book.sheets[0].state, doc.active]).toEqual(["hidden", 1]);
    // The last one shown can't be hidden.
    doc.setSheetHidden(1, true);
    expect(doc.book.sheets[1].state).toBeUndefined();
  });

  it("text with some words bold, while it is the same text", async () => {
    const runs = '<r><rPr><b/><sz val="11"/></rPr><t>Viktig:</t></r><r><rPr><sz val="11"/></rPr><t xml:space="preserve"> betal tilbake</t></r>';
    const book = await excelMade(async (part) => {
      await part("xl/sharedStrings.xml", (xml) => xml.replace(/<si><t>Bread<\/t><\/si>/, `<si>${runs}</si>`));
    });
    expect(book.sheets[0].cells.get(1 * 16384)?.v).toBe("Viktig: betal tilbake");
    expect(await saved(book, "xl/sharedStrings.xml")).toContain(runs);
    const doc = new Doc(book);
    doc.setInput(1, 0, "Something else");
    const xml = await (await JSZip.loadAsync(await writeXlsx(doc.book, doc.calc))).file("xl/sharedStrings.xml")!.async("string");
    expect([xml.includes("<b/>"), xml.includes("Something else")]).toEqual([false, true]);
  });

  it("IFS and SWITCH saved as Excel saved them, so Excel doesn't add @", async () => {
    const book = await excelMade((part) =>
      part("xl/worksheets/sheet1.xml", (xml) => xml.replace(/<c r="B4"[^]*?<\/c>/, '<c r="B4" t="str" cm="1"><f t="array" ref="B4">_xlfn.IFS(B3&gt;2,"big",TRUE,"small")</f><v>big</v></c>')),
    );
    expect(await saved(book)).toContain('<c r="B4" t="str" cm="1"><f t="array" ref="B4">_xlfn.IFS(B3&gt;2,"big",TRUE,"small")</f>');
  });
});

describe("worked out as Excel does", () => {
  it("a table's column leaves out its totals row (the sum was counted twice)", () => {
    // The table part as Excel writes it: starting with its <?xml?> line, a totals row below the data.
    const doc = new Doc(newWorkbook());
    [["Item", "Cost"], ["Bread", "2"], ["Milk", "3"], ["Tea", "4"], ["Total", "=SUBTOTAL(109,B2:B4)"]].forEach((row, r) => row.forEach((v, c) => doc.setInput(r, c, v)));
    doc.sheet.tables.push({ name: "Utgifter", range: { r1: 0, c1: 0, r2: 4, c2: 1 }, style: "TableStyleMedium2", xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Utgifter" displayName="Utgifter" ref="A1:B5" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Cost"/></tableColumns></table>' });
    doc.setInput(6, 0, "=SUM(Utgifter[Cost])");
    doc.setInput(6, 1, "=ROWS(Utgifter[Item])");
    expect([doc.value(6, 0), doc.value(6, 1)]).toEqual([9, 3]);
  });

  it("fractions, as # ?/?, # ??/?? and ?/8", () => {
    const shown = (v: number, code: string) => formatValue(v, code).text.trim();
    expect([shown(0.375, "# ?/?"), shown(2.5, "# ?/?"), shown(1 / 3, "# ??/??"), shown(0.3, "?/8"), shown(3, "# ?/?"), shown(-1.25, "# ?/?")]).toEqual(["3/8", "2 1/2", "1/3", "2/8", "3", "-1 1/4"]);
  });

  it("a General number shortened to fit its column, as Excel shows it", () => {
    const shown = (v: number, chars: number) => formatValue(v, "General", false, chars).text.replace(",", ".");
    expect([shown(0.021743434343, 8), shown(-0.000703333, 8), shown(45.990894731952, 8), shown(0.9999999993, 8), shown(1 / 3, 20), shown(123456789012, 8), shown(42, 8)]).toEqual(["0.021743", "-0.0007", "45.99089", "1", "0.333333333", "1.23E+11", "42"]);
  });
});

describe("drawn as Excel draws them", () => {
  it("Excel's data bars and icon sets, kept in the file as they were", async () => {
    const { Calculator } = await import("../src/calc");
    const { conditionalEffects } = await import("../src/conditional");
    const book = await excelMade((part) =>
      part("xl/worksheets/sheet1.xml", (xml) =>
        xml.replace(
          "</sheetData>",
          '</sheetData><conditionalFormatting sqref="B2:B4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{C3278926-6621-4CB1-B58A-28203D8EF981}</x14:id></ext></extLst></cfRule><cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>',
        ),
      ),
    );
    const fx = conditionalEffects(book.sheets[0], 0, new Calculator(book), []);
    const at = (r: number) => fx.get(r * 16384 + 1);
    expect([at(1)?.bar?.color, at(1)?.bar?.frac, at(3)?.bar?.frac]).toEqual(["#638EC6", 0.1, 1]);
    expect([at(1)?.icon, at(2)?.icon, at(3)?.icon]).toEqual([{ kind: "arrow", level: 0 }, { kind: "arrow", level: 0.5 }, { kind: "arrow", level: 1 }]);
  });
});
