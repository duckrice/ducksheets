import { describe, expect, it } from "vitest";
import { learn, run } from "../src/flashfill";

const fill = (examples: [string[], string][], rows: string[][]) => {
  const program = learn(examples.map(([inputs, output]) => ({ inputs, output })));
  return program && rows.map((r) => run(program, r));
};

describe("Flash Fill", () => {
  it("takes words, initials and parts, with the text between", () => {
    expect(fill([[["Kari Nordmann"], "Nordmann"]], [["Ola Hansen"], ["Per Olav Berg"]])).toEqual(["Hansen", "Berg"]);
    expect(fill([[["Kari Nordmann"], "K. Nordmann"]], [["Ola Hansen"]])).toEqual(["O. Hansen"]);
    expect(fill([[["kari.nordmann@uio.no"], "Kari Nordmann"]], [["ola.hansen@uio.no"]])).toEqual(["Ola Hansen"]);
    expect(fill([[["2024-03-15"], "15.03"]], [["2025-12-01"]])).toEqual(["01.12"]);
  });

  it("joins cells from several columns", () => {
    expect(fill([[["Kari", "Nordmann"], "NORDMANN, Kari"]], [["Ola", "Hansen"]])).toEqual(["HANSEN, Ola"]);
  });

  it("uses a second example to tell first from last", () => {
    // From "Kari Nordmann" alone, "the last word" and "the second word" are the same; with a middle name they aren't.
    expect(fill([[["Kari Nordmann"], "Nordmann"], [["Per Olav Berg"], "Berg"]], [["Anne Lise Moe"]])).toEqual(["Moe"]);
  });

  it("finds nothing when the result isn't made from the cells", () => {
    expect(learn([{ inputs: ["Kari"], output: "Banana" }])).toBeNull();
  });
});

describe("Flash Fill on a sheet", () => {
  it("fills the rest of the column from one example, under a heading, as one step to undo", async () => {
    const { Doc, newWorkbook } = await import("../src/doc");
    const doc = new Doc(newWorkbook());
    [["Name", "Short"], ["Kari Nordmann", "K. Nordmann"], ["Ola Hansen", ""], ["Per Olav Berg", ""]].forEach((row, r) => row.forEach((v, c) => v && doc.setInput(r, c, v)));
    expect(doc.flashFill(1, 1)).toBe(2);
    expect([doc.value(2, 1), doc.value(3, 1)]).toEqual(["O. Hansen", "P. Berg"]);
    doc.undo();
    expect([doc.value(2, 1), doc.value(3, 1)]).toEqual([null, null]);
  });
});
