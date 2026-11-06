/**
 * `=A1#*2` is a dynamic-array formula: it spills the anchor's values doubled, and it is written with the same
 * dynamic-array metadata as any other — `t="array"` over its own cell and `cm` naming `XLDAPR` — in XLSX and XLSB alike.
 * `ANCHORARRAY(A1)*2` is the same formula.
 */
import { Archive } from "@archive";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function book(formula: string) {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)" });
  Cell.setValue(ws, "C1", { formula });
  return { wb, ws };
}

const column = (ws: Parameters<typeof Cell.getResult>[0], col: string) =>
  [1, 2, 3].map(row => {
    const value = Cell.getValue(ws, `${col}${row}`);
    return typeof value === "object" && value !== null && "formula" in value ? value.result : value;
  });

describe.each(["A1#*2", "ANCHORARRAY(A1)*2", "S!A1#*2"])("%s", formula => {
  it("spills the anchor's values, doubled", () => {
    const { wb, ws } = book(formula);
    calculateFormulas(wb);
    expect(column(ws, "C")).toEqual([2, 4, 6]);
  });

  it("is written as a dynamic array to XLSX", async () => {
    const { wb } = book(formula);
    calculateFormulas(wb);
    const files = await Archive.unpackZip(await Workbook.toBuffer(wb));
    const sheet = new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
    expect(sheet).toMatch(
      /<c r="C1"[^>]* cm="1"[^>]*><f t="array" ref="C1">[^<]*ANCHORARRAY\([^<]*<\/f>/
    );
    expect(new TextDecoder().decode(files.get("xl/metadata.xml")!.data)).toContain("XLDAPR");
  });

  it.each(["xlsx", "xlsb"] as const)(
    "reads back from %s as a dynamic array that spills",
    async format => {
      const { wb } = book(formula);
      calculateFormulas(wb);
      const read = await Workbook.read(await Workbook.toBuffer(wb, { format }));
      const ws = Workbook.getWorksheet(read, "S")!;
      expect(Cell.getValue(ws, "C1")).toMatchObject({
        isDynamicArray: true,
        shareType: "array",
        ref: "C1"
      });
      calculateFormulas(read);
      expect(column(ws, "C")).toEqual([2, 4, 6]);
    }
  );
});

describe("a spill reference a parameter takes whole", () => {
  it("computes over the whole spill and does not spill itself", () => {
    const { wb, ws } = book("SUM(A1#*2)");
    Cell.setValue(ws, "D1", { formula: "SUM(ANCHORARRAY(A1)*2)" });
    Cell.setValue(ws, "E1", { formula: "SUM(A1#)" });
    calculateFormulas(wb);
    expect(column(ws, "C")).toEqual([12, null, null]);
    expect(column(ws, "D")).toEqual([12, null, null]);
    expect(column(ws, "E")).toEqual([6, null, null]);
  });
});
