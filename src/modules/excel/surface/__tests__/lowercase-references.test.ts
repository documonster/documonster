/**
 * A lowercase reference is the same reference — `b2` is `B2`, as Excel's UI reads it — at every entry point, and what
 * is stored and written is canonical.
 */
import { Cell, Column, DefinedNames, Range, Workbook, Worksheet } from "@excel/index";
import { normalizeReference } from "@excel/utils/arguments";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

describe("lowercase references", () => {
  it("address the same cells as uppercase ones", async () => {
    const { wb, ws } = sheet();
    Cell.setValue(ws, "b2", 5);
    expect(Cell.getValue(ws, "B2")).toBe(5);
    expect(Cell.has(ws, "b2")).toBe(true);
    Cell.setValue(ws, "$c$3", { formula: "B2*2" });
    expect(Cell.getFormula(ws, "C3")).toBe("B2*2");
    expect(Range.toString(Range.create("d5:e6"))).toBe("D5:E6");
    Worksheet.merge(ws, "f1:g2");
    expect(Worksheet.getMergedRegions(ws)).toEqual([{ top: 1, left: 6, bottom: 2, right: 7 }]);
    Worksheet.setAutoFilter(ws, "a1:c4");
    expect(Worksheet.getAutoFilter(ws)).toBe("A1:C4");
    Worksheet.setAutoFilter(ws, { from: "a1", to: "c4" });
    expect(Worksheet.getAutoFilter(ws)).toEqual({ from: "A1", to: "C4" });
    Worksheet.fillFormula(ws, "k1:k3", "B2");
    expect(Cell.getFormula(ws, "K1")).toBe("B2");
    DefinedNames.add(Workbook.getDefinedNames(wb), "S!$a$1:$b$2", "Block");
    expect(Cell.getNames(ws, "B2")).toEqual(["Block"]);
    expect(Column.getNumber(ws, "c")).toBe(3);

    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Cell.getValue(back, "B2")).toBe(5);
    expect(Worksheet.getAutoFilter(back)).toBe("A1:C4");
    expect(Worksheet.getMergedRegions(back)).toEqual([{ top: 1, left: 6, bottom: 2, right: 7 }]);
  });

  it("normalizeReference capitalises the references and keeps a sheet name's case", () => {
    expect(normalizeReference("b2:c3")).toBe("B2:C3");
    expect(normalizeReference("'my sheet'!a1:$b$2")).toBe("'my sheet'!A1:$B$2");
    expect(normalizeReference("data!a1 c3:d4")).toBe("data!A1 C3:D4");
  });
});
