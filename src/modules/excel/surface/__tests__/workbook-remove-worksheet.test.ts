import { Cell, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("Workbook.removeWorksheet", () => {
  it("takes a handle as well as an id or a name, and only removes a sheet of its own workbook", () => {
    const wb = Workbook.create();
    const a = Workbook.addWorksheet(wb, "A");
    const b = Workbook.addWorksheet(wb, "B");
    Cell.setValue(b, "A1", { formula: "A!A1" });
    const other = Workbook.addWorksheet(Workbook.create(), "A");

    expect(Workbook.removeWorksheet(wb, other)).toBe(false);
    expect(Workbook.removeWorksheet(wb, a)).toBe(true);
    expect(Workbook.removeWorksheet(wb, a)).toBe(false);
    expect(Workbook.listWorksheets(wb)).toEqual([b]);
    expect(Cell.getFormula(b, "A1")).toBe("#REF!");
  });

  it("is what the deprecated Worksheet.destroy does", () => {
    const wb = Workbook.create();
    const a = Workbook.addWorksheet(wb, "A");
    Worksheet.destroy(a);
    expect(Workbook.listWorksheets(wb)).toEqual([]);
  });
});
