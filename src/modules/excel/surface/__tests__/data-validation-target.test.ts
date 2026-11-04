import { Cell, DataValidation, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("DataValidation on a worksheet", () => {
  it("takes the sheet itself, as every other sheet-level namespace does", () => {
    const ws = Workbook.addWorksheet(Workbook.create(), "S");
    const rule = { type: "whole" as const, operator: "between" as const, formulae: [1, 9] };
    DataValidation.add(ws, "A1:A10", rule);
    expect(DataValidation.find(ws, "A3")).toMatchObject(rule);
    expect(Cell.getValidation(ws, "A3")).toMatchObject(rule);
    // The collection form reaches the same rules.
    expect(DataValidation.find(ws.dataValidations, "A3")).toMatchObject(rule);
    expect(DataValidation.remove(ws, "A1:A10")).toBe(true);
    expect(DataValidation.find(ws, "A3")).toBeUndefined();
  });
});
