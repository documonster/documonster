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

describe("validation rules declared as a tagged union", () => {
  const sheet = () => Workbook.addWorksheet(Workbook.create(), "S");

  it("reports an unknown type at rule.type", () => {
    expect(() => Cell.setValidation(sheet(), "A1", { type: "bogus" } as never)).toThrow(
      'Cell.setValidation: rule.type must be one of "any", "list", "whole", "decimal", "date", "time", "textLength", "custom", got "bogus"'
    );
  });

  it("checks the bounds of a rule that has them", () => {
    expect(() => Cell.setValidation(sheet(), "A1", { type: "list", formulae: [NaN] })).toThrow(
      "rule.formulae[0] must be a formula, a finite number or a valid Date, got NaN"
    );
  });

  it("does not read an operator or bounds an `any` rule carries, as no writer writes them", () => {
    expect(() =>
      Cell.setValidation(sheet(), "A1", { type: "any", operator: "x", formulae: [NaN] } as never)
    ).not.toThrow();
  });
});
