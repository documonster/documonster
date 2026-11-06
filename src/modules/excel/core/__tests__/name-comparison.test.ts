/**
 * Every place a name is looked up or kept unique compares it as Excel does (`@utils/name-key`). In particular the
 * dotless `ı` and `I` are two names everywhere, while `ß` and `SS` are one defined or table name but two sheet names.
 */
import { Cell, DefinedNames, Table, Workbook } from "@excel";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { describe, expect, it } from "vitest";

function result(formula: string, setup: (wb: Workbook.Handle, ws: Worksheet) => void): unknown {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  setup(wb, ws);
  Cell.setValue(ws, "Z1", { formula });
  calculateFormulas(wb);
  const value = Cell.getResult(ws, "Z1");
  return value !== null && typeof value === "object" && "error" in value ? value.error : value;
}

type Worksheet = ReturnType<typeof Workbook.addWorksheet>;

describe("names compared as Excel compares them", () => {
  it("defined names: ß and SS are one name, ı and I are two", () => {
    const names = (wb: Workbook.Handle) => Workbook.getDefinedNames(wb);
    expect(result("SS+1", wb => DefinedNames.addFormula(names(wb), "ß", "1"))).toBe(2);
    expect(result("I+1", wb => DefinedNames.addFormula(names(wb), "ı", "1"))).toBe("#NAME?");
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    DefinedNames.addFormula(names(wb), "Rate", "1");
    DefinedNames.addFormula(names(wb), "RATE", "2");
    DefinedNames.addFormula(names(wb), "ı", "3");
    DefinedNames.addFormula(names(wb), "I", "4");
    expect(
      DefinedNames.list(names(wb))
        .map(n => n.name)
        .sort()
    ).toEqual(["I", "RATE", "ı"]);
  });

  it("LET parameters: ς and σ are one, ß and SS are two, ı and I are two", () => {
    expect(result("LET(ς,5,σ+1)", () => {})).toBe(6);
    expect(result("LET(ß,5,SS+1)", () => {})).toBe("#NAME?");
    expect(result("LET(ı,5,I+1)", () => {})).toBe("#NAME?");
  });

  it("sheet references: ẞ finds ß, SS does not, I does not find ı", () => {
    const sheet = (name: string) => (wb: Workbook.Handle) =>
      Cell.setValue(Workbook.addWorksheet(wb, name), "A1", 7);
    expect(result("'ẞ'!A1", sheet("ß"))).toBe(7);
    expect(result("'SS'!A1", sheet("ß"))).toBe("#REF!");
    expect(result("'I'!A1", sheet("ı"))).toBe("#REF!");
  });

  it("structured references: SS finds column ß, I does not find ı", () => {
    const table = (column: string) => (_: Workbook.Handle, ws: Worksheet) => {
      Table.add(ws, { name: "T", ref: "A1", columns: [{ name: column }], rows: [[5]] });
    };
    expect(result("SUM(T[SS])", table("ß"))).toBe(5);
    expect(result("SUM(T[I])", table("ı"))).toBe("#REF!");
  });

  it("sheet names: ı and I may coexist, ß and SS may, ß and ẞ may not", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "ı");
    expect(() => Workbook.addWorksheet(wb, "I")).not.toThrow();
    Workbook.addWorksheet(wb, "ß");
    expect(() => Workbook.addWorksheet(wb, "SS")).not.toThrow();
    expect(() => Workbook.addWorksheet(wb, "ẞ")).toThrow();
    expect(() => Workbook.addWorksheet(wb, "ａ")).not.toThrow();
    expect(() => Workbook.addWorksheet(wb, "A")).toThrow();
  });
});
