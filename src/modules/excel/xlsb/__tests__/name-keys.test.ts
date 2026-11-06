/**
 * The XLSB formula encoder finds names, sheets and tables by the keys Excel compares them by (`@utils/name-key`), not
 * by `toUpperCase` or an exact match.
 *
 * `toUpperCase` folds the dotless `ı` to `I`, so with both defined every reference to `I` was written as `ı`. An exact
 * sheet match refused `s!A1` for a sheet named `S` — which the XLSX writer, and Excel, accept.
 */
import { Cell, DefinedNames, Table, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

async function roundTrip(wb: Workbook.Handle): Promise<Workbook.Handle> {
  const copy = Workbook.create();
  await Workbook.read(copy, await Workbook.toBuffer(wb, { format: "xlsb" }));
  return copy;
}

describe("XLSB formula encoding — name keys", () => {
  it("keeps `ı` and `I` apart", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "ı", "1");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "I", "2");
    Cell.setValue(ws, "A1", { formula: "ı*10+I" });
    Cell.setValue(ws, "A2", { formula: "I" });
    Cell.setValue(ws, "A3", { formula: "i" });
    const sheet = Workbook.getWorksheet(await roundTrip(wb), "S")!;
    expect(["A1", "A2", "A3"].map(address => Cell.getFormula(sheet, address))).toEqual([
      "ı*10+I",
      "I",
      "I"
    ]);
  });

  it("finds a name spelt by a different case or a compatibility form", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Maß", "1");
    Cell.setValue(ws, "A1", { formula: "MASS+1" });
    const sheet = Workbook.getWorksheet(await roundTrip(wb), "S")!;
    expect(Cell.getFormula(sheet, "A1")).toBe("Maß+1");
  });

  it.each([
    ["s!A1", "S!A1"],
    ["SUM(s!A1:A2)", "SUM(S!A1:A2)"],
    // A workbook-scoped name qualified by a sheet is the workbook's name; it reads back unqualified.
    ["s!Rate", "Rate"],
    ["'straẞe'!A1", "'Straße'!A1"]
  ])("finds the sheet of %s by the sheet-name key", async (formula, written) => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Workbook.addWorksheet(wb, "Straße");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Rate", "5");
    Cell.setValue(ws, "B1", { formula, result: 3 });
    const sheet = Workbook.getWorksheet(await roundTrip(wb), "S")!;
    expect(Cell.getFormula(sheet, "B1")).toBe(written);
  });

  it("does not equate `ß` with `SS` in a sheet name", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Workbook.addWorksheet(wb, "Straße");
    Cell.setValue(ws, "B1", { formula: "STRASSE!A1", result: 0 });
    await expect(Workbook.toBuffer(wb, { format: "xlsb" })).rejects.toThrow(/STRASSE/);
  });

  it("finds a table and its column whatever their case", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Table.add(ws, { name: "Sales", ref: "A1", columns: [{ name: "Qty" }], rows: [[1], [2]] });
    Cell.setValue(ws, "D1", { formula: "SUM(sales[qty])", result: 3 });
    const sheet = Workbook.getWorksheet(await roundTrip(wb), "S")!;
    expect(Cell.getFormula(sheet, "D1")).toBe("SUM(Sales[Qty])");
  });
});
