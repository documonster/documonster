/**
 * A formula is often given as typed into a cell, `=1+1`. A file stores it without the `=`, so every writer drops one
 * leading `=`; the in-memory value keeps what the caller gave.
 */
import { extractAll } from "@archive/unzip/extract";
import { Cell, DefinedNames, Workbook } from "@excel";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { describe, expect, it } from "vitest";

function build() {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "S");
  Cell.setValue(sheet, "A1", { formula: "=1+1" });
  Cell.setValue(sheet, "A2", { formula: "=SUM(A1,1)", result: 3 });
  Cell.setValidation(sheet, "B1", { type: "whole", operator: "greaterThan", formulae: ["=A1"] });
  DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Two", "=1+1");
  return workbook;
}

describe("a formula given with a leading =", () => {
  it("is stored without it (XLSX)", async () => {
    const workbook = build();
    const files = await extractAll(await Workbook.toBuffer(workbook));
    const sheet = new TextDecoder().decode(files.get("xl/worksheets/sheet1.xml")!.data);
    expect(sheet).toContain("<f>1+1</f>");
    expect(sheet).toContain("<f>SUM(A1,1)</f>");
    expect(sheet).toContain("<formula1>A1</formula1>");
    expect(sheet).not.toContain(">=");
    const book = new TextDecoder().decode(files.get("xl/workbook.xml")!.data);
    expect(book).toContain(">1+1</definedName>");
    expect(Cell.getFormula(Workbook.getWorksheet(workbook, "S")!, "A1")).toBe("=1+1");
  });

  it.each(["xlsx", "xlsb"] as const)("reads back without it (%s)", async format => {
    const reread = Workbook.create();
    await Workbook.read(reread, await Workbook.toBuffer(build(), { format }));
    const sheet = Workbook.getWorksheet(reread, "S")!;
    expect(Cell.getFormula(sheet, "A1")).toBe("1+1");
    expect(Cell.getFormula(sheet, "A2")).toBe("SUM(A1,1)");
  });
});

describe("a defined name given with a leading =", () => {
  it.each([
    ["=S!$A$1:$A$3", "SUM(Rng)", 6],
    ["=S!$A$1", "Rng*2", 2],
    ["=1+1", "Rng*3", 6]
  ])("%s is computed as the name it is", (expression, formula, expected) => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    [1, 2, 3].forEach((value, index) => Cell.setValue(sheet, `A${index + 1}`, value));
    DefinedNames.addFormula(Workbook.getDefinedNames(workbook), "Rng", expression);
    Cell.setValue(sheet, "C1", { formula: formula });
    calculateFormulas(workbook);
    expect(Cell.getResult(sheet, "C1")).toBe(expected);
  });
});
