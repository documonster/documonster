import {
  Cell,
  ExcelDefinedNameError,
  DefinedNames,
  ExcelArgumentError,
  Row,
  Table,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

describe("Cell.removeName / clearNames", () => {
  it("reports whether the cell carried the name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.addName(ws, "B2", "Here");
    expect(Cell.removeName(ws, "B2", "Nope")).toBe(false);
    expect(Cell.removeName(ws, "C3", "Here")).toBe(false);
    expect(Cell.removeName(ws, 2, 2, "Here")).toBe(true);
    expect(Cell.removeName(ws, "B2", "Here")).toBe(false);
    expect(Cell.getNames(ws, "B2")).toEqual([]);
  });

  it("creates neither the cell nor the row", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.removeName(ws, "Q50", "Here");
    Cell.clearNames(ws, "Q60");
    expect(Row.getCount(ws)).toBe(1);
    expect(Cell.has(ws, "Q50")).toBe(false);
  });

  it("leaves a formula-defined name of the same name in place", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Rate", "0.05");
    expect(Cell.removeName(ws, "A1", "Rate")).toBe(false);

    const back = Workbook.create();
    await Workbook.read(back, await Workbook.toBuffer(wb));
    expect(DefinedNames.list(Workbook.getDefinedNames(back)).map(n => n.name)).toContain("Rate");
  });
});

describe("a name given through Cell is checked as DefinedNames.add checks it", () => {
  it.each<[string, (ws: ReturnType<typeof Workbook.addWorksheet>) => unknown]>([
    ["addName reading as a cell reference", ws => Cell.addName(ws, "B1", "A1")],
    ["addName starting with a digit", ws => Cell.addName(ws, "B1", "1abc")],
    ["setName holding a space", ws => Cell.setName(ws, "B1", "has space")],
    ["addName that is not a string", ws => Cell.addName(ws, "B1", 42 as never)],
    ["setNames with one bad name", ws => Cell.setNames(ws, "B1", ["Good", 1 as never])],
    ["addName of a table's name", ws => Cell.addName(ws, "B1", "sales")]
  ])("%s is refused and changes nothing", (_name, edit) => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.addMany(ws, [["h"], [1]]);
    Table.add(ws, {
      name: "Sales",
      ref: "D1",
      headerRow: true,
      columns: [{ name: "h" }],
      rows: [[1]]
    });
    Cell.setNames(ws, "Z9", ["Keep1", "Keep2"]);
    expect(() => edit(ws)).toThrow(ExcelDefinedNameError);
    expect(Cell.has(ws, "B1")).toBe(false);
    expect(Cell.getNames(ws, "Z9")).toEqual(["Keep1", "Keep2"]);
    expect(DefinedNames.list(Workbook.getDefinedNames(wb)).map(n => n.name)).toEqual([
      "Keep1",
      "Keep2"
    ]);
  });

  it("setNames that is not an array is an argument error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => Cell.setNames(ws, "B1", "Good" as never)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "B1")).toBe(false);
  });

  it("setNames refused part-way keeps the names the cell had", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setNames(ws, "A7", ["Keep1", "Keep2"]);
    expect(() => Cell.setNames(ws, "A7", ["New", "1bad"])).toThrow(ExcelDefinedNameError);
    expect(Cell.getNames(ws, "A7")).toEqual(["Keep1", "Keep2"]);
  });

  it("naming a cell with a formula-defined name turns it into a reference, so removing the cell removes the name", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const dn = Workbook.getDefinedNames(wb);
    DefinedNames.addFormula(dn, "Tax", "0.2");
    Cell.addName(ws, "A6", "Tax");
    expect(DefinedNames.getModel(dn)).toEqual([{ name: "Tax", ranges: ["S!$A$6"] }]);
    Cell.removeName(ws, "A6", "Tax");
    DefinedNames.addFormula(dn, "Tax", "0.3");
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(DefinedNames.getModel(Workbook.getDefinedNames(back))).toEqual([
      { name: "Tax", ranges: ["0.3"], formulaExpression: "0.3" }
    ]);
  });
});

describe("the deprecated Cell name forms", () => {
  it("setName replaces the names as setNames with one name does", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setNames(ws, "A1", ["OldOne", "OldTwo"]);
    Cell.setNames(ws, "B1", ["OldOne", "OldTwo"]);
    Cell.setName(ws, "A1", "Only");
    Cell.setNames(ws, "B1", ["Only"]);
    expect(Cell.getNames(ws, "A1")).toEqual(["Only"]);
    expect(Cell.getNames(ws, "B1")).toEqual(["Only"]);
    Cell.setName(ws, 1, 3, "ByPosition");
    expect(Cell.getNames(ws, "C1")).toEqual(["ByPosition"]);
  });

  it("removeAllNames is clearNames", () => {
    expect(Cell.removeAllNames).toBe(Cell.clearNames);
  });
});
