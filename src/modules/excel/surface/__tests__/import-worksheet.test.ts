/**
 * `Workbook.importWorksheet` checks its source before adding anything, and a copied table whose name is taken is
 * renamed as Excel's "Move or Copy" renames it — with the copied sheet's own references to it following.
 */
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Row,
  Table,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function source() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Row.addMany(ws, [
    ["Region", "Amt"],
    ["N", 1],
    ["S", 2]
  ]);
  Table.add(ws, {
    name: "Sales",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "Region" }, { name: "Amt" }],
    rows: [
      ["N", 1],
      ["S", 2]
    ]
  });
  Cell.setValue(ws, "D1", { formula: "SUM(Sales[Amt])" });
  Cell.setValue(ws, "D2", { formula: "ROWS(Sales)" });
  Cell.setValue(ws, "D3", { formula: 'COUNTIF(Sales[Region],"N")&"Sales[x]"' });
  ConditionalFormatting.add(ws, {
    ref: "E1",
    rules: [{ type: "expression", formulae: ["SUM(Sales[Amt])>2"], priority: 1 }]
  });
  DataValidation.add(ws, "F1", {
    type: "list",
    allowBlank: true,
    formulae: ['INDIRECT("Sales[Region]")']
  });
  return { wb, ws };
}

function tableNames(wb: ReturnType<typeof Workbook.create>) {
  return Workbook.listWorksheets(wb).map(ws => [
    Worksheet.getName(ws),
    Table.list(ws).map(t => Table.getName(t))
  ]);
}

describe("Workbook.importWorksheet", () => {
  it("refuses a source that is not a worksheet, and adds nothing", () => {
    const { wb } = source();
    for (const bad of [42, null, undefined, {}, "Data"]) {
      expect(() => Workbook.importWorksheet(wb, bad as never, "X")).toThrow(ExcelArgumentError);
    }
    expect(Workbook.listWorksheets(wb).map(Worksheet.getName)).toEqual(["Data"]);
  });

  it("refuses a name that is not a string, and adds nothing", () => {
    const { wb, ws } = source();
    expect(() => Workbook.importWorksheet(wb, ws, 42 as never)).toThrow();
    expect(Workbook.listWorksheets(wb).map(Worksheet.getName)).toEqual(["Data"]);
  });

  it("renames a copied table whose name is taken, and the copy's formulas follow it", async () => {
    const { wb, ws } = source();
    const copy = Workbook.importWorksheet(wb, ws, "Copy");
    expect(tableNames(wb)).toEqual([
      ["Data", ["Sales"]],
      ["Copy", ["Sales2"]]
    ]);
    expect(Cell.getFormula(copy, "D1")).toBe("SUM(Sales2[Amt])");
    expect(Cell.getFormula(copy, "D2")).toBe("ROWS(Sales2)");
    // A string literal that only looks like a reference is text, and is left alone.
    expect(Cell.getFormula(copy, "D3")).toBe('COUNTIF(Sales2[Region],"N")&"Sales[x]"');
    expect(ConditionalFormatting.list(copy)[0].rules[0]).toMatchObject({
      formulae: ["SUM(Sales2[Amt])>2"]
    });
    // The source keeps its own.
    expect(Cell.getFormula(ws, "D1")).toBe("SUM(Sales[Amt])");
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    expect(tableNames(back)).toEqual([
      ["Data", ["Sales"]],
      ["Copy", ["Sales2"]]
    ]);
  });

  it("takes the next free name when the first is taken too, and across workbooks", async () => {
    const { wb, ws } = source();
    Workbook.importWorksheet(wb, ws, "Copy");
    const third = Workbook.importWorksheet(wb, ws, "Third");
    expect(Table.list(third).map(t => Table.getName(t))).toEqual(["Sales3"]);

    const other = Workbook.create();
    const mine = Workbook.addWorksheet(other, "Mine");
    Row.addMany(mine, [["h"], [1]]);
    Table.add(mine, {
      name: "sales",
      ref: "A1",
      headerRow: true,
      columns: [{ name: "h" }],
      rows: [[1]]
    });
    const imported = Workbook.importWorksheet(other, ws);
    expect(Table.list(imported).map(t => Table.getName(t))).toEqual(["Sales2"]);
    expect(Cell.getFormula(imported, "D1")).toBe("SUM(Sales2[Amt])");
    await expect(Workbook.toBuffer(other)).resolves.toBeDefined();
  });

  it("keeps a table's name when it is free", () => {
    const { ws } = source();
    const other = Workbook.create();
    const imported = Workbook.importWorksheet(other, ws);
    expect(Table.list(imported).map(t => Table.getName(t))).toEqual(["Sales"]);
    expect(Cell.getFormula(imported, "D1")).toBe("SUM(Sales[Amt])");
  });
});
