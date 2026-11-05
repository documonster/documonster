/**
 * Tables and defined names share one namespace in Excel, compared by `excelNameKey`: a table is found whatever the
 * case it is spelt in, and a defined name may not take a table's name, nor a table a defined name's. Excel's UI refuses
 * both, and a file holding both makes every formula that uses the name ambiguous.
 *
 * Reading such a file keeps it as read — refusing would make the file unreadable — and the validator reports it.
 */
import {
  Cell,
  ExcelDefinedNameError,
  DefinedNames,
  Table,
  ExcelTableError,
  Workbook
} from "@excel/index";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

function workbookWithTable(name = "Maß"): {
  wb: Workbook.Handle;
  ws: ReturnType<typeof Workbook.addWorksheet>;
} {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Table.add(ws, { name, ref: "A1", columns: [{ name: "x" }], rows: [[1]] });
  return { wb, ws };
}

describe("table lookup by name", () => {
  it("finds a table whatever its case or compatibility form", () => {
    const { ws } = workbookWithTable("Sales2");
    expect(Table.get(ws, "sales2")).toBe(Table.get(ws, "SALES2"));
    expect(Table.get(ws, "Sales2")).toBeDefined();
    const { ws: sales } = workbookWithTable("Sales");
    expect(Table.get(sales, "ＳＡＬＥＳ")).toBe(Table.get(sales, "Sales"));
    expect(Table.get(sales, "Sales")).toBeDefined();
    const { ws: other } = workbookWithTable("Maß");
    expect(Table.get(other, "MASS")).toBe(Table.get(other, "Maß"));
  });

  it("removes a table named in another case", () => {
    const { ws } = workbookWithTable("Sales2");
    expect(Table.remove(ws, "sales2")).toBe(true);
    expect(Table.list(ws)).toEqual([]);
  });

  it("keeps `ı` and `I` apart", () => {
    const { ws } = workbookWithTable("Iı");
    expect(Table.get(ws, "II")).toBeUndefined();
  });
});

describe("one namespace for tables and defined names", () => {
  it("refuses a defined name that is a table's name", () => {
    const { wb } = workbookWithTable("Maß");
    const names = Workbook.getDefinedNames(wb);
    expect(() => DefinedNames.addFormula(names, "MASS", "1")).toThrow(ExcelDefinedNameError);
    expect(() => DefinedNames.add(names, "S!$A$1", "maß")).toThrow(ExcelDefinedNameError);
    expect(DefinedNames.list(names)).toEqual([]);
  });

  it("refuses a table whose name is a defined name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Rate", "1");
    expect(() =>
      Table.add(ws, { name: "RATE", ref: "A1", columns: [{ name: "x" }], rows: [[1]] })
    ).toThrow(ExcelTableError);
    expect(Table.list(ws)).toEqual([]);
    const table = Table.add(ws, {
      name: "Other",
      ref: "A1",
      columns: [{ name: "x" }],
      rows: [[1]]
    });
    expect(() => Table.setName(table, "rate")).toThrow(ExcelTableError);
    expect(Table.getName(table)).toBe("Other");
  });

  it("still accepts a name that only looks alike", () => {
    const { wb } = workbookWithTable("I");
    expect(() => DefinedNames.addFormula(Workbook.getDefinedNames(wb), "ı", "1")).not.toThrow();
  });

  it("keeps a collision read from a file, and the validator reports it", async () => {
    const { wb } = workbookWithTable("Sales");
    // A model load is the read path; it must not refuse what a file holds.
    DefinedNames.setModel(Workbook.getDefinedNames(wb), [
      { name: "SALES", ranges: ["1"], formulaExpression: "1" }
    ]);
    const bytes = await Workbook.toBuffer(wb);
    const copy = Workbook.create();
    await Workbook.read(copy, bytes);
    expect(DefinedNames.list(Workbook.getDefinedNames(copy)).map(entry => entry.name)).toEqual([
      "SALES"
    ]);
    expect(Table.get(Workbook.getWorksheet(copy, "S")!, "Sales")).toBeDefined();
    const report = await validateXlsxBuffer(bytes, { includeWarnings: true });
    expect(report.problems.map(problem => problem.kind)).toContain("table-name-is-defined-name");
    Cell.setValue(Workbook.getWorksheet(copy, "S")!, "C1", 1);
    await expect(Workbook.toBuffer(copy)).resolves.toBeDefined();
  });
});
