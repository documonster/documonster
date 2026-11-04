/**
 * The `setModel` members apply a whole model: a value no workbook can hold is refused before anything changes, and a
 * model that fails part-way leaves the sheet, the cell or the workbook as it was.
 */
import {
  Cell,
  DefinedNames,
  ExcelArgumentError,
  Table,
  Workbook,
  Worksheet,
  WorksheetNameError
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Cell.setValue(ws, "A5", 1);
  Worksheet.merge(ws, "B1:C1");
  return { wb, ws };
}

describe("Cell.setModel", () => {
  it("refuses a model for another address", async () => {
    const { wb, ws } = sheet();
    expect(() => Cell.setModel(ws, "D4", { address: "Z99", type: 2, value: 5 } as never)).toThrow(
      ExcelArgumentError
    );
    expect(Cell.has(ws, "D4")).toBe(false);
    await expect(Workbook.read(await Workbook.toBuffer(wb))).resolves.toBeDefined();
  });

  it.each([
    ["NaN", { type: 2, value: NaN }],
    ["Infinity", { type: 2, value: Infinity }],
    ["a string as a number", { type: 2, value: "abc" }],
    ["a number as a string", { type: 3, value: 5 }],
    ["an invalid Date", { type: 4, value: new Date(NaN) }],
    ["a boolean that is not one", { type: 9, value: "yes" }],
    ["an unknown type", { type: 99 }],
    ["a formula result of NaN", { type: 6, formula: "1+1", result: NaN }],
    ["a formula that is not text", { type: 6, formula: 5 }]
  ])("refuses %s and leaves the cell as it was", (_name, partial) => {
    const { ws } = sheet();
    expect(() => Cell.setModel(ws, "A5", { address: "A5", ...partial } as never)).toThrow(
      ExcelArgumentError
    );
    expect(Cell.getValue(ws, "A5")).toBe(1);
  });

  it("refuses something that is not a model", () => {
    const { ws } = sheet();
    expect(() => Cell.setModel(ws, "A5", null as never)).toThrow(ExcelArgumentError);
    expect(() => Cell.setModel(ws, "A6", 5 as never)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "A6")).toBe(false);
  });

  it("still applies a model read from another cell", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "B5", { formula: "A5*2", result: 2 });
    const model = { ...Cell.getModel(ws, "B5"), address: "B6" };
    Cell.setModel(ws, "B6", model);
    expect(Cell.getFormula(ws, "B6")).toBe("A5*2");
  });
});

describe("Worksheet.setModel", () => {
  it("refuses a model missing what it must carry, and changes nothing", () => {
    const { wb, ws } = sheet();
    for (const model of [null, {}, { name: "X" }, 42]) {
      expect(() => Worksheet.setModel(ws, model as never)).toThrow(ExcelArgumentError);
    }
    expect(Worksheet.getName(ws)).toBe("Data");
    expect(Cell.getValue(ws, "A5")).toBe(1);
    expect(Worksheet.getMergedRegions(ws)).toEqual([{ top: 1, left: 2, bottom: 1, right: 3 }]);
    expect(Workbook.getWorksheet(wb, "Data")).toBe(ws);
  });

  it("refuses a name another sheet has before changing anything", () => {
    const { wb, ws } = sheet();
    Workbook.addWorksheet(wb, "B");
    const model = { ...Worksheet.getModel(ws), name: "B", rows: [] };
    expect(() => Worksheet.setModel(ws, model)).toThrow(WorksheetNameError);
    expect(Cell.getValue(ws, "A5")).toBe(1);
    expect(Worksheet.getName(ws)).toBe("Data");
  });

  it("refuses a cell no workbook can hold, and changes nothing", () => {
    const { ws } = sheet();
    const model = Worksheet.getModel(ws);
    const bad = {
      ...model,
      rows: [{ number: 5, cells: [{ address: "A5", type: 2, value: NaN }] }]
    } as never;
    expect(() => Worksheet.setModel(ws, bad)).toThrow(ExcelArgumentError);
    expect(Cell.getValue(ws, "A5")).toBe(1);
  });

  it("refuses a cell filed under another row", () => {
    const { ws } = sheet();
    const bad = {
      ...Worksheet.getModel(ws),
      rows: [{ number: 4, cells: [{ address: "Z99", type: 2, value: 5 }] }]
    } as never;
    expect(() => Worksheet.setModel(ws, bad)).toThrow(ExcelArgumentError);
    expect(Cell.getValue(ws, "A5")).toBe(1);
  });

  it("leaves the sheet as it was when a merge in the model conflicts", () => {
    const { ws } = sheet();
    const model = { ...Worksheet.getModel(ws), mergeCells: ["E1:F2", "F2:G3"] };
    expect(() => Worksheet.setModel(ws, model)).toThrow();
    expect(Cell.getValue(ws, "A5")).toBe(1);
    expect(Worksheet.getMergedRegions(ws)).toEqual([{ top: 1, left: 2, bottom: 1, right: 3 }]);
  });

  it("frees the names of the tables the model drops", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "H1", "h");
    Cell.setValue(ws, "H2", 1);
    Table.add(ws, {
      name: "T_a",
      ref: "H1",
      headerRow: true,
      columns: [{ name: "h" }],
      rows: [[1]]
    });
    Worksheet.setModel(ws, { ...Worksheet.getModel(ws), tables: [] });
    expect(() =>
      Table.add(ws, {
        name: "T_a",
        ref: "J1",
        headerRow: true,
        columns: [{ name: "h" }],
        rows: [[1]]
      })
    ).not.toThrow();
  });

  it("renames as Worksheet.setName does, so references follow", () => {
    const wb = Workbook.create();
    const a = Workbook.addWorksheet(wb, "A");
    const b = Workbook.addWorksheet(wb, "B");
    Cell.setValue(a, "A1", 5);
    Cell.setValue(b, "A1", { formula: "A!A1*2" });
    DefinedNames.add(Workbook.getDefinedNames(wb), "A!$A$1", "Five");
    Worksheet.setModel(a, { ...Worksheet.getModel(a), name: "Renamed" });
    expect(Workbook.getWorksheet(wb, "Renamed")).toBe(a);
    expect(Cell.getFormula(b, "A1")).toBe("Renamed!A1*2");
    expect(DefinedNames.getModel(Workbook.getDefinedNames(wb))).toEqual([
      { name: "Five", ranges: ["Renamed!$A$1"] }
    ]);
    expect(Cell.getValue(a, "A1")).toBe(5);
  });
});

describe("Workbook.setModel", () => {
  it("refuses a model missing what it must carry, and keeps every sheet", () => {
    const { wb, ws } = sheet();
    for (const model of [null, {}, 42]) {
      expect(() => Workbook.setModel(wb, model as never)).toThrow(ExcelArgumentError);
    }
    expect(Workbook.listWorksheets(wb)).toEqual([ws]);
    expect(Cell.getValue(ws, "A5")).toBe(1);
  });

  it("leaves the workbook as it was when a sheet in the model fails", () => {
    const { wb, ws } = sheet();
    const model = Workbook.getModel(wb);
    const broken = {
      ...model,
      worksheets: [
        { ...model.worksheets[0], name: "One" },
        { ...model.worksheets[0], id: 2, name: "Two", mergeCells: ["E1:F2", "F2:G3"] }
      ]
    };
    expect(() => Workbook.setModel(wb, broken)).toThrow();
    expect(Workbook.listWorksheets(wb).map(Worksheet.getName)).toEqual(["Data"]);
    expect(Cell.getValue(Workbook.getWorksheet(wb, "Data")!, "A5")).toBe(1);
    expect(ws).toBeDefined();
  });

  it("still applies a model it produced", async () => {
    const { wb } = sheet();
    const other = Workbook.create();
    Workbook.setModel(other, Workbook.getModel(wb));
    expect(Cell.getValue(Workbook.getWorksheet(other, "Data")!, "A5")).toBe(1);
    await expect(Workbook.toBuffer(other)).resolves.toBeDefined();
  });
});
