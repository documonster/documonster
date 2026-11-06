import { Cell, ExcelArgumentError, Row, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.setValues(ws, 1, ["a", "b", "c"]);
  Cell.setFont(ws, "B1", { bold: true });
  Cell.setNote(ws, "B1", "note");
  return ws;
}

// Expected values are what exceljs 4.4.0's `row.splice` produces for the same sheet.
describe("Row.spliceCells", () => {
  it("moves values, styles and notes left when removing", () => {
    const ws = sheet();
    Row.spliceCells(ws, 1, 1, 1);
    expect(Row.getValues(ws, 1)).toEqual(["b", "c"]);
    expect(Cell.getFont(ws, "A1")).toEqual({ bold: true });
    expect(Cell.getNote(ws, "A1")).toBe("note");
    expect(Cell.getNote(ws, "B1")).toBeUndefined();
    expect(Cell.has(ws, "C1")).toBe(false);
  });

  it("moves cells right and inserts unstyled values", () => {
    const ws = sheet();
    Row.spliceCells(ws, 1, 2, 0, "x", "y");
    expect(Row.getValues(ws, 1)).toEqual(["a", "x", "y", "b", "c"]);
    expect(Cell.getFont(ws, "B1")).toBeUndefined();
    expect(Cell.getNote(ws, "B1")).toBeUndefined();
    expect(Cell.getFont(ws, "D1")).toEqual({ bold: true });
    expect(Cell.getNote(ws, "D1")).toBe("note");
  });

  it("replaces in place when as many values are inserted as removed", () => {
    const ws = sheet();
    Row.spliceCells(ws, 1, 2, 1, "z");
    expect(Row.getValues(ws, 1)).toEqual(["a", "z", "c"]);
    expect(Cell.getFont(ws, "B1")).toBeUndefined();
  });

  it("refuses a wrong argument before changing anything", () => {
    const ws = sheet();
    expect(() => Row.spliceCells(ws, 1, 0, 1)).toThrow(ExcelArgumentError);
    expect(() => Row.spliceCells(ws, 1, 1, -1)).toThrow(ExcelArgumentError);
    expect(() => Row.spliceCells(ws, 1, 1, 0, Symbol("x") as never)).toThrow(ExcelArgumentError);
    expect(Row.getValues(ws, 1)).toEqual(["a", "b", "c"]);
  });

  it("refuses to put a value or move a cell past XFD before it changes anything", () => {
    const ws = sheet();
    Cell.setValue(ws, "XFD1", "edge");
    // The second inserted value lands past XFD; the first used to be written before that was found.
    expect(() => Row.spliceCells(ws, 2, 16384, 0, "first", "second")).toThrow(ExcelArgumentError);
    expect(Row.getValues(ws, 2)).toEqual([]);
    expect(Cell.has(ws, "XFD2")).toBe(false);
    // XFD1 would move past the edge.
    expect(() => Row.spliceCells(ws, 1, 2, 0, "x")).toThrow(/past 16384/);
    expect(Cell.getValue(ws, "XFD1")).toBe("edge");
    expect(Row.getValues(ws, 1).slice(0, 3)).toEqual(["a", "b", "c"]);
    // Deleting first makes the room an insertion needs.
    Row.spliceCells(ws, 1, 2, 1, "x");
    expect(Row.getValues(ws, 1).slice(0, 3)).toEqual(["a", "x", "c"]);
    expect(Cell.getValue(ws, "XFD1")).toBe("edge");
  });
});
