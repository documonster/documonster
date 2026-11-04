import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

describe("spill-range operator", () => {
  function sheetWithSequence() {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    return { wb, ws };
  }

  it("`A1#` is the range the anchor spills into", () => {
    const { wb, ws } = sheetWithSequence();
    Cell.setValue(ws, "C1", { formula: "SUM(A1#)", result: 0 });
    Cell.setValue(ws, "C2", { formula: "ROWS(A1#)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toBe(15);
    expect(Cell.getResult(ws, "C2")).toBe(5);
  });

  it("`_xlfn.ANCHORARRAY(A1)`, the form a file stores, means the same", () => {
    const { wb, ws } = sheetWithSequence();
    Cell.setValue(ws, "C1", { formula: "SUM(_xlfn.ANCHORARRAY(A1))", result: 0 });
    Cell.setValue(ws, "C2", { formula: "ROWS(ANCHORARRAY(S!A1))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toBe(15);
    expect(Cell.getResult(ws, "C2")).toBe(5);
  });

  it("follows the anchor when it spills further", () => {
    const { wb, ws } = sheetWithSequence();
    Cell.setValue(ws, "B1", 7);
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(B1)", result: 0 });
    Cell.setValue(ws, "C1", { formula: "ROWS(A1#)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toBe(7);
  });

  it("is #REF! on a cell that does not spill", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 3);
    Cell.setValue(ws, "A2", { formula: "1+1", result: 0 });
    Cell.setValue(ws, "C1", { formula: "SUM(A1#)", result: 0 });
    Cell.setValue(ws, "C2", { formula: "SUM(A2#)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toEqual({ error: "#REF!" });
    expect(Cell.getResult(ws, "C2")).toEqual({ error: "#REF!" });
  });
});
