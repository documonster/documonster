/**
 * Unit tests for the writeback-plan builder. These exercise the plan
 * structure directly (not via Workbook) to verify the materialize layer
 * correctly turns a RuntimeValue into ScalarWrite / SpillWrite /
 * CSEWrite / SpillErrorWrite / CleanupWrite operations.
 */

import { calculateFormulas } from "@excel/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("build-writeback-plan: scalar write", () => {
  it("ScalarWrite produced for scalar formula result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "2+3", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(5);
  });

  it("scalar write handles string result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: 'UPPER("abc")', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe("ABC");
  });

  it("scalar write handles boolean result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "TRUE", result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(true);
  });
});

describe("build-writeback-plan: spill plans", () => {
  it("source cell is set via result, ghost cells via value", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)", result: 0 });
    calculateFormulas(wb);
    // Source cell A1: result is set (formula cell)
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getFormula(ws, "A1")).toBe("SEQUENCE(3)");
    // Ghost cells A2, A3: just raw values (not formulas)
    expect(Cell.getValue(ws, "A2")).toBe(2);
    expect(Cell.getFormula(ws, "A2")).toBeFalsy();
    expect(Cell.getValue(ws, "A3")).toBe(3);
  });

  it("2D spill fills rectangle", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "B2", { formula: "SEQUENCE(2,2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B2")).toBe(1);
    expect(Cell.getValue(ws, "C2")).toBe(2);
    expect(Cell.getValue(ws, "B3")).toBe(3);
    expect(Cell.getValue(ws, "C3")).toBe(4);
  });

  it("spill region shrinks on recalc — old ghosts cleaned up", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 5);
    Cell.setValue(ws, "B1", { formula: "SEQUENCE(A1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "B5")).toBe(5);

    // Shrink
    Cell.setValue(ws, "A1", 2);
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(1);
    expect(Cell.getValue(ws, "B2")).toBe(2);
    // B3, B4, B5 should be cleared
    expect(Cell.getValue(ws, "B3")).toBeFalsy();
    expect(Cell.getValue(ws, "B5")).toBeFalsy();
  });

  it("#SPILL! when a target cell has a value", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A3", "block");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
  });

  it("#SPILL! does not overwrite the blocking cell", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A3", "block");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A3")).toBe("block");
  });
});

describe("build-writeback-plan: persistence across calcs", () => {
  it("user modifying a ghost cell turns off spill and surfaces #SPILL!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A2")).toBe(2);

    // User modifies a ghost
    Cell.setValue(ws, "A2", "user-content");
    calculateFormulas(wb);
    // SEQUENCE(3) can no longer spill — #SPILL!
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
    expect(Cell.getValue(ws, "A2")).toBe("user-content");
  });
});

describe("build-writeback-plan: error cached results", () => {
  it("error result wraps as SnapshotErrorValue", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SQRT(-1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NUM!" });
  });
});
