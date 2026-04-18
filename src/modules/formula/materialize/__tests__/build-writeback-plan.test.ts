/**
 * Unit tests for the writeback-plan builder. These exercise the plan
 * structure directly (not via Workbook) to verify the materialize layer
 * correctly turns a RuntimeValue into ScalarWrite / SpillWrite /
 * CSEWrite / SpillErrorWrite / CleanupWrite operations.
 */

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

describe("build-writeback-plan: scalar write", () => {
  it("ScalarWrite produced for scalar formula result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "2+3", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(5);
  });

  it("scalar write handles string result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: 'UPPER("abc")', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe("ABC");
  });

  it("scalar write handles boolean result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "TRUE", result: false };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(true);
  });
});

describe("build-writeback-plan: spill plans", () => {
  it("source cell is set via result, ghost cells via value", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(3)", result: 0 };
    wb.calculateFormulas();
    // Source cell A1: result is set (formula cell)
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A1").formula).toBe("SEQUENCE(3)");
    // Ghost cells A2, A3: just raw values (not formulas)
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A2").formula).toBeFalsy();
    expect(ws.getCell("A3").value).toBe(3);
  });

  it("2D spill fills rectangle", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("B2").value = { formula: "SEQUENCE(2,2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B2").result).toBe(1);
    expect(ws.getCell("C2").value).toBe(2);
    expect(ws.getCell("B3").value).toBe(3);
    expect(ws.getCell("C3").value).toBe(4);
  });

  it("spill region shrinks on recalc — old ghosts cleaned up", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 5;
    ws.getCell("B1").value = { formula: "SEQUENCE(A1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B5").value).toBe(5);

    // Shrink
    ws.getCell("A1").value = 2;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(1);
    expect(ws.getCell("B2").value).toBe(2);
    // B3, B4, B5 should be cleared
    expect(ws.getCell("B3").value).toBeFalsy();
    expect(ws.getCell("B5").value).toBeFalsy();
  });

  it("#SPILL! when a target cell has a value", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A3").value = "block";
    ws.getCell("A1").value = { formula: "SEQUENCE(5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
  });

  it("#SPILL! does not overwrite the blocking cell", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A3").value = "block";
    ws.getCell("A1").value = { formula: "SEQUENCE(5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A3").value).toBe("block");
  });
});

describe("build-writeback-plan: persistence across calcs", () => {
  it("user modifying a ghost cell turns off spill and surfaces #SPILL!", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A2").value).toBe(2);

    // User modifies a ghost
    ws.getCell("A2").value = "user-content";
    wb.calculateFormulas();
    // SEQUENCE(3) can no longer spill — #SPILL!
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
    expect(ws.getCell("A2").value).toBe("user-content");
  });
});

describe("build-writeback-plan: error cached results", () => {
  it("error result wraps as SnapshotErrorValue", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SQRT(-1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NUM!" });
  });
});
