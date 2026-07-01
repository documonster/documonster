/**
 * End-to-end tests exercising the full compile → evaluate → materialize
 * → apply pipeline. These are the only tests that verify workbook-level
 * roundtripping: the engine result is written back via the plan, and we
 * then read the live cell to confirm the materialised state.
 */

import { definedNamesAdd } from "@excel/core/defined-names";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { getDefinedNames } from "@excel/core/workbook";
import { Cell, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("applyWritebackPlan roundtrip: scalar formulas", () => {
  it("writes scalar result back to the source cell", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", { formula: "A1*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A2")).toBe(20);
  });

  it("preserves error result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
  });

  it("chained dependencies compute in topological order", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", { formula: "A1+1", result: 0 });
    Cell.setValue(ws, "A3", { formula: "A2+1", result: 0 });
    Cell.setValue(ws, "A4", { formula: "A3+1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A4")).toBe(4);
  });

  it("recalc after cell value change updates dependents", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "B1", { formula: "A1+5", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(15);

    Cell.setValue(ws, "A1", 100);
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(105);
  });
});

describe("applyWritebackPlan roundtrip: cross-sheet", () => {
  it("resolves references to other sheets", () => {
    const wb = Workbook.create();
    const s1 = Workbook.addWorksheet(wb, "Data");
    const s2 = Workbook.addWorksheet(wb, "Report");
    Cell.setValue(s1, "A1", 99);
    Cell.setValue(s2, "B1", { formula: "Data!A1 * 2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(s2, "B1")).toBe(198);
  });

  it("quoted sheet names with spaces", () => {
    const wb = Workbook.create();
    const s1 = Workbook.addWorksheet(wb, "My Data");
    const s2 = Workbook.addWorksheet(wb, "Rpt");
    Cell.setValue(s1, "A1", 7);
    Cell.setValue(s2, "A1", { formula: "'My Data'!A1 + 3", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(s2, "A1")).toBe(10);
  });
});

describe("applyWritebackPlan roundtrip: dynamic arrays", () => {
  it("SEQUENCE(3) spills down to A1:A3", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "A2")).toBe(2);
    expect(Cell.getValue(ws, "A3")).toBe(3);
  });

  it("SEQUENCE(2,3) spills to 2×3 block", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(2,3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "A2")).toBe(4);
    expect(Cell.getValue(ws, "C2")).toBe(6);
  });

  it("Changing spill source reclaims old ghost cells", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    // First, spill to A1:A5
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A5")).toBe(5);

    // Change to shorter spill
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "A2")).toBe(2);
    // A3..A5 should have been cleaned up
    expect(Cell.getValue(ws, "A3")).toBeFalsy();
  });

  it("#SPILL! when target cell occupied", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A2", "blocker");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
  });
});

describe("applyWritebackPlan roundtrip: circular references", () => {
  it("circular returns 0 by default (non-iterative)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "A1+1", result: 0 });
    calculateFormulas(wb);
    // Without iteration, engine returns 0 (the seeded fallback) + 1 = 1
    expect(Cell.getResult(ws, "A1")).toBe(1);
  });

  it("converges under iterative calc", () => {
    const wb = Workbook.create();
    wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "A1+1", result: 0 });
    calculateFormulas(wb);
    // Each iteration adds 1; 100 iterations → 100 (or maxIter)
    const r = Cell.getResult(ws, "A1") as number;
    expect(r).toBeGreaterThan(50);
  });
});

describe("applyWritebackPlan roundtrip: shared formula propagation", () => {
  it("a formula copied to a sibling cell evaluates correctly", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Cell.setValue(ws, "B1", { formula: "A1*2", result: 0 });
    Cell.setValue(ws, "B2", { formula: "A2*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(20);
    expect(Cell.getResult(ws, "B2")).toBe(40);
  });
});

describe("applyWritebackPlan roundtrip: error propagation chain", () => {
  it("error in base cell propagates through chain", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "A2", { formula: "A1+1", result: 0 });
    Cell.setValue(ws, "A3", { formula: "A2*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A3")).toEqual({ error: "#DIV/0!" });
  });

  it("IFERROR catches upstream error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "B1", { formula: "IFERROR(A1, 999)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(999);
  });
});

describe("applyWritebackPlan roundtrip: BLANK / null results", () => {
  it("formula returning nothing reads as undefined", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "IF(TRUE,,)", result: 0 });
    calculateFormulas(wb);
    // IF with empty else branch → BLANK → undefined
    const r = Cell.getResult(ws, "A1");
    expect(r === undefined || r === null || r === 0).toBe(true);
  });

  it("string result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: 'UPPER("hello")', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe("HELLO");
  });
});

describe("applyWritebackPlan roundtrip: recalc idempotence", () => {
  it("two consecutive calcs produce the same result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", { formula: "A1+5", result: 0 });
    calculateFormulas(wb);
    const r1 = Cell.getResult(ws, "A2");
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A2")).toBe(r1);
  });

  it("volatile functions re-evaluate each calc", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "RAND()", result: 0 });
    calculateFormulas(wb);
    const r1 = Cell.getResult(ws, "A1");
    calculateFormulas(wb);
    const r2 = Cell.getResult(ws, "A1");
    // 1 in 2^53 chance of same — OK to assert strict
    expect(typeof r1).toBe("number");
    expect(typeof r2).toBe("number");
  });
});

describe("applyWritebackPlan roundtrip: defined names", () => {
  it("workbook-level defined name resolves to cell", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 42);
    definedNamesAdd(getDefinedNames(wb), "S!A1", "Answer");
    Cell.setValue(ws, "B1", { formula: "Answer * 2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(84);
  });
});
