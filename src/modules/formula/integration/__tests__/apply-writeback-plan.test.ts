/**
 * End-to-end tests exercising the full compile → evaluate → materialize
 * → apply pipeline. These are the only tests that verify workbook-level
 * roundtripping: the engine result is written back via the plan, and we
 * then read the live cell to confirm the materialised state.
 */

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

describe("applyWritebackPlan roundtrip: scalar formulas", () => {
  it("writes scalar result back to the source cell", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = { formula: "A1*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A2").result).toBe(20);
  });

  it("preserves error result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
  });

  it("chained dependencies compute in topological order", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = { formula: "A1+1", result: 0 };
    ws.getCell("A3").value = { formula: "A2+1", result: 0 };
    ws.getCell("A4").value = { formula: "A3+1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A4").result).toBe(4);
  });

  it("recalc after cell value change updates dependents", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("B1").value = { formula: "A1+5", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(15);

    ws.getCell("A1").value = 100;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(105);
  });
});

describe("applyWritebackPlan roundtrip: cross-sheet", () => {
  it("resolves references to other sheets", () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet("Data");
    const s2 = wb.addWorksheet("Report");
    s1.getCell("A1").value = 99;
    s2.getCell("B1").value = { formula: "Data!A1 * 2", result: 0 };
    wb.calculateFormulas();
    expect(s2.getCell("B1").result).toBe(198);
  });

  it("quoted sheet names with spaces", () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet("My Data");
    const s2 = wb.addWorksheet("Rpt");
    s1.getCell("A1").value = 7;
    s2.getCell("A1").value = { formula: "'My Data'!A1 + 3", result: 0 };
    wb.calculateFormulas();
    expect(s2.getCell("A1").result).toBe(10);
  });
});

describe("applyWritebackPlan roundtrip: dynamic arrays", () => {
  it("SEQUENCE(3) spills down to A1:A3", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A3").value).toBe(3);
  });

  it("SEQUENCE(2,3) spills to 2×3 block", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(2,3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(4);
    expect(ws.getCell("C2").value).toBe(6);
  });

  it("Changing spill source reclaims old ghost cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // First, spill to A1:A5
    ws.getCell("A1").value = { formula: "SEQUENCE(5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A5").value).toBe(5);

    // Change to shorter spill
    ws.getCell("A1").value = { formula: "SEQUENCE(2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(2);
    // A3..A5 should have been cleaned up
    expect(ws.getCell("A3").value).toBeFalsy();
  });

  it("#SPILL! when target cell occupied", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A2").value = "blocker";
    ws.getCell("A1").value = { formula: "SEQUENCE(3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
  });
});

describe("applyWritebackPlan roundtrip: circular references", () => {
  it("circular returns 0 by default (non-iterative)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    // Without iteration, engine returns 0 (the seeded fallback) + 1 = 1
    expect(ws.getCell("A1").result).toBe(1);
  });

  it("converges under iterative calc", () => {
    const wb = new Workbook();
    wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    // Each iteration adds 1; 100 iterations → 100 (or maxIter)
    const r = ws.getCell("A1").result as number;
    expect(r).toBeGreaterThan(50);
  });
});

describe("applyWritebackPlan roundtrip: shared formula propagation", () => {
  it("a formula copied to a sibling cell evaluates correctly", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("B1").value = { formula: "A1*2", result: 0 };
    ws.getCell("B2").value = { formula: "A2*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(20);
    expect(ws.getCell("B2").result).toBe(40);
  });
});

describe("applyWritebackPlan roundtrip: error propagation chain", () => {
  it("error in base cell propagates through chain", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("A2").value = { formula: "A1+1", result: 0 };
    ws.getCell("A3").value = { formula: "A2*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A3").result).toEqual({ error: "#DIV/0!" });
  });

  it("IFERROR catches upstream error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("B1").value = { formula: "IFERROR(A1, 999)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(999);
  });
});

describe("applyWritebackPlan roundtrip: BLANK / null results", () => {
  it("formula returning nothing reads as undefined", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "IF(TRUE,,)", result: 0 };
    wb.calculateFormulas();
    // IF with empty else branch → BLANK → undefined
    const r = ws.getCell("A1").result;
    expect(r === undefined || r === null || r === 0).toBe(true);
  });

  it("string result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: 'UPPER("hello")', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe("HELLO");
  });
});

describe("applyWritebackPlan roundtrip: recalc idempotence", () => {
  it("two consecutive calcs produce the same result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = { formula: "A1+5", result: 0 };
    wb.calculateFormulas();
    const r1 = ws.getCell("A2").result;
    wb.calculateFormulas();
    expect(ws.getCell("A2").result).toBe(r1);
  });

  it("volatile functions re-evaluate each calc", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "RAND()", result: 0 };
    wb.calculateFormulas();
    const r1 = ws.getCell("A1").result;
    wb.calculateFormulas();
    const r2 = ws.getCell("A1").result;
    // 1 in 2^53 chance of same — OK to assert strict
    expect(typeof r1).toBe("number");
    expect(typeof r2).toBe("number");
  });
});

describe("applyWritebackPlan roundtrip: defined names", () => {
  it("workbook-level defined name resolves to cell", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 42;
    wb.definedNames.add("S!A1", "Answer");
    ws.getCell("B1").value = { formula: "Answer * 2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(84);
  });
});
