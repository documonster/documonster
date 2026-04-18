/**
 * Integration tests for the calculate-formulas pipeline.
 *
 * These tests exercise `calculateFormulasImpl` through the public
 * `Workbook.calculateFormulas()` entry point so that snapshot →
 * compile → topological sort → evaluate → writeback is covered
 * end-to-end.
 *
 * Focus areas:
 *   - Dependency-graph topological ordering (chains, diamonds, multi-root)
 *   - Iterative calculation convergence
 *   - Dynamic-reference discovery (INDIRECT / OFFSET) feeding back into
 *     the dependency graph via session.dynamicDeps
 *   - AST cache reuse across repeated invocations
 *   - Dynamic-array spill lifecycle (spill → resize → clear)
 *   - Formula-kind dispatch: normal, shared, CSE master/slave, dynamic array
 *   - Error surfacing from the three pipeline phases: parse, bind, evaluate
 *   - date1904 propagation to the function registry
 *   - Empty workbook / no-formula no-ops
 */

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

// ============================================================================
// Dependency Graph Topological Ordering
// ============================================================================

describe("calculate-formulas: dependency graph topological ordering", () => {
  it("evaluates a chain A1 → B1 → C1 in correct order (R2-PX-01)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "A1+10", result: 0 };
    ws.getCell("C1").value = { formula: "B1*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(11);
    expect(ws.getCell("C1").result).toBe(22);
  });

  it("evaluates a chain declared in reverse source order (R2-PX-02)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Declare C1 first — topological sort must still put A1 → B1 → C1.
    ws.getCell("C1").value = { formula: "B1*2", result: 0 };
    ws.getCell("B1").value = { formula: "A1+10", result: 0 };
    ws.getCell("A1").value = 1;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(11);
    expect(ws.getCell("C1").result).toBe(22);
  });

  it("evaluates a diamond dependency A→B, A→C, B+C→D (R2-PX-03)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 5;
    ws.getCell("B1").value = { formula: "A1*2", result: 0 }; // 10
    ws.getCell("C1").value = { formula: "A1+1", result: 0 }; // 6
    ws.getCell("D1").value = { formula: "B1+C1", result: 0 }; // 16
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(10);
    expect(ws.getCell("C1").result).toBe(6);
    expect(ws.getCell("D1").result).toBe(16);
  });

  it("evaluates multiple independent roots in one pass (R2-PX-04)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Two disjoint chains sharing no cells.
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = { formula: "A1*2", result: 0 };
    ws.getCell("B1").value = 20;
    ws.getCell("B2").value = { formula: "B1+3", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A2").result).toBe(20);
    expect(ws.getCell("B2").result).toBe(23);
  });

  it("evaluates a deep chain without stack issues (R2-PX-05)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 0;
    for (let i = 2; i <= 100; i++) {
      ws.getCell(`A${i}`).value = { formula: `A${i - 1}+1`, result: 0 };
    }
    wb.calculateFormulas();
    expect(ws.getCell("A100").result).toBe(99);
  });
});

// ============================================================================
// Iterative Calculation
// ============================================================================

describe("calculate-formulas: iterative calculation", () => {
  it("converges a single-cell self-reference with iterate=true (R2-PX-10)", () => {
    const wb = new Workbook();
    wb.calcProperties = { iterate: true, iterateCount: 200, iterateDelta: 0.0001 };
    const ws = wb.addWorksheet("Sheet1");
    // Classic converging self-ref: A1 = A1/2 + 1 → fixed point 2
    ws.getCell("A1").value = { formula: "A1/2+1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBeCloseTo(2, 3);
  });

  it("advances a non-converging self-reference up to iterateCount (R2-PX-11)", () => {
    const wb = new Workbook();
    wb.calcProperties = { iterate: true, iterateCount: 7, iterateDelta: 0 };
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    // Initial pass (circular fallback 0 → 0+1 = 1), then 7 iterations: 8
    expect(ws.getCell("A1").result).toBe(8);
  });

  it("converges a 3-cell ring A→B→C→A (R2-PX-12)", () => {
    const wb = new Workbook();
    wb.calcProperties = { iterate: true, iterateCount: 200, iterateDelta: 0.0001 };
    const ws = wb.addWorksheet("Sheet1");
    // A = C/2 + 1, B = A/2 + 1, C = B/2 + 1 → all converge to 2
    ws.getCell("A1").value = { formula: "C1/2+1", result: 0 };
    ws.getCell("B1").value = { formula: "A1/2+1", result: 0 };
    ws.getCell("C1").value = { formula: "B1/2+1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBeCloseTo(2, 3);
    expect(ws.getCell("B1").result).toBeCloseTo(2, 3);
    expect(ws.getCell("C1").result).toBeCloseTo(2, 3);
  });

  it("re-evaluates downstream non-circular cells after convergence (R2-PX-13)", () => {
    const wb = new Workbook();
    wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
    const ws = wb.addWorksheet("Sheet1");
    // A1 converges to 1, downstream B1 and C1 should reflect the converged value.
    ws.getCell("A1").value = { formula: "IF(A1>0,A1,1)", result: 0 };
    ws.getCell("B1").value = { formula: "A1*10", result: 0 };
    ws.getCell("C1").value = { formula: "B1+5", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("B1").result).toBe(10);
    expect(ws.getCell("C1").result).toBe(15);
  });

  it("returns a number (not an error) for a cycle without iterate (R2-PX-14)", () => {
    // Baseline: circular ref zero-seed fallback. Verified again here at
    // the integration layer to pin the behaviour.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "A2+1", result: 0 };
    ws.getCell("A2").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    expect(typeof ws.getCell("A1").result).toBe("number");
    expect(typeof ws.getCell("A2").result).toBe("number");
  });
});

// ============================================================================
// Dynamic Reference Discovery (INDIRECT / OFFSET)
// ============================================================================

describe("calculate-formulas: dynamic reference discovery", () => {
  it("re-evaluates INDIRECT targets when the address cell changes (R2-PX-20)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "C1";
    ws.getCell("C1").value = 100;
    ws.getCell("E1").value = 200;
    ws.getCell("B1").value = { formula: "INDIRECT(A1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(100);

    // Flip the target address and recalc — B1 must re-evaluate.
    ws.getCell("A1").value = "E1";
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(200);
  });

  it("propagates INDIRECT changes to downstream dependents (R2-PX-21)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "C1";
    ws.getCell("C1").value = 3;
    // B1 = INDIRECT(A1) → 3, D1 depends on B1.
    ws.getCell("B1").value = { formula: "INDIRECT(A1)", result: 0 };
    ws.getCell("D1").value = { formula: "B1*10", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(3);
    expect(ws.getCell("D1").result).toBe(30);
  });

  it("OFFSET-generated range is re-read after source data changes (R2-PX-22)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("B1").value = {
      formula: "SUM(OFFSET(A1,0,0,3,1))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(6);

    // Data changes; recalc should pick them up even though OFFSET's deps
    // are dynamic.
    ws.getCell("A2").value = 20;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(24);
  });
});

// ============================================================================
// AST Cache Reuse
// ============================================================================

describe("calculate-formulas: AST cache reuse across invocations", () => {
  it("produces identical results across successive recalculations (R2-PX-30)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 4;
    ws.getCell("A2").value = { formula: "A1*A1+3", result: 0 };
    wb.calculateFormulas();
    const first = ws.getCell("A2").result;
    wb.calculateFormulas();
    const second = ws.getCell("A2").result;
    wb.calculateFormulas();
    const third = ws.getCell("A2").result;
    expect(first).toBe(19);
    expect(second).toBe(19);
    expect(third).toBe(19);
  });

  it("keeps a cached AST usable after mutating dependency cells (R2-PX-31)", () => {
    // This exercises the same AST on successive calcs with different inputs.
    // The AST cache is keyed by formula text, so identical text reuses the
    // cached bound expression.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "A1*A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(1);

    ws.getCell("A1").value = 5;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(25);

    ws.getCell("A1").value = 10;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(100);
  });

  it("shares a cached AST between two cells with the same formula text (R2-PX-32)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 3;
    ws.getCell("A2").value = 5;
    // Identical formula text in different cells — AST cache hit
    ws.getCell("B1").value = { formula: "A1+A2", result: 0 };
    ws.getCell("B2").value = { formula: "A1+A2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(8);
    expect(ws.getCell("B2").result).toBe(8);
  });
});

// ============================================================================
// Dynamic Array Spill Lifecycle
// ============================================================================

describe("calculate-formulas: dynamic array spill lifecycle", () => {
  it("fills SEQUENCE(3) into A1:A3 (R2-PX-40)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SEQUENCE(3)",
      result: 0,
      shareType: "array",
      ref: "A1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A3").value).toBe(3);
  });

  it("spills a 2-column SEQUENCE to a 3×2 block (R2-PX-41)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SEQUENCE(3,2)",
      result: 0,
      shareType: "array",
      ref: "A1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("B1").value).toBe(2);
    expect(ws.getCell("A2").value).toBe(3);
    expect(ws.getCell("B2").value).toBe(4);
    expect(ws.getCell("A3").value).toBe(5);
    expect(ws.getCell("B3").value).toBe(6);
  });

  it("clears ghost cells when a dynamic array source becomes a scalar formula (R2-PX-42)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SEQUENCE(3)",
      result: 0,
      shareType: "array",
      ref: "A1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A3").value).toBe(3);

    // Replace with a scalar formula — previous ghosts must be cleared.
    ws.getCell("A1").value = { formula: "100", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(100);
    expect(ws.findCell(2, 1)?.value ?? null).toBeNull();
    expect(ws.findCell(3, 1)?.value ?? null).toBeNull();
  });

  it("shrinks the spill footprint when the output array shrinks (R2-PX-43)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("C1").value = 5;
    ws.getCell("A1").value = {
      formula: "SEQUENCE(C1)",
      result: 0,
      shareType: "array",
      ref: "A1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("A5").value).toBe(5);

    // Shrink the source — row 5 ghost should be removed
    ws.getCell("C1").value = 2;
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.findCell(3, 1)?.value ?? null).toBeNull();
    expect(ws.findCell(4, 1)?.value ?? null).toBeNull();
    expect(ws.findCell(5, 1)?.value ?? null).toBeNull();
  });
});

// ============================================================================
// Formula Kinds — Normal, Shared, CSE Master/Slave, Dynamic Array
// ============================================================================

describe("calculate-formulas: formula kind dispatch", () => {
  it("evaluates a plain scalar formula (R2-PX-50)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 6;
    ws.getCell("B1").value = { formula: "A1*7", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(42);
  });

  it("evaluates a shared (fillFormula) group with relative refs (R2-PX-51)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    // B1:B3 — master at B1, slaves at B2/B3 share the formula with slide
    ws.fillFormula("B1:B3", "A1*10", [10, 20, 30]);
    // Modify source to confirm slaves recompute, not just return cached values.
    ws.getCell("A1").value = 5;
    ws.getCell("A2").value = 6;
    ws.getCell("A3").value = 7;
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(50);
    expect(ws.getCell("B2").result).toBe(60);
    expect(ws.getCell("B3").result).toBe(70);
  });

  it("distributes CSE array formula results across the target range (R2-PX-52)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("B3").value = 30;
    // Master at C1 with range C1:C3; slaves at C2/C3
    ws.getCell("C1").value = {
      formula: "A1:A3*B1:B3",
      result: 0,
      shareType: "array",
      ref: "C1:C3"
    };
    ws.getCell("C2").value = { formula: "A1:A3*B1:B3", result: 0, shareType: "array" };
    ws.getCell("C3").value = { formula: "A1:A3*B1:B3", result: 0, shareType: "array" };
    wb.calculateFormulas();
    expect(ws.getCell("C1").result).toBe(10);
    expect(ws.getCell("C2").result).toBe(40);
    expect(ws.getCell("C3").result).toBe(90);
  });

  it("caches CSE array results for slave cells — second recalc agrees (R2-PX-53)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B1").value = 4;
    ws.getCell("B2").value = 5;
    ws.getCell("C1").value = {
      formula: "A1:A2*B1:B2",
      result: 0,
      shareType: "array",
      ref: "C1:C2"
    };
    ws.getCell("C2").value = { formula: "A1:A2*B1:B2", result: 0, shareType: "array" };
    wb.calculateFormulas();
    const first1 = ws.getCell("C1").result;
    const first2 = ws.getCell("C2").result;
    wb.calculateFormulas();
    expect(ws.getCell("C1").result).toBe(first1);
    expect(ws.getCell("C2").result).toBe(first2);
    expect(first1).toBe(8);
    expect(first2).toBe(15);
  });

  it("spills a dynamic array formula (R2-PX-54)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 3;
    ws.getCell("A2").value = 1;
    ws.getCell("A3").value = 2;
    ws.getCell("C1").value = {
      formula: "_xlfn._xlws.SORT(A1:A3)",
      result: 0,
      shareType: "array",
      ref: "C1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("C1").result).toBe(1);
    expect(ws.getCell("C2").value).toBe(2);
    expect(ws.getCell("C3").value).toBe(3);
  });
});

// ============================================================================
// Error Surfacing by Pipeline Phase
// ============================================================================

describe("calculate-formulas: error surfacing", () => {
  it("surfaces #NAME? when an unknown function is referenced (R2-PX-60)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Unknown function: tokenizes/parses fine but the runtime has no
    // matching descriptor. Omit `result:` so the engine can't preserve
    // a cached value, forcing it to surface the actual error.
    ws.getCell("A1").value = { formula: "XYZUNKNOWN(1,2,3)" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NAME?" });
  });

  it("surfaces #REF! for an external workbook reference (bind failure) (R2-PX-61)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "[Book1]Sheet1!A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
  });

  it("surfaces the specific error produced during evaluation (R2-PX-62)", () => {
    // Verifies that evaluator-phase errors propagate, not the generic
    // #CALC! catch-all.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("A2").value = { formula: "SQRT(-4)", result: 0 };
    ws.getCell("A3").value = { formula: "NA()", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
    expect(ws.getCell("A2").result).toEqual({ error: "#NUM!" });
    expect(ws.getCell("A3").result).toEqual({ error: "#N/A" });
  });

  it("preserves a cached result when a formula text is un-parseable (R2-PX-63)", () => {
    // The unsupported-formula path: if the engine can't parse or
    // recognize the formula but the author provided a cached result,
    // that result is preserved rather than overwritten.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "XYZUNKNOWN(1,2,3)",
      result: 42
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(42);
  });
});

// ============================================================================
// date1904 Propagation
// ============================================================================

describe("calculate-formulas: date1904 propagation", () => {
  it("DATE(1900,1,1) differs by exactly 1462 between 1900 and 1904 modes (R2-PX-70)", () => {
    // Same DATE() formula, two workbooks, only the date1904 flag differs.
    const wbA = new Workbook();
    wbA.properties = { date1904: false };
    const wsA = wbA.addWorksheet("Sheet1");
    wsA.getCell("A1").value = { formula: "DATE(2024,1,1)", result: 0 };
    wbA.calculateFormulas();

    const wbB = new Workbook();
    wbB.properties = { date1904: true };
    const wsB = wbB.addWorksheet("Sheet1");
    wsB.getCell("A1").value = { formula: "DATE(2024,1,1)", result: 0 };
    wbB.calculateFormulas();

    const a = wsA.getCell("A1").result as number;
    const b = wsB.getCell("A1").result as number;
    expect(typeof a).toBe("number");
    expect(typeof b).toBe("number");
    // 1900 mode serial is 1462 larger (the difference between the two epochs).
    expect(a - b).toBe(1462);
  });
});

// ============================================================================
// Empty / No-Formula Workbook
// ============================================================================

describe("calculate-formulas: empty and no-formula workbooks", () => {
  it("does nothing on an empty workbook (R2-PX-80)", () => {
    const wb = new Workbook();
    expect(() => wb.calculateFormulas()).not.toThrow();
  });

  it("does nothing on a worksheet with only data cells (R2-PX-81)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = "hello";
    ws.getCell("A3").value = true;
    expect(() => wb.calculateFormulas()).not.toThrow();
    expect(ws.getCell("A1").value).toBe(1);
    expect(ws.getCell("A2").value).toBe("hello");
    expect(ws.getCell("A3").value).toBe(true);
  });

  it("handles a single formula with no dependencies (R2-PX-82)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "2*3+4", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(10);
  });
});
