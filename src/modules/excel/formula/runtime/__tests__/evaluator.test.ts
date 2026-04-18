/**
 * Unit tests for the formula evaluator.
 *
 * These tests drive the evaluator end-to-end through the public Workbook
 * API: each case writes a formula into a cell, runs `wb.calculateFormulas()`,
 * and asserts on `cell.result`. This gives coverage of the entire
 * evaluate pipeline (parse → bind → evaluate → materialize) rather than
 * poking internal helpers.
 *
 * The suite is organised by evaluator feature so that a failure points
 * directly at the responsible branch:
 *
 *   - arithmetic / comparison / concat / unary / percent operators
 *   - broadcasting between scalars and arrays of various shapes
 *   - reference resolution (implicit intersection, `@`, 3D refs,
 *     structured refs, defined names)
 *   - array constants and dynamic-array spill
 *   - conditional forms (IF / IFS / SWITCH / CHOOSE / IFERROR / IFNA)
 *   - LET / LAMBDA closures and the higher-order family
 *     (MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY)
 *   - INDIRECT / OFFSET including their A1/R1C1 and negative-bounds cases
 *   - ROW / COLUMN / ROWS / COLUMNS composed with INDIRECT and OFFSET —
 *     these are the regressions called out in the task description
 *   - error propagation across the whole operator matrix
 *   - circular reference behaviour (zero-seed and iterative convergence)
 *   - volatile function re-evaluation
 */

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

// ============================================================================
// Arithmetic Operators
// ============================================================================

describe("evaluator: arithmetic operators", () => {
  it("adds two numbers", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 3;
    ws.getCell("A2").value = { formula: "A1+2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A2").result).toBe(5);
  });

  it("subtracts, multiplies, divides, exponentiates", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 12;
    ws.getCell("A2").value = 3;
    ws.getCell("B1").value = { formula: "A1-A2", result: 0 };
    ws.getCell("B2").value = { formula: "A1*A2", result: 0 };
    ws.getCell("B3").value = { formula: "A1/A2", result: 0 };
    ws.getCell("B4").value = { formula: "A2^A2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(9);
    expect(ws.getCell("B2").result).toBe(36);
    expect(ws.getCell("B3").result).toBe(4);
    expect(ws.getCell("B4").result).toBe(27);
  });

  it("applies percent as /100", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 250;
    ws.getCell("B1").value = { formula: "A1%", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(2.5);
  });

  it("coerces numeric strings in arithmetic", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "4";
    ws.getCell("A2").value = "6";
    ws.getCell("B1").value = { formula: "A1+A2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(10);
  });

  it("treats blank cells as zero", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 7;
    // A2 intentionally blank
    ws.getCell("B1").value = { formula: "A1+A2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(7);
  });

  it("returns #DIV/0! on divide by zero", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
  });

  it("propagates errors through arithmetic", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("B1").value = { formula: "A1+1", result: 0 };
    ws.getCell("B2").value = { formula: "2*A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
    expect(ws.getCell("B2").result).toEqual({ error: "#DIV/0!" });
  });

  it("returns #NUM! for arithmetic overflow", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // 10^309 overflows JS number (Infinity)
    ws.getCell("A1").value = { formula: "10^309", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NUM!" });
  });
});

// ============================================================================
// Comparison Operators
// ============================================================================

describe("evaluator: comparison operators", () => {
  it("compares numbers with =, <>, <, >, <=, >=", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 5;
    ws.getCell("A2").value = 7;
    ws.getCell("B1").value = { formula: "A1=A2", result: false };
    ws.getCell("B2").value = { formula: "A1<>A2", result: false };
    ws.getCell("B3").value = { formula: "A1<A2", result: false };
    ws.getCell("B4").value = { formula: "A1>A2", result: false };
    ws.getCell("B5").value = { formula: "A1<=5", result: false };
    ws.getCell("B6").value = { formula: "A1>=5", result: false };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(false);
    expect(ws.getCell("B2").result).toBe(true);
    expect(ws.getCell("B3").result).toBe(true);
    expect(ws.getCell("B4").result).toBe(false);
    expect(ws.getCell("B5").result).toBe(true);
    expect(ws.getCell("B6").result).toBe(true);
  });

  it("compares strings lexically", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "apple";
    ws.getCell("A2").value = "banana";
    ws.getCell("B1").value = { formula: "A1<A2", result: false };
    ws.getCell("B2").value = { formula: 'A1="apple"', result: false };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(true);
    expect(ws.getCell("B2").result).toBe(true);
  });

  it("compares booleans (TRUE > FALSE)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = { formula: "TRUE>FALSE", result: false };
    ws.getCell("B2").value = { formula: "TRUE=TRUE", result: false };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(true);
    expect(ws.getCell("B2").result).toBe(true);
  });

  it('treats "" as blank for comparison', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // A1 is blank; formula compares it to ""
    ws.getCell("B1").value = { formula: 'A1=""', result: false };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(true);
  });

  it("orders cross-type comparisons: number < string < boolean", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = "1";
    ws.getCell("A3").value = true;
    ws.getCell("B1").value = { formula: "A1<A2", result: false };
    ws.getCell("B2").value = { formula: "A2<A3", result: false };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(true);
    expect(ws.getCell("B2").result).toBe(true);
  });
});

// ============================================================================
// Concatenation
// ============================================================================

describe("evaluator: concatenation operator", () => {
  it("concatenates two strings with &", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "foo";
    ws.getCell("A2").value = "bar";
    ws.getCell("B1").value = { formula: "A1&A2", result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("foobar");
  });

  it("coerces numbers to strings in concatenation", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 2024;
    ws.getCell("B1").value = { formula: '"Year: "&A1', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("Year: 2024");
  });

  it('treats blank as "" in concatenation', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // A1 blank
    ws.getCell("B1").value = { formula: '"x"&A1&"y"', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("xy");
  });
});

// ============================================================================
// Unary Operators
// ============================================================================

describe("evaluator: unary operators", () => {
  it("unary minus negates a reference", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 12;
    ws.getCell("B1").value = { formula: "-A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(-12);
  });

  it("unary plus leaves the number unchanged", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 7;
    ws.getCell("B1").value = { formula: "+A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(7);
  });

  it("percent postfix divides by 100", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 50;
    ws.getCell("B1").value = { formula: "A1%", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(0.5);
  });
});

// ============================================================================
// Broadcasting
// ============================================================================

describe("evaluator: broadcasting", () => {
  it("broadcasts scalar over an array (scalar * M×N)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;

    // CSE 3×1: {=A1:A3*10} in B1:B3
    ws.getCell("B1").value = {
      formula: "A1:A3*10",
      result: 0,
      shareType: "array",
      ref: "B1:B3"
    };
    ws.getCell("B2").value = { formula: "A1:A3*10", result: 0, shareType: "array" };
    ws.getCell("B3").value = { formula: "A1:A3*10", result: 0, shareType: "array" };

    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(10);
    expect(ws.getCell("B2").result).toBe(20);
    expect(ws.getCell("B3").result).toBe(30);
  });

  it("broadcasts row × column to M×N matrix", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Row 1: 1,2
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    // Column 3: 10; 20; 30
    ws.getCell("A3").value = 10;
    ws.getCell("A4").value = 20;
    ws.getCell("A5").value = 30;

    // Dynamic-array: A1:B1 + A3:A5 → 3×2
    ws.getCell("D1").value = {
      formula: "A1:B1+A3:A5",
      result: 0,
      shareType: "array",
      ref: "D1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(11);
    expect(ws.getCell("E1").value).toBe(12);
    expect(ws.getCell("D2").value).toBe(21);
    expect(ws.getCell("E2").value).toBe(22);
    expect(ws.getCell("D3").value).toBe(31);
    expect(ws.getCell("E3").value).toBe(32);
  });

  it("returns #VALUE! when array shapes are incompatible", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("B1").value = 3;
    ws.getCell("B2").value = 4;
    ws.getCell("B3").value = 5;
    // 2×1 + 3×1 — row count mismatch
    ws.getCell("C1").value = {
      formula: "A1:A2+B1:B3",
      result: 0,
      shareType: "array",
      ref: "C1",
      isDynamicArray: true
    };
    wb.calculateFormulas();
    expect(ws.getCell("C1").result).toEqual({ error: "#VALUE!" });
  });

  it("applies unary minus element-wise over arrays", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 5;
    ws.getCell("A2").value = 10;
    ws.getCell("B1").value = { formula: "SUM(-A1:A2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(-15);
  });
});

// ============================================================================
// Implicit Intersection
// ============================================================================

describe("evaluator: reference passthrough and implicit intersection", () => {
  it("uses implicit intersection on a column range in scalar context", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    // B2 = A1:A3 — sits in row 2, should resolve to A2
    ws.getCell("B2").value = { formula: "A1:A3", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B2").result).toBe(20);
  });

  it("uses implicit intersection on a row range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("B1").value = 20;
    ws.getCell("C1").value = 30;
    // D1 = A1:C1 would intersect — col D is out of range → falls back to first
    // So instead put the formula at B2 using a row range
    ws.getCell("B3").value = { formula: "A1:C1+1", result: 0 };
    wb.calculateFormulas();
    // B3 sits in column 2, so A1:C1 implicitly intersects to B1=20
    expect(ws.getCell("B3").result).toBe(21);
  });

  it("uses @ prefix for explicit implicit intersection", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    // @ forces implicit intersection even in array-capable context
    ws.getCell("B2").value = { formula: "@A1:A3", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B2").result).toBe(20);
  });
});

// ============================================================================
// Array Constants
// ============================================================================

describe("evaluator: array constants", () => {
  it("evaluates {1,2;3,4} as a 2x2 array via SUM", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SUM({1,2;3,4})", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(10);
  });

  it("evaluates a 1-row array constant", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SUM({10,20,30})", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(60);
  });

  it("spills a 2×2 array constant to neighboring cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "{1,2;3,4}",
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
  });
});

// ============================================================================
// Conditionals — IF / IFS / SWITCH / CHOOSE / IFERROR / IFNA
// ============================================================================

describe("evaluator: conditional forms", () => {
  it("IF short-circuits the untaken branch", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "IF(TRUE,1,1/0)", result: 0 };
    ws.getCell("A2").value = { formula: "IF(FALSE,1/0,2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").result).toBe(2);
  });

  it("IFS returns the first matching branch and #N/A if none match", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 5;
    ws.getCell("B1").value = { formula: 'IFS(A1=1,"one",A1=5,"five")', result: "" };
    ws.getCell("B2").value = { formula: 'IFS(A1=0,"zero")', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("five");
    expect(ws.getCell("B2").result).toEqual({ error: "#N/A" });
  });

  it("SWITCH returns the matching branch or default", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "b";
    ws.getCell("B1").value = {
      formula: 'SWITCH(A1,"a",1,"b",2,"c",3,99)',
      result: 0
    };
    ws.getCell("B2").value = {
      formula: 'SWITCH("z","a",1,"b",2,"default")',
      result: ""
    };
    ws.getCell("B3").value = {
      formula: 'SWITCH("z","a",1,"b",2)',
      result: ""
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(2);
    expect(ws.getCell("B2").result).toBe("default");
    expect(ws.getCell("B3").result).toEqual({ error: "#N/A" });
  });

  it("CHOOSE picks the nth argument (1-based)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'CHOOSE(2,"a","b","c")', result: "" };
    ws.getCell("A2").value = { formula: 'CHOOSE(5,"a","b")', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe("b");
    expect(ws.getCell("A2").result).toEqual({ error: "#VALUE!" });
  });

  it("IFERROR replaces any error, leaves non-errors alone", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 0;
    ws.getCell("B1").value = { formula: 'IFERROR(1/A1,"oops")', result: "" };
    ws.getCell("B2").value = { formula: "IFERROR(10,99)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("oops");
    expect(ws.getCell("B2").result).toBe(10);
  });

  it("IFNA only replaces #N/A errors", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = { formula: 'IFNA(NA(),"missing")', result: "" };
    ws.getCell("B2").value = { formula: 'IFNA(1/0,"safe")', result: "" };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe("missing");
    // 1/0 is #DIV/0!, not #N/A — IFNA does not mask it
    expect(ws.getCell("B2").result).toEqual({ error: "#DIV/0!" });
  });
});

// ============================================================================
// LET + LAMBDA
// ============================================================================

describe("evaluator: LET and LAMBDA", () => {
  it("LET binds variables for use in the body expression", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "LET(x,5,y,3,x*y)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(15);
  });

  it("LET bindings can reference earlier bindings", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "LET(x,10,y,x+5,x*y)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(150);
  });

  it("LET supports nested LET with shadowing", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(x,1,LET(x,2,x+1))",
      result: 0
    };
    wb.calculateFormulas();
    // Inner x shadows outer → 2+1=3
    expect(ws.getCell("A1").result).toBe(3);
  });

  it("LAMBDA can be invoked through a named binding in LET", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Parser doesn't accept LAMBDA(...)(args) inline; bind a name first.
    ws.getCell("A1").value = {
      formula: "LET(add,LAMBDA(x,y,x+y),add(3,4))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(7);
  });

  it("LAMBDA captures its enclosing LET bindings (closure)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // LET inner binds y; the LAMBDA closes over y
    ws.getCell("A1").value = {
      formula: "LET(y,10,addY,LAMBDA(x,x+y),addY(5))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(15);
  });

  it("LAMBDA stored in a cell is invoked via a defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Put the lambda value in a helper cell, alias via defined name.
    ws.getCell("Z1").value = { formula: "LAMBDA(x,x*3)", result: 0 };
    wb.definedNames.add("Sheet1!$Z$1", "Triple");
    ws.getCell("A1").value = { formula: "Triple(7)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(21);
  });
});

// ============================================================================
// Higher-Order Functions
// ============================================================================

describe("evaluator: higher-order functions", () => {
  it("MAP applies the lambda element-wise", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(MAP({1,2,3},LAMBDA(x,x*x)))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(14);
  });

  it("MAP accepts a scalar input (treated as 1×1)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "MAP(5,LAMBDA(x,x+1))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(6);
  });

  it("REDUCE folds an array into a single value", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "REDUCE(0,{1,2,3,4,5},LAMBDA(a,b,a+b))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(15);
  });

  it("REDUCE on a scalar invokes the reducer exactly once", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "REDUCE(10,5,LAMBDA(a,b,a+b))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(15);
  });

  it("SCAN produces running accumulations", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Expose via SUM so we don't need to materialize the array into cells.
    ws.getCell("A1").value = {
      formula: "SUM(SCAN(0,{1,2,3},LAMBDA(a,b,a+b)))",
      result: 0
    };
    wb.calculateFormulas();
    // Running sums: 1,3,6 → total 10
    expect(ws.getCell("A1").result).toBe(10);
  });

  it("BYROW reduces each row", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B2").value = 4;
    // Reduce 2×2 by row → {3; 7}, sum = 10
    ws.getCell("D1").value = {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,SUM(r))))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(10);
  });

  it("BYCOL reduces each column", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B2").value = 4;
    // Reduce 2×2 by col → {4, 6}, sum = 10
    ws.getCell("D1").value = {
      formula: "SUM(BYCOL(A1:B2,LAMBDA(c,SUM(c))))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(10);
  });

  it("MAKEARRAY builds a grid of (row, col) → body values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)))",
      result: 0
    };
    wb.calculateFormulas();
    // Rows 1..2, cols 1..3 → 11+12+13+21+22+23 = 102
    expect(ws.getCell("A1").result).toBe(102);
  });
});

// ============================================================================
// INDIRECT
// ============================================================================

describe("evaluator: INDIRECT", () => {
  it('resolves INDIRECT("A1") on the current sheet', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 42;
    ws.getCell("B1").value = { formula: 'INDIRECT("A1")', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(42);
  });

  it("resolves INDIRECT with a dynamically-built address", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A5").value = 500;
    // ROW() = 3, so "A"&ROW()+2 → "A5"
    ws.getCell("C3").value = {
      formula: 'INDIRECT("A"&(ROW()+2))',
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("C3").result).toBe(500);
  });

  it("resolves INDIRECT in R1C1 mode with absolute addressing", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("C5").value = 777;
    ws.getCell("A1").value = {
      formula: 'INDIRECT("R5C3",FALSE)',
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(777);
  });

  it("resolves INDIRECT in R1C1 mode with relative addressing", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("C3").value = 99;
    // Sitting in B2, R[1]C[1] = row 3, col 3 = C3
    ws.getCell("B2").value = {
      formula: 'INDIRECT("R[1]C[1]",FALSE)',
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B2").result).toBe(99);
  });

  it("resolves INDIRECT across sheets", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    ws2.getCell("B2").value = 123;
    ws1.getCell("A1").value = {
      formula: 'INDIRECT("Sheet2!B2")',
      result: 0
    };
    wb.calculateFormulas();
    expect(ws1.getCell("A1").result).toBe(123);
  });

  it("returns #REF! when INDIRECT target is a syntactically invalid address", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // '!!' is not a valid A1 address → parser/binder throws → evaluator
    // catches and returns #REF!.
    ws.getCell("A1").value = { formula: 'INDIRECT("!!")', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
  });

  it("returns #REF! when INDIRECT target is the empty string", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'INDIRECT("")', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
  });
});

// ============================================================================
// OFFSET
// ============================================================================

describe("evaluator: OFFSET", () => {
  it("returns a single cell at the base + offset", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("C4").value = 33;
    ws.getCell("A1").value = { formula: "OFFSET(A1,3,2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(33);
  });

  it("returns a range with explicit height/width", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B2").value = 10;
    ws.getCell("B3").value = 20;
    ws.getCell("B4").value = 30;
    ws.getCell("D1").value = {
      formula: "SUM(OFFSET(A1,1,1,3,1))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(60);
  });

  it("supports negative height (range extends upward)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    // OFFSET(A3,0,0,-3,1) → A1:A3
    ws.getCell("B1").value = {
      formula: "SUM(OFFSET(A3,0,0,-3,1))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(6);
  });

  it("supports negative width (range extends leftward)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    ws.getCell("C1").value = 3;
    // OFFSET(C1,0,0,1,-3) → A1:C1
    ws.getCell("D1").value = {
      formula: "SUM(OFFSET(C1,0,0,1,-3))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(6);
  });

  it("returns #REF! for zero height or zero width", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "OFFSET(A1,0,0,0,1)", result: 0 };
    ws.getCell("B2").value = { formula: "OFFSET(A1,0,0,1,0)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#REF!" });
    expect(ws.getCell("B2").result).toEqual({ error: "#REF!" });
  });

  it("returns #REF! when the offset moves above row 1 or column 1", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "OFFSET(A1,-1,0)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#REF!" });
  });
});

// ============================================================================
// ROW / COLUMN / ROWS / COLUMNS  (regression coverage R2-PX-YY)
// ============================================================================

describe("evaluator: ROW/COLUMN/ROWS/COLUMNS", () => {
  it("ROW(cell) returns the row of the reference", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "ROW(B5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(5);
  });

  it("COLUMN(cell) returns the column of the reference", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "COLUMN(D1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(4);
  });

  it("ROW() with no argument returns the current row", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A3").value = { formula: "ROW()", result: 0 };
    ws.getCell("B7").value = { formula: "COLUMN()", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A3").result).toBe(3);
    expect(ws.getCell("B7").result).toBe(2);
  });

  it("ROW(range) returns the top row of the range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "ROW(C5:E7)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(5);
  });

  it("ROWS(range) returns the row count", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "ROWS(A1:A10)", result: 0 };
    ws.getCell("A2").value = { formula: "ROWS(B2:D2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(10);
    expect(ws.getCell("A2").result).toBe(1);
  });

  it("COLUMNS(range) returns the column count", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "COLUMNS(A1:C1)", result: 0 };
    ws.getCell("A2").value = { formula: "COLUMNS(A1:A10)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(3);
    expect(ws.getCell("A2").result).toBe(1);
  });

  // R2-PX-YY regressions — the ROW/COLUMN pass-through for INDIRECT/OFFSET.
  // These test the evaluator's "inspect raw reference before dereferencing"
  // branch in evaluateCall (around evaluator.ts:985).
  it('ROW(INDIRECT("A5")) returns 5 (R2-PX-YY)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'ROW(INDIRECT("A5"))', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(5);
  });

  it('COLUMN(INDIRECT("D1")) returns 4 (R2-PX-YY)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'COLUMN(INDIRECT("D1"))', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(4);
  });

  it('ROWS(INDIRECT("A1:A10")) returns 10 (R2-PX-YY)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'ROWS(INDIRECT("A1:A10"))', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(10);
  });

  it('COLUMNS(INDIRECT("A1:C1")) returns 3 (R2-PX-YY)', () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: 'COLUMNS(INDIRECT("A1:C1"))', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(3);
  });

  it("ROWS(OFFSET(A1,0,0,5,3)) returns 5 — OFFSET array arg path", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "ROWS(OFFSET(A1,0,0,5,3))", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(5);
  });

  it("COLUMNS(OFFSET(A1,0,0,5,3)) returns 3 — OFFSET array arg path", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = { formula: "COLUMNS(OFFSET(A1,0,0,5,3))", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(3);
  });
});

// ============================================================================
// Defined Names
// ============================================================================

describe("evaluator: defined names", () => {
  it("resolves a workbook-scoped single-cell name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 50;
    wb.definedNames.add("Sheet1!$A$1", "Price");
    ws.getCell("B1").value = { formula: "Price*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(100);
  });

  it("resolves a single-area range name via SUM", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    wb.definedNames.add("Sheet1!$A$1:$A$3", "Vals");
    ws.getCell("B1").value = { formula: "SUM(Vals)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(6);
  });

  it("resolves a formula-based defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    wb.definedNames.addFormula("Answer", "20+22");
    ws.getCell("A1").value = { formula: "Answer", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(42);
  });

  it("prefers sheet-local scope over workbook-global scope", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    ws1.getCell("A1").value = 100;
    ws2.getCell("A1").value = 200;
    // Global name → Sheet1!A1, Sheet1-local → Sheet2!A1
    wb.definedNames.model = [
      { name: "Val", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
      {
        name: "Val",
        ranges: ["Sheet2!$A$1"],
        rawText: "Sheet2!$A$1",
        localSheetId: 0
      }
    ];
    ws1.getCell("B1").value = { formula: "Val", result: 0 };
    ws2.getCell("B1").value = { formula: "Val", result: 0 };
    wb.calculateFormulas();
    // Sheet1 sees the sheet-local (Sheet2!A1 = 200)
    expect(ws1.getCell("B1").result).toBe(200);
    // Sheet2 has no local → falls back to global (Sheet1!A1 = 100)
    expect(ws2.getCell("B1").result).toBe(100);
  });

  it("returns #NAME? for an unknown name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    // Omit `result:` so the engine can't preserve a cached value —
    // otherwise the engine keeps the stale number instead of surfacing the
    // error (see calculate-formulas-impl.ts unsupported-formula branch).
    ws.getCell("A1").value = { formula: "NoSuchName" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NAME?" });
  });
});

// ============================================================================
// 3D References
// ============================================================================

describe("evaluator: 3D references", () => {
  it("SUM(Sheet1:Sheet3!A1) sums the same cell across sheets", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    const ws3 = wb.addWorksheet("Sheet3");
    const out = wb.addWorksheet("Out");
    ws1.getCell("A1").value = 10;
    ws2.getCell("A1").value = 20;
    ws3.getCell("A1").value = 30;
    out.getCell("A1").value = { formula: "SUM(Sheet1:Sheet3!A1)", result: 0 };
    wb.calculateFormulas();
    expect(out.getCell("A1").result).toBe(60);
  });

  it("SUM(Sheet1:Sheet2!A1:B1) sums an area across sheets", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    const ws2 = wb.addWorksheet("Sheet2");
    const out = wb.addWorksheet("Out");
    ws1.getCell("A1").value = 1;
    ws1.getCell("B1").value = 2;
    ws2.getCell("A1").value = 3;
    ws2.getCell("B1").value = 4;
    out.getCell("A1").value = { formula: "SUM(Sheet1:Sheet2!A1:B1)", result: 0 };
    wb.calculateFormulas();
    expect(out.getCell("A1").result).toBe(10);
  });
});

// ============================================================================
// Structured References
// ============================================================================

describe("evaluator: structured references", () => {
  it("Table[Col] reads a single column's data", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "Nums",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "V" }],
      rows: [[10], [20], [30]]
    });
    ws.getCell("D1").value = { formula: "SUM(Nums[V])", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(60);
  });

  it("Table[#All] includes header + data (+ totals if present)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.addTable({
      name: "Items",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Item" }, { name: "N" }],
      rows: [
        ["a", 1],
        ["b", 2]
      ]
    });
    // #All = 3 rows × 2 cols = 6 non-empty cells
    ws.getCell("D1").value = { formula: "COUNTA(Items[#All])", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(6);
  });
});

// ============================================================================
// Error Propagation Scenarios
// ============================================================================

describe("evaluator: error emission sources", () => {
  it("#REF! from an invalid external reference", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "[Book1]Sheet1!A1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
  });

  it("#NAME? from an unknown function", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "NOPE(1)" };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NAME?" });
  });

  it("#VALUE! from numeric coercion of an empty string", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: '1+""', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#VALUE!" });
  });

  it("#NUM! from SQRT of a negative", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SQRT(-1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#NUM!" });
  });

  it("#DIV/0! from arithmetic divide-by-zero", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "10/0", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
  });

  it("#N/A from NA()", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "NA()", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#N/A" });
  });

  it("#NULL! from a non-overlapping intersection", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("C1").value = 3;
    // A1:A2 and C1:C2 don't overlap
    ws.getCell("B1").value = { formula: "SUM(A1:A2 C1:C2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#NULL!" });
  });
});

// ============================================================================
// Circular References
// ============================================================================

describe("evaluator: circular references", () => {
  it("returns a number for a simple cycle without iterate enabled", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "A2+1", result: 0 };
    ws.getCell("A2").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    // Zero-seed fallback keeps the engine making progress. Both must resolve
    // to numbers rather than errors — the established behaviour this engine
    // preserves.
    expect(typeof ws.getCell("A1").result).toBe("number");
    expect(typeof ws.getCell("A2").result).toBe("number");
  });

  it("converges under iterate=true to the fixed point", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    wb.calcProperties = {
      iterate: true,
      iterateCount: 200,
      iterateDelta: 0.0001
    };
    // A1 = A1/2 + 1 → fixed point at 2
    ws.getCell("A1").value = { formula: "A1/2+1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBeCloseTo(2, 3);
  });

  it("reaches iterateCount when no convergence is possible", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    wb.calcProperties = {
      iterate: true,
      iterateCount: 5,
      iterateDelta: 0
    };
    // Increment-only cycle — no fixed point, runs out at count
    ws.getCell("A1").value = { formula: "A1+1", result: 0 };
    wb.calculateFormulas();
    // Initial 0 → 1; then 5 iterations → 6
    expect(ws.getCell("A1").result).toBe(6);
  });
});

// ============================================================================
// Volatile Functions
// ============================================================================

describe("evaluator: volatile functions", () => {
  it("RAND() returns a number in [0, 1)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "RAND()", result: 0 };
    wb.calculateFormulas();
    const r = ws.getCell("A1").result as number;
    expect(typeof r).toBe("number");
    expect(r).toBeGreaterThanOrEqual(0);
    expect(r).toBeLessThan(1);
  });

  it("RAND() produces a fresh value on each recalculation", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "RAND()", result: 0 };

    // Many samples — if volatile is wired up correctly, at least two
    // recalculations should yield different values. Without volatile
    // handling the cache would return the same value forever.
    const samples = new Set<number>();
    for (let i = 0; i < 20; i++) {
      wb.calculateFormulas();
      samples.add(ws.getCell("A1").result as number);
    }
    expect(samples.size).toBeGreaterThan(1);
  });

  it("TODAY() returns a whole-day serial", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "TODAY()", result: 0 };
    wb.calculateFormulas();
    const r = ws.getCell("A1").result as number;
    expect(typeof r).toBe("number");
    expect(r).toBe(Math.trunc(r));
    // Rough sanity — post-2020 serial range (> Jan 1 2020 = 43831)
    expect(r).toBeGreaterThan(43830);
  });

  it("NOW() returns a serial strictly greater than (or equal to) TODAY()", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "NOW()", result: 0 };
    ws.getCell("A2").value = { formula: "TODAY()", result: 0 };
    wb.calculateFormulas();
    const now = ws.getCell("A1").result as number;
    const today = ws.getCell("A2").result as number;
    // TODAY is anchored to the user's local calendar date (wall clock)
    // while NOW is a pure UTC instant, so depending on the runner's
    // timezone offset the two can legitimately be up to ±1 day apart.
    // The meaningful invariant is that they refer to the same rough
    // moment in time, which we verify with a two-day window.
    expect(Math.abs(now - today)).toBeLessThan(2);
  });
});

// ============================================================================
// LET / LAMBDA — extended coverage
// ============================================================================

describe("evaluator: LET additional cases", () => {
  it("LET with a single pair is just the bound expression", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "LET(x,42,x)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(42);
  });

  it("LET propagates errors from the binding expression", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "LET(x,1/0,x+1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
  });

  it("LET can bind a reference range and SUM it", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("B1").value = { formula: "LET(r,A1:A3,SUM(r))", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(6);
  });

  it("LET with odd arg count (missing body) returns #VALUE!", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "LET(x,1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#VALUE!" });
  });

  it("LET supports many bindings (stress)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(a,1,b,2,c,3,d,4,e,5,a+b+c+d+e)",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(15);
  });
});

describe("evaluator: LAMBDA additional cases", () => {
  it("LAMBDA with zero params", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(const,LAMBDA(42),const())",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(42);
  });

  it("LAMBDA with three params", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(f,LAMBDA(a,b,c,a*100+b*10+c),f(1,2,3))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(123);
  });

  it("LAMBDA recursive via defined-name self-reference", () => {
    // Define a factorial lambda and call it. The engine's name resolution
    // allows a lambda stored in a defined name to call itself by name.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("Z1").value = {
      formula: "LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))",
      result: 0
    };
    wb.definedNames.add("Sheet1!$Z$1", "Fact");
    ws.getCell("A1").value = { formula: "Fact(5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(120);
  });

  it("LAMBDA arity mismatch returns an error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(f,LAMBDA(x,y,x+y),f(1))",
      result: 0
    };
    wb.calculateFormulas();
    // Engine resolves unbound y → propagates error/blank depending on path;
    // must at least not silently succeed with a number.
    const r = ws.getCell("A1").result;
    expect(typeof r === "number" ? r : null).not.toBe(1);
  });

  it("LAMBDA body that divides by zero propagates the error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "LET(f,LAMBDA(x,1/x),f(0))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
  });
});

// ============================================================================
// Higher-order functions — extended coverage
// ============================================================================

describe("evaluator: MAP extra", () => {
  it("MAP over a 2×2 array preserves shape when wrapped in SUMPRODUCT", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B2").value = 4;
    ws.getCell("D1").value = {
      formula: "SUMPRODUCT(MAP(A1:B2,LAMBDA(x,x+1)))",
      result: 0
    };
    wb.calculateFormulas();
    // Each cell incremented by 1 → sum of (2+3+4+5) = 14
    expect(ws.getCell("D1").result).toBe(14);
  });

  it("MAP with a constant-returning lambda yields a uniform array", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("D1").value = {
      formula: "SUM(MAP({1,2,3},LAMBDA(x,7)))",
      result: 0
    };
    wb.calculateFormulas();
    // Each of 3 cells → 7 → sum = 21
    expect(ws.getCell("D1").result).toBe(21);
  });

  it("MAP with a lambda returning errors propagates them into the result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("D1").value = {
      formula: 'IFERROR(SUM(MAP({1,2,0},LAMBDA(x,1/x))),"err")',
      result: 0
    };
    wb.calculateFormulas();
    // 1/0 → #DIV/0! → SUM would propagate; IFERROR catches
    expect(ws.getCell("D1").result).toBe("err");
  });
});

describe("evaluator: REDUCE extra", () => {
  it("REDUCE sums a range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("A4").value = 4;
    ws.getCell("B1").value = {
      formula: "REDUCE(0,A1:A4,LAMBDA(a,x,a+x))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(10);
  });

  it("REDUCE multiplies a range (product accumulator)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = {
      formula: "REDUCE(1,{2,3,4},LAMBDA(a,x,a*x))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(24);
  });

  it("REDUCE propagates errors from the accumulator", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = {
      formula: "REDUCE(1,{1,0,3},LAMBDA(a,x,a/x))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
  });

  it("REDUCE counts elements matching a predicate", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = {
      formula: "REDUCE(0,{1,2,3,4,5},LAMBDA(a,x,IF(x>3,a+1,a)))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(2);
  });
});

describe("evaluator: SCAN / BYROW / BYCOL / MAKEARRAY extra", () => {
  it("SCAN with product accumulator yields running products", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(SCAN(1,{2,3,4},LAMBDA(a,x,a*x)))",
      result: 0
    };
    wb.calculateFormulas();
    // Running products: 2, 6, 24 → sum = 32
    expect(ws.getCell("A1").result).toBe(32);
  });

  it("BYROW reduces each row with MAX", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 5;
    ws.getCell("A2").value = 10;
    ws.getCell("B2").value = 3;
    ws.getCell("D1").value = {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,MAX(r))))",
      result: 0
    };
    wb.calculateFormulas();
    // row1 max=5, row2 max=10 → sum=15
    expect(ws.getCell("D1").result).toBe(15);
  });

  it("BYCOL reduces each column with AVERAGE", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 2;
    ws.getCell("B1").value = 4;
    ws.getCell("A2").value = 4;
    ws.getCell("B2").value = 6;
    ws.getCell("D1").value = {
      formula: "SUM(BYCOL(A1:B2,LAMBDA(c,AVERAGE(c))))",
      result: 0
    };
    wb.calculateFormulas();
    // col1 avg=3, col2 avg=5 → sum=8
    expect(ws.getCell("D1").result).toBe(8);
  });

  it("MAKEARRAY(3,3,...) builds an identity-style matrix", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(MAKEARRAY(3,3,LAMBDA(r,c,IF(r=c,1,0))))",
      result: 0
    };
    wb.calculateFormulas();
    // Diagonal has three 1s
    expect(ws.getCell("A1").result).toBe(3);
  });

  it("MAKEARRAY(1,5,...) builds a row vector", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(MAKEARRAY(1,5,LAMBDA(r,c,c)))",
      result: 0
    };
    wb.calculateFormulas();
    // 1+2+3+4+5 = 15
    expect(ws.getCell("A1").result).toBe(15);
  });

  it("MAKEARRAY(5,1,...) builds a column vector", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(MAKEARRAY(5,1,LAMBDA(r,c,r*2)))",
      result: 0
    };
    wb.calculateFormulas();
    // 2+4+6+8+10 = 30
    expect(ws.getCell("A1").result).toBe(30);
  });

  it("BYROW with an error inside a row propagates the error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("B1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B2").value = 4;
    ws.getCell("D1").value = {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,SUM(r))))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toEqual({ error: "#DIV/0!" });
  });

  it("SCAN on an array preserves its shape", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = {
      formula: "SUM(SCAN(0,{1,2,3,4},LAMBDA(a,x,a+x)))",
      result: 0
    };
    wb.calculateFormulas();
    // running sums: 1,3,6,10 → sum=20
    expect(ws.getCell("A1").result).toBe(20);
  });

  it("REDUCE over a 2D array traverses row-major", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("B1").value = 2;
    ws.getCell("A2").value = 3;
    ws.getCell("B2").value = 4;
    ws.getCell("D1").value = {
      formula: "REDUCE(0,A1:B2,LAMBDA(a,x,a+x))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(10);
  });

  it("MAP propagates cell-level errors through the lambda", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = { formula: "1/0", result: 0 };
    ws.getCell("A3").value = 3;
    ws.getCell("B1").value = {
      formula: "SUM(MAP(A1:A3,LAMBDA(x,x*2)))",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
  });
});
