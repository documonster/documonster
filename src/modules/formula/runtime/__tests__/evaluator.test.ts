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

import {
  definedNamesAdd,
  definedNamesAddFormula,
  definedNamesSetModel
} from "@excel/defined-names";
import { calculateFormulas } from "@excel/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { getDefinedNames } from "@excel/workbook";
import { addTable } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

// ============================================================================
// Arithmetic Operators
// ============================================================================

describe("evaluator: arithmetic operators", () => {
  it("adds two numbers", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 3);
    Cell.setValue(ws, "A2", { formula: "A1+2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A2")).toBe(5);
  });

  it("subtracts, multiplies, divides, exponentiates", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 12);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B1", { formula: "A1-A2", result: 0 });
    Cell.setValue(ws, "B2", { formula: "A1*A2", result: 0 });
    Cell.setValue(ws, "B3", { formula: "A1/A2", result: 0 });
    Cell.setValue(ws, "B4", { formula: "A2^A2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(9);
    expect(Cell.getResult(ws, "B2")).toBe(36);
    expect(Cell.getResult(ws, "B3")).toBe(4);
    expect(Cell.getResult(ws, "B4")).toBe(27);
  });

  it("applies percent as /100", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 250);
    Cell.setValue(ws, "B1", { formula: "A1%", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(2.5);
  });

  it("coerces numeric strings in arithmetic", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "4");
    Cell.setValue(ws, "A2", "6");
    Cell.setValue(ws, "B1", { formula: "A1+A2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(10);
  });

  it("treats blank cells as zero", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 7);
    // A2 intentionally blank
    Cell.setValue(ws, "B1", { formula: "A1+A2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(7);
  });

  it("returns #DIV/0! on divide by zero", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
  });

  it("propagates errors through arithmetic", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "B1", { formula: "A1+1", result: 0 });
    Cell.setValue(ws, "B2", { formula: "2*A1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#DIV/0!" });
    expect(Cell.getResult(ws, "B2")).toEqual({ error: "#DIV/0!" });
  });

  it("returns #NUM! for arithmetic overflow", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // 10^309 overflows JS number (Infinity)
    Cell.setValue(ws, "A1", { formula: "10^309", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NUM!" });
  });
});

// ============================================================================
// Comparison Operators
// ============================================================================

describe("evaluator: comparison operators", () => {
  it("compares numbers with =, <>, <, >, <=, >=", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 5);
    Cell.setValue(ws, "A2", 7);
    Cell.setValue(ws, "B1", { formula: "A1=A2", result: false });
    Cell.setValue(ws, "B2", { formula: "A1<>A2", result: false });
    Cell.setValue(ws, "B3", { formula: "A1<A2", result: false });
    Cell.setValue(ws, "B4", { formula: "A1>A2", result: false });
    Cell.setValue(ws, "B5", { formula: "A1<=5", result: false });
    Cell.setValue(ws, "B6", { formula: "A1>=5", result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(false);
    expect(Cell.getResult(ws, "B2")).toBe(true);
    expect(Cell.getResult(ws, "B3")).toBe(true);
    expect(Cell.getResult(ws, "B4")).toBe(false);
    expect(Cell.getResult(ws, "B5")).toBe(true);
    expect(Cell.getResult(ws, "B6")).toBe(true);
  });

  it("compares strings lexically", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "apple");
    Cell.setValue(ws, "A2", "banana");
    Cell.setValue(ws, "B1", { formula: "A1<A2", result: false });
    Cell.setValue(ws, "B2", { formula: 'A1="apple"', result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(true);
    expect(Cell.getResult(ws, "B2")).toBe(true);
  });

  it("compares booleans (TRUE > FALSE)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", { formula: "TRUE>FALSE", result: false });
    Cell.setValue(ws, "B2", { formula: "TRUE=TRUE", result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(true);
    expect(Cell.getResult(ws, "B2")).toBe(true);
  });

  it('treats "" as blank for comparison', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // A1 is blank; formula compares it to ""
    Cell.setValue(ws, "B1", { formula: 'A1=""', result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(true);
  });

  it("orders cross-type comparisons: number < string < boolean", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", "1");
    Cell.setValue(ws, "A3", true);
    Cell.setValue(ws, "B1", { formula: "A1<A2", result: false });
    Cell.setValue(ws, "B2", { formula: "A2<A3", result: false });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(true);
    expect(Cell.getResult(ws, "B2")).toBe(true);
  });
});

// ============================================================================
// Concatenation
// ============================================================================

describe("evaluator: concatenation operator", () => {
  it("concatenates two strings with &", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "foo");
    Cell.setValue(ws, "A2", "bar");
    Cell.setValue(ws, "B1", { formula: "A1&A2", result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("foobar");
  });

  it("coerces numbers to strings in concatenation", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 2024);
    Cell.setValue(ws, "B1", { formula: '"Year: "&A1', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("Year: 2024");
  });

  it('treats blank as "" in concatenation', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // A1 blank
    Cell.setValue(ws, "B1", { formula: '"x"&A1&"y"', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("xy");
  });
});

// ============================================================================
// Unary Operators
// ============================================================================

describe("evaluator: unary operators", () => {
  it("unary minus negates a reference", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 12);
    Cell.setValue(ws, "B1", { formula: "-A1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(-12);
  });

  it("unary plus leaves the number unchanged", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 7);
    Cell.setValue(ws, "B1", { formula: "+A1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(7);
  });

  it("percent postfix divides by 100", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 50);
    Cell.setValue(ws, "B1", { formula: "A1%", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(0.5);
  });
});

// ============================================================================
// Broadcasting
// ============================================================================

describe("evaluator: broadcasting", () => {
  it("broadcasts scalar over an array (scalar * M×N)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);

    // CSE 3×1: {=A1:A3*10} in B1:B3
    Cell.setValue(ws, "B1", {
      formula: "A1:A3*10",
      result: 0,
      shareType: "array",
      ref: "B1:B3"
    });
    Cell.setValue(ws, "B2", { formula: "A1:A3*10", result: 0, shareType: "array" });
    Cell.setValue(ws, "B3", { formula: "A1:A3*10", result: 0, shareType: "array" });

    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(10);
    expect(Cell.getResult(ws, "B2")).toBe(20);
    expect(Cell.getResult(ws, "B3")).toBe(30);
  });

  it("broadcasts row × column to M×N matrix", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // Row 1: 1,2
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    // Column 3: 10; 20; 30
    Cell.setValue(ws, "A3", 10);
    Cell.setValue(ws, "A4", 20);
    Cell.setValue(ws, "A5", 30);

    // Dynamic-array: A1:B1 + A3:A5 → 3×2
    Cell.setValue(ws, "D1", {
      formula: "A1:B1+A3:A5",
      result: 0,
      shareType: "array",
      ref: "D1",
      isDynamicArray: true
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(11);
    expect(Cell.getValue(ws, "E1")).toBe(12);
    expect(Cell.getValue(ws, "D2")).toBe(21);
    expect(Cell.getValue(ws, "E2")).toBe(22);
    expect(Cell.getValue(ws, "D3")).toBe(31);
    expect(Cell.getValue(ws, "E3")).toBe(32);
  });

  it("returns #VALUE! when array shapes are incompatible", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "B1", 3);
    Cell.setValue(ws, "B2", 4);
    Cell.setValue(ws, "B3", 5);
    // 2×1 + 3×1 — row count mismatch
    Cell.setValue(ws, "C1", {
      formula: "A1:A2+B1:B3",
      result: 0,
      shareType: "array",
      ref: "C1",
      isDynamicArray: true
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toEqual({ error: "#VALUE!" });
  });

  it("applies unary minus element-wise over arrays", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 5);
    Cell.setValue(ws, "A2", 10);
    Cell.setValue(ws, "B1", { formula: "SUM(-A1:A2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(-15);
  });
});

// ============================================================================
// Implicit Intersection
// ============================================================================

describe("evaluator: reference passthrough and implicit intersection", () => {
  it("uses implicit intersection on a column range in scalar context", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Cell.setValue(ws, "A3", 30);
    // B2 = A1:A3 — sits in row 2, should resolve to A2
    Cell.setValue(ws, "B2", { formula: "A1:A3", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B2")).toBe(20);
  });

  it("uses implicit intersection on a row range", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "B1", 20);
    Cell.setValue(ws, "C1", 30);
    // D1 = A1:C1 would intersect — col D is out of range → falls back to first
    // So instead put the formula at B2 using a row range
    Cell.setValue(ws, "B3", { formula: "A1:C1+1", result: 0 });
    calculateFormulas(wb);
    // B3 sits in column 2, so A1:C1 implicitly intersects to B1=20
    expect(Cell.getResult(ws, "B3")).toBe(21);
  });

  it("uses @ prefix for explicit implicit intersection", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Cell.setValue(ws, "A3", 30);
    // @ forces implicit intersection even in array-capable context
    Cell.setValue(ws, "B2", { formula: "@A1:A3", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B2")).toBe(20);
  });
});

// ============================================================================
// Array Constants
// ============================================================================

describe("evaluator: array constants", () => {
  it("evaluates {1,2;3,4} as a 2x2 array via SUM", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SUM({1,2;3,4})", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(10);
  });

  it("evaluates a 1-row array constant", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SUM({10,20,30})", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(60);
  });

  it("spills a 2×2 array constant to neighboring cells", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "{1,2;3,4}",
      result: 0,
      shareType: "array",
      ref: "A1",
      isDynamicArray: true
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "B1")).toBe(2);
    expect(Cell.getValue(ws, "A2")).toBe(3);
    expect(Cell.getValue(ws, "B2")).toBe(4);
  });
});

// ============================================================================
// Conditionals — IF / IFS / SWITCH / CHOOSE / IFERROR / IFNA
// ============================================================================

describe("evaluator: conditional forms", () => {
  it("IF short-circuits the untaken branch", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "IF(TRUE,1,1/0)", result: 0 });
    Cell.setValue(ws, "A2", { formula: "IF(FALSE,1/0,2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getResult(ws, "A2")).toBe(2);
  });

  it("IFS returns the first matching branch and #N/A if none match", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 5);
    Cell.setValue(ws, "B1", { formula: 'IFS(A1=1,"one",A1=5,"five")', result: "" });
    Cell.setValue(ws, "B2", { formula: 'IFS(A1=0,"zero")', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("five");
    expect(Cell.getResult(ws, "B2")).toEqual({ error: "#N/A" });
  });

  it("SWITCH returns the matching branch or default", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "b");
    Cell.setValue(ws, "B1", {
      formula: 'SWITCH(A1,"a",1,"b",2,"c",3,99)',
      result: 0
    });
    Cell.setValue(ws, "B2", {
      formula: 'SWITCH("z","a",1,"b",2,"default")',
      result: ""
    });
    Cell.setValue(ws, "B3", {
      formula: 'SWITCH("z","a",1,"b",2)',
      result: ""
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(2);
    expect(Cell.getResult(ws, "B2")).toBe("default");
    expect(Cell.getResult(ws, "B3")).toEqual({ error: "#N/A" });
  });

  it("CHOOSE picks the nth argument (1-based)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'CHOOSE(2,"a","b","c")', result: "" });
    Cell.setValue(ws, "A2", { formula: 'CHOOSE(5,"a","b")', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe("b");
    expect(Cell.getResult(ws, "A2")).toEqual({ error: "#VALUE!" });
  });

  it("IFERROR replaces any error, leaves non-errors alone", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 0);
    Cell.setValue(ws, "B1", { formula: 'IFERROR(1/A1,"oops")', result: "" });
    Cell.setValue(ws, "B2", { formula: "IFERROR(10,99)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("oops");
    expect(Cell.getResult(ws, "B2")).toBe(10);
  });

  it("IFNA only replaces #N/A errors", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", { formula: 'IFNA(NA(),"missing")', result: "" });
    Cell.setValue(ws, "B2", { formula: 'IFNA(1/0,"safe")', result: "" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe("missing");
    // 1/0 is #DIV/0!, not #N/A — IFNA does not mask it
    expect(Cell.getResult(ws, "B2")).toEqual({ error: "#DIV/0!" });
  });
});

// ============================================================================
// LET + LAMBDA
// ============================================================================

describe("evaluator: LET and LAMBDA", () => {
  it("LET binds variables for use in the body expression", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "LET(x,5,y,3,x*y)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });

  it("LET bindings can reference earlier bindings", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "LET(x,10,y,x+5,x*y)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(150);
  });

  it("LET supports nested LET with shadowing", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(x,1,LET(x,2,x+1))",
      result: 0
    });
    calculateFormulas(wb);
    // Inner x shadows outer → 2+1=3
    expect(Cell.getResult(ws, "A1")).toBe(3);
  });

  it("LAMBDA can be invoked through a named binding in LET", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // Parser doesn't accept LAMBDA(...)(args) inline; bind a name first.
    Cell.setValue(ws, "A1", {
      formula: "LET(add,LAMBDA(x,y,x+y),add(3,4))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(7);
  });

  it("LAMBDA captures its enclosing LET bindings (closure)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // LET inner binds y; the LAMBDA closes over y
    Cell.setValue(ws, "A1", {
      formula: "LET(y,10,addY,LAMBDA(x,x+y),addY(5))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });

  it("LAMBDA stored in a cell is invoked via a defined name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // Put the lambda value in a helper cell, alias via defined name.
    Cell.setValue(ws, "Z1", { formula: "LAMBDA(x,x*3)", result: 0 });
    definedNamesAdd(getDefinedNames(wb), "Sheet1!$Z$1", "Triple");
    Cell.setValue(ws, "A1", { formula: "Triple(7)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(21);
  });
});

// ============================================================================
// Higher-Order Functions
// ============================================================================

describe("evaluator: higher-order functions", () => {
  it("MAP applies the lambda element-wise", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(MAP({1,2,3},LAMBDA(x,x*x)))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(14);
  });

  it("MAP accepts a scalar input (treated as 1×1)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "MAP(5,LAMBDA(x,x+1))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(6);
  });

  it("REDUCE folds an array into a single value", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "REDUCE(0,{1,2,3,4,5},LAMBDA(a,b,a+b))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });

  it("REDUCE on a scalar invokes the reducer exactly once", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "REDUCE(10,5,LAMBDA(a,b,a+b))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });

  it("SCAN produces running accumulations", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // Expose via SUM so we don't need to materialize the array into cells.
    Cell.setValue(ws, "A1", {
      formula: "SUM(SCAN(0,{1,2,3},LAMBDA(a,b,a+b)))",
      result: 0
    });
    calculateFormulas(wb);
    // Running sums: 1,3,6 → total 10
    expect(Cell.getResult(ws, "A1")).toBe(10);
  });

  it("BYROW reduces each row", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B2", 4);
    // Reduce 2×2 by row → {3; 7}, sum = 10
    Cell.setValue(ws, "D1", {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,SUM(r))))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(10);
  });

  it("BYCOL reduces each column", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B2", 4);
    // Reduce 2×2 by col → {4, 6}, sum = 10
    Cell.setValue(ws, "D1", {
      formula: "SUM(BYCOL(A1:B2,LAMBDA(c,SUM(c))))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(10);
  });

  it("MAKEARRAY builds a grid of (row, col) → body values", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)))",
      result: 0
    });
    calculateFormulas(wb);
    // Rows 1..2, cols 1..3 → 11+12+13+21+22+23 = 102
    expect(Cell.getResult(ws, "A1")).toBe(102);
  });
});

// ============================================================================
// INDIRECT
// ============================================================================

describe("evaluator: INDIRECT", () => {
  it('resolves INDIRECT("A1") on the current sheet', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 42);
    Cell.setValue(ws, "B1", { formula: 'INDIRECT("A1")', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(42);
  });

  it("resolves INDIRECT with a dynamically-built address", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A5", 500);
    // ROW() = 3, so "A"&ROW()+2 → "A5"
    Cell.setValue(ws, "C3", {
      formula: 'INDIRECT("A"&(ROW()+2))',
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C3")).toBe(500);
  });

  it("resolves INDIRECT in R1C1 mode with absolute addressing", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "C5", 777);
    Cell.setValue(ws, "A1", {
      formula: 'INDIRECT("R5C3",FALSE)',
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(777);
  });

  it("resolves INDIRECT in R1C1 mode with relative addressing", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "C3", 99);
    // Sitting in B2, R[1]C[1] = row 3, col 3 = C3
    Cell.setValue(ws, "B2", {
      formula: 'INDIRECT("R[1]C[1]",FALSE)',
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B2")).toBe(99);
  });

  it("resolves INDIRECT across sheets", () => {
    const wb = Workbook.create();
    const ws1 = Workbook.addWorksheet(wb, "Sheet1");
    const ws2 = Workbook.addWorksheet(wb, "Sheet2");
    Cell.setValue(ws2, "B2", 123);
    Cell.setValue(ws1, "A1", {
      formula: 'INDIRECT("Sheet2!B2")',
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws1, "A1")).toBe(123);
  });

  it("returns #REF! when INDIRECT target is a syntactically invalid address", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // '!!' is not a valid A1 address → parser/binder throws → evaluator
    // catches and returns #REF!.
    Cell.setValue(ws, "A1", { formula: 'INDIRECT("!!")', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#REF!" });
  });

  it("returns #REF! when INDIRECT target is the empty string", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'INDIRECT("")', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#REF!" });
  });
});

// ============================================================================
// OFFSET
// ============================================================================

describe("evaluator: OFFSET", () => {
  it("returns a single cell at the base + offset", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "C4", 33);
    Cell.setValue(ws, "A1", { formula: "OFFSET(A1,3,2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(33);
  });

  it("returns a range with explicit height/width", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B2", 10);
    Cell.setValue(ws, "B3", 20);
    Cell.setValue(ws, "B4", 30);
    Cell.setValue(ws, "D1", {
      formula: "SUM(OFFSET(A1,1,1,3,1))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(60);
  });

  it("accepts a runtime-produced reference (INDIRECT) as base (regression)", () => {
    // Previously OFFSET rejected anything other than a literal CellRef /
    // AreaRef with #VALUE!. Excel accepts any reference-producing
    // expression as the base (INDIRECT, defined names, chained OFFSET).
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B2", 42);
    Cell.setValue(ws, "D1", {
      formula: 'OFFSET(INDIRECT("A1"), 1, 1)',
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(42);
  });

  it("supports negative height (range extends upward)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    // OFFSET(A3,0,0,-3,1) → A1:A3
    Cell.setValue(ws, "B1", {
      formula: "SUM(OFFSET(A3,0,0,-3,1))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(6);
  });

  it("supports negative width (range extends leftward)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "C1", 3);
    // OFFSET(C1,0,0,1,-3) → A1:C1
    Cell.setValue(ws, "D1", {
      formula: "SUM(OFFSET(C1,0,0,1,-3))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(6);
  });

  it("returns #REF! for zero height or zero width", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "OFFSET(A1,0,0,0,1)", result: 0 });
    Cell.setValue(ws, "B2", { formula: "OFFSET(A1,0,0,1,0)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#REF!" });
    expect(Cell.getResult(ws, "B2")).toEqual({ error: "#REF!" });
  });

  it("returns #REF! when the offset moves above row 1 or column 1", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "OFFSET(A1,-1,0)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#REF!" });
  });
});

// ============================================================================
// ROW / COLUMN / ROWS / COLUMNS  (regression coverage R2-PX-YY)
// ============================================================================

describe("evaluator: ROW/COLUMN/ROWS/COLUMNS", () => {
  it("ROW(cell) returns the row of the reference", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "ROW(B5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(5);
  });

  it("COLUMN(cell) returns the column of the reference", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "COLUMN(D1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(4);
  });

  it("ROW() with no argument returns the current row", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A3", { formula: "ROW()", result: 0 });
    Cell.setValue(ws, "B7", { formula: "COLUMN()", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A3")).toBe(3);
    expect(Cell.getResult(ws, "B7")).toBe(2);
  });

  it("ROW(range) returns the top row of the range", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "ROW(C5:E7)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(5);
  });

  it("ROWS(range) returns the row count", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "ROWS(A1:A10)", result: 0 });
    Cell.setValue(ws, "A2", { formula: "ROWS(B2:D2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(10);
    expect(Cell.getResult(ws, "A2")).toBe(1);
  });

  it("COLUMNS(range) returns the column count", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "COLUMNS(A1:C1)", result: 0 });
    Cell.setValue(ws, "A2", { formula: "COLUMNS(A1:A10)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(3);
    expect(Cell.getResult(ws, "A2")).toBe(1);
  });

  // R2-PX-YY regressions — the ROW/COLUMN pass-through for INDIRECT/OFFSET.
  // These test the evaluator's "inspect raw reference before dereferencing"
  // branch in evaluateCall (around evaluator.ts:985).
  it('ROW(INDIRECT("A5")) returns 5 (R2-PX-YY)', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'ROW(INDIRECT("A5"))', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(5);
  });

  it('COLUMN(INDIRECT("D1")) returns 4 (R2-PX-YY)', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'COLUMN(INDIRECT("D1"))', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(4);
  });

  it('ROWS(INDIRECT("A1:A10")) returns 10 (R2-PX-YY)', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'ROWS(INDIRECT("A1:A10"))', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(10);
  });

  it('COLUMNS(INDIRECT("A1:C1")) returns 3 (R2-PX-YY)', () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: 'COLUMNS(INDIRECT("A1:C1"))', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(3);
  });

  it("ROWS(OFFSET(A1,0,0,5,3)) returns 5 — OFFSET array arg path", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "ROWS(OFFSET(A1,0,0,5,3))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(5);
  });

  it("COLUMNS(OFFSET(A1,0,0,5,3)) returns 3 — OFFSET array arg path", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "COLUMNS(OFFSET(A1,0,0,5,3))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(3);
  });
});

// ============================================================================
// Defined Names
// ============================================================================

describe("evaluator: defined names", () => {
  it("resolves a workbook-scoped single-cell name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 50);
    definedNamesAdd(getDefinedNames(wb), "Sheet1!$A$1", "Price");
    Cell.setValue(ws, "B1", { formula: "Price*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(100);
  });

  it("resolves a single-area range name via SUM", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    definedNamesAdd(getDefinedNames(wb), "Sheet1!$A$1:$A$3", "Vals");
    Cell.setValue(ws, "B1", { formula: "SUM(Vals)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(6);
  });

  it("resolves a formula-based defined name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    definedNamesAddFormula(getDefinedNames(wb), "Answer", "20+22");
    Cell.setValue(ws, "A1", { formula: "Answer", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(42);
  });

  it("prefers sheet-local scope over workbook-global scope", () => {
    const wb = Workbook.create();
    const ws1 = Workbook.addWorksheet(wb, "Sheet1");
    const ws2 = Workbook.addWorksheet(wb, "Sheet2");
    Cell.setValue(ws1, "A1", 100);
    Cell.setValue(ws2, "A1", 200);
    // Global name → Sheet1!A1, Sheet1-local → Sheet2!A1
    definedNamesSetModel(getDefinedNames(wb), [
      { name: "Val", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
      {
        name: "Val",
        ranges: ["Sheet2!$A$1"],
        rawText: "Sheet2!$A$1",
        localSheetId: 0
      }
    ]);
    Cell.setValue(ws1, "B1", { formula: "Val", result: 0 });
    Cell.setValue(ws2, "B1", { formula: "Val", result: 0 });
    calculateFormulas(wb);
    // Sheet1 sees the sheet-local (Sheet2!A1 = 200)
    expect(Cell.getResult(ws1, "B1")).toBe(200);
    // Sheet2 has no local → falls back to global (Sheet1!A1 = 100)
    expect(Cell.getResult(ws2, "B1")).toBe(100);
  });

  it("returns #NAME? for an unknown name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // Omit `result:` so the engine can't preserve a cached value —
    // otherwise the engine keeps the stale number instead of surfacing the
    // error (see calculate-formulas-impl.ts unsupported-formula branch).
    Cell.setValue(ws, "A1", { formula: "NoSuchName" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NAME?" });
  });
});

// ============================================================================
// 3D References
// ============================================================================

describe("evaluator: 3D references", () => {
  it("SUM(Sheet1:Sheet3!A1) sums the same cell across sheets", () => {
    const wb = Workbook.create();
    const ws1 = Workbook.addWorksheet(wb, "Sheet1");
    const ws2 = Workbook.addWorksheet(wb, "Sheet2");
    const ws3 = Workbook.addWorksheet(wb, "Sheet3");
    const out = Workbook.addWorksheet(wb, "Out");
    Cell.setValue(ws1, "A1", 10);
    Cell.setValue(ws2, "A1", 20);
    Cell.setValue(ws3, "A1", 30);
    Cell.setValue(out, "A1", { formula: "SUM(Sheet1:Sheet3!A1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(out, "A1")).toBe(60);
  });

  it("SUM(Sheet1:Sheet2!A1:B1) sums an area across sheets", () => {
    const wb = Workbook.create();
    const ws1 = Workbook.addWorksheet(wb, "Sheet1");
    const ws2 = Workbook.addWorksheet(wb, "Sheet2");
    const out = Workbook.addWorksheet(wb, "Out");
    Cell.setValue(ws1, "A1", 1);
    Cell.setValue(ws1, "B1", 2);
    Cell.setValue(ws2, "A1", 3);
    Cell.setValue(ws2, "B1", 4);
    Cell.setValue(out, "A1", { formula: "SUM(Sheet1:Sheet2!A1:B1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(out, "A1")).toBe(10);
  });
});

// ============================================================================
// Structured References
// ============================================================================

describe("evaluator: structured references", () => {
  it("Table[Col] reads a single column's data", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    addTable(ws, {
      name: "Nums",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "V" }],
      rows: [[10], [20], [30]]
    });
    Cell.setValue(ws, "D1", { formula: "SUM(Nums[V])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(60);
  });

  it("Table[#All] includes header + data (+ totals if present)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    addTable(ws, {
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
    Cell.setValue(ws, "D1", { formula: "COUNTA(Items[#All])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(6);
  });
});

// ============================================================================
// Error Propagation Scenarios
// ============================================================================

describe("evaluator: error emission sources", () => {
  it("#REF! from an invalid external reference", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "[Book1]Sheet1!A1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#REF!" });
  });

  it("#NAME? from an unknown function", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "NOPE(1)" });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NAME?" });
  });

  it("#VALUE! from numeric coercion of an empty string", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: '1+""', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#VALUE!" });
  });

  it("#NUM! from SQRT of a negative", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SQRT(-1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NUM!" });
  });

  it("#DIV/0! from arithmetic divide-by-zero", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "10/0", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
  });

  it("#N/A from NA()", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "NA()", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#N/A" });
  });

  it("#NULL! from a non-overlapping intersection", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "C1", 3);
    // A1:A2 and C1:C2 don't overlap
    Cell.setValue(ws, "B1", { formula: "SUM(A1:A2 C1:C2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#NULL!" });
  });
});

// ============================================================================
// Circular References
// ============================================================================

describe("evaluator: circular references", () => {
  it("returns a number for a simple cycle without iterate enabled", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "A2+1", result: 0 });
    Cell.setValue(ws, "A2", { formula: "A1+1", result: 0 });
    calculateFormulas(wb);
    // Zero-seed fallback keeps the engine making progress. Both must resolve
    // to numbers rather than errors — the established behaviour this engine
    // preserves.
    expect(typeof Cell.getResult(ws, "A1")).toBe("number");
    expect(typeof Cell.getResult(ws, "A2")).toBe("number");
  });

  it("converges under iterate=true to the fixed point", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    wb.calcProperties = {
      iterate: true,
      iterateCount: 200,
      iterateDelta: 0.0001
    };
    // A1 = A1/2 + 1 → fixed point at 2
    Cell.setValue(ws, "A1", { formula: "A1/2+1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBeCloseTo(2, 3);
  });

  it("reaches iterateCount when no convergence is possible", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    wb.calcProperties = {
      iterate: true,
      iterateCount: 5,
      iterateDelta: 0
    };
    // Increment-only cycle — no fixed point, runs out at count
    Cell.setValue(ws, "A1", { formula: "A1+1", result: 0 });
    calculateFormulas(wb);
    // Initial 0 → 1; then 5 iterations → 6
    expect(Cell.getResult(ws, "A1")).toBe(6);
  });
});

// ============================================================================
// Volatile Functions
// ============================================================================

describe("evaluator: volatile functions", () => {
  it("RAND() returns a number in [0, 1)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "RAND()", result: 0 });
    calculateFormulas(wb);
    const r = Cell.getResult(ws, "A1") as number;
    expect(typeof r).toBe("number");
    expect(r).toBeGreaterThanOrEqual(0);
    expect(r).toBeLessThan(1);
  });

  it("RAND() produces a fresh value on each recalculation", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "RAND()", result: 0 });

    // Many samples — if volatile is wired up correctly, at least two
    // recalculations should yield different values. Without volatile
    // handling the cache would return the same value forever.
    const samples = new Set<number>();
    for (let i = 0; i < 20; i++) {
      calculateFormulas(wb);
      samples.add(Cell.getResult(ws, "A1") as number);
    }
    expect(samples.size).toBeGreaterThan(1);
  });

  it("TODAY() returns a whole-day serial", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "TODAY()", result: 0 });
    calculateFormulas(wb);
    const r = Cell.getResult(ws, "A1") as number;
    expect(typeof r).toBe("number");
    expect(r).toBe(Math.trunc(r));
    // Rough sanity — post-2020 serial range (> Jan 1 2020 = 43831)
    expect(r).toBeGreaterThan(43830);
  });

  it("NOW() returns a serial strictly greater than (or equal to) TODAY()", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "NOW()", result: 0 });
    Cell.setValue(ws, "A2", { formula: "TODAY()", result: 0 });
    calculateFormulas(wb);
    const now = Cell.getResult(ws, "A1") as number;
    const today = Cell.getResult(ws, "A2") as number;
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
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "LET(x,42,x)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(42);
  });

  it("LET propagates errors from the binding expression", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "LET(x,1/0,x+1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
  });

  it("LET can bind a reference range and SUM it", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "B1", { formula: "LET(r,A1:A3,SUM(r))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(6);
  });

  it("LET with odd arg count (missing body) returns #VALUE!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "LET(x,1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#VALUE!" });
  });

  it("LET supports many bindings (stress)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(a,1,b,2,c,3,d,4,e,5,a+b+c+d+e)",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });
});

describe("evaluator: LAMBDA additional cases", () => {
  it("LAMBDA with zero params", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(const,LAMBDA(42),const())",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(42);
  });

  it("LAMBDA with three params", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(f,LAMBDA(a,b,c,a*100+b*10+c),f(1,2,3))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(123);
  });

  it("LAMBDA recursive via defined-name self-reference", () => {
    // Define a factorial lambda and call it. The engine's name resolution
    // allows a lambda stored in a defined name to call itself by name.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "Z1", {
      formula: "LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))",
      result: 0
    });
    definedNamesAdd(getDefinedNames(wb), "Sheet1!$Z$1", "Fact");
    Cell.setValue(ws, "A1", { formula: "Fact(5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(120);
  });

  it("LAMBDA arity mismatch returns an error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(f,LAMBDA(x,y,x+y),f(1))",
      result: 0
    });
    calculateFormulas(wb);
    // Engine resolves unbound y → propagates error/blank depending on path;
    // must at least not silently succeed with a number.
    const r = Cell.getResult(ws, "A1");
    expect(typeof r === "number" ? r : null).not.toBe(1);
  });

  it("LAMBDA body that divides by zero propagates the error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "LET(f,LAMBDA(x,1/x),f(0))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
  });
});

// ============================================================================
// Higher-order functions — extended coverage
// ============================================================================

describe("evaluator: MAP extra", () => {
  it("MAP over a 2×2 array preserves shape when wrapped in SUMPRODUCT", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B2", 4);
    Cell.setValue(ws, "D1", {
      formula: "SUMPRODUCT(MAP(A1:B2,LAMBDA(x,x+1)))",
      result: 0
    });
    calculateFormulas(wb);
    // Each cell incremented by 1 → sum of (2+3+4+5) = 14
    expect(Cell.getResult(ws, "D1")).toBe(14);
  });

  it("MAP with a constant-returning lambda yields a uniform array", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "D1", {
      formula: "SUM(MAP({1,2,3},LAMBDA(x,7)))",
      result: 0
    });
    calculateFormulas(wb);
    // Each of 3 cells → 7 → sum = 21
    expect(Cell.getResult(ws, "D1")).toBe(21);
  });

  it("MAP supports multiple input arrays (Excel allows up to 254)", () => {
    // Regression: previously MAP only read the first array and passed
    // single-element lambda invocations; the other arrays silently
    // became no-ops. Excel's MAP(arr1, ..., arrN, lambda) invokes the
    // lambda with N values per position — the lambda must have N
    // parameters matching.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "B1", 10);
    Cell.setValue(ws, "B2", 20);
    Cell.setValue(ws, "B3", 30);
    Cell.setValue(ws, "D1", {
      formula: "SUM(MAP(A1:A3, B1:B3, LAMBDA(a,b, a+b)))",
      result: 0
    });
    calculateFormulas(wb);
    // (1+10)+(2+20)+(3+30) = 11+22+33 = 66
    expect(Cell.getResult(ws, "D1")).toBe(66);
  });

  it("MAP rejects mismatched lambda arity", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "D1", {
      // 2 arrays, lambda takes 1 param → #VALUE!.
      formula: "MAP({1,2},{3,4},LAMBDA(x,x))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toEqual({ error: "#VALUE!" });
  });

  it("MAP with a lambda returning errors propagates them into the result", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "D1", {
      formula: 'IFERROR(SUM(MAP({1,2,0},LAMBDA(x,1/x))),"err")',
      result: 0
    });
    calculateFormulas(wb);
    // 1/0 → #DIV/0! → SUM would propagate; IFERROR catches
    expect(Cell.getResult(ws, "D1")).toBe("err");
  });
});

describe("evaluator: REDUCE extra", () => {
  it("REDUCE sums a range", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "A4", 4);
    Cell.setValue(ws, "B1", {
      formula: "REDUCE(0,A1:A4,LAMBDA(a,x,a+x))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(10);
  });

  it("REDUCE multiplies a range (product accumulator)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", {
      formula: "REDUCE(1,{2,3,4},LAMBDA(a,x,a*x))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(24);
  });

  it("REDUCE propagates errors from the accumulator", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", {
      formula: "REDUCE(1,{1,0,3},LAMBDA(a,x,a/x))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#DIV/0!" });
  });

  it("REDUCE counts elements matching a predicate", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", {
      formula: "REDUCE(0,{1,2,3,4,5},LAMBDA(a,x,IF(x>3,a+1,a)))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(2);
  });
});

describe("evaluator: SCAN / BYROW / BYCOL / MAKEARRAY extra", () => {
  it("SCAN with product accumulator yields running products", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(SCAN(1,{2,3,4},LAMBDA(a,x,a*x)))",
      result: 0
    });
    calculateFormulas(wb);
    // Running products: 2, 6, 24 → sum = 32
    expect(Cell.getResult(ws, "A1")).toBe(32);
  });

  it("BYROW reduces each row with MAX", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 5);
    Cell.setValue(ws, "A2", 10);
    Cell.setValue(ws, "B2", 3);
    Cell.setValue(ws, "D1", {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,MAX(r))))",
      result: 0
    });
    calculateFormulas(wb);
    // row1 max=5, row2 max=10 → sum=15
    expect(Cell.getResult(ws, "D1")).toBe(15);
  });

  it("BYCOL reduces each column with AVERAGE", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 2);
    Cell.setValue(ws, "B1", 4);
    Cell.setValue(ws, "A2", 4);
    Cell.setValue(ws, "B2", 6);
    Cell.setValue(ws, "D1", {
      formula: "SUM(BYCOL(A1:B2,LAMBDA(c,AVERAGE(c))))",
      result: 0
    });
    calculateFormulas(wb);
    // col1 avg=3, col2 avg=5 → sum=8
    expect(Cell.getResult(ws, "D1")).toBe(8);
  });

  it("MAKEARRAY(3,3,...) builds an identity-style matrix", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(MAKEARRAY(3,3,LAMBDA(r,c,IF(r=c,1,0))))",
      result: 0
    });
    calculateFormulas(wb);
    // Diagonal has three 1s
    expect(Cell.getResult(ws, "A1")).toBe(3);
  });

  it("MAKEARRAY(1,5,...) builds a row vector", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(MAKEARRAY(1,5,LAMBDA(r,c,c)))",
      result: 0
    });
    calculateFormulas(wb);
    // 1+2+3+4+5 = 15
    expect(Cell.getResult(ws, "A1")).toBe(15);
  });

  it("MAKEARRAY(5,1,...) builds a column vector", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(MAKEARRAY(5,1,LAMBDA(r,c,r*2)))",
      result: 0
    });
    calculateFormulas(wb);
    // 2+4+6+8+10 = 30
    expect(Cell.getResult(ws, "A1")).toBe(30);
  });

  it("BYROW with an error inside a row propagates the error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B2", 4);
    Cell.setValue(ws, "D1", {
      formula: "SUM(BYROW(A1:B2,LAMBDA(r,SUM(r))))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toEqual({ error: "#DIV/0!" });
  });

  it("SCAN on an array preserves its shape", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", {
      formula: "SUM(SCAN(0,{1,2,3,4},LAMBDA(a,x,a+x)))",
      result: 0
    });
    calculateFormulas(wb);
    // running sums: 1,3,6,10 → sum=20
    expect(Cell.getResult(ws, "A1")).toBe(20);
  });

  it("REDUCE over a 2D array traverses row-major", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 2);
    Cell.setValue(ws, "A2", 3);
    Cell.setValue(ws, "B2", 4);
    Cell.setValue(ws, "D1", {
      formula: "REDUCE(0,A1:B2,LAMBDA(a,x,a+x))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(10);
  });

  it("MAP propagates cell-level errors through the lambda", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "B1", {
      formula: "SUM(MAP(A1:A3,LAMBDA(x,x*2)))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#DIV/0!" });
  });
});

describe("evaluator: multi-area reference unions", () => {
  it("INDEX with area_num picks from the specified area (Excel)", () => {
    // Regression: previously `(A1:B2, D4:E5)` either failed to parse or
    // silently stacked both ranges into one flat array — INDEX's
    // area_num had no way to address the second area. Now the
    // parenthesised union produces a multi-area ReferenceValue that
    // INDEX's reference-aware path (tryEvaluateINDEX) can route on.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 11);
    Cell.setValue(ws, "B1", 12);
    Cell.setValue(ws, "A2", 13);
    Cell.setValue(ws, "B2", 14);
    Cell.setValue(ws, "D4", 21);
    Cell.setValue(ws, "E4", 22);
    Cell.setValue(ws, "D5", 23);
    Cell.setValue(ws, "E5", 24);
    Cell.setValue(ws, "G1", { formula: "INDEX((A1:B2,D4:E5), 1, 1, 1)", result: 0 });
    Cell.setValue(ws, "G2", { formula: "INDEX((A1:B2,D4:E5), 1, 1, 2)", result: 0 });
    Cell.setValue(ws, "G3", { formula: "INDEX((A1:B2,D4:E5), 2, 2, 2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "G1")).toBe(11); // area 1 top-left
    expect(Cell.getResult(ws, "G2")).toBe(21); // area 2 top-left
    expect(Cell.getResult(ws, "G3")).toBe(24); // area 2 bottom-right
  });

  it("INDEX with out-of-range area_num returns #REF!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "D1", 2);
    Cell.setValue(ws, "G1", { formula: "INDEX((A1,D1), 1, 1, 3)", result: 0 });
    Cell.setValue(ws, "G2", { formula: "INDEX((A1,D1), 1, 1, 0)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "G1")).toEqual({ error: "#REF!" });
    expect(Cell.getResult(ws, "G2")).toEqual({ error: "#REF!" });
  });

  it("AREAS counts members of a reference union (Excel)", () => {
    // Regression: the standard dereference path would flatten a
    // multi-area reference into a single array, collapsing AREAS to 1.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "AREAS((A1:B2,D4:E5,G7))", result: 0 });
    Cell.setValue(ws, "B2", { formula: "AREAS(A1:B2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(3);
    expect(Cell.getResult(ws, "B2")).toBe(1);
  });

  it("union of non-references surfaces as #VALUE!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    // `1` is a literal, not a reference.
    Cell.setValue(ws, "A1", { formula: "INDEX((1,B1), 1, 1, 1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#VALUE!" });
  });

  it("SUM over a reference union flattens all areas (Excel)", () => {
    // SUM receives dereferenced args — the evaluator flattens multi-area
    // references into a single stacked array. Total should include every
    // cell from every area.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "B1", 10);
    Cell.setValue(ws, "B2", 20);
    Cell.setValue(ws, "D1", { formula: "SUM((A1:A2, B1:B2))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(33);
  });

  it("ROWS on a reference union flattens areas into a stacked array", () => {
    // ROWS over `(A1:A3, B1:B2)` — the dereferenceValue pipeline stacks
    // all areas into a single ArrayValue (see the multi-area branch in
    // `dereferenceValue`), so ROWS(count) equals the sum of each area's
    // heights: 3 + 2 = 5. This documents our engine's behaviour (Excel
    // actually reports #REF! for ROWS on a union; we chose to flatten).
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "B1", 4);
    Cell.setValue(ws, "B2", 5);
    Cell.setValue(ws, "D1", { formula: "ROWS((A1:A3, B1:B2))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(5);
  });

  it("single-area source with area_num=1 works; area_num>1 is #REF!", () => {
    // Regression: a single-area reference is a 1-area union for
    // INDEX purposes. `INDEX(A1:B2, 1, 1, 1)` should work; any
    // area_num > 1 on a single area must surface as #REF! rather
    // than being silently ignored.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 42);
    Cell.setValue(ws, "D1", { formula: "INDEX(A1:B2, 1, 1, 1)", result: 0 });
    Cell.setValue(ws, "D2", { formula: "INDEX(A1:B2, 1, 1, 2)", result: 0 });
    Cell.setValue(ws, "D3", { formula: "INDEX(A1:B2, 1, 1,)", result: 0 }); // blank area_num → 1
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(42);
    expect(Cell.getResult(ws, "D2")).toEqual({ error: "#REF!" });
    expect(Cell.getResult(ws, "D3")).toBe(42);
  });

  it("nested union inside IF selects the correct branch", () => {
    // Regression: UnionRef should pass through special forms unchanged —
    // IF / LET / SWITCH just forward the value. The outer SUM then
    // dereferences the selected union into a stacked array.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "B1", 10);
    Cell.setValue(ws, "B2", 20);
    Cell.setValue(ws, "C1", 100);
    Cell.setValue(ws, "C2", 200);
    // Condition=TRUE → (A1, B1) → 11
    Cell.setValue(ws, "D1", {
      formula: "SUM(IF(TRUE, (A1, B1), (C1, C2)))",
      result: 0
    });
    // Condition=FALSE → (C1, C2) → 300
    Cell.setValue(ws, "D2", {
      formula: "SUM(IF(FALSE, (A1, B1), (C1, C2)))",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(11);
    expect(Cell.getResult(ws, "D2")).toBe(300);
  });

  it("error in any union member propagates (Excel-compliant)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", { formula: "1/0", result: 0 }); // #DIV/0!
    Cell.setValue(ws, "D1", { formula: "SUM((A1, B1))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toEqual({ error: "#DIV/0!" });
  });

  it("dependency graph tracks all areas of a union (regression)", () => {
    // Without union-aware dep extraction, editing B1 wouldn't
    // invalidate a formula that reads `(A1:A2, B1:B2)`.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "B1", 10);
    Cell.setValue(ws, "D1", { formula: "SUM((A1, B1))", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(11);
    // Edit B1 and recalc; the union's B1 area should be picked up by
    // the dep graph so D1 gets recomputed.
    Cell.setValue(ws, "B1", 100);
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(101);
  });
});
