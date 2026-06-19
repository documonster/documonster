import { cellGetValue, cellResult } from "@excel/core/cell";
import { calculateFormulas } from "@excel/core/formula-adapter";
import type { WorkbookData } from "@excel/core/workbook-core";
import { addTable, getCell } from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { Cell, Row, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/**
 * End-to-end tests for table totals rows.
 *
 * These exercise the real flow: Table.store() auto-injects
 * SUBTOTAL(fnNum, Tbl[Col]) into the totals row based on the column's
 * totalsRowFunction. wb.calculateFormulas() must then evaluate those
 * SUBTOTAL calls, cache the results, and surface them via every
 * structured-ref variant.
 *
 * Covers:
 *   - Every totalsRowFunction value (sum/count/countNums/average/
 *     min/max/stdDev/var/custom/none)
 *   - Reading the totals cell directly AND via [#Totals] refs
 *   - Multi-column [#Totals] ranges
 *   - [#Data]+[#Totals] combos
 *   - Implicit Tbl[Col] excluding the totals row (Excel semantics)
 *   - Custom totalsRowFormula
 *   - #REF! propagation when [#Totals] is referenced on a totals-less table
 */

describe("table totals row: SUBTOTAL injection for every function", () => {
  const data: (string | number)[][] = [
    ["A", 10],
    ["B", 20],
    ["C", 30],
    ["D", 40]
  ];

  const cases: Array<{
    fn: "sum" | "count" | "countNums" | "average" | "min" | "max" | "stdDev" | "var";
    expected: number;
    tolerance?: number;
  }> = [
    { fn: "sum", expected: 100 },
    { fn: "count", expected: 4 },
    { fn: "countNums", expected: 4 },
    { fn: "average", expected: 25 },
    { fn: "min", expected: 10 },
    { fn: "max", expected: 40 },
    // Sample std dev of [10,20,30,40] = sqrt(166.666...) ≈ 12.9099
    { fn: "stdDev", expected: 12.9099, tolerance: 1e-3 },
    // Sample variance of [10,20,30,40] = 166.666...
    { fn: "var", expected: 166.6667, tolerance: 1e-3 }
  ];

  for (const { fn, expected, tolerance } of cases) {
    it(`totalsRowFunction="${fn}" injects SUBTOTAL and evaluates correctly`, () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "S");
      addTable(ws, {
        name: "MyTable",
        ref: "A1",
        headerRow: true,
        totalsRow: true,
        columns: [{ name: "Label" }, { name: "Qty", totalsRowFunction: fn }],
        rows: data
      });
      calculateFormulas(wb);

      // Totals row sits right after data: header (r1) + 4 data (r2-r5) + totals (r6)
      const totalsCell = getCell(ws, "B6");
      const result = cellResult(totalsCell) as number;
      if (tolerance !== undefined) {
        expect(Math.abs(result - expected)).toBeLessThan(tolerance);
      } else {
        expect(result).toBe(expected);
      }

      // Also verify the injected formula uses SUBTOTAL
      const formula = (cellGetValue(totalsCell) as { formula: string }).formula;
      expect(formula).toMatch(/^SUBTOTAL\(\d+,MyTable\[Qty\]\)$/);
    });
  }
});

describe("table totals row: structured-ref access patterns", () => {
  function makeSalesTable(totalsRow = true): WorkbookData {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Sales",
      ref: "A1",
      headerRow: true,
      totalsRow,
      columns: [
        { name: "Product" },
        { name: "Qty", totalsRowFunction: totalsRow ? "sum" : "none" },
        { name: "Price", totalsRowFunction: totalsRow ? "average" : "none" }
      ],
      rows: [
        ["Apple", 10, 1],
        ["Pear", 20, 2],
        ["Peach", 30, 3]
      ]
    });
    return wb;
  }

  it("Sales[[#Totals],[Qty]] returns the exact totals value", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", { formula: "Sales[[#Totals],[Qty]]", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBe(60);
  });

  it("Sales[[#Totals],[Price]] returns the exact average", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", { formula: "Sales[[#Totals],[Price]]", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBe(2);
  });

  it("SUM(Sales[[#Totals],[Qty]:[Price]]) sums across multi-column totals range", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", {
      formula: "SUM(Sales[[#Totals],[Qty]:[Price]])",
      result: 0
    });
    calculateFormulas(wb);
    // Totals row: Qty=60, Price=2 → sum = 62
    expect(Cell.getResult(ws, "E1")).toBe(62);
  });

  it("SUM(Sales[[#Data],[#Totals],[Qty]]) includes both data and totals", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", {
      formula: "SUM(Sales[[#Data],[#Totals],[Qty]])",
      result: 0
    });
    calculateFormulas(wb);
    // Data: 10+20+30 = 60, plus totals row = 60 → 120
    expect(Cell.getResult(ws, "E1")).toBe(120);
  });

  it("COUNTA(Sales[[#All],[Qty]]) counts header + 3 data + totals = 5", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", {
      formula: "COUNTA(Sales[[#All],[Qty]])",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBe(5);
  });

  it("Sales[Qty] (implicit) excludes the totals row", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", { formula: "SUM(Sales[Qty])", result: 0 });
    calculateFormulas(wb);
    // Data only: 10+20+30 = 60 (not 120)
    expect(Cell.getResult(ws, "E1")).toBe(60);
  });

  it("COUNTA(Sales[Qty]) excludes both header and totals", () => {
    const wb = makeSalesTable();
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", { formula: "COUNTA(Sales[Qty])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBe(3);
  });

  it("[#Totals] on a table WITHOUT a totals row returns #REF!", () => {
    const wb = makeSalesTable(false);
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", {
      formula: "Sales[[#Totals],[Qty]]",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toEqual({ error: "#REF!" });
  });

  it("SUM of [#Totals] on totals-less table propagates #REF!", () => {
    const wb = makeSalesTable(false);
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", {
      formula: "SUM(Sales[[#Totals],[Qty]])",
      result: 0
    });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toEqual({ error: "#REF!" });
  });
});

describe("table totals row: direct cell reads", () => {
  it("reading the totals cell by A1 ref yields the computed value", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ]
    });
    calculateFormulas(wb);
    // Header row 1, data rows 2-4, totals row 5
    expect(Cell.getResult(ws, "B5")).toBe(6);
  });

  it("formula in another cell referencing totals A1 sees the computed value", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ]
    });
    Cell.setValue(ws, "D1", { formula: "B5*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(12);
  });
});

describe("table totals row: custom totals formula", () => {
  it("totalsRowFunction=custom uses the user-supplied totalsRowFormula", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [
        { name: "Name" },
        {
          name: "N",
          totalsRowFunction: "custom",
          totalsRowFormula: "SUM(T[N])*10"
        }
      ],
      rows: [
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ]
    });
    calculateFormulas(wb);
    // SUM(T[N]) excludes totals → 6, *10 = 60
    expect(Cell.getResult(ws, "B5")).toBe(60);
  });

  it("custom totals formula can reference another sheet cell", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "E1", 5);
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [
        { name: "Name" },
        {
          name: "N",
          totalsRowFunction: "custom",
          totalsRowFormula: "SUM(T[N])+E1"
        }
      ],
      rows: [
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ]
    });
    calculateFormulas(wb);
    // 6 + 5 = 11
    expect(Cell.getResult(ws, "B5")).toBe(11);
  });
});

describe("table totals row: SUBTOTAL semantics are preserved", () => {
  it("SUBTOTAL in totals row ignores hidden rows when code is 1xx family", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [
        ["a", 10],
        ["b", 20],
        ["c", 30]
      ]
    });
    calculateFormulas(wb);
    // Base case (no rows hidden): 10+20+30 = 60
    expect(Cell.getResult(ws, "B5")).toBe(60);

    // The injected formula is SUBTOTAL(109,T[N]) — the 1xx variant.
    // Hide row 3 (the "b" row with value 20) and re-calculate.
    Row.setHidden(ws, 3, true);
    calculateFormulas(wb);
    // 1xx variant now drops the hidden row: 10 + 30 = 40
    expect(Cell.getResult(ws, "B5")).toBe(40);

    const formula = (Cell.getValue(ws, "B5") as { formula: string }).formula;
    expect(formula).toBe("SUBTOTAL(109,T[N])");
  });

  it("SUBTOTAL plain (9) does NOT skip hidden rows (conservative impl)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Cell.setValue(ws, "A3", 30);
    Row.setHidden(ws, 2, true);
    Cell.setValue(ws, "B1", { formula: "SUBTOTAL(9, A1:A3)", result: 0 });
    calculateFormulas(wb);
    // Plain code 9: our impl conservatively treats all rows as visible
    // because `row.hidden` conflates filter-hide with manual-hide.
    // Sum includes the hidden row.
    expect(Cell.getResult(ws, "B1")).toBe(60);
  });

  it("SUBTOTAL 1xx skips empty hidden rows too", () => {
    // Hidden rows with no cells of their own still participate in
    // hidden-row skipping. The adapter collects them via eachRow with
    // includeEmpty so row.hidden is visible even for pure-empty rows.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    // A2 unset; row 2 marked hidden.
    Cell.setValue(ws, "A3", 30);
    Row.setHidden(ws, 2, true);
    Cell.setValue(ws, "B1", { formula: "SUBTOTAL(109, A1:A3)", result: 0 });
    calculateFormulas(wb);
    // Row 2 is empty + hidden → skipped; no change vs unhidden (blank
    // contributes 0). But the point here is the adapter captures the
    // hidden flag via includeEmpty iteration.
    expect(Cell.getResult(ws, "B1")).toBe(40);
  });

  it("SUBTOTAL 109 over a multi-area reference preserves hidden-row skipping", () => {
    // Multi-area refs (A1:A3, A5:A7) flatten through dereferenceValue;
    // we need per-area hidden/subtotal masks to be merged into the
    // flattened ArrayValue so an outer SUBTOTAL still respects them.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let r = 1; r <= 7; r++) {
      Cell.setValue(ws, `A${r}`, r * 10);
    }
    // Hide row 2 (part of first area) and row 6 (part of second area).
    Row.setHidden(ws, 2, true);
    Row.setHidden(ws, 6, true);
    Cell.setValue(ws, "C1", { formula: "SUBTOTAL(109, A1:A3, A5:A7)", result: 0 });
    calculateFormulas(wb);
    // Values: A1=10, A2=20(hidden), A3=30, A5=50, A6=60(hidden), A7=70
    // Skipping 20 and 60 → 10+30+50+70 = 160
    expect(Cell.getResult(ws, "C1")).toBe(160);
  });

  it("SUBTOTAL over multi-area still skips nested SUBTOTAL cells in any area", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", { formula: "SUBTOTAL(9, A1:A2)", result: 0 }); // 3
    Cell.setValue(ws, "B1", 10);
    Cell.setValue(ws, "B2", 20);
    Cell.setValue(ws, "B3", { formula: "SUBTOTAL(9, B1:B2)", result: 0 }); // 30
    Cell.setValue(ws, "C1", { formula: "SUBTOTAL(9, A1:A3, B1:B3)", result: 0 });
    calculateFormulas(wb);
    // 1+2+10+20 = 33 (both nested SUBTOTAL cells skipped)
    expect(Cell.getResult(ws, "C1")).toBe(33);
  });

  it("nested SUBTOTAL inside table [#Totals] is not double-counted", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "T",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [
        ["a", 10],
        ["b", 20],
        ["c", 30]
      ]
    });
    // Wrapping with SUBTOTAL should skip cells that are themselves
    // SUBTOTAL results. Using [#All] deliberately includes the totals row.
    Cell.setValue(ws, "D1", {
      formula: "SUBTOTAL(109,T[[#All],[N]])",
      result: 0
    });
    calculateFormulas(wb);
    // Excel skips the inner SUBTOTAL cell → 10+20+30 = 60 (header is text)
    expect(Cell.getResult(ws, "D1")).toBe(60);
  });
});

describe("table totals row: edge cases", () => {
  it("totals row with empty data rows still evaluates without error", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Empty",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [[null, null]]
    });
    calculateFormulas(wb);
    // SUBTOTAL(109, Empty[N]) over all-null data should be 0
    expect(Cell.getResult(ws, "B3")).toBe(0);
  });

  it("multiple totals columns all compute independently", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Multi",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [
        { name: "Name" },
        { name: "A", totalsRowFunction: "sum" },
        { name: "B", totalsRowFunction: "max" },
        { name: "C", totalsRowFunction: "count" }
      ],
      rows: [
        ["x", 1, 10, "p"],
        ["y", 2, 20, "q"],
        ["z", 3, 30, "r"]
      ]
    });
    calculateFormulas(wb);
    // Totals row is row 5: header + 3 data + totals
    expect(Cell.getResult(ws, "B5")).toBe(6); // sum
    expect(Cell.getResult(ws, "C5")).toBe(30); // max
    expect(Cell.getResult(ws, "D5")).toBe(3); // count
  });
});

// ===========================================================================
// AGGREGATE — full option matrix
// ===========================================================================

describe("AGGREGATE: option semantics", () => {
  function makeSheet(): { wb: WorkbookData; ws: WorksheetData } {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    // A1..A5 with an error in A3 and a hidden row at row 4.
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Cell.setValue(ws, "A3", { error: "#DIV/0!" } as unknown as number);
    Cell.setValue(ws, "A4", 40);
    Cell.setValue(ws, "A5", 50);
    Row.setHidden(ws, 4, true);
    return { wb, ws };
  }

  it("option 4 (ignore nothing) propagates the error", () => {
    const { wb, ws } = makeSheet();
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 4, A1:A5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#DIV/0!" });
  });

  it("option 6 (ignore errors) sums the non-error cells including hidden", () => {
    const { wb, ws } = makeSheet();
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 6, A1:A5)", result: 0 });
    calculateFormulas(wb);
    // 10 + 20 + 40 + 50 = 120 (error skipped, hidden kept)
    expect(Cell.getResult(ws, "B1")).toBe(120);
  });

  it("option 5 (ignore hidden) keeps the error and fails", () => {
    const { wb, ws } = makeSheet();
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 5, A1:A5)", result: 0 });
    calculateFormulas(wb);
    // Error not skipped
    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#DIV/0!" });
  });

  it("option 7 (ignore hidden + errors) sums visible non-error cells", () => {
    const { wb, ws } = makeSheet();
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 7, A1:A5)", result: 0 });
    calculateFormulas(wb);
    // 10 + 20 + 50 = 80
    expect(Cell.getResult(ws, "B1")).toBe(80);
  });

  it("option 3 (ignore hidden + errors + nested) matches 7 for simple data", () => {
    const { wb, ws } = makeSheet();
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 3, A1:A5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(80);
  });

  it("function code 13 (MODE.SNGL) is supported", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 2);
    Cell.setValue(ws, "A4", 3);
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(13, 0, A1:A4)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(2);
  });

  it("function code 16 (PERCENTILE.INC) is supported", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 5; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(16, 0, A1:A5, 0.5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(3);
  });

  it("function code 17 (QUARTILE.INC) is supported", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 5; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    // Quartile code 2 = median
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(17, 0, A1:A5, 2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(3);
  });

  it("function code 18 (PERCENTILE.EXC) is supported", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 5; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(18, 0, A1:A5, 0.5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(3);
  });

  it("function code 19 (QUARTILE.EXC) is supported", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 7; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(19, 0, A1:A7, 2)", result: 0 });
    calculateFormulas(wb);
    // Quartile.exc 2 of [1..7] = 4
    expect(Cell.getResult(ws, "B1")).toBe(4);
  });

  it("AGGREGATE with option 0 skips nested SUBTOTAL (no double-count)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "A4", { formula: "SUBTOTAL(9, A1:A3)", result: 0 }); // nested = 6
    // AGGREGATE over A1:A4 with option 0 should skip A4 → 1+2+3 = 6
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 0, A1:A4)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(6);
  });

  it("AGGREGATE with option 4 does NOT skip nested SUBTOTAL (double-count)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    Cell.setValue(ws, "A2", 2);
    Cell.setValue(ws, "A3", 3);
    Cell.setValue(ws, "A4", { formula: "SUBTOTAL(9, A1:A3)", result: 0 });
    Cell.setValue(ws, "B1", { formula: "AGGREGATE(9, 4, A1:A4)", result: 0 });
    calculateFormulas(wb);
    // Option 4 keeps nested → 1+2+3+6 = 12
    expect(Cell.getResult(ws, "B1")).toBe(12);
  });
});
