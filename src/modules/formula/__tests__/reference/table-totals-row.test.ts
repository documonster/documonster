import { Workbook } from "@excel/workbook";
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("S");
      ws.addTable({
        name: "MyTable",
        ref: "A1",
        headerRow: true,
        totalsRow: true,
        columns: [{ name: "Label" }, { name: "Qty", totalsRowFunction: fn }],
        rows: data
      });
      wb.calculateFormulas();

      // Totals row sits right after data: header (r1) + 4 data (r2-r5) + totals (r6)
      const totalsCell = ws.getCell("B6");
      const result = totalsCell.result as number;
      if (tolerance !== undefined) {
        expect(Math.abs(result - expected)).toBeLessThan(tolerance);
      } else {
        expect(result).toBe(expected);
      }

      // Also verify the injected formula uses SUBTOTAL
      const formula = (totalsCell.value as { formula: string }).formula;
      expect(formula).toMatch(/^SUBTOTAL\(\d+,MyTable\[Qty\]\)$/);
    });
  }
});

describe("table totals row: structured-ref access patterns", () => {
  function makeSalesTable(totalsRow = true): Workbook {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = { formula: "Sales[[#Totals],[Qty]]", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBe(60);
  });

  it("Sales[[#Totals],[Price]] returns the exact average", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = { formula: "Sales[[#Totals],[Price]]", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBe(2);
  });

  it("SUM(Sales[[#Totals],[Qty]:[Price]]) sums across multi-column totals range", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = {
      formula: "SUM(Sales[[#Totals],[Qty]:[Price]])",
      result: 0
    };
    wb.calculateFormulas();
    // Totals row: Qty=60, Price=2 → sum = 62
    expect(ws.getCell("E1").result).toBe(62);
  });

  it("SUM(Sales[[#Data],[#Totals],[Qty]]) includes both data and totals", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = {
      formula: "SUM(Sales[[#Data],[#Totals],[Qty]])",
      result: 0
    };
    wb.calculateFormulas();
    // Data: 10+20+30 = 60, plus totals row = 60 → 120
    expect(ws.getCell("E1").result).toBe(120);
  });

  it("COUNTA(Sales[[#All],[Qty]]) counts header + 3 data + totals = 5", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = {
      formula: "COUNTA(Sales[[#All],[Qty]])",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBe(5);
  });

  it("Sales[Qty] (implicit) excludes the totals row", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = { formula: "SUM(Sales[Qty])", result: 0 };
    wb.calculateFormulas();
    // Data only: 10+20+30 = 60 (not 120)
    expect(ws.getCell("E1").result).toBe(60);
  });

  it("COUNTA(Sales[Qty]) excludes both header and totals", () => {
    const wb = makeSalesTable();
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = { formula: "COUNTA(Sales[Qty])", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBe(3);
  });

  it("[#Totals] on a table WITHOUT a totals row returns #REF!", () => {
    const wb = makeSalesTable(false);
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = {
      formula: "Sales[[#Totals],[Qty]]",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toEqual({ error: "#REF!" });
  });

  it("SUM of [#Totals] on totals-less table propagates #REF!", () => {
    const wb = makeSalesTable(false);
    const ws = wb.getWorksheet("S")!;
    ws.getCell("E1").value = {
      formula: "SUM(Sales[[#Totals],[Qty]])",
      result: 0
    };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toEqual({ error: "#REF!" });
  });
});

describe("table totals row: direct cell reads", () => {
  it("reading the totals cell by A1 ref yields the computed value", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    wb.calculateFormulas();
    // Header row 1, data rows 2-4, totals row 5
    expect(ws.getCell("B5").result).toBe(6);
  });

  it("formula in another cell referencing totals A1 sees the computed value", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    ws.getCell("D1").value = { formula: "B5*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(12);
  });
});

describe("table totals row: custom totals formula", () => {
  it("totalsRowFunction=custom uses the user-supplied totalsRowFormula", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    wb.calculateFormulas();
    // SUM(T[N]) excludes totals → 6, *10 = 60
    expect(ws.getCell("B5").result).toBe(60);
  });

  it("custom totals formula can reference another sheet cell", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("E1").value = 5;
    ws.addTable({
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
    wb.calculateFormulas();
    // 6 + 5 = 11
    expect(ws.getCell("B5").result).toBe(11);
  });
});

describe("table totals row: SUBTOTAL semantics are preserved", () => {
  it("SUBTOTAL in totals row ignores hidden rows when code is 1xx family", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    wb.calculateFormulas();
    // Base case (no rows hidden): 10+20+30 = 60
    expect(ws.getCell("B5").result).toBe(60);

    // The injected formula is SUBTOTAL(109,T[N]) — the 1xx variant.
    // Hide row 3 (the "b" row with value 20) and re-calculate.
    ws.getRow(3).hidden = true;
    wb.calculateFormulas();
    // 1xx variant now drops the hidden row: 10 + 30 = 40
    expect(ws.getCell("B5").result).toBe(40);

    const formula = (ws.getCell("B5").value as { formula: string }).formula;
    expect(formula).toBe("SUBTOTAL(109,T[N])");
  });

  it("SUBTOTAL plain (9) does NOT skip hidden rows (conservative impl)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    ws.getRow(2).hidden = true;
    ws.getCell("B1").value = { formula: "SUBTOTAL(9, A1:A3)", result: 0 };
    wb.calculateFormulas();
    // Plain code 9: our impl conservatively treats all rows as visible
    // because `row.hidden` conflates filter-hide with manual-hide.
    // Sum includes the hidden row.
    expect(ws.getCell("B1").result).toBe(60);
  });

  it("SUBTOTAL 1xx skips empty hidden rows too", () => {
    // Hidden rows with no cells of their own still participate in
    // hidden-row skipping. The adapter collects them via eachRow with
    // includeEmpty so row.hidden is visible even for pure-empty rows.
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    // A2 unset; row 2 marked hidden.
    ws.getCell("A3").value = 30;
    ws.getRow(2).hidden = true;
    ws.getCell("B1").value = { formula: "SUBTOTAL(109, A1:A3)", result: 0 };
    wb.calculateFormulas();
    // Row 2 is empty + hidden → skipped; no change vs unhidden (blank
    // contributes 0). But the point here is the adapter captures the
    // hidden flag via includeEmpty iteration.
    expect(ws.getCell("B1").result).toBe(40);
  });

  it("SUBTOTAL 109 over a multi-area reference preserves hidden-row skipping", () => {
    // Multi-area refs (A1:A3, A5:A7) flatten through dereferenceValue;
    // we need per-area hidden/subtotal masks to be merged into the
    // flattened ArrayValue so an outer SUBTOTAL still respects them.
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let r = 1; r <= 7; r++) {
      ws.getCell(`A${r}`).value = r * 10;
    }
    // Hide row 2 (part of first area) and row 6 (part of second area).
    ws.getRow(2).hidden = true;
    ws.getRow(6).hidden = true;
    ws.getCell("C1").value = { formula: "SUBTOTAL(109, A1:A3, A5:A7)", result: 0 };
    wb.calculateFormulas();
    // Values: A1=10, A2=20(hidden), A3=30, A5=50, A6=60(hidden), A7=70
    // Skipping 20 and 60 → 10+30+50+70 = 160
    expect(ws.getCell("C1").result).toBe(160);
  });

  it("SUBTOTAL over multi-area still skips nested SUBTOTAL cells in any area", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = { formula: "SUBTOTAL(9, A1:A2)", result: 0 }; // 3
    ws.getCell("B1").value = 10;
    ws.getCell("B2").value = 20;
    ws.getCell("B3").value = { formula: "SUBTOTAL(9, B1:B2)", result: 0 }; // 30
    ws.getCell("C1").value = { formula: "SUBTOTAL(9, A1:A3, B1:B3)", result: 0 };
    wb.calculateFormulas();
    // 1+2+10+20 = 33 (both nested SUBTOTAL cells skipped)
    expect(ws.getCell("C1").result).toBe(33);
  });

  it("nested SUBTOTAL inside table [#Totals] is not double-counted", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    ws.getCell("D1").value = {
      formula: "SUBTOTAL(109,T[[#All],[N]])",
      result: 0
    };
    wb.calculateFormulas();
    // Excel skips the inner SUBTOTAL cell → 10+20+30 = 60 (header is text)
    expect(ws.getCell("D1").result).toBe(60);
  });
});

describe("table totals row: edge cases", () => {
  it("totals row with empty data rows still evaluates without error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
      name: "Empty",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Name" }, { name: "N", totalsRowFunction: "sum" }],
      rows: [[null, null]]
    });
    wb.calculateFormulas();
    // SUBTOTAL(109, Empty[N]) over all-null data should be 0
    expect(ws.getCell("B3").result).toBe(0);
  });

  it("multiple totals columns all compute independently", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.addTable({
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
    wb.calculateFormulas();
    // Totals row is row 5: header + 3 data + totals
    expect(ws.getCell("B5").result).toBe(6); // sum
    expect(ws.getCell("C5").result).toBe(30); // max
    expect(ws.getCell("D5").result).toBe(3); // count
  });
});

// ===========================================================================
// AGGREGATE — full option matrix
// ===========================================================================

describe("AGGREGATE: option semantics", () => {
  function makeSheet(): { wb: Workbook; ws: ReturnType<Workbook["addWorksheet"]> } {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // A1..A5 with an error in A3 and a hidden row at row 4.
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = { error: "#DIV/0!" } as unknown as number;
    ws.getCell("A4").value = 40;
    ws.getCell("A5").value = 50;
    ws.getRow(4).hidden = true;
    return { wb, ws };
  }

  it("option 4 (ignore nothing) propagates the error", () => {
    const { wb, ws } = makeSheet();
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 4, A1:A5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
  });

  it("option 6 (ignore errors) sums the non-error cells including hidden", () => {
    const { wb, ws } = makeSheet();
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 6, A1:A5)", result: 0 };
    wb.calculateFormulas();
    // 10 + 20 + 40 + 50 = 120 (error skipped, hidden kept)
    expect(ws.getCell("B1").result).toBe(120);
  });

  it("option 5 (ignore hidden) keeps the error and fails", () => {
    const { wb, ws } = makeSheet();
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 5, A1:A5)", result: 0 };
    wb.calculateFormulas();
    // Error not skipped
    expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
  });

  it("option 7 (ignore hidden + errors) sums visible non-error cells", () => {
    const { wb, ws } = makeSheet();
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 7, A1:A5)", result: 0 };
    wb.calculateFormulas();
    // 10 + 20 + 50 = 80
    expect(ws.getCell("B1").result).toBe(80);
  });

  it("option 3 (ignore hidden + errors + nested) matches 7 for simple data", () => {
    const { wb, ws } = makeSheet();
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 3, A1:A5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(80);
  });

  it("function code 13 (MODE.SNGL) is supported", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 2;
    ws.getCell("A4").value = 3;
    ws.getCell("B1").value = { formula: "AGGREGATE(13, 0, A1:A4)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(2);
  });

  it("function code 16 (PERCENTILE.INC) is supported", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    ws.getCell("B1").value = { formula: "AGGREGATE(16, 0, A1:A5, 0.5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(3);
  });

  it("function code 17 (QUARTILE.INC) is supported", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    // Quartile code 2 = median
    ws.getCell("B1").value = { formula: "AGGREGATE(17, 0, A1:A5, 2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(3);
  });

  it("function code 18 (PERCENTILE.EXC) is supported", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    ws.getCell("B1").value = { formula: "AGGREGATE(18, 0, A1:A5, 0.5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(3);
  });

  it("function code 19 (QUARTILE.EXC) is supported", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 7; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    ws.getCell("B1").value = { formula: "AGGREGATE(19, 0, A1:A7, 2)", result: 0 };
    wb.calculateFormulas();
    // Quartile.exc 2 of [1..7] = 4
    expect(ws.getCell("B1").result).toBe(4);
  });

  it("AGGREGATE with option 0 skips nested SUBTOTAL (no double-count)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("A4").value = { formula: "SUBTOTAL(9, A1:A3)", result: 0 }; // nested = 6
    // AGGREGATE over A1:A4 with option 0 should skip A4 → 1+2+3 = 6
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 0, A1:A4)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(6);
  });

  it("AGGREGATE with option 4 does NOT skip nested SUBTOTAL (double-count)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getCell("A4").value = { formula: "SUBTOTAL(9, A1:A3)", result: 0 };
    ws.getCell("B1").value = { formula: "AGGREGATE(9, 4, A1:A4)", result: 0 };
    wb.calculateFormulas();
    // Option 4 keeps nested → 1+2+3+6 = 12
    expect(ws.getCell("B1").result).toBe(12);
  });
});
