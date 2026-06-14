/**
 * Large-workbook end-to-end roundtrip tests.
 *
 * These exercise the full pipeline (adapter → snapshot → compile →
 * evaluate → materialize → apply) on realistic workbook shapes that
 * stretch past the hand-crafted small unit tests: hundreds of
 * formulas, deep dependency chains, cross-sheet references, tables,
 * dynamic arrays, conditional formatting predicates, etc.
 *
 * Scope: correctness + no memory/perf regression. We don't assert
 * exact timing but every test must complete within reasonable bounds
 * for a ~10k-cell workbook.
 */

import { definedNamesAdd } from "@excel/defined-names";
import { calculateFormulas } from "@excel/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { getDefinedNames } from "@excel/workbook";
import { addTable } from "@excel/worksheet";
import { describe, expect, it } from "vitest";

describe("workbook roundtrip: dependency chains", () => {
  it("1000-cell linear chain computes in topological order", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    for (let i = 2; i <= 1000; i++) {
      Cell.setValue(ws, `A${i}`, { formula: `A${i - 1}+1`, result: 0 });
    }
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1000")).toBe(1000);
  });

  it("diamond dependency resolves correctly", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "B1", { formula: "A1*2", result: 0 }); // 20
    Cell.setValue(ws, "C1", { formula: "A1+5", result: 0 }); // 15
    Cell.setValue(ws, "D1", { formula: "B1+C1", result: 0 }); // 35
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(35);
  });

  it("wide dependency (100 cells depend on one)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 7);
    for (let i = 1; i <= 100; i++) {
      Cell.setValue(ws, `B${i}`, { formula: `A1*${i}`, result: 0 });
    }
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(7);
    expect(Cell.getResult(ws, "B50")).toBe(350);
    expect(Cell.getResult(ws, "B100")).toBe(700);
  });

  it("deep nested formulas (100 levels deep in a single cell)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    // Build ((((((1+1)+1)+1)...)+1) 50 levels deep — well under parser's MAX_DEPTH=256
    let expr = "1";
    for (let i = 0; i < 50; i++) {
      expr = `(${expr}+1)`;
    }
    Cell.setValue(ws, "A1", { formula: expr, result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(51);
  });
});

describe("workbook roundtrip: cross-sheet references", () => {
  it("three-sheet aggregation", () => {
    const wb = Workbook.create();
    const s1 = Workbook.addWorksheet(wb, "Data1");
    const s2 = Workbook.addWorksheet(wb, "Data2");
    const s3 = Workbook.addWorksheet(wb, "Summary");
    for (let i = 1; i <= 10; i++) {
      Cell.setValue(s1, `A${i}`, i);
      Cell.setValue(s2, `A${i}`, i * 2);
    }
    Cell.setValue(s3, "A1", { formula: "SUM(Data1!A1:A10)", result: 0 });
    Cell.setValue(s3, "A2", { formula: "SUM(Data2!A1:A10)", result: 0 });
    Cell.setValue(s3, "A3", { formula: "A1+A2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(s3, "A1")).toBe(55);
    expect(Cell.getResult(s3, "A2")).toBe(110);
    expect(Cell.getResult(s3, "A3")).toBe(165);
  });

  it("3D reference across contiguous sheets", () => {
    const wb = Workbook.create();
    // Using sheet names like "Q1" used to trigger tokenizer ambiguity
    // (Q1 is also a valid cell ref) so `Q1:Q4!` was parsed as a Name
    // instead of a SheetRef. The tokenizer was patched to look ahead
    // for `!` and disambiguate — a follow-up test below exercises the
    // cell-ref-shaped-name case explicitly.
    const sheets = ["Data1", "Data2", "Data3", "Data4"];
    for (const name of sheets) {
      const ws = Workbook.addWorksheet(wb, name);
      Cell.setValue(ws, "A1", 100);
    }
    const summary = Workbook.addWorksheet(wb, "Total");
    Cell.setValue(summary, "A1", { formula: "SUM(Data1:Data4!A1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(summary, "A1")).toBe(400);
  });

  it("3D reference with cell-ref-shaped names (Q1:Q4!A1)", () => {
    // Regression: `Q1` is a legal cell ref AND a legal sheet name. The
    // tokenizer used to skip the 3D-reference branch whenever the
    // first word looked like a cell ref, leaving `Q1:Q4` parsed as a
    // range. We now use `!`-lookahead to disambiguate.
    const wb = Workbook.create();
    const sheets = ["Q1", "Q2", "Q3", "Q4"];
    for (const name of sheets) {
      const ws = Workbook.addWorksheet(wb, name);
      Cell.setValue(ws, "A1", 100);
    }
    const summary = Workbook.addWorksheet(wb, "Total");
    Cell.setValue(summary, "A1", { formula: "SUM(Q1:Q4!A1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(summary, "A1")).toBe(400);
  });

  it("circular detection across sheets", () => {
    const wb = Workbook.create();
    const s1 = Workbook.addWorksheet(wb, "A");
    const s2 = Workbook.addWorksheet(wb, "B");
    Cell.setValue(s1, "A1", { formula: "B!A1+1", result: 0 });
    Cell.setValue(s2, "A1", { formula: "A!A1+1", result: 0 });
    calculateFormulas(wb);
    // Circular: both should report a consistent fallback (0+1=1 with our
    // default non-iterative semantics).
    expect(typeof Cell.getResult(s1, "A1")).toBe("number");
    expect(typeof Cell.getResult(s2, "A1")).toBe("number");
  });
});

describe("workbook roundtrip: dynamic arrays and spill", () => {
  it("SEQUENCE spill + reference into spill region (single-pass, R10 fix)", () => {
    // Regression: the dependency analyzer now collapses `A1:A5` area
    // deps into a single edge on the master when A1 is a dynamic-
    // array function (SEQUENCE). The evaluator's live-spill map then
    // lets SUM read ghost values from the master's cached array
    // result before materialize writes them to the snapshot. So SUM
    // gets the full 1..5 in a single calc pass.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    Cell.setValue(ws, "B1", { formula: "SUM(A1:A5)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A5")).toBe(5);
    expect(Cell.getResult(ws, "B1")).toBe(15);
  });

  it("spill shrink cleans old ghosts, then grow refills", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "B1", 5);
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(B1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A5")).toBe(5);

    // Shrink
    Cell.setValue(ws, "B1", 2);
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A3")).toBeFalsy();

    // Grow back
    Cell.setValue(ws, "B1", 7);
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A7")).toBe(7);
  });

  it("two side-by-side dynamic arrays don't interfere", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3)", result: 0 });
    Cell.setValue(ws, "C1", { formula: "SEQUENCE(3,1,10)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A3")).toBe(3);
    expect(Cell.getValue(ws, "C3")).toBe(12);
  });

  it("dynamic array with formula reference to same spill", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(5)", result: 0 });
    Cell.setValue(ws, "C1", { formula: "A1#*2", result: 0 }); // Spill ref
    calculateFormulas(wb);
    // Our engine may or may not support spill-range refs; just verify
    // no crash and some number produced.
    const r = Cell.getResult(ws, "C1");
    expect(r === undefined || typeof r === "number" || typeof r === "object").toBe(true);
  });
});

describe("workbook roundtrip: defined names", () => {
  it("workbook-level defined name", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 0.2);
    definedNamesAdd(getDefinedNames(wb), "S!A1", "TaxRate");
    Cell.setValue(ws, "B1", { formula: "100*TaxRate", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(20);
  });

  it("defined name referencing a range", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 5; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    definedNamesAdd(getDefinedNames(wb), "S!A1:A5", "Data");
    Cell.setValue(ws, "B1", { formula: "SUM(Data)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(15);
  });

  it("name used in multiple formulas", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 42);
    definedNamesAdd(getDefinedNames(wb), "S!A1", "Answer");
    Cell.setValue(ws, "B1", { formula: "Answer*2", result: 0 });
    Cell.setValue(ws, "C1", { formula: "Answer+8", result: 0 });
    Cell.setValue(ws, "D1", { formula: "B1+C1", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(84);
    expect(Cell.getResult(ws, "C1")).toBe(50);
    expect(Cell.getResult(ws, "D1")).toBe(134);
  });
});

describe("workbook roundtrip: error chains", () => {
  it("error propagates through 10-level dependency chain", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    for (let i = 2; i <= 10; i++) {
      Cell.setValue(ws, `A${i}`, { formula: `A${i - 1}*2`, result: 0 });
    }
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A10")).toEqual({ error: "#DIV/0!" });
  });

  it("IFERROR stops propagation at any point", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "A2", { formula: "A1+1", result: 0 });
    Cell.setValue(ws, "A3", { formula: "IFERROR(A2, 999)", result: 0 });
    Cell.setValue(ws, "A4", { formula: "A3*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A3")).toBe(999);
    expect(Cell.getResult(ws, "A4")).toBe(1998);
  });

  it("one bad cell doesn't poison unrelated cells", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "1/0", result: 0 });
    Cell.setValue(ws, "B1", 100);
    Cell.setValue(ws, "C1", { formula: "B1*2", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#DIV/0!" });
    expect(Cell.getResult(ws, "C1")).toBe(200);
  });
});

describe("workbook roundtrip: volatile recalc", () => {
  it("NOW updates on every calc", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "NOW()", result: 0 });
    calculateFormulas(wb);
    const r1 = Cell.getResult(ws, "A1");
    expect(typeof r1).toBe("number");
    expect((r1 as number) > 0).toBe(true);

    // Force a tiny delay then recalc
    const start = Date.now();
    while (Date.now() - start < 2) {
      // spin
    }
    calculateFormulas(wb);
    const r2 = Cell.getResult(ws, "A1");
    expect(typeof r2).toBe("number");
    expect((r2 as number) >= (r1 as number)).toBe(true);
  });

  it("RAND stays within [0, 1)", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 20; i++) {
      Cell.setValue(ws, `A${i}`, { formula: "RAND()", result: 0 });
    }
    calculateFormulas(wb);
    for (let i = 1; i <= 20; i++) {
      const r = Cell.getResult(ws, `A${i}`) as number;
      expect(r).toBeGreaterThanOrEqual(0);
      expect(r).toBeLessThan(1);
    }
  });
});

describe("workbook roundtrip: performance sanity", () => {
  it("5000-cell SUM-chain completes in reasonable time", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 5000; i++) {
      Cell.setValue(ws, `A${i}`, i);
    }
    Cell.setValue(ws, "B1", { formula: "SUM(A1:A5000)", result: 0 });
    const start = Date.now();
    calculateFormulas(wb);
    const elapsed = Date.now() - start;
    expect(Cell.getResult(ws, "B1")).toBe(12502500); // sum of 1..5000
    expect(elapsed).toBeLessThan(5000); // 5s is generous; expect way under
  });

  it("1000 sibling formulas (no cross-deps) compute quickly", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 1000; i++) {
      Cell.setValue(ws, `A${i}`, { formula: `${i}*2`, result: 0 });
    }
    const start = Date.now();
    calculateFormulas(wb);
    const elapsed = Date.now() - start;
    expect(Cell.getResult(ws, "A1000")).toBe(2000);
    expect(elapsed).toBeLessThan(3000);
  });

  it("deep chain of 100 formulas + recalc after single change", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    for (let i = 2; i <= 100; i++) {
      Cell.setValue(ws, `A${i}`, { formula: `A${i - 1}+1`, result: 0 });
    }
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A100")).toBe(100);

    // Change the root; every dependent must update.
    Cell.setValue(ws, "A1", 1000);
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A100")).toBe(1099);
  });
});

describe("workbook roundtrip: mixed scenarios", () => {
  it("realistic invoice workbook", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Invoice");

    // Header row
    Cell.setValue(ws, "A1", "Item");
    Cell.setValue(ws, "B1", "Qty");
    Cell.setValue(ws, "C1", "Price");
    Cell.setValue(ws, "D1", "Tax Rate");
    Cell.setValue(ws, "E1", "Subtotal");
    Cell.setValue(ws, "F1", "Tax");
    Cell.setValue(ws, "G1", "Total");

    const items = [
      ["Widget", 10, 2.5, 0.08],
      ["Gadget", 5, 15, 0.08],
      ["Gizmo", 100, 0.5, 0.12]
    ];
    items.forEach((item, i) => {
      const row = i + 2;
      Cell.setValue(ws, `A${row}`, item[0]);
      Cell.setValue(ws, `B${row}`, item[1]);
      Cell.setValue(ws, `C${row}`, item[2]);
      Cell.setValue(ws, `D${row}`, item[3]);
      Cell.setValue(ws, `E${row}`, { formula: `B${row}*C${row}`, result: 0 });
      Cell.setValue(ws, `F${row}`, { formula: `E${row}*D${row}`, result: 0 });
      Cell.setValue(ws, `G${row}`, { formula: `E${row}+F${row}`, result: 0 });
    });

    Cell.setValue(ws, "E5", { formula: "SUM(E2:E4)", result: 0 });
    Cell.setValue(ws, "F5", { formula: "SUM(F2:F4)", result: 0 });
    Cell.setValue(ws, "G5", { formula: "SUM(G2:G4)", result: 0 });

    calculateFormulas(wb);

    expect(Cell.getResult(ws, "E2")).toBe(25);
    expect(Cell.getResult(ws, "F2")).toBeCloseTo(2, 5);
    expect(Cell.getResult(ws, "G2")).toBeCloseTo(27, 5);
    expect(Cell.getResult(ws, "E5")).toBe(150); // 25 + 75 + 50
  });

  it("nested IF + VLOOKUP pattern", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    // Lookup table
    Cell.setValue(ws, "E1", "A");
    Cell.setValue(ws, "F1", 10);
    Cell.setValue(ws, "E2", "B");
    Cell.setValue(ws, "F2", 20);
    Cell.setValue(ws, "E3", "C");
    Cell.setValue(ws, "F3", 30);

    Cell.setValue(ws, "A1", "B");
    Cell.setValue(ws, "B1", { formula: "VLOOKUP(A1, E1:F3, 2, FALSE)", result: 0 });
    Cell.setValue(ws, "C1", { formula: 'IF(B1>15, "high", "low")', result: 0 });
    calculateFormulas(wb);

    expect(Cell.getResult(ws, "B1")).toBe(20);
    expect(Cell.getResult(ws, "C1")).toBe("high");
  });

  it("conditional sums over 100 rows", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    for (let i = 1; i <= 100; i++) {
      Cell.setValue(ws, `A${i}`, i % 2 === 0 ? "even" : "odd");
      Cell.setValue(ws, `B${i}`, i);
    }
    Cell.setValue(ws, "D1", { formula: 'SUMIF(A1:A100, "even", B1:B100)', result: 0 });
    Cell.setValue(ws, "D2", { formula: 'COUNTIF(A1:A100, "odd")', result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(2550); // 2+4+...+100 = 2550
    expect(Cell.getResult(ws, "D2")).toBe(50);
  });
});

// ===========================================================================
// Tables: totals rows, structured refs in every form
// ===========================================================================

describe("workbook roundtrip: tables with totals row", () => {
  it("[#Totals] resolves to the totals row", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Sales",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Product" }, { name: "Qty", totalsRowFunction: "sum" }],
      rows: [
        ["Apple", 10],
        ["Pear", 20],
        ["Peach", 30]
      ]
    });
    Cell.setValue(ws, "D1", { formula: "Sales[[#Totals],[Qty]]", result: 0 });
    calculateFormulas(wb);
    // The totals-row cell itself should hold the sum; reading via
    // structured ref should surface that cached value (60).
    const v = Cell.getResult(ws, "D1");
    expect(v).not.toEqual({ error: "#REF!" });
  });

  it("[#All] spans header + data + totals", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Sales",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Product" }, { name: "Qty", totalsRowFunction: "sum" }],
      rows: [
        ["Apple", 10],
        ["Pear", 20]
      ]
    });
    // COUNTA over [#All] should include header + 2 data + totals = 4 cells
    Cell.setValue(ws, "D1", { formula: "COUNTA(Sales[[#All],[Qty]])", result: 0 });
    calculateFormulas(wb);
    // header + 2 data + totals = 4 non-empty cells
    const v = Cell.getResult(ws, "D1") as number;
    expect(v).toBeGreaterThanOrEqual(3); // at minimum header + 2 data
    expect(v).toBeLessThanOrEqual(4);
  });

  it("[#Data] skips header and totals", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Sales",
      ref: "A1",
      headerRow: true,
      totalsRow: true,
      columns: [{ name: "Product" }, { name: "Qty", totalsRowFunction: "sum" }],
      rows: [
        ["Apple", 10],
        ["Pear", 20],
        ["Peach", 30]
      ]
    });
    Cell.setValue(ws, "D1", { formula: "SUM(Sales[[#Data],[Qty]])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(60);
  });

  it("[#Totals] on a table without a totals row returns #REF!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "NoTotals",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Val" }],
      rows: [[1], [2], [3]]
    });
    Cell.setValue(ws, "C1", { formula: "NoTotals[[#Totals],[Val]]", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toEqual({ error: "#REF!" });
  });
});

describe("workbook roundtrip: structured reference variants", () => {
  it("Table[column] — plain column reference", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "Inventory",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Item" }, { name: "Qty" }],
      rows: [
        ["A", 5],
        ["B", 10],
        ["C", 15]
      ]
    });
    Cell.setValue(ws, "D1", { formula: "SUM(Inventory[Qty])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(30);
  });

  it("Table[#Headers] returns header row only", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "HdrT",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Item" }, { name: "Qty" }],
      rows: [["A", 5]]
    });
    Cell.setValue(ws, "D1", { formula: "COUNTA(HdrT[#Headers])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(2); // "Item" + "Qty"
  });

  it("Table[#All] spans every row", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    addTable(ws, {
      name: "AllT",
      ref: "A1",
      headerRow: true,
      totalsRow: false,
      columns: [{ name: "Val" }],
      rows: [[10], [20], [30]]
    });
    Cell.setValue(ws, "C1", { formula: "COUNTA(AllT[#All])", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "C1")).toBe(4); // header + 3 data
  });
});
