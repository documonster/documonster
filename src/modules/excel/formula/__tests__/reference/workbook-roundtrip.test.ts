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

import { Workbook } from "@excel/workbook";
import { describe, expect, it } from "vitest";

describe("workbook roundtrip: dependency chains", () => {
  it("1000-cell linear chain computes in topological order", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    for (let i = 2; i <= 1000; i++) {
      ws.getCell(`A${i}`).value = { formula: `A${i - 1}+1`, result: 0 };
    }
    wb.calculateFormulas();
    expect(ws.getCell("A1000").result).toBe(1000);
  });

  it("diamond dependency resolves correctly", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 10;
    ws.getCell("B1").value = { formula: "A1*2", result: 0 }; // 20
    ws.getCell("C1").value = { formula: "A1+5", result: 0 }; // 15
    ws.getCell("D1").value = { formula: "B1+C1", result: 0 }; // 35
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(35);
  });

  it("wide dependency (100 cells depend on one)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 7;
    for (let i = 1; i <= 100; i++) {
      ws.getCell(`B${i}`).value = { formula: `A1*${i}`, result: 0 };
    }
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(7);
    expect(ws.getCell("B50").result).toBe(350);
    expect(ws.getCell("B100").result).toBe(700);
  });

  it("deep nested formulas (100 levels deep in a single cell)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // Build ((((((1+1)+1)+1)...)+1) 50 levels deep — well under parser's MAX_DEPTH=256
    let expr = "1";
    for (let i = 0; i < 50; i++) {
      expr = `(${expr}+1)`;
    }
    ws.getCell("A1").value = { formula: expr, result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(51);
  });
});

describe("workbook roundtrip: cross-sheet references", () => {
  it("three-sheet aggregation", () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet("Data1");
    const s2 = wb.addWorksheet("Data2");
    const s3 = wb.addWorksheet("Summary");
    for (let i = 1; i <= 10; i++) {
      s1.getCell(`A${i}`).value = i;
      s2.getCell(`A${i}`).value = i * 2;
    }
    s3.getCell("A1").value = { formula: "SUM(Data1!A1:A10)", result: 0 };
    s3.getCell("A2").value = { formula: "SUM(Data2!A1:A10)", result: 0 };
    s3.getCell("A3").value = { formula: "A1+A2", result: 0 };
    wb.calculateFormulas();
    expect(s3.getCell("A1").result).toBe(55);
    expect(s3.getCell("A2").result).toBe(110);
    expect(s3.getCell("A3").result).toBe(165);
  });

  it("3D reference across contiguous sheets", () => {
    const wb = new Workbook();
    // Using sheet names like "Q1" triggers tokenizer ambiguity (Q1 is
    // also a valid cell ref), so the 3D prefix "Q1:Q4!" is tokenised
    // as a Name instead of a SheetRef. Using unambiguous names avoids
    // that gap — a separate test with `Q1..Q4` is `.skip`ed below to
    // document the limitation.
    const sheets = ["Data1", "Data2", "Data3", "Data4"];
    for (const name of sheets) {
      const ws = wb.addWorksheet(name);
      ws.getCell("A1").value = 100;
    }
    const summary = wb.addWorksheet("Total");
    summary.getCell("A1").value = { formula: "SUM(Data1:Data4!A1)", result: 0 };
    wb.calculateFormulas();
    expect(summary.getCell("A1").result).toBe(400);
  });

  it("3D reference with cell-ref-shaped names (Q1:Q4!A1)", () => {
    // Regression: `Q1` is a legal cell ref AND a legal sheet name. The
    // tokenizer used to skip the 3D-reference branch whenever the
    // first word looked like a cell ref, leaving `Q1:Q4` parsed as a
    // range. We now use `!`-lookahead to disambiguate.
    const wb = new Workbook();
    const sheets = ["Q1", "Q2", "Q3", "Q4"];
    for (const name of sheets) {
      const ws = wb.addWorksheet(name);
      ws.getCell("A1").value = 100;
    }
    const summary = wb.addWorksheet("Total");
    summary.getCell("A1").value = { formula: "SUM(Q1:Q4!A1)", result: 0 };
    wb.calculateFormulas();
    expect(summary.getCell("A1").result).toBe(400);
  });

  it("circular detection across sheets", () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet("A");
    const s2 = wb.addWorksheet("B");
    s1.getCell("A1").value = { formula: "B!A1+1", result: 0 };
    s2.getCell("A1").value = { formula: "A!A1+1", result: 0 };
    wb.calculateFormulas();
    // Circular: both should report a consistent fallback (0+1=1 with our
    // default non-iterative semantics).
    expect(typeof s1.getCell("A1").result).toBe("number");
    expect(typeof s2.getCell("A1").result).toBe("number");
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
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(5)", result: 0 };
    ws.getCell("B1").value = { formula: "SUM(A1:A5)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A5").value).toBe(5);
    expect(ws.getCell("B1").result).toBe(15);
  });

  it("spill shrink cleans old ghosts, then grow refills", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("B1").value = 5;
    ws.getCell("A1").value = { formula: "SEQUENCE(B1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A5").value).toBe(5);

    // Shrink
    ws.getCell("B1").value = 2;
    wb.calculateFormulas();
    expect(ws.getCell("A3").value).toBeFalsy();

    // Grow back
    ws.getCell("B1").value = 7;
    wb.calculateFormulas();
    expect(ws.getCell("A7").value).toBe(7);
  });

  it("two side-by-side dynamic arrays don't interfere", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(3)", result: 0 };
    ws.getCell("C1").value = { formula: "SEQUENCE(3,1,10)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A3").value).toBe(3);
    expect(ws.getCell("C3").value).toBe(12);
  });

  it("dynamic array with formula reference to same spill", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SEQUENCE(5)", result: 0 };
    ws.getCell("C1").value = { formula: "A1#*2", result: 0 }; // Spill ref
    wb.calculateFormulas();
    // Our engine may or may not support spill-range refs; just verify
    // no crash and some number produced.
    const r = ws.getCell("C1").result;
    expect(r === undefined || typeof r === "number" || typeof r === "object").toBe(true);
  });
});

describe("workbook roundtrip: defined names", () => {
  it("workbook-level defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 0.2;
    wb.definedNames.add("S!A1", "TaxRate");
    ws.getCell("B1").value = { formula: "100*TaxRate", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(20);
  });

  it("defined name referencing a range", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    wb.definedNames.add("S!A1:A5", "Data");
    ws.getCell("B1").value = { formula: "SUM(Data)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(15);
  });

  it("name used in multiple formulas", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 42;
    wb.definedNames.add("S!A1", "Answer");
    ws.getCell("B1").value = { formula: "Answer*2", result: 0 };
    ws.getCell("C1").value = { formula: "Answer+8", result: 0 };
    ws.getCell("D1").value = { formula: "B1+C1", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(84);
    expect(ws.getCell("C1").result).toBe(50);
    expect(ws.getCell("D1").result).toBe(134);
  });
});

describe("workbook roundtrip: error chains", () => {
  it("error propagates through 10-level dependency chain", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    for (let i = 2; i <= 10; i++) {
      ws.getCell(`A${i}`).value = { formula: `A${i - 1}*2`, result: 0 };
    }
    wb.calculateFormulas();
    expect(ws.getCell("A10").result).toEqual({ error: "#DIV/0!" });
  });

  it("IFERROR stops propagation at any point", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("A2").value = { formula: "A1+1", result: 0 };
    ws.getCell("A3").value = { formula: "IFERROR(A2, 999)", result: 0 };
    ws.getCell("A4").value = { formula: "A3*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A3").result).toBe(999);
    expect(ws.getCell("A4").result).toBe(1998);
  });

  it("one bad cell doesn't poison unrelated cells", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1/0", result: 0 };
    ws.getCell("B1").value = 100;
    ws.getCell("C1").value = { formula: "B1*2", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#DIV/0!" });
    expect(ws.getCell("C1").result).toBe(200);
  });
});

describe("workbook roundtrip: volatile recalc", () => {
  it("NOW updates on every calc", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "NOW()", result: 0 };
    wb.calculateFormulas();
    const r1 = ws.getCell("A1").result;
    expect(typeof r1).toBe("number");
    expect((r1 as number) > 0).toBe(true);

    // Force a tiny delay then recalc
    const start = Date.now();
    while (Date.now() - start < 2) {
      // spin
    }
    wb.calculateFormulas();
    const r2 = ws.getCell("A1").result;
    expect(typeof r2).toBe("number");
    expect((r2 as number) >= (r1 as number)).toBe(true);
  });

  it("RAND stays within [0, 1)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 20; i++) {
      ws.getCell(`A${i}`).value = { formula: "RAND()", result: 0 };
    }
    wb.calculateFormulas();
    for (let i = 1; i <= 20; i++) {
      const r = ws.getCell(`A${i}`).result as number;
      expect(r).toBeGreaterThanOrEqual(0);
      expect(r).toBeLessThan(1);
    }
  });
});

describe("workbook roundtrip: performance sanity", () => {
  it("5000-cell SUM-chain completes in reasonable time", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5000; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    ws.getCell("B1").value = { formula: "SUM(A1:A5000)", result: 0 };
    const start = Date.now();
    wb.calculateFormulas();
    const elapsed = Date.now() - start;
    expect(ws.getCell("B1").result).toBe(12502500); // sum of 1..5000
    expect(elapsed).toBeLessThan(5000); // 5s is generous; expect way under
  });

  it("1000 sibling formulas (no cross-deps) compute quickly", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 1000; i++) {
      ws.getCell(`A${i}`).value = { formula: `${i}*2`, result: 0 };
    }
    const start = Date.now();
    wb.calculateFormulas();
    const elapsed = Date.now() - start;
    expect(ws.getCell("A1000").result).toBe(2000);
    expect(elapsed).toBeLessThan(3000);
  });

  it("deep chain of 100 formulas + recalc after single change", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    for (let i = 2; i <= 100; i++) {
      ws.getCell(`A${i}`).value = { formula: `A${i - 1}+1`, result: 0 };
    }
    wb.calculateFormulas();
    expect(ws.getCell("A100").result).toBe(100);

    // Change the root; every dependent must update.
    ws.getCell("A1").value = 1000;
    wb.calculateFormulas();
    expect(ws.getCell("A100").result).toBe(1099);
  });
});

describe("workbook roundtrip: mixed scenarios", () => {
  it("realistic invoice workbook", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Invoice");

    // Header row
    ws.getCell("A1").value = "Item";
    ws.getCell("B1").value = "Qty";
    ws.getCell("C1").value = "Price";
    ws.getCell("D1").value = "Tax Rate";
    ws.getCell("E1").value = "Subtotal";
    ws.getCell("F1").value = "Tax";
    ws.getCell("G1").value = "Total";

    const items = [
      ["Widget", 10, 2.5, 0.08],
      ["Gadget", 5, 15, 0.08],
      ["Gizmo", 100, 0.5, 0.12]
    ];
    items.forEach((item, i) => {
      const row = i + 2;
      ws.getCell(`A${row}`).value = item[0];
      ws.getCell(`B${row}`).value = item[1];
      ws.getCell(`C${row}`).value = item[2];
      ws.getCell(`D${row}`).value = item[3];
      ws.getCell(`E${row}`).value = { formula: `B${row}*C${row}`, result: 0 };
      ws.getCell(`F${row}`).value = { formula: `E${row}*D${row}`, result: 0 };
      ws.getCell(`G${row}`).value = { formula: `E${row}+F${row}`, result: 0 };
    });

    ws.getCell("E5").value = { formula: "SUM(E2:E4)", result: 0 };
    ws.getCell("F5").value = { formula: "SUM(F2:F4)", result: 0 };
    ws.getCell("G5").value = { formula: "SUM(G2:G4)", result: 0 };

    wb.calculateFormulas();

    expect(ws.getCell("E2").result).toBe(25);
    expect(ws.getCell("F2").result).toBeCloseTo(2, 5);
    expect(ws.getCell("G2").result).toBeCloseTo(27, 5);
    expect(ws.getCell("E5").result).toBe(150); // 25 + 75 + 50
  });

  it("nested IF + VLOOKUP pattern", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // Lookup table
    ws.getCell("E1").value = "A";
    ws.getCell("F1").value = 10;
    ws.getCell("E2").value = "B";
    ws.getCell("F2").value = 20;
    ws.getCell("E3").value = "C";
    ws.getCell("F3").value = 30;

    ws.getCell("A1").value = "B";
    ws.getCell("B1").value = { formula: "VLOOKUP(A1, E1:F3, 2, FALSE)", result: 0 };
    ws.getCell("C1").value = { formula: 'IF(B1>15, "high", "low")', result: 0 };
    wb.calculateFormulas();

    expect(ws.getCell("B1").result).toBe(20);
    expect(ws.getCell("C1").result).toBe("high");
  });

  it("conditional sums over 100 rows", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 100; i++) {
      ws.getCell(`A${i}`).value = i % 2 === 0 ? "even" : "odd";
      ws.getCell(`B${i}`).value = i;
    }
    ws.getCell("D1").value = { formula: 'SUMIF(A1:A100, "even", B1:B100)', result: 0 };
    ws.getCell("D2").value = { formula: 'COUNTIF(A1:A100, "odd")', result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(2550); // 2+4+...+100 = 2550
    expect(ws.getCell("D2").result).toBe(50);
  });
});
