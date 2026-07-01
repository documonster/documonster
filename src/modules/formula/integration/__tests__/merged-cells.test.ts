/**
 * Regression tests for merged cells — they should not double
 * their value when included in a SUM / COUNT / AVERAGE range.
 *
 * Excel's behaviour: only the top-left "master" of a merged range
 * holds the value; every other cell in the merge rectangle reads as
 * blank for formula purposes. The host's in-memory model exposes
 * `MergeValue` slaves that forward `cell.value` to the master, so the
 * snapshot builder must filter them out (otherwise a `=SUM(B2:C9)`
 * over a sheet with `B2:C2` merged sees the value at both B2 and C2
 * and counts it twice).
 */

import { ValueType } from "@excel/core/enums";
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("calculate-formulas: merged cells", () => {
  it("does not double-count a horizontally merged value in SUM", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B2", 20281.57);
    Cell.setValue(ws, "B4", 5887.5);
    Worksheet.merge(ws, "B2:C2");
    Worksheet.merge(ws, "B4:C4");
    Cell.setValue(ws, "E1", { formula: "SUM(B2:C9)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBeCloseTo(26169.07, 4);
  });

  it("does not double-count a vertically merged value in SUM", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 100);
    Worksheet.merge(ws, "A1:A3");
    Cell.setValue(ws, "B1", { formula: "SUM(A1:A3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(100);
  });

  it("counts a 2D merged value once in SUM and COUNT", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 7);
    Worksheet.merge(ws, "A1:B2"); // 4 cells, master A1
    Cell.setValue(ws, "D1", { formula: "SUM(A1:B2)", result: 0 });
    Cell.setValue(ws, "D2", { formula: "COUNT(A1:B2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(7);
    expect(Cell.getResult(ws, "D2")).toBe(1);
  });

  it("ignores merge slaves in AVERAGE", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 10);
    Cell.setValue(ws, "A2", 20);
    Worksheet.merge(ws, "A1:B1");
    // Range A1:B2 contains: A1=10, B1=slave(of A1), A2=20, B2=blank
    // Real Excel AVERAGE = (10 + 20) / 2 = 15 (not (10+10+20)/3)
    Cell.setValue(ws, "D1", { formula: "AVERAGE(A1:B2)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "D1")).toBe(15);
  });

  it("treats a merge slave as blank in COUNTA", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", "hello");
    Worksheet.merge(ws, "A1:C1"); // master A1, slaves B1+C1
    Cell.setValue(ws, "E1", { formula: "COUNTA(A1:C1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "E1")).toBe(1);
  });

  it("evaluates a formula stored on the master cell once", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", 4);
    Cell.setValue(ws, "B1", { formula: "A1*5", result: 0 }); // master, =20
    Worksheet.merge(ws, "B1:C1");
    Cell.setValue(ws, "E1", { formula: "SUM(B1:C1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "B1")).toBe(20);
    expect(Cell.getResult(ws, "E1")).toBe(20);
  });
});

describe("calculate-formulas: dynamic-array spill vs merged regions", () => {
  it("a spill that would land on a merge slave returns #SPILL!", () => {
    // Spill range [B1:B2] tries to land on B2, which is a slave of the
    // merge A2:C2 (master A2). The slave proxies value writes to the
    // master via `MergeValue`, so without merge-aware spill checking
    // the spill silently corrupts A2.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A2", "MASTER");
    Worksheet.merge(ws, "A2:C2");
    Cell.setValue(ws, "B1", { formula: "SEQUENCE(2,1)", result: 0 });

    calculateFormulas(wb);

    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#SPILL!" });
    // Master must NOT have been overwritten by the spill payload.
    expect(Cell.getValue(ws, "A2")).toBe("MASTER");
    // Merge geometry must remain intact.
    expect(Cell.getType(ws, "B2")).toBe(ValueType.Merge);
  });

  it("a spill that fully covers a merged region (master + slaves) returns #SPILL!", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "B1", "MERGED");
    Worksheet.merge(ws, "B1:C1");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(1,3)", result: 0 });

    calculateFormulas(wb);

    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
    expect(Cell.getValue(ws, "B1")).toBe("MERGED");
  });

  it("a spill that does not touch any merged region succeeds", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "F2", "X");
    Worksheet.merge(ws, "F2:G2");
    // Spill into A1:A3, nowhere near the merge
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3,1)", result: 0 });

    calculateFormulas(wb);

    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "A2")).toBe(2);
    expect(Cell.getValue(ws, "A3")).toBe(3);
  });

  it("a spill onto an empty merged region (master holds no value) still returns #SPILL!", () => {
    // Pre-fix: the snapshot for an empty merge had every slave with
    // value=null, so the value/formula occupancy check let the spill
    // through, corrupting the master with the spill payload.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Worksheet.merge(ws, "A2:C2"); // master A2 has no value
    Cell.setValue(ws, "B1", { formula: "SEQUENCE(2,1)", result: 0 });

    calculateFormulas(wb);

    expect(Cell.getResult(ws, "B1")).toEqual({ error: "#SPILL!" });
    // Master must remain unset — spill must not have written into it.
    expect(Cell.getValue(ws, "A2")).toBeNull();
  });

  it("stale-ghost cleanup must not clobber a merge created over old ghosts", () => {
    // Regression: after a SEQUENCE(3,1) spill at A1 (ghosts at A2, A3),
    // the user merges A2:B3 and writes "USER" to the new master A2.
    // Recalc with a 1x1 result triggers cleanup of the previous spill's
    // ghosts. The cleanup must NOT overwrite A2 (the new merge master)
    // nor any of its slaves (which would forward to the master via
    // `MergeValue`'s setter and silently wipe "USER").
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3,1)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getValue(ws, "A2")).toBe(2);
    expect(Cell.getValue(ws, "A3")).toBe(3);

    Worksheet.merge(ws, "A2:B3");
    Cell.setValue(ws, "A2", "USER");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(1,1)", result: 0 });
    calculateFormulas(wb);

    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "A2")).toBe("USER");
    // Slaves still proxy to master — confirms merge geometry is intact
    expect(Cell.getValue(ws, "B2")).toBe("USER");
    expect(Cell.getValue(ws, "A3")).toBe("USER");
    expect(Cell.getValue(ws, "B3")).toBe("USER");
  });

  it("unmerging unblocks a previously-blocked spill", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Worksheet.merge(ws, "B1:C1");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(1,3)", result: 0 });
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });

    Worksheet.unmerge(ws, "B1:C1");
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toBe(1);
    expect(Cell.getValue(ws, "B1")).toBe(2);
    expect(Cell.getValue(ws, "C1")).toBe(3);
  });

  it("a dynamic-array formula sitting in a merge master cannot spill into its own slaves", () => {
    // Pathological: user writes SEQUENCE(3,1) at A1 then merges A1:A3.
    // The spill targets A2 and A3 — both slaves of the same merge.
    // Excel reports #SPILL!.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3,1)", result: 0 });
    Worksheet.merge(ws, "A1:A3");
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
  });

  it("a dynamic-array source inside a merge returns #SPILL! even when ghosts land outside the merge", () => {
    // Source A1 is the master of A1:B1 (horizontal merge) and the
    // formula is SEQUENCE(3,1) which spills vertically into A2, A3.
    // The ghosts A2/A3 do NOT touch the merge, but Excel still reports
    // #SPILL! because the source itself is part of a merged region.
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "Sheet1");
    Cell.setValue(ws, "A1", { formula: "SEQUENCE(3,1)", result: 0 });
    Worksheet.merge(ws, "A1:B1");
    calculateFormulas(wb);
    expect(Cell.getResult(ws, "A1")).toEqual({ error: "#SPILL!" });
    // Ghosts must NOT have been written.
    expect(Cell.getValue(ws, "A2")).toBeNull();
    expect(Cell.getValue(ws, "A3")).toBeNull();
  });
});
