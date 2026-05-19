/**
 * Regression tests for issue #162 — merged cells should not double
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

import { ValueType } from "@excel/enums";
import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

describe("calculate-formulas: merged cells (issue #162)", () => {
  it("does not double-count a horizontally merged value in SUM", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B2").value = 20281.57;
    ws.getCell("B4").value = 5887.5;
    ws.mergeCells("B2:C2");
    ws.mergeCells("B4:C4");
    ws.getCell("E1").value = { formula: "SUM(B2:C9)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBeCloseTo(26169.07, 4);
  });

  it("does not double-count a vertically merged value in SUM", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 100;
    ws.mergeCells("A1:A3");
    ws.getCell("B1").value = { formula: "SUM(A1:A3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(100);
  });

  it("counts a 2D merged value once in SUM and COUNT", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 7;
    ws.mergeCells("A1:B2"); // 4 cells, master A1
    ws.getCell("D1").value = { formula: "SUM(A1:B2)", result: 0 };
    ws.getCell("D2").value = { formula: "COUNT(A1:B2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(7);
    expect(ws.getCell("D2").result).toBe(1);
  });

  it("ignores merge slaves in AVERAGE", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.mergeCells("A1:B1");
    // Range A1:B2 contains: A1=10, B1=slave(of A1), A2=20, B2=blank
    // Real Excel AVERAGE = (10 + 20) / 2 = 15 (not (10+10+20)/3)
    ws.getCell("D1").value = { formula: "AVERAGE(A1:B2)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("D1").result).toBe(15);
  });

  it("treats a merge slave as blank in COUNTA", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "hello";
    ws.mergeCells("A1:C1"); // master A1, slaves B1+C1
    ws.getCell("E1").value = { formula: "COUNTA(A1:C1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("E1").result).toBe(1);
  });

  it("evaluates a formula stored on the master cell once", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 4;
    ws.getCell("B1").value = { formula: "A1*5", result: 0 }; // master, =20
    ws.mergeCells("B1:C1");
    ws.getCell("E1").value = { formula: "SUM(B1:C1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("B1").result).toBe(20);
    expect(ws.getCell("E1").result).toBe(20);
  });
});

describe("calculate-formulas: dynamic-array spill vs merged regions", () => {
  it("a spill that would land on a merge slave returns #SPILL!", () => {
    // Spill range [B1:B2] tries to land on B2, which is a slave of the
    // merge A2:C2 (master A2). The slave proxies value writes to the
    // master via `MergeValue`, so without merge-aware spill checking
    // the spill silently corrupts A2 — see issue #162 follow-up.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A2").value = "MASTER";
    ws.mergeCells("A2:C2");
    ws.getCell("B1").value = { formula: "SEQUENCE(2,1)", result: 0 };

    wb.calculateFormulas();

    expect(ws.getCell("B1").result).toEqual({ error: "#SPILL!" });
    // Master must NOT have been overwritten by the spill payload.
    expect(ws.getCell("A2").value).toBe("MASTER");
    // Merge geometry must remain intact.
    expect(ws.getCell("B2").type).toBe(ValueType.Merge);
  });

  it("a spill that fully covers a merged region (master + slaves) returns #SPILL!", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = "MERGED";
    ws.mergeCells("B1:C1");
    ws.getCell("A1").value = { formula: "SEQUENCE(1,3)", result: 0 };

    wb.calculateFormulas();

    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
    expect(ws.getCell("B1").value).toBe("MERGED");
  });

  it("a spill that does not touch any merged region succeeds", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("F2").value = "X";
    ws.mergeCells("F2:G2");
    // Spill into A1:A3, nowhere near the merge
    ws.getCell("A1").value = { formula: "SEQUENCE(3,1)", result: 0 };

    wb.calculateFormulas();

    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A3").value).toBe(3);
  });

  it("a spill onto an empty merged region (master holds no value) still returns #SPILL!", () => {
    // Pre-fix: the snapshot for an empty merge had every slave with
    // value=null, so the value/formula occupancy check let the spill
    // through, corrupting the master with the spill payload.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.mergeCells("A2:C2"); // master A2 has no value
    ws.getCell("B1").value = { formula: "SEQUENCE(2,1)", result: 0 };

    wb.calculateFormulas();

    expect(ws.getCell("B1").result).toEqual({ error: "#SPILL!" });
    // Master must remain unset — spill must not have written into it.
    expect(ws.getCell("A2").value).toBeNull();
  });

  it("stale-ghost cleanup must not clobber a merge created over old ghosts", () => {
    // Regression: after a SEQUENCE(3,1) spill at A1 (ghosts at A2, A3),
    // the user merges A2:B3 and writes "USER" to the new master A2.
    // Recalc with a 1x1 result triggers cleanup of the previous spill's
    // ghosts. The cleanup must NOT overwrite A2 (the new merge master)
    // nor any of its slaves (which would forward to the master via
    // `MergeValue`'s setter and silently wipe "USER").
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SEQUENCE(3,1)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A2").value).toBe(2);
    expect(ws.getCell("A3").value).toBe(3);

    ws.mergeCells("A2:B3");
    ws.getCell("A2").value = "USER";
    ws.getCell("A1").value = { formula: "SEQUENCE(1,1)", result: 0 };
    wb.calculateFormulas();

    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("A2").value).toBe("USER");
    // Slaves still proxy to master — confirms merge geometry is intact
    expect(ws.getCell("B2").value).toBe("USER");
    expect(ws.getCell("A3").value).toBe("USER");
    expect(ws.getCell("B3").value).toBe("USER");
  });

  it("unmerging unblocks a previously-blocked spill", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.mergeCells("B1:C1");
    ws.getCell("A1").value = { formula: "SEQUENCE(1,3)", result: 0 };
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });

    ws.unMergeCells("B1:C1");
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toBe(1);
    expect(ws.getCell("B1").value).toBe(2);
    expect(ws.getCell("C1").value).toBe(3);
  });

  it("a dynamic-array formula sitting in a merge master cannot spill into its own slaves", () => {
    // Pathological: user writes SEQUENCE(3,1) at A1 then merges A1:A3.
    // The spill targets A2 and A3 — both slaves of the same merge.
    // Excel reports #SPILL!.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SEQUENCE(3,1)", result: 0 };
    ws.mergeCells("A1:A3");
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
  });

  it("a dynamic-array source inside a merge returns #SPILL! even when ghosts land outside the merge", () => {
    // Source A1 is the master of A1:B1 (horizontal merge) and the
    // formula is SEQUENCE(3,1) which spills vertically into A2, A3.
    // The ghosts A2/A3 do NOT touch the merge, but Excel still reports
    // #SPILL! because the source itself is part of a merged region.
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { formula: "SEQUENCE(3,1)", result: 0 };
    ws.mergeCells("A1:B1");
    wb.calculateFormulas();
    expect(ws.getCell("A1").result).toEqual({ error: "#SPILL!" });
    // Ghosts must NOT have been written.
    expect(ws.getCell("A2").value).toBeNull();
    expect(ws.getCell("A3").value).toBeNull();
  });
});
