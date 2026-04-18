/**
 * Unit tests for `buildWorkbookSnapshot`. The snapshot is the immutable
 * input to the entire compile → evaluate → materialize pipeline, so
 * every formula-engine behaviour ultimately depends on it capturing the
 * right shape of the live workbook.
 */

import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

import { buildWorkbookSnapshot } from "../workbook-adapter";
import { snapshotCellKey } from "../workbook-snapshot";

describe("buildWorkbookSnapshot: basic shape", () => {
  it("captures worksheet name and id", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");
    ws.getCell("A1").value = 42;
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets).toHaveLength(1);
    expect(snap.worksheets[0].name).toBe("Data");
    expect(snap.worksheets[0].id).toBe(ws.id);
  });

  it("builds worksheet name lookup (lowercase)", () => {
    const wb = new Workbook();
    wb.addWorksheet("Foo");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheetsByName.has("foo")).toBe(true);
    expect(snap.worksheetsByName.has("FOO")).toBe(false);
  });

  it("builds worksheet id lookup", () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("A");
    const ws2 = wb.addWorksheet("B");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheetsById.get(ws1.id)?.name).toBe("A");
    expect(snap.worksheetsById.get(ws2.id)?.name).toBe("B");
  });

  it("handles empty workbook (no sheets)", () => {
    const wb = new Workbook();
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets).toHaveLength(0);
    expect(snap.worksheetsByName.size).toBe(0);
  });
});

describe("buildWorkbookSnapshot: cell capture", () => {
  it("captures numeric values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 42;
    const snap = buildWorkbookSnapshot(wb);
    const cell = snap.worksheets[0].cells.get(snapshotCellKey(1, 1));
    expect(cell?.value).toBe(42);
    expect(cell?.formulaKind).toBe("none");
  });

  it("captures string values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("B2").value = "hello";
    const snap = buildWorkbookSnapshot(wb);
    const cell = snap.worksheets[0].cells.get(snapshotCellKey(2, 2));
    expect(cell?.value).toBe("hello");
  });

  it("captures boolean values", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = true;
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets[0].cells.get(snapshotCellKey(1, 1))?.value).toBe(true);
  });

  it("does not emit empty cells (sparse representation)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("C3").value = 3;
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets[0].cells.size).toBe(2);
  });

  it("captures Date values as Excel serial numbers", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = new Date(Date.UTC(2024, 0, 15));
    const snap = buildWorkbookSnapshot(wb);
    const v = snap.worksheets[0].cells.get(snapshotCellKey(1, 1))?.value;
    expect(typeof v).toBe("number");
    // 2024-01-15 = serial 45306
    expect(v).toBeGreaterThan(45000);
    expect(v).toBeLessThan(46000);
  });

  it("captures formula cells with kind=normal", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "1+1", result: 2 };
    const snap = buildWorkbookSnapshot(wb);
    const cell = snap.worksheets[0].cells.get(snapshotCellKey(1, 1));
    expect(cell?.formulaKind).toBe("normal");
    expect(cell?.formula).toBe("1+1");
    expect(cell?.cachedResult).toBe(2);
  });

  it("captures a formula cell's cached result", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { formula: "SUM(1,2,3)", result: 6 };
    const snap = buildWorkbookSnapshot(wb);
    const cell = snap.worksheets[0].cells.get(snapshotCellKey(1, 1));
    expect(cell?.cachedResult).toBe(6);
  });
});

describe("buildWorkbookSnapshot: error value sanitization (R8)", () => {
  it("preserves known Excel error codes", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = { error: "#N/A" } as unknown as string;
    const snap = buildWorkbookSnapshot(wb);
    const v = snap.worksheets[0].cells.get(snapshotCellKey(1, 1))?.value;
    expect(v).toEqual({ error: "#N/A" });
  });

  it("gates unknown error codes to #VALUE!", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    // Inject an unknown error string
    ws.getCell("A1").value = { error: "custom-err" } as unknown as string;
    const snap = buildWorkbookSnapshot(wb);
    const v = snap.worksheets[0].cells.get(snapshotCellKey(1, 1))?.value;
    expect(v).toEqual({ error: "#VALUE!" });
  });
});

describe("buildWorkbookSnapshot: defined names", () => {
  it("captures workbook-level defined name", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 100;
    wb.definedNames.add("Sheet1!A1", "TaxRate");
    const snap = buildWorkbookSnapshot(wb);
    // Name should be findable (keys are uppercase)
    const found = Array.from(snap.definedNames.keys()).find(k =>
      k.toUpperCase().includes("TAXRATE")
    );
    expect(found).toBeDefined();
  });

  it("empty workbook has no defined names", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.definedNames.size).toBe(0);
  });
});

describe("buildWorkbookSnapshot: calc properties", () => {
  it("default iterative off", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.calcProperties.iterate).toBeFalsy();
  });

  it("captures iterate flag", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    wb.calcProperties = { iterate: true, iterateCount: 50, iterateDelta: 0.0001 };
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.calcProperties.iterate).toBe(true);
    expect(snap.calcProperties.iterateCount).toBe(50);
    expect(snap.calcProperties.iterateDelta).toBe(0.0001);
  });
});

describe("buildWorkbookSnapshot: date1904 property", () => {
  it("defaults to false", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.properties.date1904).toBe(false);
  });

  it("captures date1904=true from workbook properties", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    wb.properties = { date1904: true };
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.properties.date1904).toBe(true);
  });
});

describe("buildWorkbookSnapshot: hiddenRows capture", () => {
  it("empty worksheet has empty hiddenRows set", () => {
    const wb = new Workbook();
    wb.addWorksheet("S");
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets[0].hiddenRows.size).toBe(0);
  });

  it("captures a hidden row that contains data", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    ws.getRow(2).hidden = true;
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets[0].hiddenRows.has(2)).toBe(true);
    expect(snap.worksheets[0].hiddenRows.has(1)).toBe(false);
    expect(snap.worksheets[0].hiddenRows.has(3)).toBe(false);
  });

  it("captures a hidden row that has no cells", () => {
    // Pure empty hidden row — adapter must use includeEmpty iteration
    // so the hidden flag is observed even without populated cells.
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    ws.getCell("A1").value = 1;
    ws.getCell("A3").value = 3;
    // Row 2 is empty; mark it hidden anyway.
    ws.getRow(2).hidden = true;
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheets[0].hiddenRows.has(2)).toBe(true);
  });

  it("captures multiple hidden rows across the sheet", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("S");
    for (let i = 1; i <= 5; i++) {
      ws.getCell(`A${i}`).value = i;
    }
    ws.getRow(1).hidden = true;
    ws.getRow(3).hidden = true;
    ws.getRow(5).hidden = true;
    const snap = buildWorkbookSnapshot(wb);
    const hidden = snap.worksheets[0].hiddenRows;
    expect(hidden.has(1)).toBe(true);
    expect(hidden.has(2)).toBe(false);
    expect(hidden.has(3)).toBe(true);
    expect(hidden.has(4)).toBe(false);
    expect(hidden.has(5)).toBe(true);
    expect(hidden.size).toBe(3);
  });

  it("independent per-sheet hiddenRows sets", () => {
    const wb = new Workbook();
    const s1 = wb.addWorksheet("S1");
    const s2 = wb.addWorksheet("S2");
    s1.getCell("A1").value = 1;
    s1.getRow(1).hidden = true;
    s2.getCell("A1").value = 1;
    // S2 not hidden.
    const snap = buildWorkbookSnapshot(wb);
    expect(snap.worksheetsByName.get("s1")?.hiddenRows.has(1)).toBe(true);
    expect(snap.worksheetsByName.get("s2")?.hiddenRows.has(1)).toBe(false);
  });
});
