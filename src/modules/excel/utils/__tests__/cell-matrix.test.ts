import { describe, it, expect } from "vitest";
import { CellMatrix } from "@excel/utils/cell-matrix";
import { colCache } from "@excel/utils/col-cache";

describe("CellMatrix", () => {
  it("getCell creates and caches cells", () => {
    const cm = new CellMatrix();

    const a1 = cm.getCell("Sheet1!A1");
    const a1v2 = cm.getCell("Sheet1!$A$1");
    expect(a1).toBe(a1v2);

    const b2 = cm.getCell("Sheet1!B2");
    expect(b2).not.toBe(a1);

    expect(cm.findCell("Sheet1!A1")).toBe(a1);
    expect(cm.findCell("Sheet1!B2")).toBe(b2);
  });

  it("findCell returns undefined for unknown cells", () => {
    const cm = new CellMatrix();
    expect(cm.findCell("Sheet1!A1")).toBeUndefined();
    expect(cm.findCell("Sheet1!$B$2")).toBeUndefined();
  });

  it("addCell expands ranges", () => {
    const cm = new CellMatrix();
    cm.addCell("Sheet1!A1:B2");

    // Uses 1-based row/col indices (Excel-style)
    expect(cm.findCellAt("Sheet1", 1, 1)).toBeTruthy(); // A1
    expect(cm.findCellAt("Sheet1", 1, 2)).toBeTruthy(); // B1
    expect(cm.findCellAt("Sheet1", 2, 1)).toBeTruthy(); // A2
    expect(cm.findCellAt("Sheet1", 2, 2)).toBeTruthy(); // B2
  });

  it("removeCellEx removes a known cell", () => {
    const cm = new CellMatrix();
    const cell = cm.getCell("Sheet1!C3");
    expect(cm.findCell("Sheet1!C3")).toBe(cell);

    cm.removeCellEx(cell);
    expect(cm.findCell("Sheet1!C3")).toBeUndefined();
  });

  it("spliceRows shifts row storage", () => {
    const cm = new CellMatrix();
    const r1 = cm.getCell("Sheet1!A1");
    const r2 = cm.getCell("Sheet1!A2");
    expect(cm.findCellAt("Sheet1", 1, 1)).toBe(r1);
    expect(cm.findCellAt("Sheet1", 2, 1)).toBe(r2);

    // Remove row index 1
    cm.spliceRows("Sheet1", 1, 1, 0);
    expect(cm.findCellAt("Sheet1", 1, 1)).toBe(r2);
    expect(cm.findCellAt("Sheet1", 2, 1)).toBeUndefined();
  });

  it("spliceColumns shifts column storage", () => {
    const cm = new CellMatrix();
    const c1 = cm.getCell("Sheet1!A1");
    const c2 = cm.getCell("Sheet1!B1");
    expect(cm.findCellAt("Sheet1", 1, 1)).toBe(c1);
    expect(cm.findCellAt("Sheet1", 1, 2)).toBe(c2);

    // Remove column index 1
    cm.spliceColumns("Sheet1", 1, 1, 0);
    expect(cm.findCellAt("Sheet1", 1, 1)).toBe(c2);
    expect(cm.findCellAt("Sheet1", 1, 2)).toBeUndefined();
  });

  it("forEach/map iterates all created cells", () => {
    const cm = new CellMatrix();
    cm.getCell("Sheet1!A1");
    cm.getCell("Sheet1!B2");

    const addresses = cm.map(c => c.address).sort((a, b) => a.localeCompare(b));
    expect(addresses).toEqual(["A1", "B2"]);
  });

  it("clones template deeply and filters prototype pollution keys", () => {
    const template: any = Object.create(null);
    template.style = { nested: { n: 1 } };
    template["__proto__"] = { polluted: true };

    const cm = new CellMatrix(template);
    const a1: any = cm.getCell("Sheet1!A1");
    const a2: any = cm.getCell("Sheet1!A2");

    expect(a1.style).toBeTruthy();
    expect(a2.style).toBeTruthy();
    expect(a1.style).not.toBe(a2.style);
    expect(a1.style.nested).not.toBe(a2.style.nested);

    a1.style.nested.n = 2;
    expect(a2.style.nested.n).toBe(1);

    // __proto__ should not be copied in
    expect((a1 as any).polluted).toBeUndefined();
    expect(({} as any).polluted).toBeUndefined();
  });

  it("handles __proto__ and constructor as safe sheet names via Map", () => {
    const cm = new CellMatrix();

    // Map-based storage is immune to prototype pollution, so these are valid sheet names
    const cell1 = cm.getCellAt("__proto__", 1, 1);
    expect(cell1).toBeTruthy();
    expect(cell1.sheetName).toBe("__proto__");

    const cell2 = cm.getCellAt("constructor", 1, 1);
    expect(cell2).toBeTruthy();
    expect(cell2.sheetName).toBe("constructor");

    // Ensure no prototype pollution occurred
    expect(({} as any).polluted).toBeUndefined();
  });

  it("addCellEx handles decoded ranges", () => {
    const cm = new CellMatrix();
    const decoded = colCache.decodeEx("Sheet1!A1:B2") as any;
    cm.addCellEx(decoded);

    expect(cm.findCellAt("Sheet1", 1, 1)).toBeTruthy();
    expect(cm.findCellAt("Sheet1", 2, 2)).toBeTruthy();
  });
});
