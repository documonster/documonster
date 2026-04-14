import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Style fixes", () => {
  // ---------------------------------------------------------------------------
  // Bug 1: _mergeStyle – empty object {} on row/col style should not shadow
  // a real style from the other level.
  //
  // In JavaScript `{}` is truthy, so `(rowStyle && rowStyle.font)` evaluates
  // to `{}` even when font is empty, preventing the column's real font from
  // being used. The fix guards with a check that the object has at least one
  // own property.
  // ---------------------------------------------------------------------------
  describe("_mergeStyle empty-object guard", () => {
    it("inherits column font when row font is empty {}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").font = { name: "Arial", size: 12, bold: true };
      ws.getRow(1).style = { font: {} };

      const cell = ws.getCell("A1");
      expect(cell.font).toBeDefined();
      expect(cell.font!.name).toBe("Arial");
      expect(cell.font!.bold).toBe(true);
    });

    it("inherits column alignment when row alignment is empty {}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").alignment = { horizontal: "center" as const, vertical: "middle" as const };
      ws.getRow(1).style = { alignment: {} };

      const cell = ws.getCell("A1");
      expect(cell.alignment).toBeDefined();
      expect(cell.alignment!.horizontal).toBe("center");
    });

    it("inherits column border when row border is empty {}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").border = {
        top: { style: "thin" as const },
        bottom: { style: "thin" as const }
      };
      ws.getRow(1).style = { border: {} };

      const cell = ws.getCell("A1");
      expect(cell.border).toBeDefined();
      expect(cell.border!.top!.style).toBe("thin");
    });

    it("inherits column fill when row fill is empty {}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").fill = {
        type: "pattern" as const,
        pattern: "solid" as const,
        fgColor: { argb: "FFFF0000" }
      };
      ws.getRow(1).style = { fill: {} as any };

      const cell = ws.getCell("A1");
      expect(cell.fill).toBeDefined();
      expect((cell.fill as any).type).toBe("pattern");
    });

    it("inherits column protection when row protection is empty {}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").protection = { locked: true };
      ws.getRow(1).style = { protection: {} };

      const cell = ws.getCell("A1");
      expect(cell.protection).toBeDefined();
      expect(cell.protection!.locked).toBe(true);
    });

    it("non-empty row style still takes priority over column style", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").font = { name: "Arial", size: 10 };
      ws.getRow(1).style = { font: { name: "Helvetica", size: 14 } };

      const cell = ws.getCell("A1");
      expect(cell.font!.name).toBe("Helvetica");
      expect(cell.font!.size).toBe(14);
    });

    it("column style is used when row has no style at all", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getColumn("A").font = { name: "Courier", size: 11 };

      const cell = ws.getCell("A1");
      expect(cell.font!.name).toBe("Courier");
    });
  });

  // ---------------------------------------------------------------------------
  // Bug 3: duplicateRow / spliceRows / Row.splice – styles must be deep-copied
  // so that mutating a copied row or cell style does not affect the original.
  // ---------------------------------------------------------------------------
  describe("duplicateRow deep-copies styles", () => {
    it("mutating duplicated cell style does not affect source", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "Name";
      ws.getCell("A1").font = { name: "Calibri", size: 11, bold: true };
      ws.getCell("A1").border = {
        top: { style: "thin" as const },
        bottom: { style: "thin" as const }
      };

      ws.duplicateRow(1, 1, true);

      // Values should match
      expect(ws.getCell("A2").font).toEqual(ws.getCell("A1").font);
      expect(ws.getCell("A2").border).toEqual(ws.getCell("A1").border);

      // Mutate the duplicate
      ws.getCell("A2").font = { name: "Courier", size: 14 };
      ws.getCell("A2").border = { top: { style: "double" as const } };

      // Original must be unaffected
      expect(ws.getCell("A1").font!.name).toBe("Calibri");
      expect(ws.getCell("A1").font!.bold).toBe(true);
      expect(ws.getCell("A1").border!.top!.style).toBe("thin");
      expect(ws.getCell("A1").border!.bottom!.style).toBe("thin");
    });

    it("duplicated styles are equal in value but not the same reference", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "test";
      ws.getCell("A1").fill = {
        type: "pattern" as const,
        pattern: "solid" as const,
        fgColor: { argb: "FF0000FF" }
      };

      ws.duplicateRow(1, 1, true);

      expect(ws.getCell("A2").fill).toEqual(ws.getCell("A1").fill);
      expect(ws.getCell("A2").fill).not.toBe(ws.getCell("A1").fill);
    });
  });

  describe("spliceRows deep-copies styles when shifting rows", () => {
    it("inserting rows: shifted row style is independent from original", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "Row1";
      ws.getCell("A1").font = { name: "Helvetica", size: 12, bold: true };

      ws.getCell("A2").value = "Row2";
      ws.getCell("A2").font = { name: "Georgia", size: 14 };
      ws.getCell("A2").border = {
        left: { style: "thin" as const, color: { argb: "FF00FF00" } }
      };

      // Insert 1 empty row at position 2 — pushes old row 2 down to row 3
      ws.spliceRows(2, 0, []);

      expect(ws.getCell("A3").font!.name).toBe("Georgia");
      expect(ws.getCell("A3").border!.left!.style).toBe("thin");

      // Mutate the shifted row
      ws.getCell("A3").font = { name: "Verdana", size: 10 };

      // Row 1 must be unaffected
      expect(ws.getCell("A1").font!.name).toBe("Helvetica");
    });

    it("removing rows: shifted row style is independent from original", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "Row1";
      ws.getCell("A2").value = "Row2";
      ws.getCell("A3").value = "Row3";
      ws.getCell("A3").font = { name: "Impact", size: 20, bold: true };
      ws.getCell("A3").fill = {
        type: "pattern" as const,
        pattern: "solid" as const,
        fgColor: { argb: "FF00FFFF" }
      };

      // Remove row 2 — row 3 shifts up to row 2
      ws.spliceRows(2, 1);

      expect(ws.getCell("A2").font!.name).toBe("Impact");
      expect((ws.getCell("A2").fill as any).fgColor.argb).toBe("FF00FFFF");

      // Mutate
      ws.getCell("A2").font = { name: "Comic Sans", size: 8 };

      // Row 1 must be unaffected
      expect(ws.getCell("A1").value).toBe("Row1");
    });
  });

  describe("Row.splice deep-copies styles when shifting cells", () => {
    it("inserting cells: shifted cell style is independent", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "stay";
      ws.getCell("B1").value = "shift";
      ws.getCell("B1").font = { name: "Arial", size: 12, bold: true };

      // Insert one cell at B, pushing B1 to C1
      ws.getRow(1).splice(2, 0, "inserted");

      expect(ws.getCell("C1").font!.name).toBe("Arial");
      expect(ws.getCell("C1").font!.bold).toBe(true);

      // Mutate shifted cell
      ws.getCell("C1").font = { name: "Verdana", size: 10 };

      // A1 must be unaffected (sanity)
      expect(ws.getCell("A1").value).toBe("stay");
    });

    it("removing cells: shifted cell style is independent", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet");

      ws.getCell("A1").value = "remove";
      ws.getCell("B1").value = "shift-left";
      ws.getCell("B1").font = { name: "Georgia", size: 14, italic: true };
      ws.getCell("C1").value = "also-shift";
      ws.getCell("C1").border = { top: { style: "thin" as const } };

      // Remove A1 — B1 shifts to A1, C1 shifts to B1
      ws.getRow(1).splice(1, 1);

      expect(ws.getCell("A1").font!.name).toBe("Georgia");
      expect(ws.getCell("B1").border!.top!.style).toBe("thin");

      // Mutate
      ws.getCell("A1").font = { name: "Other", size: 8 };

      // B1 must be unaffected
      expect(ws.getCell("B1").border!.top!.style).toBe("thin");
    });
  });
});
