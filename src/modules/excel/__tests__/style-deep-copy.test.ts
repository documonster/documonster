import { cellAlignment, cellBorder, cellFill, cellFont, cellProtection } from "@excel/core/cell";
import {
  columnSetAlignment,
  columnSetBorder,
  columnSetFill,
  columnSetFont,
  columnSetProtection,
  getCell,
  getColumn,
  getRow,
  rowSplice
} from "@excel/core/worksheet";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { describe, it, expect } from "vitest";

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
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetFont(getColumn(ws, "A"), { name: "Arial", size: 12, bold: true });
      getRow(ws, 1).style = { font: {} };

      const cell = getCell(ws, "A1");
      expect(cellFont(cell)).toBeDefined();
      expect(cellFont(cell)!.name).toBe("Arial");
      expect(cellFont(cell)!.bold).toBe(true);
    });

    it("inherits column alignment when row alignment is empty {}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetAlignment(getColumn(ws, "A"), {
        horizontal: "center" as const,
        vertical: "middle" as const
      });
      getRow(ws, 1).style = { alignment: {} };

      const cell = getCell(ws, "A1");
      expect(cellAlignment(cell)).toBeDefined();
      expect(cellAlignment(cell)!.horizontal).toBe("center");
    });

    it("inherits column border when row border is empty {}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetBorder(getColumn(ws, "A"), {
        top: { style: "thin" as const },
        bottom: { style: "thin" as const }
      });
      getRow(ws, 1).style = { border: {} };

      const cell = getCell(ws, "A1");
      expect(cellBorder(cell)).toBeDefined();
      expect(cellBorder(cell)!.top!.style).toBe("thin");
    });

    it("inherits column fill when row fill is empty {}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetFill(getColumn(ws, "A"), {
        type: "pattern" as const,
        pattern: "solid" as const,
        fgColor: { argb: "FFFF0000" }
      });
      getRow(ws, 1).style = { fill: {} as any };

      const cell = getCell(ws, "A1");
      expect(cellFill(cell)).toBeDefined();
      expect((cellFill(cell) as any).type).toBe("pattern");
    });

    it("inherits column protection when row protection is empty {}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetProtection(getColumn(ws, "A"), { locked: true });
      getRow(ws, 1).style = { protection: {} };

      const cell = getCell(ws, "A1");
      expect(cellProtection(cell)).toBeDefined();
      expect(cellProtection(cell)!.locked).toBe(true);
    });

    it("non-empty row style still takes priority over column style", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetFont(getColumn(ws, "A"), { name: "Arial", size: 10 });
      getRow(ws, 1).style = { font: { name: "Helvetica", size: 14 } };

      const cell = getCell(ws, "A1");
      expect(cellFont(cell)!.name).toBe("Helvetica");
      expect(cellFont(cell)!.size).toBe(14);
    });

    it("column style is used when row has no style at all", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      columnSetFont(getColumn(ws, "A"), { name: "Courier", size: 11 });

      const cell = getCell(ws, "A1");
      expect(cellFont(cell)!.name).toBe("Courier");
    });
  });

  // ---------------------------------------------------------------------------
  // Bug 3: duplicateRow / spliceRows / Row.splice – styles must be deep-copied
  // so that mutating a copied row or cell style does not affect the original.
  // ---------------------------------------------------------------------------
  describe("duplicateRow deep-copies styles", () => {
    it("mutating duplicated cell style does not affect source", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "Name");
      Cell.setStyle(ws, "A1", { font: { name: "Calibri", size: 11, bold: true } });
      Cell.setStyle(ws, "A1", {
        border: {
          top: { style: "thin" as const },
          bottom: { style: "thin" as const }
        }
      });

      Worksheet.duplicateRow(ws, 1, 1, true);

      // Values should match
      expect(Cell.getStyle(ws, "A2").font).toEqual(Cell.getStyle(ws, "A1").font);
      expect(Cell.getStyle(ws, "A2").border).toEqual(Cell.getStyle(ws, "A1").border);

      // Mutate the duplicate
      Cell.setStyle(ws, "A2", { font: { name: "Courier", size: 14 } });
      Cell.setStyle(ws, "A2", { border: { top: { style: "double" as const } } });

      // Original must be unaffected
      expect(Cell.getStyle(ws, "A1").font!.name).toBe("Calibri");
      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
      expect(Cell.getStyle(ws, "A1").border!.top!.style).toBe("thin");
      expect(Cell.getStyle(ws, "A1").border!.bottom!.style).toBe("thin");
    });

    it("duplicated styles are equal in value but not the same reference", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "test");
      Cell.setStyle(ws, "A1", {
        fill: {
          type: "pattern" as const,
          pattern: "solid" as const,
          fgColor: { argb: "FF0000FF" }
        }
      });

      Worksheet.duplicateRow(ws, 1, 1, true);

      expect(Cell.getStyle(ws, "A2").fill).toEqual(Cell.getStyle(ws, "A1").fill);
      expect(Cell.getStyle(ws, "A2").fill).not.toBe(Cell.getStyle(ws, "A1").fill);
    });
  });

  describe("spliceRows deep-copies styles when shifting rows", () => {
    it("inserting rows: shifted row style is independent from original", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "Row1");
      Cell.setStyle(ws, "A1", { font: { name: "Helvetica", size: 12, bold: true } });

      Cell.setValue(ws, "A2", "Row2");
      Cell.setStyle(ws, "A2", { font: { name: "Georgia", size: 14 } });
      Cell.setStyle(ws, "A2", {
        border: {
          left: { style: "thin" as const, color: { argb: "FF00FF00" } }
        }
      });

      // Insert 1 empty row at position 2 — pushes old row 2 down to row 3
      Worksheet.spliceRows(ws, 2, 0, []);

      expect(Cell.getStyle(ws, "A3").font!.name).toBe("Georgia");
      expect(Cell.getStyle(ws, "A3").border!.left!.style).toBe("thin");

      // Mutate the shifted row
      Cell.setStyle(ws, "A3", { font: { name: "Verdana", size: 10 } });

      // Row 1 must be unaffected
      expect(Cell.getStyle(ws, "A1").font!.name).toBe("Helvetica");
    });

    it("removing rows: shifted row style is independent from original", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "Row1");
      Cell.setValue(ws, "A2", "Row2");
      Cell.setValue(ws, "A3", "Row3");
      Cell.setStyle(ws, "A3", { font: { name: "Impact", size: 20, bold: true } });
      Cell.setStyle(ws, "A3", {
        fill: {
          type: "pattern" as const,
          pattern: "solid" as const,
          fgColor: { argb: "FF00FFFF" }
        }
      });

      // Remove row 2 — row 3 shifts up to row 2
      Worksheet.spliceRows(ws, 2, 1);

      expect(Cell.getStyle(ws, "A2").font!.name).toBe("Impact");
      expect((Cell.getStyle(ws, "A2").fill as any).fgColor.argb).toBe("FF00FFFF");

      // Mutate
      Cell.setStyle(ws, "A2", { font: { name: "Comic Sans", size: 8 } });

      // Row 1 must be unaffected
      expect(Cell.getValue(ws, "A1")).toBe("Row1");
    });
  });

  describe("Row.splice deep-copies styles when shifting cells", () => {
    it("inserting cells: shifted cell style is independent", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "stay");
      Cell.setValue(ws, "B1", "shift");
      Cell.setStyle(ws, "B1", { font: { name: "Arial", size: 12, bold: true } });

      // Insert one cell at B, pushing B1 to C1
      rowSplice(Worksheet.getRow(ws, 1), 2, 0, "inserted");

      expect(Cell.getStyle(ws, "C1").font!.name).toBe("Arial");
      expect(Cell.getStyle(ws, "C1").font!.bold).toBe(true);

      // Mutate shifted cell
      Cell.setStyle(ws, "C1", { font: { name: "Verdana", size: 10 } });

      // A1 must be unaffected (sanity)
      expect(Cell.getValue(ws, "A1")).toBe("stay");
    });

    it("removing cells: shifted cell style is independent", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet");

      Cell.setValue(ws, "A1", "remove");
      Cell.setValue(ws, "B1", "shift-left");
      Cell.setStyle(ws, "B1", { font: { name: "Georgia", size: 14, italic: true } });
      Cell.setValue(ws, "C1", "also-shift");
      Cell.setStyle(ws, "C1", { border: { top: { style: "thin" as const } } });

      // Remove A1 — B1 shifts to A1, C1 shifts to B1
      rowSplice(Worksheet.getRow(ws, 1), 1, 1);

      expect(Cell.getStyle(ws, "A1").font!.name).toBe("Georgia");
      expect(Cell.getStyle(ws, "B1").border!.top!.style).toBe("thin");

      // Mutate
      Cell.setStyle(ws, "A1", { font: { name: "Other", size: 8 } });

      // B1 must be unaffected
      expect(Cell.getStyle(ws, "B1").border!.top!.style).toBe("thin");
    });
  });
});
