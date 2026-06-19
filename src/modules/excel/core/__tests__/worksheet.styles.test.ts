import { testUtils } from "@excel/__tests__/shared";
import { cellAlignment, cellFill, cellFont, cellGetValue, cellNumFmt } from "@excel/core/cell";
import {
  rowSetAlignment,
  rowSetBorder,
  rowSetFill,
  rowSetFont,
  rowSetNumFmt
} from "@excel/core/row";
import {
  columnSetAlignment,
  columnSetBorder,
  columnSetFill,
  columnSetFont,
  columnSetNumFmt,
  findCell,
  getCell,
  getColumn
} from "@excel/core/worksheet";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Styles", () => {
    // =========================================================================
    // Row Style Inheritance
    // =========================================================================

    it("sets row styles — all cells in row inherit", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "basket");

      Cell.setValue(ws, "A1", 5);
      Cell.setStyle(ws, "A1", { numFmt: testUtils.styles.numFmts.numFmt1 });
      Cell.setStyle(ws, "A1", { font: testUtils.styles.fonts.arialBlackUI14 });

      Cell.setValue(ws, "C1", "Hello, World!");
      Cell.setStyle(ws, "C1", { alignment: testUtils.styles.namedAlignments.bottomRight });
      Cell.setStyle(ws, "C1", { border: testUtils.styles.borders.doubleRed });
      Cell.setStyle(ws, "C1", { fill: testUtils.styles.fills.redDarkVertical });

      rowSetNumFmt(Worksheet.getRow(ws, 1), testUtils.styles.numFmts.numFmt2);
      rowSetFont(Worksheet.getRow(ws, 1), testUtils.styles.fonts.comicSansUdB16);
      rowSetAlignment(Worksheet.getRow(ws, 1), testUtils.styles.namedAlignments.middleCentre);
      rowSetBorder(Worksheet.getRow(ws, 1), testUtils.styles.borders.thin);
      rowSetFill(Worksheet.getRow(ws, 1), testUtils.styles.fills.redGreenDarkTrellis);

      // Existing cell A1 should be overwritten by row style
      expect(Cell.getStyle(ws, "A1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "A1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "A1").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "A1").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "A1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // B1 didn't exist, should not be materialized
      expect(findCell(ws, "B1")).toBeUndefined();

      // Existing cell C1 should also be overwritten
      expect(Cell.getStyle(ws, "C1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "C1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "C1").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "C1").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "C1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // New cell B1 (created via getCell) should inherit the row styles
      expect(Cell.getStyle(ws, "B1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "B1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "B1").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "B1").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "B1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });

    // =========================================================================
    // Column Style Inheritance
    // =========================================================================

    it("sets col styles — all cells in column inherit", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "basket");

      Cell.setValue(ws, "A1", 5);
      Cell.setStyle(ws, "A1", { numFmt: testUtils.styles.numFmts.numFmt1 });
      Cell.setStyle(ws, "A1", { font: testUtils.styles.fonts.arialBlackUI14 });

      Cell.setValue(ws, "A3", "Hello, World!");
      Cell.setStyle(ws, "A3", { alignment: testUtils.styles.namedAlignments.bottomRight });
      Cell.setStyle(ws, "A3", { border: testUtils.styles.borders.doubleRed });
      Cell.setStyle(ws, "A3", { fill: testUtils.styles.fills.redDarkVertical });

      columnSetNumFmt(getColumn(ws, "A"), testUtils.styles.numFmts.numFmt2);
      columnSetFont(getColumn(ws, "A"), testUtils.styles.fonts.comicSansUdB16);
      columnSetAlignment(getColumn(ws, "A"), testUtils.styles.namedAlignments.middleCentre);
      columnSetBorder(getColumn(ws, "A"), testUtils.styles.borders.thin);
      columnSetFill(getColumn(ws, "A"), testUtils.styles.fills.redGreenDarkTrellis);

      // Existing cell A1 should be overwritten
      expect(Cell.getStyle(ws, "A1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "A1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "A1").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "A1").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "A1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // Row 2 didn't exist, should not be materialized
      expect(Worksheet.findRow(ws, 2)).toBeUndefined();

      // Existing cell A3 should be overwritten
      expect(Cell.getStyle(ws, "A3").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "A3").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "A3").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "A3").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "A3").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // New cell A2 (created via getCell) should inherit column styles
      expect(Cell.getStyle(ws, "A2").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(Cell.getStyle(ws, "A2").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws, "A2").alignment).toEqual(
        testUtils.styles.namedAlignments.middleCentre
      );
      expect(Cell.getStyle(ws, "A2").border).toEqual(testUtils.styles.borders.thin);
      expect(Cell.getStyle(ws, "A2").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });

    // =========================================================================
    // Cell-Level Styles
    // =========================================================================

    it("cell styles are independent from each other", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setStyle(ws, "A1", { font: { bold: true } });
      Cell.setStyle(ws, "A2", { font: { italic: true } });

      expect(Cell.getStyle(ws, "A1").font!.bold).toBe(true);
      expect(Cell.getStyle(ws, "A1").font!.italic).toBeUndefined();
      expect(Cell.getStyle(ws, "A2").font!.italic).toBe(true);
      expect(Cell.getStyle(ws, "A2").font!.bold).toBeUndefined();
    });

    it("cell numFmt can be set and read independently", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", 0.5);
      Cell.setStyle(ws, "A1", { numFmt: "0.00%" });
      Cell.setValue(ws, "B1", 1234);
      Cell.setStyle(ws, "B1", { numFmt: "#,##0" });

      expect(Cell.getStyle(ws, "A1").numFmt).toBe("0.00%");
      expect(Cell.getStyle(ws, "B1").numFmt).toBe("#,##0");
    });

    it("cell alignment properties work", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setStyle(ws, "A1", {
        alignment: {
          horizontal: "center",
          vertical: "middle",
          wrapText: true
        }
      });

      expect(Cell.getStyle(ws, "A1").alignment!.horizontal).toBe("center");
      expect(Cell.getStyle(ws, "A1").alignment!.vertical).toBe("middle");
      expect(Cell.getStyle(ws, "A1").alignment!.wrapText).toBe(true);
    });

    it("cell fill properties work", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setStyle(ws, "A1", {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        }
      });

      const fill = Cell.getStyle(ws, "A1").fill as {
        type: string;
        pattern: string;
        fgColor: { argb: string };
      };
      expect(fill.type).toBe("pattern");
      expect(fill.pattern).toBe("solid");
      expect(fill.fgColor.argb).toBe("FFFF0000");
    });

    it("cell border properties work", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setStyle(ws, "A1", {
        border: {
          top: { style: "thin" },
          bottom: { style: "double", color: { argb: "FF0000FF" } }
        }
      });

      expect(Cell.getStyle(ws, "A1").border!.top).toEqual({ style: "thin" });
      expect(Cell.getStyle(ws, "A1").border!.bottom).toEqual({
        style: "double",
        color: { argb: "FF0000FF" }
      });
    });

    // =========================================================================
    // Style Precedence
    // =========================================================================

    it("row style does not affect cells in other rows", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "row 1");
      Cell.setValue(ws, "A2", "row 2");

      rowSetFont(Worksheet.getRow(ws, 1), testUtils.styles.fonts.arialBlackUI14);

      expect(Cell.getStyle(ws, "A1").font).toEqual(testUtils.styles.fonts.arialBlackUI14);
      // Row 2 should NOT have the font
      expect(Cell.getStyle(ws, "A2").font).toBeUndefined();
    });

    it("column style does not affect cells in other columns", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "col A");
      Cell.setValue(ws, "B1", "col B");

      columnSetFont(getColumn(ws, "A"), testUtils.styles.fonts.arialBlackUI14);

      expect(Cell.getStyle(ws, "A1").font).toEqual(testUtils.styles.fonts.arialBlackUI14);
      // Column B should NOT have the font
      expect(Cell.getStyle(ws, "B1").font).toBeUndefined();
    });

    // =========================================================================
    // Style Round-Trip
    // =========================================================================

    it("cell styles survive XLSX round-trip", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", 42);
      Cell.setStyle(ws, "A1", { font: { bold: true, size: 16, name: "Arial" } });
      Cell.setStyle(ws, "A1", { numFmt: "0.00" });
      Cell.setStyle(ws, "A1", { alignment: { horizontal: "center" } });
      Cell.setStyle(ws, "A1", {
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF00FF00" }
        }
      });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const cell = getCell(Workbook.getWorksheet(wb2, "test")!, "A1");
      expect(cellGetValue(cell)).toBe(42);
      expect(cellFont(cell)!.bold).toBe(true);
      expect(cellFont(cell)!.size).toBe(16);
      expect(cellNumFmt(cell)).toBe("0.00");
      expect(cellAlignment(cell)!.horizontal).toBe("center");
      const fill = cellFill(cell) as { type: string; pattern: string; fgColor: { argb: string } };
      expect(fill.pattern).toBe("solid");
      expect(fill.fgColor.argb).toBe("FF00FF00");
    });

    it("row numFmt survives XLSX round-trip", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", 1.5);
      rowSetNumFmt(Worksheet.getRow(ws, 1), "0.000");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      expect(Cell.getStyle(Workbook.getWorksheet(wb2, "test")!, "A1").numFmt).toBe("0.000");
    });
  });
});
