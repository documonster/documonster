import { testUtils } from "@excel/__tests__/shared";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Worksheet", () => {
  describe("Styles", () => {
    // =========================================================================
    // Row Style Inheritance
    // =========================================================================

    it("sets row styles — all cells in row inherit", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("basket");

      ws.getCell("A1").value = 5;
      ws.getCell("A1").numFmt = testUtils.styles.numFmts.numFmt1;
      ws.getCell("A1").font = testUtils.styles.fonts.arialBlackUI14;

      ws.getCell("C1").value = "Hello, World!";
      ws.getCell("C1").alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getCell("C1").border = testUtils.styles.borders.doubleRed;
      ws.getCell("C1").fill = testUtils.styles.fills.redDarkVertical;

      ws.getRow(1).numFmt = testUtils.styles.numFmts.numFmt2;
      ws.getRow(1).font = testUtils.styles.fonts.comicSansUdB16;
      ws.getRow(1).alignment = testUtils.styles.namedAlignments.middleCentre;
      ws.getRow(1).border = testUtils.styles.borders.thin;
      ws.getRow(1).fill = testUtils.styles.fills.redGreenDarkTrellis;

      // Existing cell A1 should be overwritten by row style
      expect(ws.getCell("A1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("A1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("A1").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("A1").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("A1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // B1 didn't exist, should not be materialized
      expect(ws.findCell("B1")).toBeUndefined();

      // Existing cell C1 should also be overwritten
      expect(ws.getCell("C1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("C1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("C1").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("C1").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("C1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // New cell B1 (created via getCell) should inherit the row styles
      expect(ws.getCell("B1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("B1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("B1").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("B1").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("B1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });

    // =========================================================================
    // Column Style Inheritance
    // =========================================================================

    it("sets col styles — all cells in column inherit", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("basket");

      ws.getCell("A1").value = 5;
      ws.getCell("A1").numFmt = testUtils.styles.numFmts.numFmt1;
      ws.getCell("A1").font = testUtils.styles.fonts.arialBlackUI14;

      ws.getCell("A3").value = "Hello, World!";
      ws.getCell("A3").alignment = testUtils.styles.namedAlignments.bottomRight;
      ws.getCell("A3").border = testUtils.styles.borders.doubleRed;
      ws.getCell("A3").fill = testUtils.styles.fills.redDarkVertical;

      ws.getColumn("A").numFmt = testUtils.styles.numFmts.numFmt2;
      ws.getColumn("A").font = testUtils.styles.fonts.comicSansUdB16;
      ws.getColumn("A").alignment = testUtils.styles.namedAlignments.middleCentre;
      ws.getColumn("A").border = testUtils.styles.borders.thin;
      ws.getColumn("A").fill = testUtils.styles.fills.redGreenDarkTrellis;

      // Existing cell A1 should be overwritten
      expect(ws.getCell("A1").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("A1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("A1").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("A1").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("A1").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // Row 2 didn't exist, should not be materialized
      expect(ws.findRow(2)).toBeUndefined();

      // Existing cell A3 should be overwritten
      expect(ws.getCell("A3").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("A3").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("A3").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("A3").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("A3").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);

      // New cell A2 (created via getCell) should inherit column styles
      expect(ws.getCell("A2").numFmt).toEqual(testUtils.styles.numFmts.numFmt2);
      expect(ws.getCell("A2").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws.getCell("A2").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell("A2").border).toEqual(testUtils.styles.borders.thin);
      expect(ws.getCell("A2").fill).toEqual(testUtils.styles.fills.redGreenDarkTrellis);
    });

    // =========================================================================
    // Cell-Level Styles
    // =========================================================================

    it("cell styles are independent from each other", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").font = { bold: true };
      ws.getCell("A2").font = { italic: true };

      expect(ws.getCell("A1").font!.bold).toBe(true);
      expect(ws.getCell("A1").font!.italic).toBeUndefined();
      expect(ws.getCell("A2").font!.italic).toBe(true);
      expect(ws.getCell("A2").font!.bold).toBeUndefined();
    });

    it("cell numFmt can be set and read independently", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = 0.5;
      ws.getCell("A1").numFmt = "0.00%";
      ws.getCell("B1").value = 1234;
      ws.getCell("B1").numFmt = "#,##0";

      expect(ws.getCell("A1").numFmt).toBe("0.00%");
      expect(ws.getCell("B1").numFmt).toBe("#,##0");
    });

    it("cell alignment properties work", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true
      };

      expect(ws.getCell("A1").alignment!.horizontal).toBe("center");
      expect(ws.getCell("A1").alignment!.vertical).toBe("middle");
      expect(ws.getCell("A1").alignment!.wrapText).toBe(true);
    });

    it("cell fill properties work", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" }
      };

      const fill = ws.getCell("A1").fill as {
        type: string;
        pattern: string;
        fgColor: { argb: string };
      };
      expect(fill.type).toBe("pattern");
      expect(fill.pattern).toBe("solid");
      expect(fill.fgColor.argb).toBe("FFFF0000");
    });

    it("cell border properties work", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").border = {
        top: { style: "thin" },
        bottom: { style: "double", color: { argb: "FF0000FF" } }
      };

      expect(ws.getCell("A1").border!.top).toEqual({ style: "thin" });
      expect(ws.getCell("A1").border!.bottom).toEqual({
        style: "double",
        color: { argb: "FF0000FF" }
      });
    });

    // =========================================================================
    // Style Precedence
    // =========================================================================

    it("row style does not affect cells in other rows", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "row 1";
      ws.getCell("A2").value = "row 2";

      ws.getRow(1).font = testUtils.styles.fonts.arialBlackUI14;

      expect(ws.getCell("A1").font).toEqual(testUtils.styles.fonts.arialBlackUI14);
      // Row 2 should NOT have the font
      expect(ws.getCell("A2").font).toBeUndefined();
    });

    it("column style does not affect cells in other columns", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "col A";
      ws.getCell("B1").value = "col B";

      ws.getColumn("A").font = testUtils.styles.fonts.arialBlackUI14;

      expect(ws.getCell("A1").font).toEqual(testUtils.styles.fonts.arialBlackUI14);
      // Column B should NOT have the font
      expect(ws.getCell("B1").font).toBeUndefined();
    });

    // =========================================================================
    // Style Round-Trip
    // =========================================================================

    it("cell styles survive XLSX round-trip", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = 42;
      ws.getCell("A1").font = { bold: true, size: 16, name: "Arial" };
      ws.getCell("A1").numFmt = "0.00";
      ws.getCell("A1").alignment = { horizontal: "center" };
      ws.getCell("A1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF00FF00" }
      };

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const cell = wb2.getWorksheet("test")!.getCell("A1");
      expect(cell.value).toBe(42);
      expect(cell.font!.bold).toBe(true);
      expect(cell.font!.size).toBe(16);
      expect(cell.numFmt).toBe("0.00");
      expect(cell.alignment!.horizontal).toBe("center");
      const fill = cell.fill as { type: string; pattern: string; fgColor: { argb: string } };
      expect(fill.pattern).toBe("solid");
      expect(fill.fgColor.argb).toBe("FF00FF00");
    });

    it("row numFmt survives XLSX round-trip", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = 1.5;
      ws.getRow(1).numFmt = "0.000";

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.getWorksheet("test")!.getCell("A1").numFmt).toBe("0.000");
    });
  });
});
