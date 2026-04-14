import { testUtils } from "@excel/__tests__/shared";
import { Anchor } from "@excel/anchor";
import { describe, it, expect, beforeEach } from "vitest";

describe("Anchor", () => {
  describe("colWidth", () => {
    it("should colWidth equals 640000 when worksheet is undefined", () => {
      const anchor = new Anchor();
      expect(anchor.colWidth).toBe(640000);
    });
    it("should colWidth equals 640000 when column has not set custom width", () => {
      const anchor = new Anchor(testUtils.createSheetMock());
      expect(anchor.colWidth).toBe(640000);
    });
    it("should colWidth equals column width", () => {
      const worksheet = testUtils.createSheetMock();
      const anchor = new Anchor(worksheet);
      worksheet.addColumn(anchor.nativeCol + 1, {
        width: 10
      });
      expect(anchor.colWidth).toBe(worksheet.getColumn(anchor.nativeCol + 1).width * 10000);
    });
  });
  describe("rowHeight", () => {
    it("should rowHeight equals 180000 when worksheet is undefined", () => {
      const anchor = new Anchor();
      expect(anchor.rowHeight).toBe(180000);
    });
    it("should rowHeight equals 180000 when row has not set height", () => {
      const anchor = new Anchor(testUtils.createSheetMock());
      expect(anchor.rowHeight).toBe(180000);
    });
    it("should rowHeight equals row height", () => {
      const worksheet = testUtils.createSheetMock();
      worksheet.getRow(1).height = 10;

      const anchor = new Anchor(worksheet);
      expect(anchor.rowHeight).toBe(worksheet.getRow(1).height * 10000);
    });
  });
  describe("resize worksheet`s cells", () => {
    let worksheet: any;
    let anchor: any;

    beforeEach(() => {
      worksheet = testUtils.createSheetMock();
      worksheet.getColumn(1).width = 20;
      worksheet.getRow(1).height = 20;

      anchor = new Anchor(worksheet, { col: 0.6, row: 0.6 });
    });

    it("should update colWidth", () => {
      const pre = anchor.colWidth;
      worksheet.getColumn(1).width *= 2;
      expect(anchor.colWidth).not.toBe(pre);
      expect(anchor.colWidth).toBe(pre * 2);
    });
    it("should update rowHeight", () => {
      const pre = anchor.rowHeight;
      worksheet.getRow(1).height *= 2;
      expect(anchor.rowHeight).not.toBe(pre);
      expect(anchor.rowHeight).toBe(pre * 2);
    });
    it("should recalculate col", () => {
      const pre = anchor.col;
      worksheet.getColumn(1).width *= 2;
      expect(anchor.col).not.toBe(pre);
    });
    it("should recalculate row", () => {
      const pre = anchor.row;
      worksheet.getRow(1).height *= 2;
      expect(anchor.row).not.toBe(pre);
    });
    it("should integer part of row and nativeRow should always be equal", () => {
      expect(Math.floor(anchor.row)).toBe(Math.floor(anchor.nativeRow));
      worksheet.getRow(1).height *= 2;
      expect(Math.floor(anchor.row)).toBe(Math.floor(anchor.nativeRow));
      worksheet.getRow(1).height /= 4;
      expect(Math.floor(anchor.row)).toBe(Math.floor(anchor.nativeRow));
      worksheet.getRow(1).height = 0.1;
      expect(Math.floor(anchor.row)).toBe(Math.floor(anchor.nativeRow));
      worksheet.getRow(1).height = 9999;
      expect(Math.floor(anchor.row)).toBe(Math.floor(anchor.nativeRow));
    });
    it("should integer part of col and colOff should be always equals", () => {
      expect(Math.floor(anchor.col)).toBe(Math.floor(anchor.nativeCol));
      worksheet.getColumn(1).width *= 2;
      expect(Math.floor(anchor.col)).toBe(Math.floor(anchor.nativeCol));
      worksheet.getColumn(1).width /= 4;
      expect(Math.floor(anchor.col)).toBe(Math.floor(anchor.nativeCol));
      worksheet.getColumn(1).width = 0.1;
      expect(Math.floor(anchor.col)).toBe(Math.floor(anchor.nativeCol));
      worksheet.getColumn(1).width = 9999;
      expect(Math.floor(anchor.col)).toBe(Math.floor(anchor.nativeCol));
    });
    it("should update nativeColOff after col has been changed", () => {
      const pre = anchor.nativeColOff;
      anchor.col -= 0.321;
      expect(anchor.nativeColOff).not.toBe(pre);
    });
    it("should update nativeRowOff after row has been changed", () => {
      const pre = anchor.nativeRowOff;
      anchor.row -= 0.321;
      expect(anchor.nativeRowOff).not.toBe(pre);
    });
  });

  describe("integer short-circuit", () => {
    it("set col/row with integer value does not access colWidth/rowHeight", () => {
      // worksheet is undefined — colWidth/rowHeight getters would access
      // worksheet.getColumn()/getRow() and fail if they were called.
      // With the short-circuit fix, integer values bypass these getters entirely.
      const anchor = new Anchor(undefined, { col: 3, row: 5 });
      expect(anchor.nativeCol).toBe(3);
      expect(anchor.nativeColOff).toBe(0);
      expect(anchor.nativeRow).toBe(5);
      expect(anchor.nativeRowOff).toBe(0);
    });

    it("get col/row with zero offset does not access colWidth/rowHeight", () => {
      const anchor = new Anchor(undefined, {
        nativeCol: 3,
        nativeColOff: 0,
        nativeRow: 5,
        nativeRowOff: 0
      });
      // If the getters accessed colWidth/rowHeight with undefined worksheet,
      // they would attempt worksheet.getColumn()/getRow() on undefined.
      // With the short-circuit fix, zero offsets bypass these getters.
      expect(anchor.col).toBe(3);
      expect(anchor.row).toBe(5);
    });

    it("set col/row with fractional value still computes offset correctly", () => {
      // With a real worksheet mock, fractional values should produce non-zero offsets.
      const ws = testUtils.createSheetMock();
      const anchor = new Anchor(ws, { col: 1.5, row: 2.75 });
      expect(anchor.nativeCol).toBe(1);
      expect(anchor.nativeColOff).toBeGreaterThan(0);
      expect(anchor.nativeRow).toBe(2);
      expect(anchor.nativeRowOff).toBeGreaterThan(0);
    });
  });
});
