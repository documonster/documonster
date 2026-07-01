import { testUtils } from "@excel/__tests__/shared";
import {
  anchorCol,
  anchorColWidth,
  anchorCreate,
  anchorRow,
  anchorRowHeight,
  anchorSetCol,
  anchorSetRow
} from "@excel/core/anchor";
import { getColumn } from "@excel/core/worksheet-core";
import { Worksheet } from "@excel/index";
import { describe, it, expect, beforeEach } from "vitest";

describe("Anchor", () => {
  describe("colWidth", () => {
    it("should colWidth equals 640000 when worksheet is undefined", () => {
      const anchor = anchorCreate();
      expect(anchorColWidth(anchor)).toBe(640000);
    });
    it("should colWidth equals 640000 when column has not set custom width", () => {
      const anchor = anchorCreate(testUtils.createSheetMock());
      expect(anchorColWidth(anchor)).toBe(640000);
    });
    it("should colWidth equals column width", () => {
      const worksheet = testUtils.createSheetMock();
      const anchor = anchorCreate(worksheet);
      getColumn(worksheet, anchor.nativeCol + 1).width! = 10;
      expect(anchorColWidth(anchor)).toBe(
        getColumn(worksheet, anchor.nativeCol + 1).width! * 10000
      );
    });
  });
  describe("rowHeight", () => {
    it("should rowHeight equals 180000 when worksheet is undefined", () => {
      const anchor = anchorCreate();
      expect(anchorRowHeight(anchor)).toBe(180000);
    });
    it("should rowHeight equals 180000 when row has not set height", () => {
      const anchor = anchorCreate(testUtils.createSheetMock());
      expect(anchorRowHeight(anchor)).toBe(180000);
    });
    it("should rowHeight equals row height", () => {
      const worksheet = testUtils.createSheetMock();
      Worksheet.getRow(worksheet, 1).height! = 10;

      const anchor = anchorCreate(worksheet);
      expect(anchorRowHeight(anchor)).toBe(Worksheet.getRow(worksheet, 1).height! * 10000);
    });
  });
  describe("resize worksheet`s cells", () => {
    let worksheet: any;
    let anchor: any;

    beforeEach(() => {
      worksheet = testUtils.createSheetMock();
      getColumn(worksheet, 1).width! = 20;
      Worksheet.getRow(worksheet, 1).height! = 20;

      anchor = anchorCreate(worksheet, { col: 0.6, row: 0.6 });
    });

    it("should update colWidth", () => {
      const pre = anchorColWidth(anchor);
      getColumn(worksheet, 1).width! *= 2;
      expect(anchorColWidth(anchor)).not.toBe(pre);
      expect(anchorColWidth(anchor)).toBe(pre * 2);
    });
    it("should update rowHeight", () => {
      const pre = anchorRowHeight(anchor);
      Worksheet.getRow(worksheet, 1).height! *= 2;
      expect(anchorRowHeight(anchor)).not.toBe(pre);
      expect(anchorRowHeight(anchor)).toBe(pre * 2);
    });
    it("should recalculate col", () => {
      const pre = anchorCol(anchor);
      getColumn(worksheet, 1).width! *= 2;
      expect(anchorCol(anchor)).not.toBe(pre);
    });
    it("should recalculate row", () => {
      const pre = anchorRow(anchor);
      Worksheet.getRow(worksheet, 1).height! *= 2;
      expect(anchorRow(anchor)).not.toBe(pre);
    });
    it("should integer part of row and nativeRow should always be equal", () => {
      expect(Math.floor(anchorRow(anchor))).toBe(Math.floor(anchor.nativeRow));
      Worksheet.getRow(worksheet, 1).height! *= 2;
      expect(Math.floor(anchorRow(anchor))).toBe(Math.floor(anchor.nativeRow));
      Worksheet.getRow(worksheet, 1).height! /= 4;
      expect(Math.floor(anchorRow(anchor))).toBe(Math.floor(anchor.nativeRow));
      Worksheet.getRow(worksheet, 1).height! = 0.1;
      expect(Math.floor(anchorRow(anchor))).toBe(Math.floor(anchor.nativeRow));
      Worksheet.getRow(worksheet, 1).height! = 9999;
      expect(Math.floor(anchorRow(anchor))).toBe(Math.floor(anchor.nativeRow));
    });
    it("should integer part of col and colOff should be always equals", () => {
      expect(Math.floor(anchorCol(anchor))).toBe(Math.floor(anchor.nativeCol));
      getColumn(worksheet, 1).width! *= 2;
      expect(Math.floor(anchorCol(anchor))).toBe(Math.floor(anchor.nativeCol));
      getColumn(worksheet, 1).width! /= 4;
      expect(Math.floor(anchorCol(anchor))).toBe(Math.floor(anchor.nativeCol));
      getColumn(worksheet, 1).width! = 0.1;
      expect(Math.floor(anchorCol(anchor))).toBe(Math.floor(anchor.nativeCol));
      getColumn(worksheet, 1).width! = 9999;
      expect(Math.floor(anchorCol(anchor))).toBe(Math.floor(anchor.nativeCol));
    });
    it("should update nativeColOff after col has been changed", () => {
      const pre = anchor.nativeColOff;
      anchorSetCol(anchor, anchorCol(anchor) - 0.321);
      expect(anchor.nativeColOff).not.toBe(pre);
    });
    it("should update nativeRowOff after row has been changed", () => {
      const pre = anchor.nativeRowOff;
      anchorSetRow(anchor, anchorRow(anchor) - 0.321);
      expect(anchor.nativeRowOff).not.toBe(pre);
    });
  });

  describe("integer short-circuit", () => {
    it("set col/row with integer value does not access colWidth/rowHeight", () => {
      // worksheet is undefined — colWidth/rowHeight would access
      // getColumn(worksheet, )/getRow() and fail if they were called.
      // With the short-circuit fix, integer values bypass these entirely.
      const anchor = anchorCreate(undefined, { col: 3, row: 5 });
      expect(anchor.nativeCol).toBe(3);
      expect(anchor.nativeColOff).toBe(0);
      expect(anchor.nativeRow).toBe(5);
      expect(anchor.nativeRowOff).toBe(0);
    });

    it("get col/row with zero offset does not access colWidth/rowHeight", () => {
      const anchor = anchorCreate(undefined, {
        nativeCol: 3,
        nativeColOff: 0,
        nativeRow: 5,
        nativeRowOff: 0
      });
      // If col/row accessed colWidth/rowHeight with undefined worksheet,
      // they would attempt getColumn(worksheet, )/getRow() on undefined.
      // With the short-circuit fix, zero offsets bypass these.
      expect(anchorCol(anchor)).toBe(3);
      expect(anchorRow(anchor)).toBe(5);
    });

    it("set col/row with fractional value still computes offset correctly", () => {
      // With a real worksheet mock, fractional values should produce non-zero offsets.
      const ws = testUtils.createSheetMock();
      const anchor = anchorCreate(ws, { col: 1.5, row: 2.75 });
      expect(anchor.nativeCol).toBe(1);
      expect(anchor.nativeColOff).toBeGreaterThan(0);
      expect(anchor.nativeRow).toBe(2);
      expect(anchor.nativeRowOff).toBeGreaterThan(0);
    });
  });
});
