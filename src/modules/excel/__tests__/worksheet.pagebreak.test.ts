import { columnAddPageBreak } from "@excel/column";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { rowAddPageBreak } from "@excel/row";
import { getColumn, getSheetModel } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Page Breaks", () => {
    // =========================================================================
    // Row Breaks
    // =========================================================================

    it("adds multiple row breaks", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Cell.setValue(ws, "A1", "A1");
      Cell.setValue(ws, "B1", "B1");
      Cell.setValue(ws, "A2", "A2");
      Cell.setValue(ws, "B2", "B2");
      Cell.setValue(ws, "A3", "A3");
      Cell.setValue(ws, "B3", "B3");

      rowAddPageBreak(Worksheet.getRow(ws, 1));
      rowAddPageBreak(Worksheet.getRow(ws, 2));

      expect(ws.rowBreaks.length).toBe(2);
    });

    it("adds a single row break with correct structure", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "data");
      rowAddPageBreak(Worksheet.getRow(ws, 1));

      expect(ws.rowBreaks.length).toBe(1);
      expect(ws.rowBreaks[0]).toHaveProperty("id");
      expect(ws.rowBreaks[0]).toHaveProperty("man");
      expect(ws.rowBreaks[0].man).toBe(1);
    });

    it("rowBreaks starts as empty array", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      expect(ws.rowBreaks).toEqual([]);
    });

    it("row breaks survive XLSX round-trip", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "above break");
      Cell.setValue(ws, "A2", "below break");
      rowAddPageBreak(Worksheet.getRow(ws, 1));

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "test")!;
      expect(ws2.rowBreaks.length).toBe(1);
    });

    it("rowBreaks is included in worksheet model", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "data");
      rowAddPageBreak(Worksheet.getRow(ws, 1));
      rowAddPageBreak(Worksheet.getRow(ws, 3));

      const model = getSheetModel(ws);
      expect(model.rowBreaks).toBeDefined();
      expect(model.rowBreaks!.length).toBe(2);
    });

    // =========================================================================
    // Column Breaks
    // =========================================================================

    it("adds a single column break", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "A1");
      Cell.setValue(ws, "B1", "B1");
      Cell.setValue(ws, "C1", "C1");

      const col = getColumn(ws, 1);
      columnAddPageBreak(col);

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0]).toEqual({
        id: 1,
        max: 1048575,
        man: 1
      });
    });

    it("adds multiple column breaks", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      for (let col = 1; col <= 10; col++) {
        Cell.setValue(ws, 1, col, `Col ${col}`);
      }

      columnAddPageBreak(getColumn(ws, 3));
      columnAddPageBreak(getColumn(ws, 6));
      columnAddPageBreak(getColumn(ws, 9));

      expect(ws.colBreaks.length).toBe(3);
      expect(ws.colBreaks[0].id).toBe(3);
      expect(ws.colBreaks[1].id).toBe(6);
      expect(ws.colBreaks[2].id).toBe(9);
    });

    it("adds column break with row constraints", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "A1");

      columnAddPageBreak(getColumn(ws, "B"), 5, 100);

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0]).toEqual({
        id: 2,
        max: 99, // 100 - 1 (0-indexed)
        min: 4, // 5 - 1 (0-indexed)
        man: 1
      });
    });

    it("adds column break using column letter", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      columnAddPageBreak(getColumn(ws, "D"));

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0].id).toBe(4); // D is column 4
    });

    it("initializes colBreaks as empty array", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      expect(ws.colBreaks).toEqual([]);
    });

    it("colBreaks is included in worksheet model", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      columnAddPageBreak(getColumn(ws, 2));

      const model = getSheetModel(ws);
      expect(model.colBreaks).toEqual([{ id: 2, max: 1048575, man: 1 }]);
    });

    // =========================================================================
    // Row + Column Coexistence
    // =========================================================================

    it("row and column breaks can coexist", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      Cell.setValue(ws, "A1", "data");
      rowAddPageBreak(Worksheet.getRow(ws, 1));
      columnAddPageBreak(getColumn(ws, 1));

      expect(ws.rowBreaks.length).toBe(1);
      expect(ws.colBreaks.length).toBe(1);
    });
  });
});
