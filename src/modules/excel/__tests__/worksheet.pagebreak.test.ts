import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Worksheet", () => {
  describe("Page Breaks", () => {
    // =========================================================================
    // Row Breaks
    // =========================================================================

    it("adds multiple row breaks", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.getCell("A1").value = "A1";
      ws.getCell("B1").value = "B1";
      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";
      ws.getCell("A3").value = "A3";
      ws.getCell("B3").value = "B3";

      ws.getRow(1).addPageBreak();
      ws.getRow(2).addPageBreak();

      expect(ws.rowBreaks.length).toBe(2);
    });

    it("adds a single row break with correct structure", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "data";
      ws.getRow(1).addPageBreak();

      expect(ws.rowBreaks.length).toBe(1);
      expect(ws.rowBreaks[0]).toHaveProperty("id");
      expect(ws.rowBreaks[0]).toHaveProperty("man");
      expect(ws.rowBreaks[0].man).toBe(1);
    });

    it("rowBreaks starts as empty array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      expect(ws.rowBreaks).toEqual([]);
    });

    it("row breaks survive XLSX round-trip", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "above break";
      ws.getCell("A2").value = "below break";
      ws.getRow(1).addPageBreak();

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("test")!;
      expect(ws2.rowBreaks.length).toBe(1);
    });

    it("rowBreaks is included in worksheet model", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "data";
      ws.getRow(1).addPageBreak();
      ws.getRow(3).addPageBreak();

      const model = ws.model;
      expect(model.rowBreaks).toBeDefined();
      expect(model.rowBreaks!.length).toBe(2);
    });

    // =========================================================================
    // Column Breaks
    // =========================================================================

    it("adds a single column break", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "A1";
      ws.getCell("B1").value = "B1";
      ws.getCell("C1").value = "C1";

      const col = ws.getColumn(1);
      col.addPageBreak();

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0]).toEqual({
        id: 1,
        max: 1048575,
        man: 1
      });
    });

    it("adds multiple column breaks", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      for (let col = 1; col <= 10; col++) {
        ws.getCell(1, col).value = `Col ${col}`;
      }

      ws.getColumn(3).addPageBreak();
      ws.getColumn(6).addPageBreak();
      ws.getColumn(9).addPageBreak();

      expect(ws.colBreaks.length).toBe(3);
      expect(ws.colBreaks[0].id).toBe(3);
      expect(ws.colBreaks[1].id).toBe(6);
      expect(ws.colBreaks[2].id).toBe(9);
    });

    it("adds column break with row constraints", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "A1";

      ws.getColumn("B").addPageBreak(5, 100);

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0]).toEqual({
        id: 2,
        max: 99, // 100 - 1 (0-indexed)
        min: 4, // 5 - 1 (0-indexed)
        man: 1
      });
    });

    it("adds column break using column letter", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getColumn("D").addPageBreak();

      expect(ws.colBreaks.length).toBe(1);
      expect(ws.colBreaks[0].id).toBe(4); // D is column 4
    });

    it("initializes colBreaks as empty array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      expect(ws.colBreaks).toEqual([]);
    });

    it("colBreaks is included in worksheet model", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getColumn(2).addPageBreak();

      const model = ws.model;
      expect(model.colBreaks).toEqual([{ id: 2, max: 1048575, man: 1 }]);
    });

    // =========================================================================
    // Row + Column Coexistence
    // =========================================================================

    it("row and column breaks can coexist", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.getCell("A1").value = "data";
      ws.getRow(1).addPageBreak();
      ws.getColumn(1).addPageBreak();

      expect(ws.rowBreaks.length).toBe(1);
      expect(ws.colBreaks.length).toBe(1);
    });
  });
});
