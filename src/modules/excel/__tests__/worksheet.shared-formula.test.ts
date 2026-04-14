import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Worksheet", () => {
  describe("Shared Formulae", () => {
    it("Fills formula using 2D array values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("A1:B2", "ROW()+COLUMN()", [
        [2, 3],
        [3, 4]
      ]);
      expect(ws.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(ws.getCell("B1").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("A2").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("B2").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Translates formulae to slave cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 2;
      ws.fillFormula("A2:B3", "A1", [
        [1, 2],
        [1, 2]
      ]);
      expect(ws.getCell("A2").value).toEqual({
        formula: "A1",
        shareType: "shared",
        ref: "A2:B3",
        result: 1
      });

      expect(ws.getCell("B2").value).toEqual({
        sharedFormula: "A2",
        result: 2
      });
      expect(ws.getCell("B2").formula).toBe("B1");

      expect(ws.getCell("A3").value).toEqual({
        sharedFormula: "A2",
        result: 1
      });
      expect(ws.getCell("A3").formula).toBe("A2");

      expect(ws.getCell("B3").value).toEqual({
        sharedFormula: "A2",
        result: 2
      });
      expect(ws.getCell("B3").formula).toBe("B2");
    });

    it("Fills formula down using 1D array values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("A1:A4", "ROW()+COLUMN()", [2, 3, 4, 5]);
      expect(ws.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:A4",
        result: 2
      });
      expect(ws.getCell("A2").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("A3").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
      expect(ws.getCell("A4").value).toEqual({
        sharedFormula: "A1",
        result: 5
      });
    });

    it("Fills formula across using 1D array values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("A1:D1", "ROW()+COLUMN()", [2, 3, 4, 5]);
      expect(ws.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:D1",
        result: 2
      });
      expect(ws.getCell("B1").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("C1").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
      expect(ws.getCell("D1").value).toEqual({
        sharedFormula: "A1",
        result: 5
      });
    });

    it("Fills formula down and across using 1D array values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("A1:B2", "ROW()+COLUMN()", [2, 3, 3, 4]);
      expect(ws.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(ws.getCell("B1").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("A2").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("B2").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Fills formula using function", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("A1:B2", "ROW()+COLUMN()", (r, c) => r + c);
      expect(ws.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(ws.getCell("B1").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("A2").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws.getCell("B2").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Fills formula for a single cell range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.fillFormula("C3:C3", "SUM(A3:B3)", [10]);
      expect(ws.getCell("C3").value).toEqual({
        formula: "SUM(A3:B3)",
        shareType: "shared",
        ref: "C3:C3",
        result: 10
      });
    });

    it("fillFormula with callback receives correct row and column", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      const calls: Array<[number, number]> = [];
      ws.fillFormula("B2:C3", "ROW()", (r, c) => {
        calls.push([r, c]);
        return r * 10 + c;
      });

      expect(calls).toEqual([
        [2, 2],
        [2, 3],
        [3, 2],
        [3, 3]
      ]);
      expect(ws.getCell("B2").value).toMatchObject({ result: 22 });
      expect(ws.getCell("C3").value).toMatchObject({ result: 33 });
    });
  });
});
