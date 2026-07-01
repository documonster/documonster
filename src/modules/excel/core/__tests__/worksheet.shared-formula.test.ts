import { Cell, Workbook, Worksheet } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Shared Formulae", () => {
    it("Fills formula using 2D array values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "A1:B2", "ROW()+COLUMN()", [
        [2, 3],
        [3, 4]
      ]);
      expect(Cell.getValue(ws, "A1")).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(Cell.getValue(ws, "B1")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "A2")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "B2")).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Translates formulae to slave cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "B1", 2);
      Worksheet.fillFormula(ws, "A2:B3", "A1", [
        [1, 2],
        [1, 2]
      ]);
      expect(Cell.getValue(ws, "A2")).toEqual({
        formula: "A1",
        shareType: "shared",
        ref: "A2:B3",
        result: 1
      });

      expect(Cell.getValue(ws, "B2")).toEqual({
        sharedFormula: "A2",
        result: 2
      });
      expect(Cell.getFormula(ws, "B2")).toBe("B1");

      expect(Cell.getValue(ws, "A3")).toEqual({
        sharedFormula: "A2",
        result: 1
      });
      expect(Cell.getFormula(ws, "A3")).toBe("A2");

      expect(Cell.getValue(ws, "B3")).toEqual({
        sharedFormula: "A2",
        result: 2
      });
      expect(Cell.getFormula(ws, "B3")).toBe("B2");
    });

    it("Fills formula down using 1D array values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "A1:A4", "ROW()+COLUMN()", [2, 3, 4, 5]);
      expect(Cell.getValue(ws, "A1")).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:A4",
        result: 2
      });
      expect(Cell.getValue(ws, "A2")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "A3")).toEqual({
        sharedFormula: "A1",
        result: 4
      });
      expect(Cell.getValue(ws, "A4")).toEqual({
        sharedFormula: "A1",
        result: 5
      });
    });

    it("Fills formula across using 1D array values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "A1:D1", "ROW()+COLUMN()", [2, 3, 4, 5]);
      expect(Cell.getValue(ws, "A1")).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:D1",
        result: 2
      });
      expect(Cell.getValue(ws, "B1")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "C1")).toEqual({
        sharedFormula: "A1",
        result: 4
      });
      expect(Cell.getValue(ws, "D1")).toEqual({
        sharedFormula: "A1",
        result: 5
      });
    });

    it("Fills formula down and across using 1D array values", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "A1:B2", "ROW()+COLUMN()", [2, 3, 3, 4]);
      expect(Cell.getValue(ws, "A1")).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(Cell.getValue(ws, "B1")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "A2")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "B2")).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Fills formula using function", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "A1:B2", "ROW()+COLUMN()", (r, c) => r + c);
      expect(Cell.getValue(ws, "A1")).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(Cell.getValue(ws, "B1")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "A2")).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(Cell.getValue(ws, "B2")).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("Fills formula for a single cell range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      Worksheet.fillFormula(ws, "C3:C3", "SUM(A3:B3)", [10]);
      expect(Cell.getValue(ws, "C3")).toEqual({
        formula: "SUM(A3:B3)",
        shareType: "shared",
        ref: "C3:C3",
        result: 10
      });
    });

    it("fillFormula with callback receives correct row and column", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb);

      const calls: Array<[number, number]> = [];
      Worksheet.fillFormula(ws, "B2:C3", "ROW()", (r, c) => {
        calls.push([r, c]);
        return r * 10 + c;
      });

      expect(calls).toEqual([
        [2, 2],
        [2, 3],
        [3, 2],
        [3, 3]
      ]);
      expect(Cell.getValue(ws, "B2")).toMatchObject({ result: 22 });
      expect(Cell.getValue(ws, "C3")).toMatchObject({ result: 33 });
    });
  });
});
