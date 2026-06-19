import { cellGetValue } from "@excel/core/cell";
import { getColumn, getCell } from "@excel/core/worksheet";
import { Workbook, Worksheet } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, it, expect, beforeAll } from "vitest";

const excelTestDataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");

// =============================================================================
// This spec is based around a gold standard Excel workbook 'gold-standard.xlsx'

describe("Workbook", () => {
  describe("Gold fixture", () => {
    describe("Read", () => {
      let wb;
      beforeAll(() => {
        wb = Workbook.create();
        return Workbook.readFile(wb, excelTestDataPath("gold-standard.xlsx"));
      });

      it("reads Values sheet", () => {
        const ws = Workbook.getWorksheet(wb, "Values")!;

        expect(cellGetValue(getCell(ws, "B1"))).toBe("I am Text");
        expect(cellGetValue(getCell(ws, "B2"))).toBe(3.14);
        expect(cellGetValue(getCell(ws, "B3"))).toBe(5);
        expect(cellGetValue(getCell(ws, "B4"))).toEqual(new Date("2016-05-17T00:00:00.000Z"));
        expect(cellGetValue(getCell(ws, "B5"))).toEqual({
          formula: "B1",
          result: "I am Text"
        });

        expect(cellGetValue(getCell(ws, "B6"))).toEqual({
          hyperlink: "https://www.npmjs.com/package/documonster",
          text: "documonster"
        });

        expect(Worksheet.lastColumn(ws)).toBe(getColumn(ws, 2));
        expect(Worksheet.lastRow(ws)).toBe(Worksheet.getRow(ws, 6));
      });
    });
  });
});
