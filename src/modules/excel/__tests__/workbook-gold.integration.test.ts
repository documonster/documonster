import { makeTestDataPath } from "@test/utils";
import { describe, it, expect, beforeAll } from "vitest";

import { Workbook } from "../../../index";

const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");

// =============================================================================
// This spec is based around a gold standard Excel workbook 'gold-standard.xlsx'

describe("Workbook", () => {
  describe("Gold fixture", () => {
    describe("Read", () => {
      let wb;
      beforeAll(() => {
        wb = new Workbook();
        return wb.xlsx.readFile(excelTestDataPath("gold-standard.xlsx"));
      });

      it("reads Values sheet", () => {
        const ws = wb.getWorksheet("Values");

        expect(ws.getCell("B1").value).toBe("I am Text");
        expect(ws.getCell("B2").value).toBe(3.14);
        expect(ws.getCell("B3").value).toBe(5);
        expect(ws.getCell("B4").value).toEqual(new Date("2016-05-17T00:00:00.000Z"));
        expect(ws.getCell("B5").value).toEqual({
          formula: "B1",
          result: "I am Text"
        });

        expect(ws.getCell("B6").value).toEqual({
          hyperlink: "https://www.npmjs.com/package/excelts",
          text: "excelts"
        });

        expect(ws.lastColumn).toBe(ws.getColumn(2));
        expect(ws.lastRow).toBe(ws.getRow(6));
      });
    });
  });
});
