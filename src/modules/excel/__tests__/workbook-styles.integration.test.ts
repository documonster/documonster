import { columnAlignment, columnFont } from "@excel/column";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { rowFont, rowSetFont } from "@excel/row";
import { testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

const TEST_XLSX_FILE_NAME = testFilePath("workbook-styles.test");

// =============================================================================
// Sample Data
import { richTextSample } from "@excel/__tests__/data/rich-text-sample";
import richTextSampleA1 from "@excel/__tests__/data/rich-text-sample-a1.json" with { type: "json" };
import { testUtils } from "@excel/__tests__/shared";
import { getWorksheets } from "@excel/workbook";
import { getColumn } from "@excel/worksheet";
import { PassThrough } from "@stream";

// =============================================================================
// Tests

describe("Workbook", () => {
  describe("Styles", () => {
    it("row styles and columns properly", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Worksheet.setColumns(ws, [
        { header: "A1", width: 10 },
        {
          header: "B1",
          width: 20,
          style: {
            font: testUtils.styles.fonts.comicSansUdB16,
            alignment: testUtils.styles.alignments[1].alignment
          }
        },
        { header: "C1", width: 30 }
      ]);

      rowSetFont(Worksheet.getRow(ws, 2), testUtils.styles.fonts.broadwayRedOutline20);

      Cell.setValue(ws, "A2", "A2");
      Cell.setValue(ws, "B2", "B2");
      Cell.setValue(ws, "C2", "C2");
      Cell.setValue(ws, "A3", "A3");
      Cell.setValue(ws, "B3", "B3");
      Cell.setValue(ws, "C3", "C3");

      await Workbook.writeFile(wb, TEST_XLSX_FILE_NAME);
      await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, TEST_XLSX_FILE_NAME);

      const ws2 = Workbook.getWorksheet(wb2, "blort")!;
      ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"].forEach(address => {
        expect(Cell.getValue(ws2, address)).toBe(address);
      });
      expect(Cell.getStyle(ws2, "B1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws2, "B1").alignment).toEqual(testUtils.styles.alignments[1].alignment);
      expect(Cell.getStyle(ws2, "A2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(Cell.getStyle(ws2, "B2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(Cell.getStyle(ws2, "C2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(Cell.getStyle(ws2, "B3").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(Cell.getStyle(ws2, "B3").alignment).toEqual(testUtils.styles.alignments[1].alignment);

      expect(columnFont(getColumn(ws2, 2))).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(columnAlignment(getColumn(ws2, 2))).toEqual(testUtils.styles.alignments[1].alignment);

      expect(rowFont(Worksheet.getRow(ws2, 2))).toEqual(
        testUtils.styles.fonts.broadwayRedOutline20
      );
    });

    it("in-cell formats properly in xlsx file", async () => {
      const testData = Buffer.from(richTextSample, "base64");
      const bufferStream = new PassThrough();
      bufferStream.write(testData);
      bufferStream.end();

      const wb = Workbook.create();
      await Workbook.readStream(wb, bufferStream);

      const ws = getWorksheets(wb)[0];
      expect(Cell.getValue(ws, "A1")).toEqual(richTextSampleA1);
      expect(Cell.getText(ws, "A1")).toBe(Cell.getValue(ws, "A2"));
    });

    it("null cells retain style", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "blort");

      Cell.setValue(ws, "B2", "hello");
      Cell.setStyle(ws, "B4", { fill: testUtils.styles.fills.redDarkVertical });
      Cell.setStyle(ws, "B4", { font: testUtils.styles.fonts.broadwayRedOutline20 });

      await Workbook.writeFile(wb, TEST_XLSX_FILE_NAME);
      await expectValidXlsx(new Uint8Array(await Workbook.toBuffer(wb)));
      const wb2 = Workbook.create();
      await Workbook.readFile(wb2, TEST_XLSX_FILE_NAME);

      const ws2 = Workbook.getWorksheet(wb2, "blort")!;
      expect(Cell.getStyle(ws2, "B4").fill).toEqual(testUtils.styles.fills.redDarkVertical);
      expect(Cell.getStyle(ws2, "B4").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
    });
  });
});
