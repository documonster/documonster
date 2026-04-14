import { testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

const TEST_XLSX_FILE_NAME = testFilePath("workbook-styles.test");

// =============================================================================
// Sample Data
import { richTextSample } from "@excel/__tests__/data/rich-text-sample";
import richTextSampleA1 from "@excel/__tests__/data/rich-text-sample-a1.json" with { type: "json" };
import { testUtils } from "@excel/__tests__/shared";
import { PassThrough } from "@stream";

// =============================================================================
// Tests

describe("Workbook", () => {
  describe("Styles", () => {
    it("row styles and columns properly", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.columns = [
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
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";
      ws.getCell("C2").value = "C2";
      ws.getCell("A3").value = "A3";
      ws.getCell("B3").value = "B3";
      ws.getCell("C3").value = "C3";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort")!;
      ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"].forEach(address => {
        expect(ws2.getCell(address).value).toBe(address);
      });
      expect(ws2.getCell("B1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getCell("B1").alignment).toEqual(testUtils.styles.alignments[1].alignment);
      expect(ws2.getCell("A2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("B2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("C2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("B3").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getCell("B3").alignment).toEqual(testUtils.styles.alignments[1].alignment);

      expect(ws2.getColumn(2).font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getColumn(2).alignment).toEqual(testUtils.styles.alignments[1].alignment);

      expect(ws2.getRow(2).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
    });

    it("in-cell formats properly in xlsx file", async () => {
      const testData = Buffer.from(richTextSample, "base64");
      const bufferStream = new PassThrough();
      bufferStream.write(testData);
      bufferStream.end();

      const wb = new Workbook();
      await wb.xlsx.read(bufferStream);

      const ws = wb.worksheets[0];
      expect(ws.getCell("A1").value).toEqual(richTextSampleA1);
      expect(ws.getCell("A1").text).toBe(ws.getCell("A2").value);
    });

    it("null cells retain style", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.getCell("B2").value = "hello";
      ws.getCell("B4").fill = testUtils.styles.fills.redDarkVertical;
      ws.getCell("B4").font = testUtils.styles.fonts.broadwayRedOutline20;

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort")!;
      expect(ws2.getCell("B4").fill).toEqual(testUtils.styles.fills.redDarkVertical);
      expect(ws2.getCell("B4").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
    });
  });
});
