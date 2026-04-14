import fs from "fs";

import { testUtils } from "@excel/__tests__/shared";
import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect, beforeAll, beforeEach, afterEach } from "vitest";

import { ValueType, Workbook, WorkbookReader, WorkbookWriter } from "../../../../index";

const streamTestDataPath = makeTestDataPath(import.meta.url, "./data");

const TEST_FILE_NAME = testFilePath("wb-xlsx-reader.test");

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe("WorkbookReader", () => {
  describe("Serialise", () => {
    it("xlsx file", async () => {
      const wb = testUtils.createTestBook(new Workbook(), "xlsx");

      await wb.xlsx.writeFile(TEST_FILE_NAME);
      await testUtils.checkTestBookReader(TEST_FILE_NAME);
    }, 30000);
  });

  describe("#readFile", () => {
    describe("Row limit", () => {
      it("should bail out if the file contains more rows than the limit", async () => {
        const workbook = new Workbook();
        // The Fibonacci sheet has 19 rows
        await expect(
          workbook.xlsx.readFile(streamTestDataPath("fibonacci.xlsx"), { maxRows: 10 })
        ).rejects.toThrow("Max row count (10) exceeded");
      });

      it("should fail fast on a huge file", async () => {
        const workbook = new Workbook();
        await expect(
          workbook.xlsx.readFile(streamTestDataPath("huge.xlsx"), { maxRows: 100 })
        ).rejects.toThrow("Max row count (100) exceeded");
      }, 30000);

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = new Workbook();
        await workbook.xlsx.readFile(streamTestDataPath("fibonacci.xlsx"), { maxRows: 20 });
        expect(workbook.worksheets.length).toBeGreaterThan(0);
      });
    });

    describe("Column limit", () => {
      it("should bail out if the file contains more cells than the limit", async () => {
        const workbook = new Workbook();
        // The many-columns sheet has 20 columns in row 2
        await expect(
          workbook.xlsx.readFile(streamTestDataPath("many-columns.xlsx"), { maxCols: 15 })
        ).rejects.toThrow("Max column count (15) exceeded");
      });

      it("should fail fast on a huge file", async () => {
        const workbook = new Workbook();
        await expect(
          workbook.xlsx.readFile(streamTestDataPath("huge.xlsx"), { maxCols: 10 })
        ).rejects.toThrow("Max column count (10) exceeded");
      }, 30000);

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = new Workbook();
        await workbook.xlsx.readFile(streamTestDataPath("many-columns.xlsx"), { maxCols: 40 });
        expect(workbook.worksheets.length).toBeGreaterThan(0);
      });
    });
  });

  describe("#read", () => {
    describe("Row limit", () => {
      it("should bail out if the file contains more rows than the limit", async () => {
        const workbook = new Workbook();
        // The Fibonacci sheet has 19 rows
        await expect(
          workbook.xlsx.read(fs.createReadStream(streamTestDataPath("fibonacci.xlsx")), {
            maxRows: 10
          })
        ).rejects.toThrow("Max row count (10) exceeded");
      });

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = new Workbook();
        await workbook.xlsx.read(fs.createReadStream(streamTestDataPath("fibonacci.xlsx")), {
          maxRows: 20
        });
        expect(workbook.worksheets.length).toBeGreaterThan(0);
      });
    });
  });

  describe("edit styles in existing file", () => {
    let wb;

    beforeEach(async () => {
      wb = new Workbook();
      await wb.xlsx.readFile(streamTestDataPath("test-row-styles.xlsx"));
    });

    it("edit styles of single row instead of all", () => {
      const ws = wb.getWorksheet(1);

      ws.eachRow((row, rowNo) => {
        if (rowNo % 5 === 0) {
          row.font = { color: { argb: "00ff00" } };
        }
      });

      expect(ws.getRow(3).font.color.argb).toBe(ws.getRow(6).font.color.argb);
      expect(ws.getRow(6).font.color.argb).toBe(ws.getRow(9).font.color.argb);
      expect(ws.getRow(9).font.color.argb).toBe(ws.getRow(12).font.color.argb);
      expect(ws.getRow(12).font.color.argb).not.toBe(ws.getRow(15).font.color.argb);
      expect(ws.getRow(15).font.color.argb).not.toBe(ws.getRow(18).font.color.argb);
      expect(ws.getRow(15).font.color.argb).toBe(ws.getRow(10).font.color.argb);
      expect(ws.getRow(10).font.color.argb).toBe(ws.getRow(5).font.color.argb);
    });
  });

  describe("with a spreadsheet that contains formulas", () => {
    let worksheet;
    let cell;

    beforeAll(async () => {
      const workbook = new Workbook();
      await workbook.xlsx.read(fs.createReadStream(streamTestDataPath("formulas.xlsx")));
      worksheet = workbook.getWorksheet();
    });

    describe("with a cell that contains a regular formula", () => {
      beforeEach(() => {
        cell = worksheet.getCell("A2");
      });

      it("should be classified as a formula cell", () => {
        expect(cell.type).toBe(ValueType.Formula);
      });

      it("should have text corresponding to the evaluated formula result", () => {
        expect(cell.text).toBe("someone@example.com");
      });

      it("should have the formula source", () => {
        expect(cell.model.formula).toBe('_xlfn.CONCAT("someone","@example.com")');
      });
    });

    describe("with a cell that contains a hyperlinked formula", () => {
      beforeEach(() => {
        cell = worksheet.getCell("A1");
      });

      it("should be classified as a formula cell", () => {
        expect(cell.type).toBe(ValueType.Hyperlink);
      });

      it("should have text corresponding to the evaluated formula result", () => {
        expect(cell.value.text).toBe("someone@example.com");
      });

      it("should have the formula source", () => {
        expect(cell.model.formula).toBe('_xlfn.CONCAT("someone","@example.com")');
      });

      it("should contain the linked url", () => {
        expect(cell.value.hyperlink).toBe("mailto:someone@example.com");
        expect(cell.hyperlink).toBe("mailto:someone@example.com");
      });
    });
  });

  describe("with a spreadsheet that contains a shared string with an escaped underscore", () => {
    let worksheet;

    beforeAll(async () => {
      const workbook = new Workbook();
      await workbook.xlsx.read(
        fs.createReadStream(streamTestDataPath("shared_string_with_escape.xlsx"))
      );
      worksheet = workbook.getWorksheet();
    });

    it("should decode the underscore", () => {
      const cell = worksheet.getCell("A1");
      expect(cell.value).toBe("_x000D_");
    });
  });

  describe("streaming reader should decode OOXML escapes in cached shared strings", () => {
    it("should decode _x005F_x000D_ to literal _x000D_", async () => {
      const workbookReader = new WorkbookReader(
        streamTestDataPath("shared_string_with_escape.xlsx"),
        {
          entries: "emit",
          sharedStrings: "cache",
          styles: "cache",
          worksheets: "emit"
        }
      );

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet =>
          worksheet.on("row", row => {
            expect(row.values[1]).toBe("_x000D_");
            resolve();
          })
        );
        workbookReader.on("error", reject);
        workbookReader.read();
      });
    });

    it("should roundtrip literal _xHHHH_ patterns through streaming write + read", async () => {
      const testFile = testFilePath("ooxml-escape-streaming.test");

      const workbook = new WorkbookWriter({
        filename: testFile,
        useSharedStrings: true
      });

      const sheet = workbook.addWorksheet("data");
      sheet.addRow(["_x000D_", "Normal", "_x000a_test"]).commit();
      sheet.commit();
      await workbook.commit();

      const workbookReader = new WorkbookReader(testFile, {
        entries: "emit",
        sharedStrings: "cache",
        styles: "cache",
        worksheets: "emit"
      });

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet =>
          worksheet.on("row", row => {
            expect(row.values[1]).toBe("_x000D_");
            expect(row.values[2]).toBe("Normal");
            expect(row.values[3]).toBe("_x000a_test");
            resolve();
          })
        );
        workbookReader.on("error", reject);
        workbookReader.read();
      });
    });
  });

  describe("with a spreadsheet that contains shared formulas", () => {
    it("should read shared formula models from a file", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(streamTestDataPath("fibonacci.xlsx"));

      const ws = wb.getWorksheet("fib");

      expect(ws!.getCell("A4").value).toEqual({
        formula: "A3+1",
        shareType: "shared",
        ref: "A4:A19",
        result: 4
      });
      expect(ws!.getCell("A5").value).toEqual({ sharedFormula: "A4", result: 5 });

      expect(ws!.getCell("A4").type).toBe(ValueType.Formula);
      expect(ws!.getCell("A5").type).toBe(ValueType.Formula);
    });
  });

  describe("hyperlinks cache", () => {
    it("should resolve worksheet hyperlink rIds to targets", async () => {
      const worksheets: any[] = [];

      const workbookReader = new WorkbookReader(
        fs.createReadStream(streamTestDataPath("hyperlinks-cache.xlsx")),
        {
          worksheets: "emit",
          sharedStrings: "cache",
          styles: "ignore",
          hyperlinks: "cache",
          entries: "ignore"
        }
      );

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet => {
          worksheets.push(worksheet);
        });
        workbookReader.on("end", resolve);
        workbookReader.on("error", reject);
        workbookReader.read();
      });

      const targets: string[] = [];
      for (const worksheet of worksheets) {
        const sheetNo = worksheet.sheetNo ?? worksheet.id;
        const hyperlinks = worksheet.hyperlinks as undefined | Record<string, { rId: string }>;
        if (!hyperlinks) {
          continue;
        }

        for (const ref in hyperlinks) {
          const rId = hyperlinks[ref].rId;
          const target = (workbookReader as any).getHyperlinkTarget?.(sheetNo, rId) as
            | string
            | undefined;
          if (target) {
            targets.push(target);
          }
        }
      }

      expect(targets.length).toBeGreaterThan(0);
      expect(targets[0]).toBeTypeOf("string");
    });
  });

  describe("cached styles", () => {
    it("should emit Date objects when styles are cached", async () => {
      const rows: any[] = [];

      const workbookReader = new WorkbookReader(
        fs.createReadStream(streamTestDataPath("date-styles.xlsx")),
        {
          worksheets: "emit",
          styles: "cache",
          sharedStrings: "cache",
          hyperlinks: "ignore",
          entries: "ignore"
        }
      );

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet => {
          worksheet.on("row", row => rows.push(row.values[1]));
        });
        workbookReader.on("end", resolve);
        workbookReader.on("error", reject);
        workbookReader.read();
      });

      expect(rows).toEqual(["Date", new Date("2020-11-20T00:00:00.000Z")]);
    });
  });

  describe("worksheet names", () => {
    it("should preserve worksheet name on streaming XLSX reader", async () => {
      const names: string[] = [];
      const workbookReader = new WorkbookReader(
        streamTestDataPath("worksheet-name-preserved.xlsx"),
        {}
      );

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet => {
          names.push(worksheet.name);
        });
        workbookReader.on("end", resolve);
        workbookReader.on("error", reject);
        workbookReader.read();
      });

      expect(names).toContain("Sum Worksheet");
    });
  });

  describe("rich text within shared strings", () => {
    it("streaming reader should handle rich text within shared strings", async () => {
      const testFile = testFilePath("pr-1431.stream-reader.test");

      const rowData = [
        {
          richText: [
            { font: { bold: true }, text: "This should " },
            { font: { italic: true }, text: "be one shared string value" }
          ]
        },
        "this should be the second shared string"
      ];

      const workbook = new WorkbookWriter({
        filename: testFile,
        useSharedStrings: true
      });

      const sheet = workbook.addWorksheet("data");
      sheet.addRow(rowData).commit();
      sheet.commit();
      await workbook.commit();

      const workbookReader = new WorkbookReader(testFile, {
        entries: "emit",
        hyperlinks: "cache",
        sharedStrings: "cache",
        styles: "cache",
        worksheets: "emit"
      });

      await new Promise<void>((resolve, reject) => {
        workbookReader.on("worksheet", worksheet =>
          worksheet.on("row", row => {
            expect(row.values[1]).toEqual(rowData[0]);
            expect(row.values[2]).toBe(rowData[1]);
            resolve();
          })
        );
        workbookReader.on("error", reject);
        workbookReader.read();
      });
    });
  });

  describe("with a spreadsheet that has an XML parse error in a worksheet", () => {
    let unhandledRejection;
    function unhandledRejectionHandler(err) {
      unhandledRejection = err;
    }
    beforeEach(() => {
      process.on("unhandledRejection", unhandledRejectionHandler);
    });
    afterEach(() => {
      process.removeListener("unhandledRejection", unhandledRejectionHandler);
    });

    it("should reject the promise with the XML parse error", async () => {
      const workbook = new Workbook();
      await expect(workbook.xlsx.readFile(streamTestDataPath("invalid-xml.xlsx"))).rejects.toThrow(
        "3:1: text data outside of root node."
      );

      // Wait a tick to verify no unhandled rejection fires
      await new Promise(setImmediate);
      expect(unhandledRejection).toBeUndefined();
    });
  });

  describe("with a spreadsheet that is missing some files in the zip container", () => {
    it("should not break", async () => {
      const workbook = new Workbook();
      await workbook.xlsx.readFile(streamTestDataPath("missing-bits.xlsx"));
      // Should load gracefully despite missing zip entries
      expect(workbook.worksheets.length).toBeGreaterThanOrEqual(0);
    });
  });

  describe("with a spreadsheet that contains images", () => {
    let worksheet;

    beforeAll(async () => {
      const workbook = new Workbook();
      await workbook.xlsx.read(fs.createReadStream(streamTestDataPath("images.xlsx")));
      worksheet = workbook.getWorksheet();
    });

    describe("with image`s tl anchor", () => {
      it("should have at least one image in the test file", () => {
        expect(worksheet.getImages().length).toBeGreaterThan(0);
      });

      it("Should integer part of col equals nativeCol", () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.col)).toBe(image.range.tl.nativeCol);
        });
      });
      it("Should integer part of row equals nativeRow", () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.row)).toBe(image.range.tl.nativeRow);
        });
      });
      it("Should anchor width equals to column width when custom", () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.tl.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.tl.colWidth).toBe(Math.floor(col.width * 10000));
          } else {
            expect(image.range.tl.colWidth).toBe(640000);
          }
        });
      });
      it("Should anchor height equals to row height", () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.tl.nativeRow + 1);

          if (row.height != null) {
            expect(image.range.tl.rowHeight).toBe(Math.floor(row.height * 10000));
          } else {
            expect(image.range.tl.rowHeight).toBe(180000);
          }
        });
      });
    });

    describe("with image`s br anchor", () => {
      it("should have at least one image in the test file", () => {
        expect(worksheet.getImages().length).toBeGreaterThan(0);
      });

      it("Should integer part of col equals nativeCol", () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.col)).toBe(image.range.br.nativeCol);
        });
      });
      it("Should integer part of row equals nativeRow", () => {
        worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.row)).toBe(image.range.br.nativeRow);
        });
      });
      it("Should anchor width equals to column width when custom", () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.br.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.br.colWidth).toBe(Math.floor(col.width * 10000));
          } else {
            expect(image.range.br.colWidth).toBe(640000);
          }
        });
      });
      it("Should anchor height equals to row height", () => {
        const ws = worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.br.nativeRow + 1);

          if (row.height != null) {
            expect(image.range.br.rowHeight).toBe(Math.floor(row.height * 10000));
          } else {
            expect(image.range.br.rowHeight).toBe(180000);
          }
        });
      });
    });
  });
  describe("with a spreadsheet containing a defined name that kinda looks like it contains a range", () => {
    it("should not crash", async () => {
      const workbook = new Workbook();
      await workbook.xlsx.read(fs.createReadStream(streamTestDataPath("bogus-defined-name.xlsx")));
      expect(workbook.worksheets.length).toBeGreaterThanOrEqual(0);
    });
  });
});
