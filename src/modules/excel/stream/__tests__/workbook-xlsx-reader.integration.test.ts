import fs from "fs";

import { testUtils } from "@excel/__tests__/shared";
import { anchorCol, anchorRow, anchorColWidth, anchorRowHeight } from "@excel/anchor";
import { cellGetValue, cellType, cellText, cellGetModel, cellHyperlink } from "@excel/cell";
import { columnIsCustomWidth } from "@excel/column";
import { Cell, Image, Workbook, Worksheet } from "@excel/index";
import { rowFont, rowSetFont, rowValues } from "@excel/row";
import type { CellHyperlinkValue } from "@excel/types";
import { getWorksheets } from "@excel/workbook";
import { getCell, rowCommit, getColumn } from "@excel/worksheet";
import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect, beforeAll, beforeEach, afterEach } from "vitest";

import { ValueType, WorkbookReader, WorkbookWriter } from "../../../../index";

const streamTestDataPath = makeTestDataPath(import.meta.url, "./data");

const TEST_FILE_NAME = testFilePath("wb-xlsx-reader.test");

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe("WorkbookReader", () => {
  describe("Serialise", () => {
    it("xlsx file", async () => {
      const wb = testUtils.createTestBook(Workbook.create(), "xlsx");

      await Workbook.writeXlsx(wb, TEST_FILE_NAME);
      await testUtils.checkTestBookReader(TEST_FILE_NAME);
    }, 30000);
  });

  describe("#readFile", () => {
    describe("Row limit", () => {
      it("should bail out if the file contains more rows than the limit", async () => {
        const workbook = Workbook.create();
        // The Fibonacci sheet has 19 rows
        await expect(
          Workbook.readXlsxFile(workbook, streamTestDataPath("fibonacci.xlsx"), { maxRows: 10 })
        ).rejects.toThrow("Max row count (10) exceeded");
      });

      it("should fail fast on a huge file", async () => {
        const workbook = Workbook.create();
        await expect(
          Workbook.readXlsxFile(workbook, streamTestDataPath("huge.xlsx"), { maxRows: 100 })
        ).rejects.toThrow("Max row count (100) exceeded");
      }, 30000);

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = Workbook.create();
        await Workbook.readXlsxFile(workbook, streamTestDataPath("fibonacci.xlsx"), {
          maxRows: 20
        });
        expect(getWorksheets(workbook).length).toBeGreaterThan(0);
      });
    });

    describe("Column limit", () => {
      it("should bail out if the file contains more cells than the limit", async () => {
        const workbook = Workbook.create();
        // The many-columns sheet has 20 columns in row 2
        await expect(
          Workbook.readXlsxFile(workbook, streamTestDataPath("many-columns.xlsx"), { maxCols: 15 })
        ).rejects.toThrow("Max column count (15) exceeded");
      });

      it("should fail fast on a huge file", async () => {
        const workbook = Workbook.create();
        await expect(
          Workbook.readXlsxFile(workbook, streamTestDataPath("huge.xlsx"), { maxCols: 10 })
        ).rejects.toThrow("Max column count (10) exceeded");
      }, 30000);

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = Workbook.create();
        await Workbook.readXlsxFile(workbook, streamTestDataPath("many-columns.xlsx"), {
          maxCols: 40
        });
        expect(getWorksheets(workbook).length).toBeGreaterThan(0);
      });
    });
  });

  describe("#read", () => {
    describe("Row limit", () => {
      it("should bail out if the file contains more rows than the limit", async () => {
        const workbook = Workbook.create();
        // The Fibonacci sheet has 19 rows
        await expect(
          Workbook.readXlsxStream(
            workbook,
            fs.createReadStream(streamTestDataPath("fibonacci.xlsx")),
            {
              maxRows: 10
            }
          )
        ).rejects.toThrow("Max row count (10) exceeded");
      });

      it("should parse fine if the limit is not exceeded", async () => {
        const workbook = Workbook.create();
        await Workbook.readXlsxStream(
          workbook,
          fs.createReadStream(streamTestDataPath("fibonacci.xlsx")),
          {
            maxRows: 20
          }
        );
        expect(getWorksheets(workbook).length).toBeGreaterThan(0);
      });
    });
  });

  describe("edit styles in existing file", () => {
    let wb;

    beforeEach(async () => {
      wb = Workbook.create();
      await Workbook.readXlsxFile(wb, streamTestDataPath("test-row-styles.xlsx"));
    });

    it("edit styles of single row instead of all", () => {
      const ws = Workbook.getWorksheet(wb, 1)!;

      Worksheet.eachRow(ws, (row, rowNo) => {
        if (rowNo % 5 === 0) {
          rowSetFont(row, { color: { argb: "00ff00" } });
        }
      });

      expect(rowFont(Worksheet.getRow(ws, 3))!.color!.argb).toBe(
        rowFont(Worksheet.getRow(ws, 6))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 6))!.color!.argb).toBe(
        rowFont(Worksheet.getRow(ws, 9))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 9))!.color!.argb).toBe(
        rowFont(Worksheet.getRow(ws, 12))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 12))!.color!.argb).not.toBe(
        rowFont(Worksheet.getRow(ws, 15))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 15))!.color!.argb).not.toBe(
        rowFont(Worksheet.getRow(ws, 18))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 15))!.color!.argb).toBe(
        rowFont(Worksheet.getRow(ws, 10))!.color!.argb
      );
      expect(rowFont(Worksheet.getRow(ws, 10))!.color!.argb).toBe(
        rowFont(Worksheet.getRow(ws, 5))!.color!.argb
      );
    });
  });

  describe("with a spreadsheet that contains formulas", () => {
    let worksheet;
    let cell;

    beforeAll(async () => {
      const workbook = Workbook.create();
      await Workbook.readXlsxStream(
        workbook,
        fs.createReadStream(streamTestDataPath("formulas.xlsx"))
      );
      worksheet = Workbook.getWorksheet(workbook)!;
    });

    describe("with a cell that contains a regular formula", () => {
      beforeEach(() => {
        cell = getCell(worksheet, "A2");
      });

      it("should be classified as a formula cell", () => {
        expect(cellType(cell)).toBe(ValueType.Formula);
      });

      it("should have text corresponding to the evaluated formula result", () => {
        expect(cellText(cell)).toBe("someone@example.com");
      });

      it("should have the formula source", () => {
        expect(cellGetModel(cell).formula).toBe('_xlfn.CONCAT("someone","@example.com")');
      });
    });

    describe("with a cell that contains a hyperlinked formula", () => {
      beforeEach(() => {
        cell = getCell(worksheet, "A1");
      });

      it("should be classified as a formula cell", () => {
        expect(cellType(cell)).toBe(ValueType.Hyperlink);
      });

      it("should have text corresponding to the evaluated formula result", () => {
        expect((cellGetValue(cell) as CellHyperlinkValue).text).toBe("someone@example.com");
      });

      it("should have the formula source", () => {
        expect(cellGetModel(cell).formula).toBe('_xlfn.CONCAT("someone","@example.com")');
      });

      it("should contain the linked url", () => {
        expect((cellGetValue(cell) as CellHyperlinkValue).hyperlink).toBe(
          "mailto:someone@example.com"
        );
        expect(cellHyperlink(cell)).toBe("mailto:someone@example.com");
      });
    });
  });

  describe("with a spreadsheet that contains a shared string with an escaped underscore", () => {
    let worksheet;

    beforeAll(async () => {
      const workbook = Workbook.create();
      await Workbook.readXlsxStream(
        workbook,
        fs.createReadStream(streamTestDataPath("shared_string_with_escape.xlsx"))
      );
      worksheet = Workbook.getWorksheet(workbook)!;
    });

    it("should decode the underscore", () => {
      const cell = getCell(worksheet, "A1");
      expect(cellGetValue(cell)).toBe("_x000D_");
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
            expect(rowValues(row)[1]).toBe("_x000D_");
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
      rowCommit(sheet.addRow(["_x000D_", "Normal", "_x000a_test"]));
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
            expect(rowValues(row)[1]).toBe("_x000D_");
            expect(rowValues(row)[2]).toBe("Normal");
            expect(rowValues(row)[3]).toBe("_x000a_test");
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
      const wb = Workbook.create();
      await Workbook.readXlsxFile(wb, streamTestDataPath("fibonacci.xlsx"));

      const ws = Workbook.getWorksheet(wb, "fib")!;

      expect(Cell.getValue(ws!, "A4")).toEqual({
        formula: "A3+1",
        shareType: "shared",
        ref: "A4:A19",
        result: 4
      });
      expect(Cell.getValue(ws!, "A5")).toEqual({ sharedFormula: "A4", result: 5 });

      expect(Cell.getType(ws!, "A4")).toBe(ValueType.Formula);
      expect(Cell.getType(ws!, "A5")).toBe(ValueType.Formula);
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
          worksheet.on("row", row => rows.push(rowValues(row)[1]));
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
      rowCommit(sheet.addRow(rowData));
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
            expect(rowValues(row)[1]).toEqual(rowData[0]);
            expect(rowValues(row)[2]).toBe(rowData[1]);
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
      const workbook = Workbook.create();
      await expect(
        Workbook.readXlsxFile(workbook, streamTestDataPath("invalid-xml.xlsx"))
      ).rejects.toThrow("3:1: text data outside of root node.");

      // Wait a tick to verify no unhandled rejection fires
      await new Promise(setImmediate);
      expect(unhandledRejection).toBeUndefined();
    });
  });

  describe("with a spreadsheet that is missing some files in the zip container", () => {
    it("should not break", async () => {
      const workbook = Workbook.create();
      await Workbook.readXlsxFile(workbook, streamTestDataPath("missing-bits.xlsx"));
      // Should load gracefully despite missing zip entries
      expect(getWorksheets(workbook).length).toBeGreaterThanOrEqual(0);
    });
  });

  describe("with a spreadsheet that contains images", () => {
    let worksheet;

    beforeAll(async () => {
      const workbook = Workbook.create();
      await Workbook.readXlsxStream(
        workbook,
        fs.createReadStream(streamTestDataPath("images.xlsx"))
      );
      worksheet = Workbook.getWorksheet(workbook)!;
    });

    describe("with image`s tl anchor", () => {
      it("should have at least one image in the test file", () => {
        expect(Image.list(worksheet).length).toBeGreaterThan(0);
      });

      it("Should integer part of col equals nativeCol", () => {
        Image.list(worksheet).forEach(image => {
          expect(Math.floor(anchorCol(image.range!.tl))).toBe(image.range!.tl.nativeCol);
        });
      });
      it("Should integer part of row equals nativeRow", () => {
        Image.list(worksheet).forEach(image => {
          expect(Math.floor(anchorRow(image.range!.tl))).toBe(image.range!.tl.nativeRow);
        });
      });
      it("Should anchor width equals to column width when custom", () => {
        const ws = worksheet;

        Image.list(ws).forEach(image => {
          const col = getColumn(ws, image.range!.tl.nativeCol + 1);

          if (columnIsCustomWidth(col)) {
            expect(anchorColWidth(image.range!.tl)).toBe(Math.floor(col.width! * 10000));
          } else {
            expect(anchorColWidth(image.range!.tl)).toBe(640000);
          }
        });
      });
      it("Should anchor height equals to row height", () => {
        const ws = worksheet;

        Image.list(ws).forEach(image => {
          const row = Worksheet.getRow(ws, image.range!.tl.nativeRow + 1);

          if (row.height != null) {
            expect(anchorRowHeight(image.range!.tl)).toBe(Math.floor(row.height * 10000));
          } else {
            expect(anchorRowHeight(image.range!.tl)).toBe(180000);
          }
        });
      });
    });

    describe("with image`s br anchor", () => {
      it("should have at least one image in the test file", () => {
        expect(Image.list(worksheet).length).toBeGreaterThan(0);
      });

      it("Should integer part of col equals nativeCol", () => {
        Image.list(worksheet).forEach(image => {
          expect(Math.floor(anchorCol(image.range!.br!))).toBe(image.range!.br!.nativeCol);
        });
      });
      it("Should integer part of row equals nativeRow", () => {
        Image.list(worksheet).forEach(image => {
          expect(Math.floor(anchorRow(image.range!.br!))).toBe(image.range!.br!.nativeRow);
        });
      });
      it("Should anchor width equals to column width when custom", () => {
        const ws = worksheet;

        Image.list(ws).forEach(image => {
          const col = getColumn(ws, image.range!.br!.nativeCol + 1);

          if (columnIsCustomWidth(col)) {
            expect(anchorColWidth(image.range!.br!)).toBe(Math.floor(col.width! * 10000));
          } else {
            expect(anchorColWidth(image.range!.br!)).toBe(640000);
          }
        });
      });
      it("Should anchor height equals to row height", () => {
        const ws = worksheet;

        Image.list(ws).forEach(image => {
          const row = Worksheet.getRow(ws, image.range!.br!.nativeRow + 1);

          if (row.height != null) {
            expect(anchorRowHeight(image.range!.br!)).toBe(Math.floor(row.height * 10000));
          } else {
            expect(anchorRowHeight(image.range!.br!)).toBe(180000);
          }
        });
      });
    });
  });
  describe("with a spreadsheet containing a defined name that kinda looks like it contains a range", () => {
    it("should not crash", async () => {
      const workbook = Workbook.create();
      await Workbook.readXlsxStream(
        workbook,
        fs.createReadStream(streamTestDataPath("bogus-defined-name.xlsx"))
      );
      expect(getWorksheets(workbook).length).toBeGreaterThanOrEqual(0);
    });
  });
});
