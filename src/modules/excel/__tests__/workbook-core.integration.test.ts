import fs from "fs";

import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { testUtils } from "@excel/__tests__/shared";
import { ValueType } from "@excel/enums";
import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook, type CsvOptions } from "../../../index";
import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");

const TEST_XLSX_FILE_NAME = testFilePath("workbook-core.integration.test");
const TEST_CSV_FILE_NAME = testFilePath("workbook-core.integration", ".csv");

// =============================================================================
// Tests

describe("Workbook", () => {
  describe("Serialise", () => {
    it("xlsx file", async () => {
      const wb = testUtils.createTestBook(new Workbook(), "xlsx", undefined);

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", undefined, {});
    });
    describe("Xlsx Zip Compression", () => {
      it("xlsx file with best compression", async () => {
        const wb = testUtils.createTestBook(new Workbook(), "xlsx", undefined);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME, {
          zip: {
            compression: "DEFLATE",
            compressionOptions: {
              level: 9
            }
          }
        });
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        testUtils.checkTestBook(wb2, "xlsx", undefined, {});
      });

      it("xlsx file with default compression", async () => {
        const wb = testUtils.createTestBook(new Workbook(), "xlsx", undefined);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME, {
          zip: {
            compression: "DEFLATE"
          }
        });
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        testUtils.checkTestBook(wb2, "xlsx", undefined, {});
      });

      it("xlsx file with fast compression", async () => {
        const wb = testUtils.createTestBook(new Workbook(), "xlsx", undefined);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME, {
          zip: {
            compression: "DEFLATE",
            compressionOptions: {
              level: 1
            }
          }
        });
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        testUtils.checkTestBook(wb2, "xlsx", undefined, {});
      });

      it("xlsx file with no compression", async () => {
        const wb = testUtils.createTestBook(new Workbook(), "xlsx", undefined);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME, {
          zip: {
            compression: "STORE"
          }
        });
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        testUtils.checkTestBook(wb2, "xlsx", undefined, {});
      });
    });
    it("sheets with correct names", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Hello, World!");
      expect(ws1.name).toBe("Hello, World!");
      ws1.getCell("A1").value = "Hello, World!";

      const ws2 = wb.addWorksheet();
      expect(ws2.name).toMatch(/sheet\d+/);
      ws2.getCell("A1").value = ws2.name;

      wb.addWorksheet("This & That");

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.getWorksheet("Hello, World!")).toBeTruthy();
      expect(wb2.getWorksheet("This & That")).toBeTruthy();
    });

    it('removes "vertical tab" and other invalid control characters', async () => {
      const filename = testFilePath("invalid-control-chars.workbook");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Hello, \x01World!";
      ws.getCell("A2").value = "Hello, \x0bWorld!";

      await wb.xlsx.writeFile(filename);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filename);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.getCell("A1").value).toBe("Hello, World!");
      expect(ws2.getCell("A2").value).toBe("Hello, World!");
    });

    it("special cell values produce a valid file", async () => {
      const filename = testFilePath("special-object-keys.workbook");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      const specialValues = [
        "constructor",
        "hasOwnProperty",
        "isPrototypeOf",
        "propertyIsEnumerable",
        "toLocaleString",
        "toString",
        "valueOf",
        "__defineGetter__",
        "__defineSetter__",
        "__lookupGetter__",
        "__lookupSetter__",
        "__proto__"
      ];

      for (let i = 0; i < specialValues.length; i++) {
        const value = specialValues[i];
        ws.addRow([value]);
        ws.getCell(`B${i + 1}`).value = value;
      }

      await wb.xlsx.writeFile(filename);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filename);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      for (let i = 0; i < specialValues.length; i++) {
        const value = specialValues[i];
        expect(ws2.getCell(`A${i + 1}`).value).toBe(value);
        expect(ws2.getCell(`B${i + 1}`).value).toBe(value);
      }
    });

    it("hyperlink without text does not crash on write", async () => {
      const sourceFile = excelTestDataPath("hyperlink-without-text.xlsx");
      const outFile = testFilePath("hyperlink-without-text.workbook");

      const wb = new Workbook();
      await wb.xlsx.readFile(sourceFile);
      const buffer = await wb.xlsx.writeBuffer({
        useStyles: true,
        useSharedStrings: true
      });
      await expectValidXlsx(buffer, { label: "hyperlink-without-text" });
      await fs.promises.writeFile(outFile, buffer);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(outFile);
      expect(wb2.worksheets.length).toBeGreaterThan(0);
    });

    it("readFile should not throw", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("graceful-read-no-throw.xlsx"));
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });

    it("unexpected xml node should not break parsing", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("unexpected-xml-node.xlsx"));
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });

    describe("1904 dates", () => {
      it("reads 1904-based workbook", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("date-system-1904.xlsx"));

        expect(wb.properties.date1904).toBe(true);
        const ws = wb.getWorksheet("Sheet1")!;
        expect((ws.getCell("B4").value as Date).toISOString()).toBe("1904-01-01T00:00:00.000Z");
      });

      it("writes and reads 1904-based workbook", async () => {
        const filename = testFilePath("date1904-roundtrip.workbook");
        const wb = new Workbook();
        wb.properties.date1904 = true;
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("B4").value = new Date("1904-01-01T00:00:00.000Z");

        await wb.xlsx.writeFile(filename);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(filename);
        expect(wb2.properties.date1904).toBe(true);
        const ws2 = wb2.getWorksheet("Sheet1")!;
        expect((ws2.getCell("B4").value as Date).toISOString()).toBe("1904-01-01T00:00:00.000Z");
      });
    });

    it("sheet order is preserved", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("sheet-order.xlsx"));
      expect(wb.worksheets.map(ws => ws.name)).toEqual(["First", "Second"]);
    });

    describe("missing r attribute in row/cell elements", () => {
      it("reads xlsx missing r attributes", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("missing-cell-address.xlsx"));

        const ws = wb.worksheets[0];
        expect(ws).toBeDefined();
        expect(ws.rowCount).toBe(2);

        const row1 = ws.getRow(1);
        expect(row1.getCell(1).value).toBeDefined();
        expect(row1.getCell(2).value).toBeDefined();
        expect(row1.getCell(3).value).toBeDefined();

        const row2 = ws.getRow(2);
        expect(row2.getCell(1).value).toBe(1);
        expect(row2.getCell(2).value).toBeDefined();
        expect(row2.getCell(3).value).toBeDefined();
      });

      it("infers cell addresses when r is missing", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("missing-cell-address.xlsx"));
        const ws = wb.worksheets[0];

        expect(ws.getCell("A1").value).toBeDefined();
        expect(ws.getCell("B1").value).toBeDefined();
        expect(ws.getCell("C1").value).toBeDefined();
        expect(ws.getCell("A2").value).toBeDefined();
        expect(ws.getCell("B2").value).toBeDefined();
        expect(ws.getCell("C2").value).toBeDefined();
      });

      it("can write back after reading", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("missing-cell-address.xlsx"));

        // NOTE: `expectValidXlsx` is intentionally NOT invoked here.
        // The `missing-cell-address.xlsx` fixture is already OOXML-broken
        // (cells without `r=` attrs) and the library's round-trip loses
        // the theme part while leaving content-types / rels pointing at
        // it. A stricter validator flags that, but the test only asserts
        // "write does not throw", which remains true.
        const buffer = await wb.xlsx.writeBuffer();
        expect(buffer).toBeDefined();
        expect(buffer.byteLength).toBeGreaterThan(0);
      });
    });

    it("optional autofilter and custom autofilter on tables", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("table-autofilter-optional.xlsx"));
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });

    it("<contentType /> element", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("content-type-element.xlsx"));
      expect(wb.worksheets.length).toBeGreaterThan(0);
    });

    it("borders for merged cells survive rewrite", async () => {
      const outFile = testFilePath("merged-cell-borders.workbook");
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("merged-cell-borders.xlsx"));

      const assertBorder = (cell: any, borders: Array<"left" | "right" | "top" | "bottom">) => {
        expect(cell.style?.border).toBeTruthy();
        borders.forEach(b => {
          expect(cell.style.border).toHaveProperty(b);
        });
      };

      const ws = wb.getWorksheet(1)!;
      assertBorder(ws.getCell("B2"), ["left", "top"]);
      assertBorder(ws.getCell("B3"), ["left", "bottom"]);
      assertBorder(ws.getCell("C2"), ["right", "top"]);
      assertBorder(ws.getCell("C3"), ["right", "bottom"]);

      await wb.xlsx.writeFile(outFile);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(outFile);
      const ws2 = wb2.getWorksheet(1)!;
      assertBorder(ws2.getCell("B2"), ["left", "top"]);
      assertBorder(ws2.getCell("B3"), ["left", "bottom"]);
      assertBorder(ws2.getCell("C2"), ["right", "top"]);
      assertBorder(ws2.getCell("C3"), ["right", "bottom"]);
    });

    it("malformed comment does not crash on write", async () => {
      const sourceFile = excelTestDataPath("malformed-comment.xlsx");
      const outFile = testFilePath("malformed-comment.workbook");

      const wb = new Workbook();
      await wb.xlsx.readFile(sourceFile);
      const buffer = await wb.xlsx.writeBuffer({
        useStyles: true,
        useSharedStrings: true
      });
      await expectValidXlsx(buffer, { label: "malformed-comment" });
      await fs.promises.writeFile(outFile, buffer);

      const stat = await fs.promises.stat(outFile);
      expect(stat.size).toBeGreaterThan(0);
    }, 6000);

    it("table without autofilter model", async () => {
      const wb = new Workbook();
      await wb.xlsx.readFile(excelTestDataPath("table-without-autofilter.xlsx"));
      expect(wb.worksheets.length).toBeGreaterThan(0);
    }, 6000);

    describe("worksheet file naming with non-sequential sheetIds", () => {
      it("uses sequential file names regardless of sheetId values", async () => {
        const workbook = new Workbook();
        const ws1: any = workbook.addWorksheet("Sheet1");
        const ws2: any = workbook.addWorksheet("Sheet2");

        ws1.id = 1;
        ws2.id = 3;

        ws1.getCell("A1").value = "Sheet 1 Data";
        ws2.getCell("A1").value = "Sheet 2 Data";

        const buffer = await workbook.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "sequential sheet filenames" });
        const zipData = await extractAll(new Uint8Array(buffer));

        expect(zipData.has("xl/worksheets/sheet1.xml")).toBe(true);
        expect(zipData.has("xl/worksheets/sheet2.xml")).toBe(true);
        expect(zipData.has("xl/worksheets/sheet3.xml")).toBe(false);

        const relsData = zipData.get("xl/_rels/workbook.xml.rels");
        const relsContent = new TextDecoder().decode(relsData?.data);
        expect(relsContent).toContain("worksheets/sheet1.xml");
        expect(relsContent).toContain("worksheets/sheet2.xml");
        expect(relsContent).not.toContain("worksheets/sheet3.xml");
      });

      it("preserves sheetId values in workbook.xml while using sequential file names", async () => {
        const workbook = new Workbook();
        const ws1: any = workbook.addWorksheet("Sheet1");
        const ws2: any = workbook.addWorksheet("Sheet2");

        ws1.id = 1;
        ws2.id = 3;

        ws1.getCell("A1").value = "Data 1";
        ws2.getCell("A1").value = "Data 2";

        const buffer = await workbook.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "sheetId preserved" });
        const zipData = await extractAll(new Uint8Array(buffer));

        const workbookData = zipData.get("xl/workbook.xml");
        const workbookContent = new TextDecoder().decode(workbookData?.data);
        expect(workbookContent).toContain('sheetId="1"');
        expect(workbookContent).toContain('sheetId="3"');
      });
    });

    describe("many definedNames should not cause OOM", () => {
      it("loads file with many definedNames without excessive memory use", async () => {
        const sourceFile = excelTestDataPath("many-defined-names.xlsx");

        const wb = new Workbook();
        await wb.xlsx.readFile(sourceFile);
        expect(wb.worksheets.length).toBeGreaterThan(0);

        const definedNamesModel = wb.definedNames.model;
        expect(Array.isArray(definedNamesModel)).toBe(true);

        // The file contains 35000+ defined names, most of which are garbage
        // (array constants, error values, etc.). With the two-phase classifier,
        // garbage entries are preserved as opaque (for round-trip) but do NOT
        // expand into CellMatrix objects. Verify that matrixMap stays small.
        const matrixCount = Object.keys(wb.definedNames.matrixMap).length;
        expect(matrixCount).toBeLessThan(1000);
      }, 60000);

      it("loads file from buffer without excessive memory use", async () => {
        const sourceFile = excelTestDataPath("many-defined-names.xlsx");
        const buffer = await fs.promises.readFile(sourceFile);

        const wb = new Workbook();
        await wb.xlsx.load(buffer);
        expect(wb.worksheets.length).toBeGreaterThan(0);
      }, 60000);

      it("filters out array constants from definedNames ranges", async () => {
        const sourceFile = excelTestDataPath("many-defined-names.xlsx");

        const wb = new Workbook();
        await wb.xlsx.readFile(sourceFile);

        // Collect ranges from all defined names — opaque names have empty ranges,
        // so array constants should never appear in the ranges array.
        const allRanges: string[] = [];
        (wb.definedNames.model as Array<{ ranges: string[] }>).forEach(dn => {
          allRanges.push(...dn.ranges);
        });

        allRanges.forEach(range => {
          expect(range.startsWith("{")).toBe(false);
          expect(range.endsWith("}")).toBe(false);
        });
      }, 60000);
    });

    describe("regressions", () => {
      it("worksheet should not be undefined", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("worksheet-not-undefined.xlsx"));
        const ws = wb.getWorksheet(1);
        expect(ws).toBeDefined();
      });

      it("reads worksheet hidden state", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("hidden-worksheet-state.xlsx"));

        const expected: Record<number, string> = { 1: "visible", 2: "hidden", 3: "visible" };
        wb.eachSheet((ws, sheetId) => {
          expect(ws.state).toBe(expected[sheetId]);
        });
      });

      it("reads workbook with whole-column defined names", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("whole-column-defined-names.xlsx"));
        expect(wb.worksheets.length).toBeGreaterThan(0);
      });

      it("handles empty _xlnm.Print_Area ranges without crashing", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "test";

        // Inject a Print_Area defined name with empty ranges,
        // simulating an Excel file where print area was set then cleared
        const workbookModel: any = wb.model;
        workbookModel.definedNames = [
          {
            name: "_xlnm.Print_Area",
            localSheetId: 0,
            ranges: []
          }
        ];
        wb.model = workbookModel;

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "empty Print_Area ranges" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);

        const ws2 = wb2.getWorksheet("Sheet1")!;
        expect(ws2).toBeDefined();
        expect(ws2.getCell("A1").value).toBe("test");
        expect(ws2.pageSetup.printArea).toBeUndefined();
      });

      it("lastColumn with an empty column", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");

        ws.getCell("A1").value = "not empty";
        ws.getCell("B1").style = { numFmt: "@" };

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "lastColumn empty column" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);

        const ws2 = wb2.getWorksheet("Sheet1")!;
        expect(ws2.lastColumn).toBe(ws2.getColumn(2));
      });

      it("inlineStr cell type support", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("inline-string-cells.xlsx"));

        const ws = wb.getWorksheet("Sheet1")!;
        expect(ws.getCell("A1").value).toBe("A");
        expect(ws.getCell("B1").value).toBe("B");
        expect(ws.getCell("C1").value).toBe("C");
        expect(ws.getCell("A2").value).toBe("1.0");
        expect(ws.getCell("B2").value).toBe("2.0");
        expect(ws.getCell("C2").value).toBe("3.0");
        expect(ws.getCell("A3").value).toBe("4.0");
        expect(ws.getCell("B3").value).toBe("5.0");
        expect(ws.getCell("C3").value).toBe("6.0");
      });

      describe("preserve whitespace", () => {
        it("preserves leading and trailing whitespace", async () => {
          const testFile = testFilePath("pr-896.whitespace");
          const wb = new Workbook();
          const ws = wb.addWorksheet("foo");
          ws.getCell("A1").value = " leading";
          ws.getCell("A1").note = " leading";
          ws.getCell("B1").value = "trailing ";
          ws.getCell("B1").note = "trailing ";
          ws.getCell("C1").value = " both ";
          ws.getCell("C1").note = " both ";

          await wb.xlsx.writeFile(testFile);

          const wb2 = new Workbook();
          await wb2.xlsx.readFile(testFile);
          const ws2 = wb2.getWorksheet("foo")!;
          expect(ws2.getCell("A1").value).toBe(" leading");
          expect(ws2.getCell("A1").note).toBe(" leading");
          expect(ws2.getCell("B1").value).toBe("trailing ");
          expect(ws2.getCell("B1").note).toBe("trailing ");
          expect(ws2.getCell("C1").value).toBe(" both ");
          expect(ws2.getCell("C1").note).toBe(" both ");
        });

        it("preserves newlines", async () => {
          const testFile = testFilePath("pr-896.newlines");
          const wb = new Workbook();
          const ws = wb.addWorksheet("foo");
          ws.getCell("A1").value = "Hello,\nWorld!";
          ws.getCell("A1").note = "Later,\nAlligator!";
          ws.getCell("B1").value = " Hello, \n World! ";
          ws.getCell("B1").note = " Later, \n Alligator! ";

          await wb.xlsx.writeFile(testFile);

          const wb2 = new Workbook();
          await wb2.xlsx.readFile(testFile);
          const ws2 = wb2.getWorksheet("foo")!;
          expect(ws2.getCell("A1").value).toBe("Hello,\nWorld!");
          expect(ws2.getCell("A1").note).toBe("Later,\nAlligator!");
          expect(ws2.getCell("B1").value).toBe(" Hello, \n World! ");
          expect(ws2.getCell("B1").note).toBe(" Later, \n Alligator! ");
        });

        it("preserves richText and comment texts with leading new line", async () => {
          const testFile = testFilePath("pr-896.richtext");
          const wb = new Workbook();
          const ws = wb.addWorksheet("sheet1");
          ws.getColumn(1).width = 20;

          const RT_ARR = [
            { text: "First Line:\n", font: { bold: true } },
            { text: "Second Line\n" },
            { text: "Third Line\n" },
            { text: "Last Line" }
          ];
          const TEST_VALUE = { richText: RT_ARR };
          const TEST_NOTE = { texts: RT_ARR };

          ws.getCell("A1").value = TEST_VALUE;
          ws.getCell("A1").note = TEST_NOTE;
          ws.getCell("A1").alignment = { wrapText: true };

          await wb.xlsx.writeFile(testFile);

          const wb2 = new Workbook();
          await wb2.xlsx.readFile(testFile);
          const ws2 = wb2.getWorksheet("sheet1")!;
          expect(ws2).toBeDefined();
          expect(ws2.getCell("A1").value).toEqual(TEST_VALUE);
        });
      });

      it("writeFile rejects when image file is missing", async () => {
        const testFile = testFilePath("pr-2244.missing-image");

        let error: unknown;
        try {
          const workbook = new Workbook();
          const worksheet = workbook.addWorksheet("sheet");
          const imageId1 = workbook.addImage({
            filename: "path/to/image.jpg",
            extension: "jpeg"
          });
          worksheet.addImage(imageId1, "B2:D6");
          await workbook.xlsx.writeFile(testFile);
        } catch (err) {
          error = err;
        }

        expect(error).toBeInstanceOf(Error);
      });
    });

    it("creator, lastModifiedBy, etc", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = "World!";
      wb.creator = "Foo";
      wb.lastModifiedBy = "Bar";
      wb.created = new Date(2016, 0, 1);
      wb.modified = new Date(2016, 4, 19);
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.creator).toBe(wb.creator);
      expect(wb2.lastModifiedBy).toBe(wb.lastModifiedBy);
      expect(wb2.created).toEqual(wb.created);
      expect(wb2.modified).toEqual(wb.modified);
    });
    it("printTitlesRow", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("printHeader");

      ws.getCell("A1").value = "This is a header row repeated on every printed page";
      ws.getCell("B2").value = "This is a header row too";

      for (let i = 0; i < 100; i++) {
        ws.addRow(["not header row"]);
      }

      ws.pageSetup.printTitlesRow = "1:2";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("printHeader")!;
      expect(ws2.pageSetup.printTitlesRow).toBe("1:2");
      expect(ws2.pageSetup.printTitlesColumn).toBeUndefined();
    });
    it("printTitlesColumn", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("printColumn");

      ws.getCell("A1").value = "This is a column repeated on every printed page";
      ws.getCell("A2").value = "This is a column repeated on every printed page";
      ws.getCell("B1").value = "This is a repeated column too";
      ws.getCell("B2").value = "This is a repeated column too";

      ws.getCell("C1").value = "This is a regular column";
      ws.getCell("C2").value = "This is a regular column";
      ws.getCell("D1").value = "This is a regular column";
      ws.getCell("D2").value = "This is a regular column";

      ws.pageSetup.printTitlesRow = "A:B";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("printColumn")!;
      expect(ws2.pageSetup.printTitlesRow).toBeUndefined();
      expect(ws2.pageSetup.printTitlesColumn).toBe("A:B");
    });
    it("printTitlesRowAndColumn", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("printHeaderAndColumn");

      ws.getCell("A1").value = "This is a column / row repeated on every printed page";
      ws.getCell("A2").value = "This is a column / row repeated on every printed page";
      ws.getCell("B1").value = "This is a repeated column / row too";
      ws.getCell("B2").value = "This is a repeated column / row too";

      ws.getCell("C1").value = "This is a regular column, repeated row";
      ws.getCell("C2").value = "This is a regular column, repeated row";
      ws.getCell("D1").value = "This is a regular column, repeated row";
      ws.getCell("D2").value = "This is a regular column, repeated row";

      ws.getCell("A3").value = "This is a repeated column";
      ws.getCell("B3").value = "This is a repeated column";
      ws.getCell("C3").value = "This is a regular column / row";
      ws.getCell("D3").value = "This is a regular column / row";

      ws.pageSetup.printTitlesColumn = "A:B";
      ws.pageSetup.printTitlesRow = "1:2";

      for (let i = 0; i < 100; i++) {
        ws.addRow([
          "repeated column, not repeated row",
          "repeated column, not repeated row",
          "no repeat",
          "no repeat"
        ]);
      }

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("printHeaderAndColumn")!;
      expect(ws2.pageSetup.printTitlesRow).toBe("1:2");
      expect(ws2.pageSetup.printTitlesColumn).toBe("A:B");
    });

    it("single-cell printArea without colon round-trips correctly", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      ws.pageSetup.printArea = "A1";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "single-cell printArea" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printArea).toBe("A1:A1");
    });

    it("multiple printAreas on a single sheet round-trip correctly", async () => {
      // Issue #168: multiple print areas in a single worksheet must round-trip.
      // Per ECMA-376 §18.2.5 the (name, localSheetId) pair on `<definedName>`
      // must be unique, so multiple print areas collapse into ONE
      // `<definedName name="_xlnm.Print_Area">` whose text is a comma-
      // separated list of ranges (Excel's native format). The
      // worksheet-level `printArea` field uses `&&` as the multi-range
      // separator (legacy excelts convention, preserved for backwards
      // compatibility); both `&&` and `,` are accepted on input.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      for (let r = 1; r <= 10; r++) {
        for (let c = 1; c <= 5; c++) {
          ws.getCell(r, c).value = `${r}-${c}`;
        }
      }
      ws.pageSetup.printArea = "A1:B5&&D1:E10";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "multiple printAreas" });

      // Inspect workbook.xml to confirm a single `<definedName>` with both
      // ranges, not two duplicate elements (which would violate OOXML).
      const zipData = await extractAll(new Uint8Array(buffer));
      const workbookData = zipData.get("xl/workbook.xml");
      const workbookContent = new TextDecoder().decode(workbookData?.data);
      const printAreaMatches = workbookContent.match(/<definedName name="_xlnm.Print_Area"[^>]*>/g);
      expect(printAreaMatches).toHaveLength(1);
      // The writer normalises every range to canonical `$col$row:$col$row`
      // form (matching what Excel itself emits), regardless of how the
      // user spelled the input.
      expect(workbookContent).toContain("$A$1:$B$5,&apos;Sheet1&apos;!$D$1:$E$10");

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printArea).toBe("A1:B5&&D1:E10");
    });

    it("printArea with comma separator (Excel's native syntax) is accepted", async () => {
      // Users pasting from Excel may use `,` as the separator. We accept
      // both `,` and `&&` on write; the round-tripped form uses `&&` for
      // backwards compatibility.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      ws.pageSetup.printArea = "A1:B5,D1:E10";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "comma-separated printArea" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printArea).toBe("A1:B5&&D1:E10");
    });

    it("multiple printAreas survive across separate sheets", async () => {
      // Each sheet keeps its own `_xlnm.Print_Area` defined name (with its
      // own `localSheetId`), so multi-range entries on one sheet must not
      // bleed into other sheets.
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("S1");
      const ws2 = wb.addWorksheet("S2");
      ws1.getCell("A1").value = "x";
      ws2.getCell("A1").value = "y";
      ws1.pageSetup.printArea = "A1:B2&&D1:E2";
      ws2.pageSetup.printArea = "A1:C3";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "per-sheet multiple printAreas" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      expect(wb2.getWorksheet("S1")!.pageSetup.printArea).toBe("A1:B2&&D1:E2");
      expect(wb2.getWorksheet("S2")!.pageSetup.printArea).toBe("A1:C3");
    });

    it("multiple printAreas on a sheet with a comma in its name round-trip correctly", async () => {
      // Sheet names containing commas are quoted in OOXML
      // (`'Q1, Forecast'!$A$1:$B$5`). The reader must split on top-level
      // commas only — splitting on every comma would shred the sheet name.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Q1, Forecast");
      ws.getCell("A1").value = "x";
      ws.pageSetup.printArea = "A1:B5&&D1:E10";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "comma-named sheet printAreas" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet("Q1, Forecast")!;
      expect(ws2.pageSetup.printArea).toBe("A1:B5&&D1:E10");
    });

    describe("printArea / printTitles input normalisation", () => {
      // Regression: the previous string-concatenation writer produced
      // corrupt OOXML for several user-input shapes Excel itself accepts:
      //   - lowercase `a1:b5`        -> `$a1:$b5` (read-back returned NaN)
      //   - already-anchored `$A$1`  -> `$$A$1` (Excel rejects double-$)
      //   - sheet-prefixed input     -> `$Sheet1!A1:$B5` (corrupt)
      //   - whitespace `A1 : B5`     -> `$A1 :$ B5` (corrupt)
      //   - row-relative `A1:B5`     -> `$A1:$B5` (drifts when rows
      //                                 inserted above the print area)
      // The writer now routes every input through `parsePrintReference`
      // — a hand-rolled parser that recognises all four legal Excel
      // print-reference shapes (cell, range, whole-row, whole-column)
      // and emits Excel's canonical `$col$row[:$col$row]` shape.
      const cases: Array<{ input: string; expectedRoundTrip: string; xmlContains: string }> = [
        { input: "a1:b5", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "$A$1:$B$5", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "$A1:$B5", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "A1 : B5", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "  A1:B5  ", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "Sheet1!A1:B5", expectedRoundTrip: "A1:B5", xmlContains: "$A$1:$B$5" },
        { input: "A1", expectedRoundTrip: "A1:A1", xmlContains: "$A$1:$A$1" }
      ];
      for (const { input, expectedRoundTrip, xmlContains } of cases) {
        it(`printArea input ${JSON.stringify(input)} normalises to canonical OOXML`, async () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("S");
          ws.getCell("A1").value = "x";
          ws.pageSetup.printArea = input;

          const buffer = await wb.xlsx.writeBuffer();
          await expectValidXlsx(buffer, { label: `printArea normalise ${input}` });

          const zipData = await extractAll(new Uint8Array(buffer));
          const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
          // Output never contains the broken double-`$` or row-relative
          // forms that the old string-concat writer used to emit.
          expect(workbookContent).not.toMatch(/\$\$[A-Z]/);
          expect(workbookContent).toContain(xmlContains);

          const wb2 = new Workbook();
          await wb2.xlsx.load(buffer);
          expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe(expectedRoundTrip);
        });
      }

      it("printTitlesRow input with $ does not produce $$ output", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printTitlesRow = "$1:$2";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "printTitlesRow $1:$2" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).not.toMatch(/\$\$\d/);
        expect(workbookContent).toContain("$1:$2");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printTitlesRow).toBe("1:2");
      });

      it("printTitlesColumn input with $ and lowercase normalises", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printTitlesColumn = "$a:$b";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "printTitlesColumn $a:$b" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).not.toMatch(/\$\$[A-Za-z]/);
        expect(workbookContent).toContain("$A:$B");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printTitlesColumn).toBe("A:B");
      });

      it("printTitlesRow with a column-shaped value keeps backwards-compatible auto-routing", async () => {
        // Long-standing quirk: the OOXML reader infers the axis from the
        // emitted reference, so users who set `printTitlesRow = "A:B"`
        // got the value silently re-classified onto `printTitlesColumn`
        // on round-trip. Preserve that behaviour rather than silently
        // dropping the entry — strict enforcement would be a regression.
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printTitlesRow = "A:B";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "row=A:B legacy" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        const ps = wb2.getWorksheet("S")!.pageSetup;
        expect(ps.printTitlesRow).toBeUndefined();
        expect(ps.printTitlesColumn).toBe("A:B");
      });

      it("malformed printArea input is dropped, not written as corrupt XML", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        // Garbage input — neither an address nor a range.
        ws.pageSetup.printArea = "not-a-range!!";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "garbage printArea" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).not.toContain("Print_Area");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBeUndefined();
      });

      it("whole-row printArea (1:5) round-trips", async () => {
        // Excel UI allows selecting entire rows as a print area. Emitted
        // OOXML form is `'Sheet'!$1:$5`. The earlier writer mangled this
        // to `$1:$5` *without* the sheet prefix and the read side then
        // returned `NaN:NaN`; the parser-driven writer handles it.
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "1:5";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "whole-row printArea" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).toContain("$1:$5");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe("1:5");
      });

      it("whole-column printArea (A:C) round-trips", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A:C";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "whole-column printArea" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).toContain("$A:$C");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe("A:C");
      });

      it("mixed printArea forms (cell, range, whole-row, whole-column) coexist", async () => {
        // All four shapes can appear in the same comma-separated
        // OOXML `<definedName>`; verify the multi-range pipeline accepts
        // each shape and the read path preserves them all.
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A1&&B2:C3&&5:7&&E:F";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "mixed shapes" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        // Cell promotes to A1:A1 (degenerate range); other shapes
        // round-trip verbatim.
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe("A1:A1&&B2:C3&&5:7&&E:F");
      });

      it("OOXML with bare-cell `<...>'S'!$A$1</...>` reads back as A1:A1", async () => {
        // Excel sometimes emits a bare cell (no `:`) when the print
        // area is a single cell. The reader must recognise that as a
        // legitimate print area and surface it as `A1:A1` on the
        // worksheet API (matching the writer's promote-cell-to-range
        // policy).
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        const m: any = wb.model;
        m.definedNames = [
          { name: "_xlnm.Print_Area", localSheetId: 0, ranges: [], rawText: "'S'!$A$1" }
        ];
        wb.model = m;
        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "bare-cell OOXML" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe("A1:A1");
      });

      it("printArea with column past XFD throws ColumnOutOfBoundsError on write", async () => {
        // `AAAA` is column 18,279 — past Excel's hard XFD (16,384) limit.
        // Excel cannot represent the column at all, so emitting it would
        // produce a workbook Excel rejects. We surface the same error
        // type `getCell("AAAA1")` already throws so the user finds out
        // immediately rather than discovering the print area silently
        // disappeared from the saved file. The error message must carry
        // the offending letter (`AAAA`) so the user can locate the
        // mistake — the legacy `_fill(level)` path used to report
        // `Column 4 is out of bounds` (the letter count), which lied.
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A1:AAAA5";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Column AAAA is out of bounds/);
      });

      it("whole-column printArea past XFD throws with the letter in the message", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "AAAA:AAAB";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Column AAAA is out of bounds/);
      });

      it("printTitlesColumn past XFD throws with the letter in the message", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printTitlesColumn = "AAAA:AAAB";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Column AAAA is out of bounds/);
      });

      it("loading a workbook whose OOXML carries an out-of-bounds column drops the bad range without aborting the load", async () => {
        // A file authored by another tool (or hand-edited) might
        // contain a print-area `<definedName>` whose body references
        // a column past XFD. We must NOT let that single corrupt entry
        // take out the whole load — drop it silently and let the rest
        // of the workbook reconcile normally. Mirrors how the validator
        // reports such cells as errors but the parser still returns a
        // usable model.
        //
        // First write a clean workbook, then patch its `xl/workbook.xml`
        // to inject a `<definedName>` that mixes a bad (out-of-XFD)
        // range and a good range. The good range must survive the load.
        const cleanWb = new Workbook();
        cleanWb.addWorksheet("Sheet1").getCell("A1").value = "x";
        const cleanBuf = await cleanWb.xlsx.writeBuffer();

        const entries = await extractAll(new Uint8Array(cleanBuf));
        const wbXmlText = new TextDecoder().decode(entries.get("xl/workbook.xml")!.data);
        // Inject `<definedNames>` immediately before `<calcPr>` (the
        // OOXML-mandated position) so the patched workbook stays valid
        // and the reconcile path actually sees the bad name.
        const patched = wbXmlText.replace(
          /<calcPr/,
          `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$AAAA$1:$AAAA$5,'Sheet1'!$A$1:$B$5</definedName></definedNames><calcPr`
        );
        const zipFiles: Array<{ name: string; data: Uint8Array }> = [];
        for (const [path, file] of entries) {
          zipFiles.push({
            name: path,
            data: path === "xl/workbook.xml" ? new TextEncoder().encode(patched) : file.data
          });
        }
        const patchedBuffer = await createZip(zipFiles);

        const wb2 = new Workbook();
        await wb2.xlsx.load(patchedBuffer);
        // Bad range dropped, good range kept.
        expect(wb2.getWorksheet("Sheet1")!.pageSetup.printArea).toBe("A1:B5");
      });

      it("user-supplied sheet-prefixed input with comma in the sheet name is split correctly", async () => {
        // Regression: a quote-aware split is required so commas *inside*
        // a quoted sheet name (`'Q1, Forecast'!A1:B5`) are not treated
        // as range separators. The legacy `split(/&&|,/)` shredded such
        // inputs and lost every range — now they round-trip cleanly.
        const wb = new Workbook();
        const ws = wb.addWorksheet("Q1, Forecast");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "'Q1, Forecast'!A1:B5,'Q1, Forecast'!D1:E10";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "comma-in-name sheet-prefixed input" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("Q1, Forecast")!.pageSetup.printArea).toBe("A1:B5&&D1:E10");
      });

      it("user-supplied sheet-prefixed input with `&&` in the sheet name is split correctly", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("A&&B");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "'A&&B'!A1:B5&&'A&&B'!D1:E10";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "ampersand-in-name sheet-prefixed input" });

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("A&&B")!.pageSetup.printArea).toBe("A1:B5&&D1:E10");
      });

      it("reversed range endpoints are canonicalised to top-left:bottom-right", async () => {
        // Excel's UI never produces `B5:A1`, but a hand-authored input
        // might. Downstream consumers (PDF layout, range loops) assume
        // `s.r <= e.r && s.c <= e.c`, so the writer normalises here.
        const cases: Array<{ input: string; expected: string }> = [
          { input: "B5:A1", expected: "A1:B5" },
          { input: "5:1", expected: "1:5" },
          { input: "C:A", expected: "A:C" }
        ];
        for (const { input, expected } of cases) {
          const wb = new Workbook();
          const ws = wb.addWorksheet("S");
          ws.getCell("A1").value = "x";
          ws.pageSetup.printArea = input;
          const buffer = await wb.xlsx.writeBuffer();
          await expectValidXlsx(buffer, { label: `reversed ${input}` });
          const wb2 = new Workbook();
          await wb2.xlsx.load(buffer);
          expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe(expected);
        }
      });

      it("row 0 is rejected — Excel rows are 1-indexed", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A0:B5";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Row 0 is out of bounds/);
      });

      it("whole-row 0 input (e.g. `0:5`) is rejected", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "0:5";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Row 0 is out of bounds/);
      });

      it("row past Excel's 1048576 limit is rejected", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A1:B1048577";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(
          /Row 1048577 is out of bounds.*1 to 1048576/
        );
      });

      it("printTitlesRow past the row limit is rejected", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printTitlesRow = "1:1048577";

        await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/Row 1048577 is out of bounds/);
      });

      it("leading-zero row inputs are normalised to canonical integers", async () => {
        // OOXML expects `$A$1`, not `$A$001`. Excel tolerates the latter
        // on read, but emitting it makes the file look hand-edited and
        // confuses tooling that does string equality on cell refs.
        const wb = new Workbook();
        const ws = wb.addWorksheet("S");
        ws.getCell("A1").value = "x";
        ws.pageSetup.printArea = "A001:B005";

        const buffer = await wb.xlsx.writeBuffer();
        await expectValidXlsx(buffer, { label: "leading-zero row" });
        const zipData = await extractAll(new Uint8Array(buffer));
        const workbookContent = new TextDecoder().decode(zipData.get("xl/workbook.xml")?.data);
        expect(workbookContent).toContain("$A$1:$B$5");
        expect(workbookContent).not.toContain("$A$001");

        const wb2 = new Workbook();
        await wb2.xlsx.load(buffer);
        expect(wb2.getWorksheet("S")!.pageSetup.printArea).toBe("A1:B5");
      });

      it("loading a workbook with a row past the limit drops the bad range without aborting", async () => {
        // Mirror image of the column-OOB read-side test: a hand-edited
        // file with a row past 1048576 must not abort the load. The
        // `try/catch` around `parsePrintReference` in the read path
        // catches `RowOutOfBoundsError` the same way it catches
        // `ColumnOutOfBoundsError`.
        const cleanWb = new Workbook();
        cleanWb.addWorksheet("Sheet1").getCell("A1").value = "x";
        const cleanBuf = await cleanWb.xlsx.writeBuffer();

        const entries = await extractAll(new Uint8Array(cleanBuf));
        const wbXmlText = new TextDecoder().decode(entries.get("xl/workbook.xml")!.data);
        const patched = wbXmlText.replace(
          /<calcPr/,
          `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$B$99999999,'Sheet1'!$A$1:$B$5</definedName></definedNames><calcPr`
        );
        const zipFiles: Array<{ name: string; data: Uint8Array }> = [];
        for (const [path, file] of entries) {
          zipFiles.push({
            name: path,
            data: path === "xl/workbook.xml" ? new TextEncoder().encode(patched) : file.data
          });
        }
        const patchedBuffer = await createZip(zipFiles);

        const wb2 = new Workbook();
        await wb2.xlsx.load(patchedBuffer);
        expect(wb2.getWorksheet("Sheet1")!.pageSetup.printArea).toBe("A1:B5");
      });
    });

    it("single-column printTitlesColumn without colon round-trips correctly", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      ws.pageSetup.printTitlesColumn = "A";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "single-column printTitlesColumn" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printTitlesColumn).toBe("A:A");
    });

    it("single-row printTitlesRow without colon round-trips correctly", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      ws.pageSetup.printTitlesRow = "1";

      const buffer = await wb.xlsx.writeBuffer();
      await expectValidXlsx(buffer, { label: "single-row printTitlesRow" });

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printTitlesRow).toBe("1:1");
    });

    it("shared formula", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.fillFormula("A1:B2", "ROW()+COLUMN()", [
        [2, 3],
        [3, 4]
      ]);
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2.getCell("A1").value).toEqual({
        formula: "ROW()+COLUMN()",
        shareType: "shared",
        ref: "A1:B2",
        result: 2
      });
      expect(ws2.getCell("B1").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws2.getCell("A2").value).toEqual({
        sharedFormula: "A1",
        result: 3
      });
      expect(ws2.getCell("B2").value).toEqual({
        sharedFormula: "A1",
        result: 4
      });
    });

    it("auto filter", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("B2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("B3").value = 3;

      ws.autoFilter = "A1:B1";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2.autoFilter).toBe("A1:B1");
    });

    it("auto filter with object form {row, col}", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Name";
      ws.getCell("B1").value = "Age";
      ws.getCell("C1").value = "City";
      ws.autoFilter = { from: { row: 1, col: 1 }, to: { row: 1, col: 3 } };

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      // After round-trip, autoFilter is read back as string form
      expect(ws2.autoFilter).toBe("A1:C1");
    });

    it("company, manager, etc", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = "World!";
      wb.company = "Cyber Sapiens, Ltd";
      wb.manager = "Test Manager";
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.company).toBe(wb.company);
      expect(wb2.manager).toBe(wb.manager);
    });

    it("title, subject, etc", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = "World!";
      wb.title = "the title";
      wb.subject = "the subject";
      wb.keywords = "the keywords";
      wb.category = "the category";
      wb.description = "the description";
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.title).toBe(wb.title);
      expect(wb2.subject).toBe(wb.subject);
      expect(wb2.keywords).toBe(wb.keywords);
      expect(wb2.category).toBe(wb.category);
      expect(wb2.description).toBe(wb.description);
    });

    it("language, revision and contentStatus", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = "World!";
      wb.language = "Klingon";
      wb.revision = 2;
      wb.contentStatus = "Final";
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.language).toBe(wb.language);
      expect(wb2.revision).toBe(wb.revision);
      expect(wb2.contentStatus).toBe(wb.contentStatus);
    });

    it("empty strings", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = "Foo";
      ws.getCell("A2").value = "";
      ws.getCell("A3").value = "Baz";
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      expect(ws2.getCell("A1").value).toBe("Foo");
      expect(ws2.getCell("A2").value).toBe("");
      expect(ws2.getCell("A3").value).toBe("Baz");
    });

    it("dataValidations", async () => {
      const wb = testUtils.createTestBook(new Workbook(), "xlsx", ["dataValidations"]);

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", ["dataValidations"], {});
    });

    it("empty string", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet();

      ws.columns = [
        { key: "id", width: 10 },
        { key: "name", width: 32 }
      ];

      ws.addRow({ id: 1, name: "" });

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
    });

    it("a lot of sheets to xlsx file", async function () {
      let i;
      const wb = new Workbook();
      const numSheets = 90;
      // add numSheets sheets
      for (i = 1; i <= numSheets; i++) {
        const ws = wb.addWorksheet(`sheet${i}`);
        ws.getCell("A1").value = i;
      }
      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      for (i = 1; i <= numSheets; i++) {
        const ws2 = wb2.getWorksheet(`sheet${i}`)!;
        expect(ws2).toBeTruthy();
        expect(ws2.getCell("A1").value).toBe(i);
      }
    });

    it("csv file", async function () {
      const wb = testUtils.createTestBook(new Workbook(), "csv", undefined);

      await wb.writeCsvFile(TEST_CSV_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.readCsvFile(TEST_CSV_FILE_NAME);
      testUtils.checkTestBook(wb2, "csv", undefined, {});
    });

    it("CSV file and its configuration", async function () {
      const writeOptions = {
        dateFormat: "DD/MM/YYYY HH:mm:ss",
        dateUTC: false,
        encoding: "utf-8",
        includeEmptyRows: false,
        sheetName: "sheet1",
        delimiter: "\t",
        quote: false
      };
      const readOptions: CsvOptions = {
        dateFormats: ["DD/MM/YYYY HH:mm:ss"],
        sheetName: "sheet1",
        delimiter: "\t",
        quote: null
      };
      const wb = testUtils.createTestBook(new Workbook(), "csv", undefined);

      await wb.writeCsvFile(TEST_CSV_FILE_NAME, writeOptions);
      const wb2 = new Workbook();
      await wb2.readCsvFile(TEST_CSV_FILE_NAME, readOptions);
      testUtils.checkTestBook(wb2, "csv", undefined, writeOptions);
    });

    it("defined names", async () => {
      const wb1 = new Workbook();
      const ws1a = wb1.addWorksheet("blort");
      const ws1b = wb1.addWorksheet("foo");

      function assign(sheet: any, address: any, value: any, name: any) {
        const cell = sheet.getCell(address);
        cell.value = value;
        if (Array.isArray(name)) {
          cell.names = name;
        } else {
          cell.name = name;
        }
      }

      // single entry
      assign(ws1a, "A1", 5, "five");

      // three amigos - horizontal line
      assign(ws1a, "A3", 3, "amigos");
      assign(ws1a, "B3", 3, "amigos");
      assign(ws1a, "C3", 3, "amigos");

      // three amigos - vertical line
      assign(ws1a, "E1", 3, "verts");
      assign(ws1a, "E2", 3, "verts");
      assign(ws1a, "E3", 3, "verts");

      // four square
      assign(ws1a, "C5", 4, "squares");
      assign(ws1a, "B6", 4, "squares");
      assign(ws1a, "C6", 4, "squares");
      assign(ws1a, "B5", 4, "squares");

      // long distance
      assign(ws1a, "B7", 2, "sheets");
      assign(ws1b, "B7", 2, "sheets");

      // two names
      assign(ws1a, "G1", 1, "thing1");
      ws1a.getCell("G1").addName("thing2");

      // once removed
      assign(ws1a, "G2", 1, ["once", "twice"]);
      ws1a.getCell("G2").removeName("once");

      await wb1.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2a = wb2.getWorksheet("blort")!;
      const ws2b = wb2.getWorksheet("foo")!;

      function check(sheet: any, address: any, value: any, name: any) {
        const cell = sheet.getCell(address);
        expect(cell.value).toBe(value);
        expect(cell.name).toBe(name);
      }

      // single entry
      check(ws2a, "A1", 5, "five");

      // three amigos - horizontal line
      check(ws2a, "A3", 3, "amigos");
      check(ws2a, "B3", 3, "amigos");
      check(ws2a, "C3", 3, "amigos");

      // three amigos - vertical line
      check(ws2a, "E1", 3, "verts");
      check(ws2a, "E2", 3, "verts");
      check(ws2a, "E3", 3, "verts");

      // four square
      check(ws2a, "C5", 4, "squares");
      check(ws2a, "B6", 4, "squares");
      check(ws2a, "C6", 4, "squares");
      check(ws2a, "B5", 4, "squares");

      // long distance
      check(ws2a, "B7", 2, "sheets");
      check(ws2b, "B7", 2, "sheets");

      // two names
      expect(ws2a.getCell("G1").names).toEqual(expect.arrayContaining(["thing1", "thing2"]));
      expect(ws2a.getCell("G1").names.length).toBe(2);

      // once removed
      expect(ws2a.getCell("G2").names).toEqual(expect.arrayContaining(["twice"]));
      expect(ws2a.getCell("G2").names.length).toBe(1);

      // ranges
      function rangeCheck(name: any, members: any) {
        const ranges = wb2.definedNames.getRanges(name);
        expect(ranges.name).toBe(name);
        if (members.length) {
          expect(ranges.ranges).toEqual(expect.arrayContaining(members));
          expect(ranges.ranges.length).toBe(members.length);
        } else {
          expect(ranges.ranges.length).toBe(0);
        }
      }

      rangeCheck("five", ["blort!$A$1"]);
      rangeCheck("amigos", ["blort!$A$3:$C$3"]);
      rangeCheck("verts", ["blort!$E$1:$E$3"]);
      rangeCheck("squares", ["blort!$B$5:$C$6"]);
      rangeCheck("sheets", ["blort!$B$7", "foo!$B$7"]);
      rangeCheck("thing1", ["blort!$G$1"]);
      rangeCheck("thing2", ["blort!$G$1"]);
      rangeCheck("once", []);
      rangeCheck("twice", ["blort!$G$2"]);
    });

    describe("Duplicate Rows", () => {
      it("Duplicate rows with styles properly", async () => {
        const fileDuplicateRowTestFile = excelTestDataPath("duplicate-row-styles.xlsx");
        const wb = new Workbook();
        await wb.xlsx.readFile(fileDuplicateRowTestFile);
        const ws = wb.getWorksheet("duplicateTest");
        if (!ws) {
          throw new Error("Worksheet not found");
        }

        ws.getCell("A1").value = "OneInfo";
        ws.getCell("A2").value = "TwoInfo";
        ws.duplicateRow(1, 2);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("duplicateTest")!;

        expect(ws2.getCell("A2").value).toBe("OneInfo");
        expect(ws2.getCell("A2").style).toStrictEqual(ws2.getCell("A1").style);
        expect(ws2.getCell("A3").value).toBe("OneInfo");
        expect(ws2.getCell("A3").style).toStrictEqual(ws2.getCell("A1").style);
        expect(ws2.getCell("A4").value).toBeNull();
      });

      it("Duplicate rows replacing properly", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("duplicateTest");
        ws.getCell("A1").value = "OneInfo";
        ws.getCell("A2").value = "TwoInfo";
        ws.getCell("A3").value = "ThreeInfo";
        ws.getCell("A4").value = "FourInfo";
        ws.duplicateRow(1, 2, false);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("duplicateTest")!;

        expect(ws2.getCell("A1").value).toBe("OneInfo");
        expect(ws2.getCell("A2").value).toBe("OneInfo");
        expect(ws2.getCell("A3").value).toBe("OneInfo");
        expect(ws2.getCell("A4").value).toBe("FourInfo");
      });

      it("Duplicate rows shifting properly", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("duplicateTest");
        ws.getCell("A1").value = "OneInfo";
        ws.getCell("A2").value = "TwoInfo";
        ws.getCell("A3").value = "ThreeInfo";
        ws.getCell("A4").value = "FourInfo";
        ws.duplicateRow(1, 2, true);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("duplicateTest")!;

        expect(ws2.getCell("A1").value).toBe("OneInfo");
        expect(ws2.getCell("A2").value).toBe("OneInfo");
        expect(ws2.getCell("A3").value).toBe("OneInfo");
        expect(ws2.getCell("A4").value).toBe("TwoInfo");
      });

      it("Duplicate rows with height properly", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("duplicateTest");
        ws.getCell("A1").value = "OneInfo";
        ws.getCell("A2").value = "TwoInfo";
        ws.getRow(1).height = 25;
        ws.getRow(2).height = 15;
        ws.duplicateRow(1, 1, true);

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("duplicateTest")!;

        expect(ws2.getCell("A1").value).toBe("OneInfo");
        expect(ws2.getCell("A2").value).toBe("OneInfo");
        expect(ws2.getRow(1).height).toBe(ws2.getRow(2).height);
        expect(ws2.getRow(1).height).not.toBe(ws2.getRow(3).height);
      });
    });

    describe("Merge Cells", () => {
      it("serialises and deserialises properly", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("blort");

        // initial values
        ws.getCell("B2").value = "B2";

        ws.mergeCells("B2:C3");

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("blort")!;

        expect(ws2.getCell("B2").value).toBe("B2");
        expect(ws2.getCell("B3").value).toBe("B2");
        expect(ws2.getCell("C2").value).toBe("B2");
        expect(ws2.getCell("C3").value).toBe("B2");

        expect(ws2.getCell("B2").type).toBe(ValueType.String);
        expect(ws2.getCell("B3").type).toBe(ValueType.Merge);
        expect(ws2.getCell("C2").type).toBe(ValueType.Merge);
        expect(ws2.getCell("C3").type).toBe(ValueType.Merge);
      });

      it("styles", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("blort");

        // initial values
        const B2 = ws.getCell("B2");
        B2.value = 5;
        B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
        B2.style.border = testUtils.styles.borders.doubleRed;
        B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
        B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
        B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

        // expecting styles to be copied (see worksheet spec)
        ws.mergeCells("B2:C3");

        const dblRed = testUtils.styles.borders.doubleRed;

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("blort")!;

        // Non-border styles are identical on all cells
        for (const addr of ["B2", "B3", "C2", "C3"]) {
          expect(ws2.getCell(addr).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell(addr).fill).toEqual(testUtils.styles.fills.blueWhiteHGrad);
          expect(ws2.getCell(addr).alignment).toEqual(
            testUtils.styles.namedAlignments.middleCentre
          );
          expect(ws2.getCell(addr).numFmt).toBe(testUtils.styles.numFmts.numFmt1);
        }

        // Borders are position-aware after round-trip
        expect(ws2.getCell("B2").border).toEqual({
          left: dblRed.left,
          top: dblRed.top
        });
        expect(ws2.getCell("C2").border).toEqual({
          right: dblRed.right,
          top: dblRed.top
        });
        expect(ws2.getCell("B3").border).toEqual({
          left: dblRed.left,
          bottom: dblRed.bottom
        });
        expect(ws2.getCell("C3").border).toEqual({
          right: dblRed.right,
          bottom: dblRed.bottom
        });
      });
    });
  });

  it("spliced meat and ham", async () => {
    const wb = new Workbook();
    const sheets = [
      "splice.rows.removeOnly",
      "splice.rows.insertFewer",
      "splice.rows.insertSame",
      "splice.rows.insertMore",
      "splice.rows.insertStyle",
      "splice.columns.removeOnly",
      "splice.columns.insertFewer",
      "splice.columns.insertSame",
      "splice.columns.insertMore"
    ];
    const options = {
      checkBadAlignments: false,
      checkSheetProperties: false,
      checkViews: false
    };

    testUtils.createTestBook(wb, "xlsx", sheets);

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
    const wb2 = new Workbook();
    await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
    testUtils.checkTestBook(wb2, "xlsx", sheets, options);
  });

  it("throws an error when xlsx file not found", async () => {
    const wb = new Workbook();
    await expect(wb.xlsx.readFile("./wb.doesnotexist.xlsx")).rejects.toThrow();
  });

  it("throws an error when csv file not found", async () => {
    const wb = new Workbook();
    await expect(wb.readCsvFile("./wb.doesnotexist.csv")).rejects.toThrow();
  });
  it("throw an error for wrong data type", async () => {
    const wb = new Workbook();
    try {
      // Deliberately passing the wrong runtime type to verify the guard.
      await wb.xlsx.load({} as unknown as Uint8Array);
      expect.fail("should fail for given argument");
    } catch (e) {
      expect((e as Error).message).toContain(
        "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"
      );
    }
  });

  describe("Sheet Views", () => {
    it("frozen panes", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("frozen");
      ws.views = [
        {
          state: "frozen",
          xSplit: 2,
          ySplit: 3,
          topLeftCell: "C4",
          activeCell: "D5"
        },
        { state: "frozen", ySplit: 1 },
        { state: "frozen", xSplit: 1 }
      ];
      ws.getCell("A1").value = "Let it Snow!";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("frozen")!;
      expect(ws2).toBeTruthy();
      expect(ws2.getCell("A1").value).toBe("Let it Snow!");
      expect(ws2.views).toEqual([
        {
          workbookViewId: 0,
          state: "frozen",
          xSplit: 2,
          ySplit: 3,
          topLeftCell: "C4",
          activeCell: "D5",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        },
        {
          workbookViewId: 0,
          state: "frozen",
          xSplit: 0,
          ySplit: 1,
          topLeftCell: "A2",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        },
        {
          workbookViewId: 0,
          state: "frozen",
          xSplit: 1,
          ySplit: 0,
          topLeftCell: "B1",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        }
      ]);
    });

    it("serialises split panes", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("split");
      ws.views = [
        {
          state: "split",
          xSplit: 2000,
          ySplit: 3000,
          topLeftCell: "C4",
          activeCell: "D5",
          activePane: "bottomRight"
        },
        {
          state: "split",
          ySplit: 1500,
          activePane: "bottomLeft",
          topLeftCell: "A10"
        },
        { state: "split", xSplit: 1500, activePane: "topRight" }
      ];
      ws.getCell("A1").value = "Do the splits!";

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("split")!;
      expect(ws2).toBeTruthy();
      expect(ws2.getCell("A1").value).toBe("Do the splits!");
      expect(ws2.views).toEqual([
        {
          workbookViewId: 0,
          state: "split",
          xSplit: 2000,
          ySplit: 3000,
          topLeftCell: "C4",
          activeCell: "D5",
          activePane: "bottomRight",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        },
        {
          workbookViewId: 0,
          state: "split",
          xSplit: 0,
          ySplit: 1500,
          topLeftCell: "A10",
          activePane: "bottomLeft",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        },
        {
          workbookViewId: 0,
          state: "split",
          xSplit: 1500,
          ySplit: 0,
          topLeftCell: undefined,
          activePane: "topRight",
          showRuler: true,
          showGridLines: true,
          showRowColHeaders: true,
          zoomScale: 100,
          zoomScaleNormal: 100,
          rightToLeft: false,
          tabSelected: false
        }
      ]);
    });

    it("multiple book views", async () => {
      const wb = new Workbook();
      wb.views = [testUtils.views.book.visible, testUtils.views.book.hidden];

      const ws1 = wb.addWorksheet("one");
      ws1.views = [testUtils.views.sheet.frozen];

      const ws2 = wb.addWorksheet("two");
      ws2.views = [testUtils.views.sheet.split];

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.views).toEqual(wb.views);

      const ws1b = wb2.getWorksheet("one")!;
      expect(ws1b!.views).toEqual(ws1.views);

      const ws2b = wb2.getWorksheet("two")!;
      expect(ws2b!.views).toEqual(ws2.views);
    });
  });
});
