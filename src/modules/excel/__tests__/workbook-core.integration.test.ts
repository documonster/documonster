import { describe, it, expect } from "vitest";
import fs from "fs";
import { testUtils } from "@excel/__tests__/shared";
import { Workbook, type CsvOptions } from "../../../index";
import { ValueType } from "@excel/enums";
import { makeTestDataPath, testFilePath } from "@test/utils";
import { extractAll } from "@archive/unzip/extract";

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
        expect(definedNamesModel.length).toBeLessThan(1000);
      }, 60000);

      it("loads file from buffer without excessive memory use", async () => {
        const sourceFile = excelTestDataPath("many-defined-names.xlsx");
        const buffer = await fs.promises.readFile(sourceFile);

        const wb = new Workbook();
        await wb.xlsx.load(buffer);
        expect(wb.worksheets.length).toBeGreaterThan(0);
      }, 60000);

      it("filters out array constants from definedNames", async () => {
        const sourceFile = excelTestDataPath("many-defined-names.xlsx");

        const wb = new Workbook();
        await wb.xlsx.readFile(sourceFile);

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

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.pageSetup.printArea).toBe("A1:A1");
    });

    it("single-column printTitlesColumn without colon round-trips correctly", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "test";
      ws.pageSetup.printTitlesColumn = "A";

      const buffer = await wb.xlsx.writeBuffer();

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
      await wb.xlsx.load({});
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
