import { describe, it, expect } from "vitest";
import fs from "fs";
import { promisify } from "util";
import { testUtils } from "@excel/__tests__/shared";
import { Workbook, WorkbookWriter } from "../../../../index";
import type { NoteConfig } from "@excel/cell";
import { makeTestDataPath, testFilePath } from "@test/utils";

const streamTestDataPath = makeTestDataPath(import.meta.url, "./data");

const TEST_XLSX_FILE_NAME = testFilePath("wb-xlsx-writer.test");
const IMAGE_FILENAME = streamTestDataPath("image.png");
const fsReadFileAsync = promisify(fs.readFile);

describe("WorkbookWriter", () => {
  it("creates sheets with correct names", () => {
    const wb = new WorkbookWriter();
    const ws1 = wb.addWorksheet("Hello, World!");
    expect(ws1.name).toBe("Hello, World!");

    const ws2 = wb.addWorksheet();
    expect(ws2.name).toMatch(/sheet\d+/);
  });

  describe("Serialise", () => {
    it("xlsx file", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true
      };
      const wb = testUtils.createTestBook(new WorkbookWriter(options), "xlsx");

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx");
    });

    it("hyperlink with query arguments corrupts workbook", async () => {
      const filename = testFilePath("hyperlink-query-args.workbook-writer");
      const wb = new WorkbookWriter({
        filename,
        useStyles: true
      });
      const ws = wb.addWorksheet("Sheet1");

      const hyperlink = {
        text: "Somewhere with query params",
        hyperlink: 'www.somewhere.com?a=1&b=2&c=<>&d="\'"'
      };

      ws.getCell("A1").value = hyperlink;
      ws.commit();

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filename);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2.getCell("A1").value).toEqual(hyperlink);
    });

    it("special cell values produce a valid file", async () => {
      const filename = testFilePath("special-object-keys.workbook-writer");
      const wb = new WorkbookWriter({
        filename,
        useStyles: true,
        useSharedStrings: true
      });
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

      for (let i = 0, len = specialValues.length; i < len; i++) {
        const value = specialValues[i];
        ws.addRow([value]);
        ws.getCell(`B${i + 1}`).value = value;
      }

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(filename);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      for (let i = 0, len = specialValues.length; i < len; i++) {
        const value = specialValues[i];
        expect(ws2.getCell(`A${i + 1}`).value).toBe(value);
        expect(ws2.getCell(`B${i + 1}`).value).toBe(value);
      }
    });

    it("shared formula", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = {
        formula: "ROW()+COLUMN()",
        ref: "A1:B2",
        result: 2
      };
      ws.getCell("B1").value = { sharedFormula: "A1", result: 3 };
      ws.getCell("A2").value = { sharedFormula: "A1", result: 3 };
      ws.getCell("B2").value = { sharedFormula: "A1", result: 4 };

      ws.commit();
      await wb.commit();
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
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("B2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("B3").value = 3;

      ws.autoFilter = "A1:B1";
      ws.commit();

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2.autoFilter).toBe("A1:B1");
    });

    it("Without styles", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: false
      };
      const wb = testUtils.createTestBook(new WorkbookWriter(options), "xlsx");

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", undefined, {
        checkStyles: false
      });
    });

    it("serializes row styles and columns properly", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("blort");

      const colStyle = {
        font: testUtils.styles.fonts.comicSansUdB16,
        alignment: testUtils.styles.namedAlignments.middleCentre
      };
      ws.columns = [
        { header: "A1", width: 10 },
        { header: "B1", style: colStyle },
        { header: "C1", width: 30 },
        { header: "D1" }
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";
      ws.getCell("C2").value = "C2";
      ws.getCell("A3").value = "A3";
      ws.getCell("B3").value = "B3";
      ws.getCell("C3").value = "C3";

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("blort")!;
      ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"].forEach(address => {
        expect(ws2.getCell(address).value).toBe(address);
      });
      expect(ws2.getCell("B1").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getCell("B1").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws2.getCell("A2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("B2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("C2").font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell("B3").font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getCell("B3").alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);

      expect(ws2.getColumn(2).font).toEqual(testUtils.styles.fonts.comicSansUdB16);
      expect(ws2.getColumn(2).alignment).toEqual(testUtils.styles.namedAlignments.middleCentre);
      expect(ws2.getColumn(2).width).toBe(9);

      expect(ws2.getColumn(4).width).toBe(undefined);

      expect(ws2.getRow(2).font).toEqual(testUtils.styles.fonts.broadwayRedOutline20);
    });

    it("rich text", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      ws.getCell("A1").value = {
        richText: [
          {
            font: { color: { argb: "FF0000" } },
            text: "red "
          },
          {
            font: { color: { argb: "00FF00" }, bold: true },
            text: " bold green"
          }
        ]
      };

      ws.getCell("B1").value = "plain text";

      ws.commit();
      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2.getCell("A1").value).toEqual({
        richText: [
          {
            font: { color: { argb: "FF0000" } },
            text: "red "
          },
          {
            font: { color: { argb: "00FF00" }, bold: true },
            text: " bold green"
          }
        ]
      });
      expect(ws2.getCell("B1").value).toBe("plain text");
    });

    it("A lot of sheets", async () => {
      const wb = new WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME
      });
      const numSheets = 90;
      for (let i = 1; i <= numSheets; i++) {
        const ws = wb.addWorksheet(`sheet${i}`);
        ws.getCell("A1").value = i;
      }
      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      for (let i = 1; i <= numSheets; i++) {
        const ws2 = wb2.getWorksheet(`sheet${i}`)!;
        expect(ws2).toBeTruthy();
        expect(ws2.getCell("A1").value).toBe(i);
      }
    });

    it("addRow", async () => {
      const options = {
        stream: fs.createWriteStream(TEST_XLSX_FILE_NAME, { flags: "w" }),
        useStyles: true,
        useSharedStrings: true
      };
      const workbook = new WorkbookWriter(options);
      const worksheet = workbook.addWorksheet("test");
      const newRow = worksheet.addRow(["hello"]);
      newRow.commit();
      worksheet.commit();
      await workbook.commit();

      // Verify the written file is a valid XLSX
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.getWorksheet("test")!.getCell("A1").value).toBe("hello");
    });

    it("defined names", async () => {
      const wb = new WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME
      });
      const ws = wb.addWorksheet("blort");
      ws.getCell("A1").value = 5;
      ws.getCell("A1").name = "five";

      ws.getCell("A3").value = "drei";
      ws.getCell("A3").name = "threes";
      ws.getCell("B3").value = "trois";
      ws.getCell("B3").name = "threes";
      ws.getCell("B3").value = "san";
      ws.getCell("B3").name = "threes";

      ws.getCell("E1").value = "grün";
      ws.getCell("E1").name = "greens";
      ws.getCell("E2").value = "vert";
      ws.getCell("E2").name = "greens";
      ws.getCell("E3").value = "verde";
      ws.getCell("E3").name = "greens";

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("blort")!;
      expect(ws2.getCell("A1").name).toBe("five");

      expect(ws2.getCell("A3").name).toBe("threes");
      expect(ws2.getCell("B3").name).toBe("threes");
      expect(ws2.getCell("B3").name).toBe("threes");

      expect(ws2.getCell("E1").name).toBe("greens");
      expect(ws2.getCell("E2").name).toBe("greens");
      expect(ws2.getCell("E3").name).toBe("greens");
    });

    it("does not escape special xml characters", async () => {
      const wb = new WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
        useSharedStrings: true
      });
      const ws = wb.addWorksheet("blort");
      const xmlCharacters = 'xml characters: & < > "';

      ws.getCell("A1").value = xmlCharacters;

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("blort")!;
      expect(ws2.getCell("A1").value).toBe(xmlCharacters);
    });

    it("serializes and deserializes dataValidations", async () => {
      const options = { filename: TEST_XLSX_FILE_NAME };
      const wb = testUtils.createTestBook(new WorkbookWriter(options), "xlsx", ["dataValidations"]);

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", ["dataValidations"]);
    });

    it("with zip compression option", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
        zip: {
          zlib: { level: 9 }
        }
      };
      const wb = testUtils.createTestBook(new WorkbookWriter(options), "xlsx", ["dataValidations"]);

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", ["dataValidations"]);
    });

    it("writes notes", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");
      ws.getCell("B2").value = 5;
      ws.getCell("B2").note = "five";

      const note: NoteConfig = {
        texts: [
          {
            font: {
              size: 12,
              color: { argb: "FFFF6600" },
              name: "Calibri",
              scheme: "minor"
            },
            text: "seven"
          }
        ],
        margins: {
          insetmode: "auto",
          inset: [0.13, 0.13, 0.25, 0.25]
        },
        protection: {
          locked: "True",
          lockText: "True"
        },
        editAs: "twoCells"
      };
      ws.getCell("D2").value = 7;
      ws.getCell("D2").note = note;

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      expect(ws2.getCell("B2").value).toBe(5);
      expect(ws2.getCell("B2").note).toBe("five");
      expect(ws2.getCell("D2").value).toBe(7);
      const note2 = ws2.getCell("D2").note as typeof note;
      expect(note2.texts).toEqual(note.texts);
      expect(note2.margins).toEqual(note.margins);
      expect(note2.protection).toEqual(note.protection);
      expect(note2.editAs).toEqual(note.editAs);
    });

    it("Cell annotation supports setting margins and protection properties", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");
      ws.getCell("B2").value = 5;
      ws.getCell("B2").note = "five";
      const note: NoteConfig = {
        texts: [
          {
            font: {
              size: 12,
              color: { argb: "FFFF6600" },
              name: "Calibri",
              scheme: "minor"
            },
            text: "seven"
          }
        ],
        margins: {
          insetmode: "custom",
          inset: [0.25, 0.25, 0.35, 0.35]
        },
        protection: {
          locked: "False",
          lockText: "False"
        },
        editAs: "oneCells"
      };
      ws.getCell("D2").value = 7;
      ws.getCell("D2").note = note;

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2.getCell("B2").value).toBe(5);
      expect(ws2.getCell("B2").note).toBe("five");

      expect(ws2.getCell("D2").value).toBe(7);
      const note2 = ws2.getCell("D2").note as typeof note;
      expect(note2.texts).toEqual(note.texts);
      expect(note2.margins).toEqual(note.margins);
      expect(note2.protection).toEqual(note.protection);
      expect(note2.editAs).toEqual(note.editAs);
    });

    it("with background image", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });
      ws.getCell("A1").value = "Hello, World!";
      ws.addBackgroundImage(imageId);

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const backgroundId2 = ws2.getBackgroundImageId();
      const image = wb2.getImage(backgroundId2!);
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("with background image where worksheet is commited in advance", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });
      ws.getCell("A1").value = "Hello, World!";
      ws.addBackgroundImage(imageId);

      ws.commit();
      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const backgroundId2 = ws2.getBackgroundImageId();
      const image = wb2.getImage(backgroundId2!);
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("with conditional formatting", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
        useSharedStrings: true
      };
      const wb = testUtils.createTestBook(new WorkbookWriter(options), "xlsx", [
        "conditionalFormatting"
      ]);

      await wb.commit();
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      testUtils.checkTestBook(wb2, "xlsx", ["conditionalFormatting"]);
    });

    it("with conditional formatting that contains numFmt (#1814)", async () => {
      const sheet = "conditionalFormatting";
      const options = { filename: TEST_XLSX_FILE_NAME, useStyles: true };

      // generate file with conditional formatting that contains styles with numFmt
      const wb1 = new WorkbookWriter(options);
      const ws1 = wb1.addWorksheet(sheet);
      const cf1 = testUtils.conditionalFormatting.abbreviation;
      ws1.addConditionalFormatting(cf1);
      await wb1.commit();

      // read generated file and extract saved conditional formatting rule
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet(sheet)!;
      const [cf2] = ws2.conditionalFormattings;

      // verify that rules from generated file contain styles with valid numFmt
      expect(cf2.rules.length).toBeGreaterThan(0);
      cf2.rules.forEach(rule => {
        const numFmt = rule.style?.numFmt;
        expect(numFmt).toBeDefined();
        // After reading from file, numFmt is always a NumFmt object (not string)
        expect(typeof numFmt).toBe("object");
        if (typeof numFmt === "object" && numFmt !== null) {
          expect(numFmt.id).toBeTypeOf("number");
          expect(numFmt.formatCode).toBeTypeOf("string");
        }
      });
    });
  });

  // ==========================================================================
  // Regression tests for Issue #88 (memory leak) and Issue #89 (RangeError)
  // ==========================================================================

  describe("Streaming memory behavior", () => {
    /** Write to a WorkbookWriter that streams to an in-memory buffer, then read it back. */
    async function writeAndReadBack(
      options: Record<string, any>,
      populate: (ws: any) => void
    ): Promise<any> {
      const { PassThrough } = await import("@stream");
      const output = new PassThrough();
      const chunks: Uint8Array[] = [];
      output.on("data", (chunk: Uint8Array) => chunks.push(chunk));

      const workbook = new WorkbookWriter({
        stream: output,
        useSharedStrings: false,
        ...options
      });
      const worksheet = workbook.addWorksheet("Sheet 1");
      populate(worksheet);
      await workbook.commit();

      const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
      const xlsxBuffer = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        xlsxBuffer.set(chunk, offset);
        offset += chunk.length;
      }

      const readBack = new Workbook();
      await readBack.xlsx.load(xlsxBuffer);
      return readBack.getWorksheet("Sheet 1");
    }

    it("does not accumulate worksheet data in memory with trueStreaming (#88)", async () => {
      const cellValue = "abcdefghij".repeat(40); // 400 chars per cell
      const ws = await writeAndReadBack({ trueStreaming: true }, worksheet => {
        for (let i = 0; i < 5000; i++) {
          const row = worksheet.getRow(i + 1);
          for (let c = 1; c <= 9; c++) {
            row.getCell(c).value = cellValue;
          }
          row.commit();
        }
      });

      expect(ws.rowCount).toBe(5000);
      expect(ws.getCell("A1").value).toBe(cellValue);
    }, 30000);

    it("does not accumulate worksheet data in memory with default streaming (#88)", async () => {
      const ws = await writeAndReadBack({}, worksheet => {
        for (let i = 0; i < 5000; i++) {
          const row = worksheet.getRow(i + 1);
          for (let c = 1; c <= 9; c++) {
            row.getCell(c).value = "abcdefghij".repeat(40);
          }
          row.commit();
        }
      });

      expect(ws.rowCount).toBe(5000);
    }, 30000);

    it("handles very large cell values in trueStreaming mode without crashing (#89)", async () => {
      const largeCellValue = "x".repeat(36_000); // ~36KB per cell × 9 cells = ~324KB per row
      const ws = await writeAndReadBack({ trueStreaming: true }, worksheet => {
        for (let i = 0; i < 100; i++) {
          const row = worksheet.getRow(i + 1);
          for (let c = 1; c <= 9; c++) {
            row.getCell(c).value = largeCellValue;
          }
          row.commit();
        }
      });

      expect(ws.rowCount).toBe(100);
      expect(ws.getCell("A1").value).toBe(largeCellValue);
    }, 30000);

    it("maintains constant memory during streaming with large row count (#88)", async () => {
      const { PassThrough } = await import("@stream");
      const output = new PassThrough();
      const chunks: Uint8Array[] = [];
      output.on("data", (chunk: Uint8Array) => chunks.push(chunk));

      const workbook = new WorkbookWriter({
        stream: output,
        useSharedStrings: false,
        trueStreaming: true
      });
      const worksheet = workbook.addWorksheet("Sheet 1");

      const bigValue = "abcdefghij".repeat(40); // 400 chars per cell

      // Warm up — write initial rows to stabilize JIT, GC, and internal structures
      for (let i = 0; i < 2000; i++) {
        const row = worksheet.getRow(i + 1);
        for (let c = 1; c <= 9; c++) {
          row.getCell(c).value = bigValue;
        }
        row.commit();
      }
      if (global.gc) {
        global.gc();
      }
      const baselineRSS = process.memoryUsage().rss;

      // Steady-state — write many more rows
      for (let i = 2000; i < 20_000; i++) {
        const row = worksheet.getRow(i + 1);
        for (let c = 1; c <= 9; c++) {
          row.getCell(c).value = bigValue;
        }
        row.commit();
      }
      if (global.gc) {
        global.gc();
      }
      const finalRSS = process.memoryUsage().rss;
      const growthMB = (finalRSS - baselineRSS) / 1024 / 1024;

      await workbook.commit();

      // Verify the file is valid by reading it back
      const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
      expect(totalLength).toBeGreaterThan(0);

      const xlsxBuffer = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        xlsxBuffer.set(chunk, offset);
        offset += chunk.length;
      }

      const readBack = new Workbook();
      await readBack.xlsx.load(xlsxBuffer);
      const ws = readBack.getWorksheet("Sheet 1");
      expect(ws!.rowCount).toBe(20_000);
      expect(ws!.getCell("A1").value).toBe(bigValue);
      expect(ws!.getCell("I20000").value).toBe(bigValue);

      // Memory assertion: 18K rows × 9 cells × 400 chars should NOT cause
      // significant memory growth when streaming. Before the fix, this would
      // accumulate ~300MB+ in push chain closures. With the fix, growth should
      // be minimal (< 100MB allows for GC timing, V8 overhead, and output buffer).
      expect(growthMB).toBeLessThan(100);
    }, 30000);

    it("correctly handles small worksheets via pushSync smart-store path (#88)", async () => {
      // Small worksheets (< 16KB total) exercise the smart-store sampling path
      // in pushSync where data is buffered in _pendingChunks before compression
      // mode is decided. This verifies the sampling → decision → flush cycle.
      const ws = await writeAndReadBack({ trueStreaming: true }, worksheet => {
        for (let i = 0; i < 5; i++) {
          const row = worksheet.getRow(i + 1);
          row.getCell(1).value = "tiny";
          row.commit();
        }
      });

      expect(ws.rowCount).toBe(5);
      expect(ws.getCell("A1").value).toBe("tiny");
      expect(ws.getCell("A5").value).toBe("tiny");
    });
  });

  // ==========================================================================
  // WorksheetWriter.addImage tests (Issue #108)
  // ==========================================================================

  describe("WorksheetWriter.addImage", () => {
    it("stores embedded image with string range (two-cell anchor)", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.getCell("A1").value = "Hello, World!";
      ws.addImage(imageId, "C3:E6");

      await wb.commit();

      // Read back and verify
      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;
      expect(ws2).toBeDefined();

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      // String range "C3:E6" => tl: col 2, row 2 (with offset -1: nativeCol=2, nativeRow=2)
      expect(imageDesc.range!.tl.nativeCol).toBe(2);
      expect(imageDesc.range!.tl.nativeRow).toBe(2);
      expect(imageDesc.range!.br!.nativeCol).toBe(5);
      expect(imageDesc.range!.br!.nativeRow).toBe(6);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("stores embedded image with object range (two-cell anchor)", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        br: { col: 2.101046875, row: 3.4 },
        editAs: "oneCell"
      });

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("oneCell");

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("stores embedded image with one-cell anchor (ext)", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 },
        editAs: "oneCell"
      });

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("oneCell");
      expect(imageDesc.range!.ext!.width).toBe(100);
      expect(imageDesc.range!.ext!.height).toBe(100);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("stores embedded image with hyperlinks", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 },
        editAs: "absolute",
        hyperlinks: {
          hyperlink: "http://www.somewhere.com",
          tooltip: "www.somewhere.com"
        }
      });

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("absolute");
      expect(imageDesc.range!.hyperlinks).toEqual({
        hyperlink: "http://www.somewhere.com",
        tooltip: "www.somewhere.com"
      });

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("stores multiple images on the same worksheet", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId1 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });
      const imageId2 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.addImage(imageId1, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 }
      });
      ws.addImage(imageId2, {
        tl: { col: 0.1125, row: 0.4 },
        br: { col: 2.101046875, row: 3.4 },
        editAs: "oneCell"
      });

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(2);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image1 = wb2.getImage(images[0].imageId!);
      const image2 = wb2.getImage(images[1].imageId!);
      expect(Buffer.compare(imageData, image1!.buffer!)).toBe(0);
      expect(Buffer.compare(imageData, image2!.buffer!)).toBe(0);
    });

    it("stores images on multiple worksheets", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      const ws1 = wb.addWorksheet("Sheet1");
      ws1.getCell("A1").value = "Sheet 1";
      ws1.addImage(imageId, "A1:B2");

      const ws2 = wb.addWorksheet("Sheet2");
      ws2.getCell("A1").value = "Sheet 2";
      ws2.addImage(imageId, "C3:D4");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2Sheet1 = wb2.getWorksheet("Sheet1")!;
      const ws2Sheet2 = wb2.getWorksheet("Sheet2")!;

      expect(ws2Sheet1.getImages().length).toBe(1);
      expect(ws2Sheet2.getImages().length).toBe(1);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const img1 = wb2.getImage(ws2Sheet1.getImages()[0].imageId!);
      const img2 = wb2.getImage(ws2Sheet2.getImages()[0].imageId!);
      expect(Buffer.compare(imageData, img1!.buffer!)).toBe(0);
      expect(Buffer.compare(imageData, img2!.buffer!)).toBe(0);
    });

    it("works when worksheet is committed before workbook", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.getCell("A1").value = "Hello, World!";
      ws.addImage(imageId, "C3:E6");
      ws.commit();

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works with both background and embedded images", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.getCell("A1").value = "Hello, World!";
      ws.addBackgroundImage(imageId);
      ws.addImage(imageId, "C3:E6");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      // Check background image
      const backgroundId = ws2.getBackgroundImageId();
      expect(backgroundId).toBeDefined();

      // Check embedded image
      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works with base64 image data", async () => {
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const base64 = imageData.toString("base64");

      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        base64,
        extension: "png"
      });

      ws.addImage(imageId, "A1:B2");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works with buffer image data", async () => {
      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        buffer: imageData,
        extension: "png"
      });

      ws.addImage(imageId, "A1:B2");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("deduplicates drawing rels for same imageId used twice", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      // Add the same image twice with different positions
      ws.addImage(imageId, "A1:B2");
      ws.addImage(imageId, "D4:E5");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(2);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      for (const img of images) {
        const imgBuffer = wb2.getImage(img.imageId!);
        expect(Buffer.compare(imageData, imgBuffer!.buffer!)).toBe(0);
      }
    });

    it("works on a worksheet with only images and no cell data", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Empty");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      // No cell data — only an image
      ws.addImage(imageId, "A1:B2");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Empty")!;
      expect(ws2).toBeDefined();

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works with native anchor coordinates (nativeCol/nativeRow/nativeColOff/nativeRowOff)", async () => {
      const options = {
        filename: TEST_XLSX_FILE_NAME
      };
      const wb = new WorkbookWriter(options);
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      ws.addImage(imageId, {
        tl: {
          nativeCol: 1,
          nativeRow: 2,
          nativeColOff: 100000,
          nativeRowOff: 50000
        } as any,
        br: {
          nativeCol: 3,
          nativeRow: 5,
          nativeColOff: 200000,
          nativeRowOff: 80000
        } as any,
        editAs: "twoCell"
      });

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.tl.nativeCol).toBe(1);
      expect(imageDesc.range!.tl.nativeRow).toBe(2);
      expect(imageDesc.range!.tl.nativeColOff).toBe(100000);
      expect(imageDesc.range!.tl.nativeRowOff).toBe(50000);
      expect(imageDesc.range!.br!.nativeCol).toBe(3);
      expect(imageDesc.range!.br!.nativeRow).toBe(5);
    });

    it("works via stream output (non-filename)", async () => {
      const { PassThrough } = await import("@stream");
      const output = new PassThrough();
      const chunks: Uint8Array[] = [];
      output.on("data", (chunk: Uint8Array) => chunks.push(chunk));

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const wb = new WorkbookWriter({ stream: output });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        buffer: imageData,
        extension: "png"
      });

      ws.getCell("A1").value = "Hello";
      ws.addImage(imageId, "C3:E6");

      await wb.commit();

      const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
      const xlsxBuffer = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        xlsxBuffer.set(chunk, offset);
        offset += chunk.length;
      }

      const wb2 = new Workbook();
      await wb2.xlsx.load(xlsxBuffer);
      const ws2 = wb2.getWorksheet("Hello")!;

      expect(ws2.getCell("A1").value).toBe("Hello");
      const images = ws2.getImages();
      expect(images.length).toBe(1);

      const image = wb2.getImage(images[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("getImages() returns added images before commit", () => {
      const wb = new WorkbookWriter({ filename: TEST_XLSX_FILE_NAME });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      expect(ws.getImages().length).toBe(0);

      ws.addImage(imageId, "A1:B2");
      expect(ws.getImages().length).toBe(1);
      expect(ws.getImages()[0].type).toBe("image");
      expect(ws.getImages()[0].imageId).toBe(String(imageId));

      ws.addImage(imageId, "C3:D4");
      expect(ws.getImages().length).toBe(2);
    });

    it("throws error for invalid single-cell string range", () => {
      const wb = new WorkbookWriter({ filename: TEST_XLSX_FILE_NAME });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      // "A1" is a single cell, not a range — should throw
      expect(() => ws.addImage(imageId, "A1")).toThrow('Invalid image range: "A1"');
    });

    it("addBackgroundImage accepts string imageId", async () => {
      const wb = new WorkbookWriter({ filename: TEST_XLSX_FILE_NAME });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.getCell("A1").value = "Hello";
      // Pass imageId as string (same as Worksheet API allows)
      ws.addBackgroundImage(String(imageId));

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      const backgroundId = ws2.getBackgroundImageId();
      expect(backgroundId).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(backgroundId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works when individual rows are committed before addImage (string range)", async () => {
      const wb = new WorkbookWriter({ filename: TEST_XLSX_FILE_NAME });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      // Commit rows individually first — this advances _rowZero
      const row1 = ws.getRow(1);
      row1.getCell(1).value = "data1";
      row1.commit();
      const row2 = ws.getRow(2);
      row2.getCell(1).value = "data2";
      row2.commit();

      // Should NOT throw RowOutOfBoundsError even though rows 1-2 are committed
      ws.addImage(imageId, "A1:C3");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      expect(ws2.getImages().length).toBe(1);

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);
      const image = wb2.getImage(ws2.getImages()[0].imageId!);
      expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
    });

    it("works when all rows are committed before addImage via addRow pattern", async () => {
      const wb = new WorkbookWriter({ filename: TEST_XLSX_FILE_NAME });
      const ws = wb.addWorksheet("Hello");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      // Typical streaming pattern: addRow auto-commits previous rows
      for (let i = 1; i <= 10; i++) {
        ws.addRow([`row ${i}`]).commit();
      }

      // Add image referencing rows that are all committed
      ws.addImage(imageId, "B2:D8");

      await wb.commit();

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet("Hello")!;

      expect(ws2.getImages().length).toBe(1);
      expect(ws2.getCell("A1").value).toBe("row 1");
      expect(ws2.getCell("A10").value).toBe("row 10");
    });
  });
});
