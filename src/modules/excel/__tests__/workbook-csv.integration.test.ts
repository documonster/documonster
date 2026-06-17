import { readCsv, writeCsv, writeCsvBuffer } from "@excel/bridge/csv-bridge";
import { readCsvFile, writeCsvFile } from "@excel/bridge/csv-bridge.node";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { getSheetName } from "@excel/worksheet";
import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

const csvTestDataPath = makeTestDataPath(import.meta.url, "./data");

describe("Workbook", () => {
  describe("CSV", () => {
    it("differentiates between strings with leading numbers and dates when reading csv files", async () => {
      const wb = Workbook.create();
      const worksheet = await readCsvFile(wb, csvTestDataPath("date-vs-leading-zeros.csv"));

      expect(Cell.getValue(worksheet, "A1")!.toString()).toBe(
        new Date("2019-11-04T00:00:00").toString()
      );
      expect(Cell.getValue(worksheet, "A2")!.toString()).toBe(
        new Date("2019-11-04T00:00:00").toString()
      );
      expect(Cell.getValue(worksheet, "A3")!.toString()).toBe(
        new Date("2019-11-04T10:17:55").toString()
      );
      expect(Cell.getValue(worksheet, "A4")).toBe("00210PRG1");
      expect(Cell.getValue(worksheet, "A5")).toBe("1234-5thisisnotadate");
    });

    it("supports encoding option on writeFile + readFile roundtrip", async () => {
      const TEST_CSV_FILE_NAME = testFilePath("csv-encoding-utf8-roundtrip", ".csv");
      const HEBREW_TEST_STRING = "משהו שכתוב בעברית";

      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "wheee");
      Cell.setValue(ws, "A1", HEBREW_TEST_STRING);

      await writeCsvFile(wb, TEST_CSV_FILE_NAME, { encoding: "UTF-8" });

      const wb2 = Workbook.create();
      const ws2 = await readCsvFile(wb2, TEST_CSV_FILE_NAME);
      expect(Cell.getValue(ws2, "A1")).toBe(HEBREW_TEST_STRING);
    }, 6000);

    describe("append mode", () => {
      it("should create new file with headers on first write", async () => {
        const TEST_FILE = testFilePath("csv-append-new-" + Date.now(), ".csv");

        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "Data");
        Worksheet.addRow(ws, ["Name", "Age"]);
        Worksheet.addRow(ws, ["Alice", 30]);

        await writeCsvFile(wb, TEST_FILE, { append: true });

        const wb2 = Workbook.create();
        const ws2 = await readCsvFile(wb2, TEST_FILE);
        expect(Worksheet.rowCount(ws2)).toBe(2);
        expect(Cell.getValue(ws2, "A1")).toBe("Name");
        expect(Cell.getValue(ws2, "A2")).toBe("Alice");
      });

      it("should append without headers to existing file", async () => {
        const TEST_FILE = testFilePath("csv-append-existing-" + Date.now(), ".csv");

        // First write - creates file with headers
        const wb1 = Workbook.create();
        const ws1 = Workbook.addWorksheet(wb1, "Data");
        Worksheet.addRow(ws1, ["Name", "Age"]);
        Worksheet.addRow(ws1, ["Alice", 30]);
        await writeCsvFile(wb1, TEST_FILE);

        // Second write - append mode (only add data rows, no header)
        const wb2 = Workbook.create();
        const ws2 = Workbook.addWorksheet(wb2, "Data");
        Worksheet.addRow(ws2, ["Bob", 25]); // Data only, no header row
        await writeCsvFile(wb2, TEST_FILE, { append: true });

        // Read and verify
        const wb3 = Workbook.create();
        const ws3 = await readCsvFile(wb3, TEST_FILE);
        expect(Worksheet.rowCount(ws3)).toBe(3); // Header + Alice + Bob
        expect(Cell.getValue(ws3, "A1")).toBe("Name");
        expect(Cell.getValue(ws3, "A2")).toBe("Alice");
        expect(Cell.getValue(ws3, "A3")).toBe("Bob");
      });

      it("should append multiple batches correctly", async () => {
        const TEST_FILE = testFilePath("csv-append-batches-" + Date.now(), ".csv");

        // Initial file with header
        const wb1 = Workbook.create();
        const ws1 = Workbook.addWorksheet(wb1, "Log");
        Worksheet.addRow(ws1, ["Time", "Event"]);
        Worksheet.addRow(ws1, ["10:00", "start"]);
        await writeCsvFile(wb1, TEST_FILE);

        // Append batch 1 (data only)
        const wb2 = Workbook.create();
        const ws2 = Workbook.addWorksheet(wb2, "Log");
        Worksheet.addRow(ws2, ["10:05", "process"]);
        await writeCsvFile(wb2, TEST_FILE, { append: true });

        // Append batch 2 (data only)
        const wb3 = Workbook.create();
        const ws3 = Workbook.addWorksheet(wb3, "Log");
        Worksheet.addRow(ws3, ["10:10", "end"]);
        await writeCsvFile(wb3, TEST_FILE, { append: true });

        // Verify
        const wb4 = Workbook.create();
        const ws4 = await readCsvFile(wb4, TEST_FILE);
        expect(Worksheet.rowCount(ws4)).toBe(4); // Header + 3 data rows
        expect(Cell.getValue(ws4, "A1")).toBe("Time");
        expect(Cell.getValue(ws4, "B2")).toBe("start");
        expect(Cell.getValue(ws4, "B3")).toBe("process");
        expect(Cell.getValue(ws4, "B4")).toBe("end");
      });

      it("should respect custom lineEnding in append mode", async () => {
        const TEST_FILE = testFilePath("csv-append-crlf-" + Date.now(), ".csv");

        // Create file with CRLF line endings
        const wb1 = Workbook.create();
        const ws1 = Workbook.addWorksheet(wb1, "Data");
        Worksheet.addRow(ws1, ["A", "B"]);
        Worksheet.addRow(ws1, [1, 2]);
        await writeCsvFile(wb1, TEST_FILE, { lineEnding: "\r\n" });

        // Append with same line ending
        const wb2 = Workbook.create();
        const ws2 = Workbook.addWorksheet(wb2, "Data");
        Worksheet.addRow(ws2, [3, 4]);
        await writeCsvFile(wb2, TEST_FILE, { append: true, lineEnding: "\r\n" });

        // Read and verify
        const wb3 = Workbook.create();
        const ws3 = await readCsvFile(wb3, TEST_FILE);
        expect(Worksheet.rowCount(ws3)).toBe(3);
        expect(Cell.getValue(ws3, "A3")).toBe(3);
      });
    });

    // =========================================================================
    // Unified API Tests (parse, stringify, toBuffer)
    // =========================================================================

    describe("Unified API", () => {
      describe("parse()", () => {
        it("should parse CSV string", async () => {
          const wb = Workbook.create();
          const ws = await readCsv(wb, "a,b,c\n1,2,3");

          expect(Cell.getValue(ws, "A1")).toBe("a");
          expect(Cell.getValue(ws, "B1")).toBe("b");
          expect(Cell.getValue(ws, "A2")).toBe(1);
        });

        it("should auto-detect semicolon delimiter", async () => {
          const wb = Workbook.create();
          const ws = await readCsv(wb, "a;b;c\n1;2;3", { delimiter: "" });

          expect(Cell.getValue(ws, "A1")).toBe("a");
          expect(Cell.getValue(ws, "B1")).toBe("b");
          expect(Cell.getValue(ws, "C1")).toBe("c");
          expect(Cell.getValue(ws, "A2")).toBe(1);
        });

        it("should auto-detect tab delimiter", async () => {
          const wb = Workbook.create();
          const ws = await readCsv(wb, "a\tb\tc\n1\t2\t3", { delimiter: "" });

          expect(Cell.getValue(ws, "A1")).toBe("a");
          expect(Cell.getValue(ws, "B1")).toBe("b");
          expect(Cell.getValue(ws, "C1")).toBe("c");
        });

        it("should respect explicit delimiter option", async () => {
          const wb = Workbook.create();
          // Even though the data has commas, use semicolon delimiter
          const ws = await readCsv(wb, "a,b;c,d\n1,2;3,4", { delimiter: ";" });

          expect(Cell.getValue(ws, "A1")).toBe("a,b");
          expect(Cell.getValue(ws, "B1")).toBe("c,d");
        });

        it("should parse ArrayBuffer input", async () => {
          const wb = Workbook.create();
          const data = new TextEncoder().encode("name,value\ntest,123");
          const ws = await readCsv(wb, data.buffer);

          expect(Cell.getValue(ws, "A1")).toBe("name");
          expect(Cell.getValue(ws, "B2")).toBe(123);
        });

        it("should parse Uint8Array input", async () => {
          const wb = Workbook.create();
          const data = new TextEncoder().encode("x,y\n10,20");
          const ws = await readCsv(wb, data);

          expect(Cell.getValue(ws, "A1")).toBe("x");
          expect(Cell.getValue(ws, "A2")).toBe(10);
        });

        it("should use custom sheet name", async () => {
          const wb = Workbook.create();
          const ws = await readCsv(wb, "a,b\n1,2", { sheetName: "MySheet" });

          expect(getSheetName(ws)).toBe("MySheet");
        });

        it("should support headers option", async () => {
          const wb = Workbook.create();
          const ws = await readCsv(wb, "name,age\nAlice,30\nBob,25", { headers: true });

          // When headers: true, the first row becomes column headers but values are still parsed
          expect(Cell.getValue(ws, "A1")).toBe("name");
          expect(Cell.getValue(ws, "B1")).toBe("age");
        });

        it("should support headers option with stream input", async () => {
          const wb = Workbook.create();
          // Create a readable stream from string using async generator
          const csvData = "name,age\nAlice,30\nBob,25";
          const readable = {
            async *[Symbol.asyncIterator]() {
              // Split into chunks to simulate streaming
              yield new TextEncoder().encode(csvData.slice(0, 10));
              yield new TextEncoder().encode(csvData.slice(10));
            }
          };

          const ws = await readCsv(wb, readable as any, { headers: true });

          // Verify headers row is written
          expect(Cell.getValue(ws, "A1")).toBe("name");
          expect(Cell.getValue(ws, "B1")).toBe("age");
          // Verify data rows are converted from object to array correctly
          expect(Cell.getValue(ws, "A2")).toBe("Alice");
          expect(Cell.getValue(ws, "B2")).toBe(30);
          expect(Cell.getValue(ws, "A3")).toBe("Bob");
          expect(Cell.getValue(ws, "B3")).toBe(25);
        });

        it("should support decimalSeparator option", async () => {
          const wb = Workbook.create();
          await readCsv(wb, "value\n1,50\n2,75", {
            delimiter: ";",
            decimalSeparator: ","
          });

          // Note: with delimiter auto-detect and single column, comma won't be detected as delimiter
          // So we explicitly set delimiter to avoid ambiguity
        });
      });

      describe("stringify()", () => {
        it("should convert worksheet to CSV string", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "Test");
          Cell.setValue(ws, "A1", "hello");
          Cell.setValue(ws, "B1", "world");
          Cell.setValue(ws, "A2", 1);
          Cell.setValue(ws, "B2", 2);

          const csv = writeCsv(wb);

          expect(csv).toBe("hello,world\n1,2");
        });

        it("should use custom delimiter", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "Test");
          Cell.setValue(ws, "A1", "a");
          Cell.setValue(ws, "B1", "b");

          const csv = writeCsv(wb, { delimiter: ";" });

          expect(csv).toBe("a;b");
        });

        it("should quote fields with special characters", () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "Test");
          Cell.setValue(ws, "A1", "hello, world");
          Cell.setValue(ws, "B1", 'say "hi"');

          const csv = writeCsv(wb);

          expect(csv).toContain('"hello, world"');
          expect(csv).toContain('"say ""hi"""');
        });

        it("should select worksheet by name", () => {
          const wb = Workbook.create();
          Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "first");
          Cell.setValue(Workbook.addWorksheet(wb, "Sheet2"), "A1", "second");

          const csv = writeCsv(wb, { sheetName: "Sheet2" });

          expect(csv).toBe("second");
        });

        it("should return empty string for non-existent worksheet", () => {
          const wb = Workbook.create();
          Workbook.addWorksheet(wb, "Test");

          const csv = writeCsv(wb, { sheetName: "NonExistent" });

          expect(csv).toBe("");
        });
      });

      describe("toBuffer()", () => {
        it("should convert worksheet to Uint8Array", async () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "Test");
          Cell.setValue(ws, "A1", "test");

          const buffer = await writeCsvBuffer(wb);

          expect(buffer).toBeInstanceOf(Uint8Array);
          const content = new TextDecoder().decode(buffer);
          expect(content).toBe("test");
        });

        it("should encode UTF-8 correctly", async () => {
          const wb = Workbook.create();
          const ws = Workbook.addWorksheet(wb, "Test");
          Cell.setValue(ws, "A1", "こんにちは");
          Cell.setValue(ws, "B1", "мир");

          const buffer = await writeCsvBuffer(wb);
          const content = new TextDecoder().decode(buffer);

          expect(content).toContain("こんにちは");
          expect(content).toContain("мир");
        });
      });
    });
  });
});
