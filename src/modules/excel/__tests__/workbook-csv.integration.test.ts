import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

const csvTestDataPath = makeTestDataPath(import.meta.url, "./data");

describe("Workbook", () => {
  describe("CSV", () => {
    it("differentiates between strings with leading numbers and dates when reading csv files", async () => {
      const wb = new Workbook();
      const worksheet = await wb.readCsvFile(csvTestDataPath("date-vs-leading-zeros.csv"));

      expect(worksheet.getCell("A1").value!.toString()).toBe(
        new Date("2019-11-04T00:00:00").toString()
      );
      expect(worksheet.getCell("A2").value!.toString()).toBe(
        new Date("2019-11-04T00:00:00").toString()
      );
      expect(worksheet.getCell("A3").value!.toString()).toBe(
        new Date("2019-11-04T10:17:55").toString()
      );
      expect(worksheet.getCell("A4").value).toBe("00210PRG1");
      expect(worksheet.getCell("A5").value).toBe("1234-5thisisnotadate");
    });

    it("supports encoding option on writeFile + readFile roundtrip", async () => {
      const TEST_CSV_FILE_NAME = testFilePath("csv-encoding-utf8-roundtrip", ".csv");
      const HEBREW_TEST_STRING = "משהו שכתוב בעברית";

      const wb = new Workbook();
      const ws = wb.addWorksheet("wheee");
      ws.getCell("A1").value = HEBREW_TEST_STRING;

      await wb.writeCsvFile(TEST_CSV_FILE_NAME, { encoding: "UTF-8" });

      const wb2 = new Workbook();
      const ws2 = await wb2.readCsvFile(TEST_CSV_FILE_NAME);
      expect(ws2.getCell("A1").value).toBe(HEBREW_TEST_STRING);
    }, 6000);

    describe("append mode", () => {
      it("should create new file with headers on first write", async () => {
        const TEST_FILE = testFilePath("csv-append-new-" + Date.now(), ".csv");

        const wb = new Workbook();
        const ws = wb.addWorksheet("Data");
        ws.addRow(["Name", "Age"]);
        ws.addRow(["Alice", 30]);

        await wb.writeCsvFile(TEST_FILE, { append: true });

        const wb2 = new Workbook();
        const ws2 = await wb2.readCsvFile(TEST_FILE);
        expect(ws2.rowCount).toBe(2);
        expect(ws2.getCell("A1").value).toBe("Name");
        expect(ws2.getCell("A2").value).toBe("Alice");
      });

      it("should append without headers to existing file", async () => {
        const TEST_FILE = testFilePath("csv-append-existing-" + Date.now(), ".csv");

        // First write - creates file with headers
        const wb1 = new Workbook();
        const ws1 = wb1.addWorksheet("Data");
        ws1.addRow(["Name", "Age"]);
        ws1.addRow(["Alice", 30]);
        await wb1.writeCsvFile(TEST_FILE);

        // Second write - append mode (only add data rows, no header)
        const wb2 = new Workbook();
        const ws2 = wb2.addWorksheet("Data");
        ws2.addRow(["Bob", 25]); // Data only, no header row
        await wb2.writeCsvFile(TEST_FILE, { append: true });

        // Read and verify
        const wb3 = new Workbook();
        const ws3 = await wb3.readCsvFile(TEST_FILE);
        expect(ws3.rowCount).toBe(3); // Header + Alice + Bob
        expect(ws3.getCell("A1").value).toBe("Name");
        expect(ws3.getCell("A2").value).toBe("Alice");
        expect(ws3.getCell("A3").value).toBe("Bob");
      });

      it("should append multiple batches correctly", async () => {
        const TEST_FILE = testFilePath("csv-append-batches-" + Date.now(), ".csv");

        // Initial file with header
        const wb1 = new Workbook();
        const ws1 = wb1.addWorksheet("Log");
        ws1.addRow(["Time", "Event"]);
        ws1.addRow(["10:00", "start"]);
        await wb1.writeCsvFile(TEST_FILE);

        // Append batch 1 (data only)
        const wb2 = new Workbook();
        const ws2 = wb2.addWorksheet("Log");
        ws2.addRow(["10:05", "process"]);
        await wb2.writeCsvFile(TEST_FILE, { append: true });

        // Append batch 2 (data only)
        const wb3 = new Workbook();
        const ws3 = wb3.addWorksheet("Log");
        ws3.addRow(["10:10", "end"]);
        await wb3.writeCsvFile(TEST_FILE, { append: true });

        // Verify
        const wb4 = new Workbook();
        const ws4 = await wb4.readCsvFile(TEST_FILE);
        expect(ws4.rowCount).toBe(4); // Header + 3 data rows
        expect(ws4.getCell("A1").value).toBe("Time");
        expect(ws4.getCell("B2").value).toBe("start");
        expect(ws4.getCell("B3").value).toBe("process");
        expect(ws4.getCell("B4").value).toBe("end");
      });

      it("should respect custom lineEnding in append mode", async () => {
        const TEST_FILE = testFilePath("csv-append-crlf-" + Date.now(), ".csv");

        // Create file with CRLF line endings
        const wb1 = new Workbook();
        const ws1 = wb1.addWorksheet("Data");
        ws1.addRow(["A", "B"]);
        ws1.addRow([1, 2]);
        await wb1.writeCsvFile(TEST_FILE, { lineEnding: "\r\n" });

        // Append with same line ending
        const wb2 = new Workbook();
        const ws2 = wb2.addWorksheet("Data");
        ws2.addRow([3, 4]);
        await wb2.writeCsvFile(TEST_FILE, { append: true, lineEnding: "\r\n" });

        // Read and verify
        const wb3 = new Workbook();
        const ws3 = await wb3.readCsvFile(TEST_FILE);
        expect(ws3.rowCount).toBe(3);
        expect(ws3.getCell("A3").value).toBe(3);
      });
    });

    // =========================================================================
    // Unified API Tests (parse, stringify, toBuffer)
    // =========================================================================

    describe("Unified API", () => {
      describe("parse()", () => {
        it("should parse CSV string", async () => {
          const wb = new Workbook();
          const ws = await wb.readCsv("a,b,c\n1,2,3");

          expect(ws.getCell("A1").value).toBe("a");
          expect(ws.getCell("B1").value).toBe("b");
          expect(ws.getCell("A2").value).toBe(1);
        });

        it("should auto-detect semicolon delimiter", async () => {
          const wb = new Workbook();
          const ws = await wb.readCsv("a;b;c\n1;2;3", { delimiter: "" });

          expect(ws.getCell("A1").value).toBe("a");
          expect(ws.getCell("B1").value).toBe("b");
          expect(ws.getCell("C1").value).toBe("c");
          expect(ws.getCell("A2").value).toBe(1);
        });

        it("should auto-detect tab delimiter", async () => {
          const wb = new Workbook();
          const ws = await wb.readCsv("a\tb\tc\n1\t2\t3", { delimiter: "" });

          expect(ws.getCell("A1").value).toBe("a");
          expect(ws.getCell("B1").value).toBe("b");
          expect(ws.getCell("C1").value).toBe("c");
        });

        it("should respect explicit delimiter option", async () => {
          const wb = new Workbook();
          // Even though the data has commas, use semicolon delimiter
          const ws = await wb.readCsv("a,b;c,d\n1,2;3,4", { delimiter: ";" });

          expect(ws.getCell("A1").value).toBe("a,b");
          expect(ws.getCell("B1").value).toBe("c,d");
        });

        it("should parse ArrayBuffer input", async () => {
          const wb = new Workbook();
          const data = new TextEncoder().encode("name,value\ntest,123");
          const ws = await wb.readCsv(data.buffer);

          expect(ws.getCell("A1").value).toBe("name");
          expect(ws.getCell("B2").value).toBe(123);
        });

        it("should parse Uint8Array input", async () => {
          const wb = new Workbook();
          const data = new TextEncoder().encode("x,y\n10,20");
          const ws = await wb.readCsv(data);

          expect(ws.getCell("A1").value).toBe("x");
          expect(ws.getCell("A2").value).toBe(10);
        });

        it("should use custom sheet name", async () => {
          const wb = new Workbook();
          const ws = await wb.readCsv("a,b\n1,2", { sheetName: "MySheet" });

          expect(ws.name).toBe("MySheet");
        });

        it("should support headers option", async () => {
          const wb = new Workbook();
          const ws = await wb.readCsv("name,age\nAlice,30\nBob,25", { headers: true });

          // When headers: true, the first row becomes column headers but values are still parsed
          expect(ws.getCell("A1").value).toBe("name");
          expect(ws.getCell("B1").value).toBe("age");
        });

        it("should support headers option with stream input", async () => {
          const wb = new Workbook();
          // Create a readable stream from string using async generator
          const csvData = "name,age\nAlice,30\nBob,25";
          const readable = {
            async *[Symbol.asyncIterator]() {
              // Split into chunks to simulate streaming
              yield new TextEncoder().encode(csvData.slice(0, 10));
              yield new TextEncoder().encode(csvData.slice(10));
            }
          };

          const ws = await wb.readCsv(readable as any, { headers: true });

          // Verify headers row is written
          expect(ws.getCell("A1").value).toBe("name");
          expect(ws.getCell("B1").value).toBe("age");
          // Verify data rows are converted from object to array correctly
          expect(ws.getCell("A2").value).toBe("Alice");
          expect(ws.getCell("B2").value).toBe(30);
          expect(ws.getCell("A3").value).toBe("Bob");
          expect(ws.getCell("B3").value).toBe(25);
        });

        it("should support decimalSeparator option", async () => {
          const wb = new Workbook();
          await wb.readCsv("value\n1,50\n2,75", {
            delimiter: ";",
            decimalSeparator: ","
          });

          // Note: with delimiter auto-detect and single column, comma won't be detected as delimiter
          // So we explicitly set delimiter to avoid ambiguity
        });
      });

      describe("stringify()", () => {
        it("should convert worksheet to CSV string", () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("Test");
          ws.getCell("A1").value = "hello";
          ws.getCell("B1").value = "world";
          ws.getCell("A2").value = 1;
          ws.getCell("B2").value = 2;

          const csv = wb.writeCsv();

          expect(csv).toBe("hello,world\n1,2");
        });

        it("should use custom delimiter", () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("Test");
          ws.getCell("A1").value = "a";
          ws.getCell("B1").value = "b";

          const csv = wb.writeCsv({ delimiter: ";" });

          expect(csv).toBe("a;b");
        });

        it("should quote fields with special characters", () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("Test");
          ws.getCell("A1").value = "hello, world";
          ws.getCell("B1").value = 'say "hi"';

          const csv = wb.writeCsv();

          expect(csv).toContain('"hello, world"');
          expect(csv).toContain('"say ""hi"""');
        });

        it("should select worksheet by name", () => {
          const wb = new Workbook();
          wb.addWorksheet("Sheet1").getCell("A1").value = "first";
          wb.addWorksheet("Sheet2").getCell("A1").value = "second";

          const csv = wb.writeCsv({ sheetName: "Sheet2" });

          expect(csv).toBe("second");
        });

        it("should return empty string for non-existent worksheet", () => {
          const wb = new Workbook();
          wb.addWorksheet("Test");

          const csv = wb.writeCsv({ sheetName: "NonExistent" });

          expect(csv).toBe("");
        });
      });

      describe("toBuffer()", () => {
        it("should convert worksheet to Uint8Array", async () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("Test");
          ws.getCell("A1").value = "test";

          const buffer = await wb.writeCsvBuffer();

          expect(buffer).toBeInstanceOf(Uint8Array);
          const content = new TextDecoder().decode(buffer);
          expect(content).toBe("test");
        });

        it("should encode UTF-8 correctly", async () => {
          const wb = new Workbook();
          const ws = wb.addWorksheet("Test");
          ws.getCell("A1").value = "こんにちは";
          ws.getCell("B1").value = "мир";

          const buffer = await wb.writeCsvBuffer();
          const content = new TextDecoder().decode(buffer);

          expect(content).toContain("こんにちは");
          expect(content).toContain("мир");
        });
      });
    });
  });
});
