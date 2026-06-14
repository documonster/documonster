import {
  cellDataValidation,
  cellFont,
  cellFormula,
  cellGetValue,
  cellHyperlink,
  cellIsMerged,
  cellNumFmt,
  cellResult,
  cellSetDataValidation,
  cellSetFill,
  cellSetFont,
  cellSetNumFmt,
  cellSetValue,
  cellText,
  cellSetName
} from "@excel/cell";
import { readCsv, writeCsv, writeCsvBuffer } from "@excel/csv-bridge";
import { Cell, Image, Worksheet } from "@excel/index";
import { getCell, getSheetName, getColumn } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

declare const ExcelTS: {
  Workbook: any;
};

describe("ExcelTS Browser Tests", () => {
  it("should read and write xlsx via binary buffer", async () => {
    const { Workbook } = ExcelTS;
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "blort");

    cellSetValue(getCell(ws, "A1"), "Hello, World!");
    cellSetValue(getCell(ws, "A2"), 7);

    const buffer = await Workbook.toXlsxBuffer(wb);

    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, buffer);

    const ws2 = Workbook.getWorksheet(wb2, "blort")!;
    expect(ws2).toBeTruthy();
    expect(ws2!.getCell("A1").value).toEqual("Hello, World!");
    expect(ws2!.getCell("A2").value).toEqual(7);
  });

  it("should read and write xlsx via base64 buffer", async () => {
    const { Workbook } = ExcelTS;
    const options = {
      base64: true
    };
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "blort");

    cellSetValue(getCell(ws, "A1"), "Hello, World!");
    cellSetValue(getCell(ws, "A2"), 7);

    const buffer = await Workbook.toXlsxBuffer(wb, options);

    // Convert Uint8Array to base64 string
    const base64String = btoa(String.fromCharCode(...buffer));

    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, base64String, options);

    const ws2 = Workbook.getWorksheet(wb2, "blort")!;
    expect(ws2).toBeTruthy();
    expect(ws2!.getCell("A1").value).toEqual("Hello, World!");
    expect(ws2!.getCell("A2").value).toEqual(7);
  });

  // CSV support is now available in browser using native RFC 4180 implementation
  it("should write csv via buffer (browser)", async () => {
    const { Workbook } = ExcelTS;
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "blort");

    cellSetValue(getCell(ws, "A1"), "Hello, World!");
    cellSetValue(getCell(ws, "B1"), "What time is it?");
    cellSetValue(getCell(ws, "A2"), 7);
    cellSetValue(getCell(ws, "B2"), "12pm");

    const buffer = await writeCsvBuffer(wb);

    // In browser, buffer is Uint8Array; use TextDecoder to convert to string
    const content = new TextDecoder().decode(buffer);
    // Uses \n as row delimiter, trailingNewline defaults to false
    expect(content).toEqual('"Hello, World!",What time is it?\n7,12pm');
  });

  // Test crypto polyfill - worksheet protection uses crypto.randomBytes and crypto.createHash
  it("should support worksheet protection with password (crypto polyfill)", async () => {
    const { Workbook } = ExcelTS;
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "protected");

    cellSetValue(getCell(ws, "A1"), "Protected Data");

    // This uses crypto.randomBytes() and crypto.createHash() internally
    // Use low spinCount for faster test execution (default is 100000 which is slow)
    await ws.protect("password123", { sheet: true, spinCount: 1000 });

    expect(ws.sheetProtection).toBeTruthy();
    expect(ws.sheetProtection.sheet).toBe(true);
    expect(ws.sheetProtection.algorithmName).toBe("SHA-512");
    expect(ws.sheetProtection.saltValue).toBeTruthy();
    expect(ws.sheetProtection.hashValue).toBeTruthy();
    expect(ws.sheetProtection.spinCount).toBe(1000);

    // Verify we can write and read back the protected workbook
    const buffer = await Workbook.toXlsxBuffer(wb);
    const wb2 = Workbook.create();
    await Workbook.loadXlsx(wb2, buffer);

    const ws2 = Workbook.getWorksheet(wb2, "protected")!;
    expect(ws2).toBeTruthy();
    expect(ws2!.sheetProtection).toBeTruthy();
    expect(ws2!.sheetProtection.sheet).toBe(true);
  });

  // =========================================================================
  // CSV Browser Tests
  // =========================================================================

  describe("CSV Operations", () => {
    it("should load CSV from string", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const csvContent = "Name,Age,City\nAlice,30,New York\nBob,25,Los Angeles";

      const ws = await readCsv(wb, csvContent);

      expect(Cell.getValue(ws, "A1")).toBe("Name");
      expect(Cell.getValue(ws, "B1")).toBe("Age");
      expect(Cell.getValue(ws, "C1")).toBe("City");
      expect(Cell.getValue(ws, "A2")).toBe("Alice");
      // CSV numbers are auto-converted to numbers by the worksheet
      expect(Cell.getValue(ws, "B2")).toBe(30);
      expect(Cell.getValue(ws, "C2")).toBe("New York");
      expect(Cell.getValue(ws, "A3")).toBe("Bob");
    });

    it("should load CSV from ArrayBuffer", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const csvContent = "Col1,Col2\nA,B\nC,D";
      const buffer = new TextEncoder().encode(csvContent);

      const ws = await readCsv(wb, buffer);

      expect(Cell.getValue(ws, "A1")).toBe("Col1");
      expect(Cell.getValue(ws, "B2")).toBe("B");
    });

    it("should handle quoted fields with commas", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const csvContent = 'Name,Address\n"Smith, John","123 Main St, Apt 4"';

      const ws = await readCsv(wb, csvContent);

      expect(Cell.getValue(ws, "A2")).toBe("Smith, John");
      expect(Cell.getValue(ws, "B2")).toBe("123 Main St, Apt 4");
    });

    it("should handle quoted fields with newlines", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const csvContent = 'Description\n"Line 1\nLine 2\nLine 3"';

      const ws = await readCsv(wb, csvContent);

      expect(Cell.getValue(ws, "A2")).toBe("Line 1\nLine 2\nLine 3");
    });

    it("should handle escaped quotes", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const csvContent = 'Quote\n"He said ""Hello"""';

      const ws = await readCsv(wb, csvContent);

      expect(Cell.getValue(ws, "A2")).toBe('He said "Hello"');
    });

    it("should write CSV with proper quoting", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      cellSetValue(getCell(ws, "A1"), "Name");
      cellSetValue(getCell(ws, "B1"), "Quote");
      cellSetValue(getCell(ws, "A2"), "Smith, John");
      cellSetValue(getCell(ws, "B2"), 'He said "Hi"');

      const content = writeCsv(wb);

      expect(content).toContain('"Smith, John"');
      expect(content).toContain('"He said ""Hi"""');
    });

    it("should write CSV to buffer", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      cellSetValue(getCell(ws, "A1"), "Test");
      cellSetValue(getCell(ws, "B1"), "Data");

      const buffer = await writeCsvBuffer(wb);

      expect(buffer).toBeInstanceOf(Uint8Array);
      const content = new TextDecoder().decode(buffer);
      expect(content).toBe("Test,Data");
    });

    it("should support tab delimiters", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "tab");

      cellSetValue(getCell(ws, "A1"), "Col1");
      cellSetValue(getCell(ws, "B1"), "Col2");
      cellSetValue(getCell(ws, "A2"), "A");
      cellSetValue(getCell(ws, "B2"), "B");

      // Write with tab delimiter
      const output = writeCsv(wb, {
        sheetName: ws.name,
        delimiter: "\t"
      });
      expect(output).toBe("Col1\tCol2\nA\tB");
    });

    it("should round-trip CSV data", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const originalCsv = 'Name,Value\nTest,123\n"Quoted, Value",456';

      const ws = await readCsv(wb, originalCsv);
      const outputCsv = writeCsv(wb, { sheetName: getSheetName(ws) });

      // Load the output back and verify
      const wb2 = Workbook.create();
      const ws2 = await readCsv(wb2, outputCsv);

      expect(Cell.getValue(ws2, "A1")).toBe("Name");
      expect(Cell.getValue(ws2, "B1")).toBe("Value");
      expect(Cell.getValue(ws2, "A2")).toBe("Test");
      // Numbers are auto-converted
      expect(Cell.getValue(ws2, "B2")).toBe(123);
      expect(Cell.getValue(ws2, "A3")).toBe("Quoted, Value");
      expect(Cell.getValue(ws2, "B3")).toBe(456);
    });
  });

  // =========================================================================
  // XLSX/ZIP Browser Tests
  // =========================================================================

  describe("XLSX/ZIP Operations", () => {
    it("should handle multiple worksheets", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();

      const ws1 = Workbook.addWorksheet(wb, "Sheet1");
      cellSetValue(getCell(ws1, "A1"), "Sheet 1 Data");

      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      cellSetValue(getCell(ws2, "A1"), "Sheet 2 Data");

      const ws3 = Workbook.addWorksheet(wb, "Sheet3");
      cellSetValue(getCell(ws3, "A1"), "Sheet 3 Data");

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      expect(wb2.worksheets.length).toBe(3);
      expect(Workbook.getWorksheet(wb2, "Sheet1")!.getCell("A1").value).toBe("Sheet 1 Data");
      expect(Workbook.getWorksheet(wb2, "Sheet2")!.getCell("A1").value).toBe("Sheet 2 Data");
      expect(Workbook.getWorksheet(wb2, "Sheet3")!.getCell("A1").value).toBe("Sheet 3 Data");
    });

    it("should preserve cell styles", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "styled");

      cellSetValue(getCell(ws, "A1"), "Bold");
      cellSetFont(getCell(ws, "A1"), { bold: true });

      cellSetValue(getCell(ws, "B1"), "Red");
      cellSetFont(getCell(ws, "B1"), { color: { argb: "FFFF0000" } });

      cellSetValue(getCell(ws, "C1"), "Big");
      cellSetFont(getCell(ws, "C1"), { size: 20 });

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "styled")!;
      expect(cellFont(getCell(ws2, "A1"))?.bold).toBe(true);
      expect(cellFont(getCell(ws2, "B1"))?.color?.argb).toBe("FFFF0000");
      expect(cellFont(getCell(ws2, "C1"))?.size).toBe(20);
    });

    it("should preserve cell number formats", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "formats");

      cellSetValue(getCell(ws, "A1"), 1234.5678);
      cellSetNumFmt(getCell(ws, "A1"), "#,##0.00");

      cellSetValue(getCell(ws, "B1"), 0.75);
      cellSetNumFmt(getCell(ws, "B1"), "0%");

      cellSetValue(getCell(ws, "C1"), new Date(2024, 11, 25));
      cellSetNumFmt(getCell(ws, "C1"), "yyyy-mm-dd");

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "formats")!;
      expect(cellNumFmt(getCell(ws2, "A1"))).toBe("#,##0.00");
      expect(cellNumFmt(getCell(ws2, "B1"))).toBe("0%");
      expect(cellNumFmt(getCell(ws2, "C1"))).toBe("yyyy-mm-dd");
    });

    it("should preserve merged cells", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "merged");

      cellSetValue(getCell(ws, "A1"), "Merged Header");
      Worksheet.merge(ws, "A1:D1");

      cellSetValue(getCell(ws, "A2"), "Another Merge");
      Worksheet.merge(ws, "A2:B3");

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "merged")!;
      // Check that merge info is preserved
      expect(cellGetValue(getCell(ws2, "A1"))).toBe("Merged Header");
      expect(cellGetValue(getCell(ws2, "A2"))).toBe("Another Merge");
      // B1, C1, D1 should be merge slaves
      expect(cellIsMerged(getCell(ws2, "B1"))).toBe(true);
      expect(cellIsMerged(getCell(ws2, "C1"))).toBe(true);
    });

    it("should preserve formulas", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "formulas");

      cellSetValue(getCell(ws, "A1"), 10);
      cellSetValue(getCell(ws, "A2"), 20);
      cellSetValue(getCell(ws, "A3"), { formula: "SUM(A1:A2)", result: 30 });
      cellSetValue(getCell(ws, "B1"), { formula: "A1*2", result: 20 });

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "formulas")!;
      expect(cellFormula(getCell(ws2, "A3"))).toBe("SUM(A1:A2)");
      expect(cellResult(getCell(ws2, "A3"))).toBe(30);
      expect(cellFormula(getCell(ws2, "B1"))).toBe("A1*2");
    });

    it("should handle large data sets", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "large");

      // Create 1000 rows x 10 columns
      const rows = 1000;
      const cols = 10;
      for (let r = 1; r <= rows; r++) {
        for (let c = 1; c <= cols; c++) {
          cellSetValue(getCell(ws, r, c), `R${r}C${c}`);
        }
      }

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "large")!;
      expect(cellGetValue(getCell(ws2, 1, 1))).toBe("R1C1");
      expect(cellGetValue(getCell(ws2, 500, 5))).toBe("R500C5");
      expect(cellGetValue(getCell(ws2, 1000, 10))).toBe("R1000C10");
    });

    it("should preserve hyperlinks", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "links");

      cellSetValue(getCell(ws, "A1"), {
        text: "Google",
        hyperlink: "https://www.google.com"
      });
      cellSetValue(getCell(ws, "A2"), {
        text: "Email",
        hyperlink: "mailto:test@example.com"
      });

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "links")!;
      expect(cellText(getCell(ws2, "A1"))).toBe("Google");
      expect(cellHyperlink(getCell(ws2, "A1"))).toBe("https://www.google.com");
      expect(cellHyperlink(getCell(ws2, "A2"))).toBe("mailto:test@example.com");
    });

    it("should preserve column widths and row heights", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "dimensions");

      getColumn(ws, "A").width = 25;
      getColumn(ws, "B").width = 50;
      Worksheet.getRow(ws, 1).height = 30;
      Worksheet.getRow(ws, 2).height = 40;

      cellSetValue(getCell(ws, "A1"), "Wide column");
      cellSetValue(getCell(ws, "B1"), "Wider column");

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "dimensions")!;
      expect(getColumn(ws2, "A").width).toBe(25);
      expect(getColumn(ws2, "B").width).toBe(50);
      expect(Worksheet.getRow(ws2, 1).height).toBe(30);
      expect(Worksheet.getRow(ws2, 2).height).toBe(40);
    });

    it("should preserve data validation", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "validation");

      cellSetValue(getCell(ws, "A1"), "Yes");
      cellSetDataValidation(getCell(ws, "A1"), {
        type: "list",
        allowBlank: true,
        formulae: ['"Yes,No,Maybe"']
      });

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "validation")!;
      expect(cellDataValidation(getCell(ws2, "A1"))).toBeTruthy();
      expect(cellDataValidation(getCell(ws2, "A1"))?.type).toBe("list");
    });

    it("should handle workbook with defined names", async () => {
      const { Workbook } = ExcelTS;
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "names");

      cellSetValue(getCell(ws, "A1"), 100);
      cellSetName(getCell(ws, "A1"), "MyValue");

      cellSetValue(getCell(ws, "B1"), { formula: "MyValue * 2", result: 200 });

      const buffer = await Workbook.toXlsxBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      // Check that the value and formula are preserved
      const ws2 = Workbook.getWorksheet(wb2, "names")!;
      expect(cellGetValue(getCell(ws2, "A1"))).toBe(100);
      expect(cellFormula(getCell(ws2, "B1"))).toBe("MyValue * 2");
      expect(cellResult(getCell(ws2, "B1"))).toBe(200);
    });
  });

  // =========================================================================
  // DEFLATE Fallback Tests (simulating old browsers without CompressionStream)
  // =========================================================================
  describe("DEFLATE Fallback (simulating old browsers)", () => {
    it("should work when CompressionStream is disabled", async () => {
      // Save original APIs
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      // Disable native compression APIs to simulate old browsers
      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      try {
        const { Workbook } = ExcelTS;
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "fallback-test");

        // Add various data types
        cellSetValue(getCell(ws, "A1"), "Hello, World!");
        cellSetValue(getCell(ws, "A2"), 12345);
        cellSetValue(getCell(ws, "A3"), new Date("2024-01-01"));
        cellSetValue(getCell(ws, "A4"), { formula: "A2*2", result: 24690 });

        // Write using JS fallback compression
        const buffer = await Workbook.toXlsxBuffer(wb);
        expect(buffer).toBeTruthy();
        expect(buffer.byteLength).toBeGreaterThan(0);

        // Read using JS fallback decompression
        const wb2 = Workbook.create();
        await Workbook.loadXlsx(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "fallback-test")!;
        expect(ws2).toBeTruthy();
        expect(ws2!.getCell("A1").value).toBe("Hello, World!");
        expect(ws2!.getCell("A2").value).toBe(12345);
        expect(ws2!.getCell("A4").formula).toBe("A2*2");
      } finally {
        // Restore original APIs
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should handle large workbook with fallback compression", async () => {
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      try {
        const { Workbook } = ExcelTS;
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "large-data");

        // Create a larger dataset (500 rows)
        for (let i = 1; i <= 500; i++) {
          cellSetValue(getCell(ws, `A${i}`), `Row ${i}`);
          cellSetValue(getCell(ws, `B${i}`), i * 100);
          cellSetValue(getCell(ws, `C${i}`), `Data ${i} with some repeated text`.repeat(3));
        }

        const buffer = await Workbook.toXlsxBuffer(wb);
        expect(buffer.byteLength).toBeGreaterThan(0);

        const wb2 = Workbook.create();
        await Workbook.loadXlsx(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "large-data")!;
        expect(cellGetValue(getCell(ws2, "A1"))).toBe("Row 1");
        expect(cellGetValue(getCell(ws2, "B500"))).toBe(50000);
        expect(cellGetValue(getCell(ws2, "A500"))).toBe("Row 500");
      } finally {
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should handle styles and formatting with fallback", async () => {
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      try {
        const { Workbook } = ExcelTS;
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "styled");

        cellSetValue(getCell(ws, "A1"), "Bold Text");
        cellSetFont(getCell(ws, "A1"), { bold: true, size: 14 });
        cellSetFill(getCell(ws, "A1"), {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        });

        cellSetValue(getCell(ws, "B1"), 1234.56);
        cellSetNumFmt(getCell(ws, "B1"), "$#,##0.00");

        const buffer = await Workbook.toXlsxBuffer(wb);

        const wb2 = Workbook.create();
        await Workbook.loadXlsx(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "styled")!;
        expect(cellFont(getCell(ws2, "A1"))?.bold).toBe(true);
        expect(cellNumFmt(getCell(ws2, "B1"))).toBe("$#,##0.00");
      } finally {
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should read file created with native compression using fallback decompression", async () => {
      const { Workbook } = ExcelTS;

      // First, create a file with native compression (if available)
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "native-created");
      cellSetValue(getCell(ws, "A1"), "Created with native compression");
      cellSetValue(getCell(ws, "A2"), 42);

      const buffer = await Workbook.toXlsxBuffer(wb);

      // Now disable native APIs and try to read
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      try {
        const wb2 = Workbook.create();
        await Workbook.loadXlsx(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "native-created")!;
        expect(cellGetValue(getCell(ws2, "A1"))).toBe("Created with native compression");
        expect(cellGetValue(getCell(ws2, "A2"))).toBe(42);
      } finally {
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should create file with fallback that native compression can read", async () => {
      const { Workbook } = ExcelTS;

      // Disable native APIs
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      let buffer: ArrayBuffer;
      try {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "fallback-created");
        cellSetValue(getCell(ws, "A1"), "Created with JS fallback");
        cellSetValue(getCell(ws, "A2"), 123);

        buffer = await Workbook.toXlsxBuffer(wb);
      } finally {
        // Restore native APIs
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }

      // Now read with native compression restored
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "fallback-created")!;
      expect(cellGetValue(getCell(ws2, "A1"))).toBe("Created with JS fallback");
      expect(cellGetValue(getCell(ws2, "A2"))).toBe(123);
    });

    // Regression test: loading files with drawings via loadFromFiles path
    // Previously, _processDrawingEntry would fail because it tried to collect
    // data from an already-consumed text stream instead of using the provided rawData
    it("should load files with embedded images via buffer (loadFromFiles path)", async () => {
      const { Workbook } = ExcelTS;

      // Create a workbook with an embedded image
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "with-image");
      cellSetValue(getCell(ws, "A1"), "Image Test");

      // Add a simple 1x1 PNG image (smallest valid PNG)
      // This is a 1x1 red pixel PNG
      const pngData = new Uint8Array([
        0x89,
        0x50,
        0x4e,
        0x47,
        0x0d,
        0x0a,
        0x1a,
        0x0a, // PNG signature
        0x00,
        0x00,
        0x00,
        0x0d,
        0x49,
        0x48,
        0x44,
        0x52, // IHDR chunk length + type
        0x00,
        0x00,
        0x00,
        0x01,
        0x00,
        0x00,
        0x00,
        0x01, // 1x1 dimensions
        0x08,
        0x02,
        0x00,
        0x00,
        0x00,
        0x90,
        0x77,
        0x53, // bit depth, color type, etc
        0xde,
        0x00,
        0x00,
        0x00,
        0x0c,
        0x49,
        0x44,
        0x41, // IDAT chunk
        0x54,
        0x08,
        0xd7,
        0x63,
        0xf8,
        0xcf,
        0xc0,
        0x00, // compressed data
        0x00,
        0x00,
        0x03,
        0x00,
        0x01,
        0x00,
        0x05,
        0xfe, //
        0xd4,
        0xef,
        0x00,
        0x00,
        0x00,
        0x00,
        0x49,
        0x45, // IEND chunk
        0x4e,
        0x44,
        0xae,
        0x42,
        0x60,
        0x82
      ]);

      const imageId = wb.addImage({
        buffer: pngData,
        extension: "png"
      });
      Image.place(ws, imageId, "B2:D6");

      // Write to buffer
      const buffer = await Workbook.toXlsxBuffer(wb);

      // Load via xlsx.load() which uses loadFromFiles internally
      const wb2 = Workbook.create();
      await Workbook.loadXlsx(wb2, buffer);

      // Verify data
      const ws2 = Workbook.getWorksheet(wb2, "with-image")!;
      expect(ws2).toBeTruthy();
      expect(ws2!.getCell("A1").value).toBe("Image Test");

      // Verify image was loaded
      const images = ws2!.getImages();
      expect(images.length).toBe(1);
      expect(images[0].range.tl.col).toBe(1); // B column = 1
      expect(images[0].range.tl.row).toBe(1); // Row 2 = 1 (0-indexed)
    });
  });
});
