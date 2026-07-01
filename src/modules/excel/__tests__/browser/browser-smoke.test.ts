import { describe, it, expect, beforeAll } from "vitest";

declare const Documonster: any;

// The `Documonster.Excel` global is injected by the IIFE bundle loaded in the
// browser-setup `beforeAll` (src/test/browser/setup.ts), which runs before
// this file's `beforeAll`. Bind the namespaces lazily — destructuring at
// module top-level would run at import time, before the global exists.
let Workbook: any, Worksheet: any, Cell: any, Column: any, Image: any;
beforeAll(() => {
  ({ Workbook, Worksheet, Cell, Column, Image } = Documonster.Excel);
});

describe("Documonster.Excel Browser Tests", () => {
  it("should read and write xlsx via binary buffer", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "blort");

    Cell.setValue(Workbook.getWorksheet(wb, "blort"), "A1", "Hello, World!");
    Cell.setValue(Workbook.getWorksheet(wb, "blort"), "A2", 7);

    const buffer = await Workbook.toBuffer(wb);

    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);

    const ws2 = Workbook.getWorksheet(wb2, "blort")!;
    expect(ws2).toBeTruthy();
    expect(Cell.getValue(ws2!, "A1")).toEqual("Hello, World!");
    expect(Cell.getValue(ws2!, "A2")).toEqual(7);
  });

  it("should read and write xlsx via base64 buffer", async () => {
    const options = {
      base64: true
    };
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "blort");

    Cell.setValue(ws, "A1", "Hello, World!");
    Cell.setValue(ws, "A2", 7);

    const buffer = await Workbook.toBuffer(wb, options);

    // Convert Uint8Array to base64 string
    const base64String = btoa(String.fromCharCode(...buffer));

    const wb2 = Workbook.create();
    await Workbook.read(wb2, base64String, options);

    const ws2 = Workbook.getWorksheet(wb2, "blort")!;
    expect(ws2).toBeTruthy();
    expect(Cell.getValue(ws2!, "A1")).toEqual("Hello, World!");
    expect(Cell.getValue(ws2!, "A2")).toEqual(7);
  });

  // Test crypto polyfill - worksheet protection uses crypto.randomBytes and crypto.createHash
  it("should support worksheet protection with password (crypto polyfill)", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "protected");

    Cell.setValue(ws, "A1", "Protected Data");

    // This uses crypto.randomBytes() and crypto.createHash() internally
    // Use low spinCount for faster test execution (default is 100000 which is slow)
    await Worksheet.protect(ws, "password123", { sheet: true, spinCount: 1000 });

    expect(ws.sheetProtection).toBeTruthy();
    expect(ws.sheetProtection.sheet).toBe(true);
    expect(ws.sheetProtection.algorithmName).toBe("SHA-512");
    expect(ws.sheetProtection.saltValue).toBeTruthy();
    expect(ws.sheetProtection.hashValue).toBeTruthy();
    expect(ws.sheetProtection.spinCount).toBe(1000);

    // Verify we can write and read back the protected workbook
    const buffer = await Workbook.toBuffer(wb);
    const wb2 = Workbook.create();
    await Workbook.read(wb2, buffer);

    const ws2 = Workbook.getWorksheet(wb2, "protected")!;
    expect(ws2).toBeTruthy();
    expect(ws2!.sheetProtection).toBeTruthy();
    expect(ws2!.sheetProtection.sheet).toBe(true);
  });

  // =========================================================================
  // XLSX/ZIP Browser Tests
  // =========================================================================

  describe("XLSX/ZIP Operations", () => {
    it("should handle multiple worksheets", async () => {
      const wb = Workbook.create();

      const ws1 = Workbook.addWorksheet(wb, "Sheet1");
      Cell.setValue(ws1, "A1", "Sheet 1 Data");

      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      Cell.setValue(ws2, "A1", "Sheet 2 Data");

      const ws3 = Workbook.addWorksheet(wb, "Sheet3");
      Cell.setValue(ws3, "A1", "Sheet 3 Data");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      expect(Workbook.getWorksheets(wb2).length).toBe(3);
      expect(Cell.getValue(Workbook.getWorksheet(wb2, "Sheet1")!, "A1")).toBe("Sheet 1 Data");
      expect(Cell.getValue(Workbook.getWorksheet(wb2, "Sheet2")!, "A1")).toBe("Sheet 2 Data");
      expect(Cell.getValue(Workbook.getWorksheet(wb2, "Sheet3")!, "A1")).toBe("Sheet 3 Data");
    });

    it("should preserve cell styles", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "styled");

      Cell.setValue(ws, "A1", "Bold");
      Cell.setFont(ws, "A1", { bold: true });

      Cell.setValue(ws, "B1", "Red");
      Cell.setFont(ws, "B1", { color: { argb: "FFFF0000" } });

      Cell.setValue(ws, "C1", "Big");
      Cell.setFont(ws, "C1", { size: 20 });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "styled")!;
      expect(Cell.getFont(ws2, "A1")?.bold).toBe(true);
      expect(Cell.getFont(ws2, "B1")?.color?.argb).toBe("FFFF0000");
      expect(Cell.getFont(ws2, "C1")?.size).toBe(20);
    });

    it("should preserve cell number formats", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "formats");

      Cell.setValue(ws, "A1", 1234.5678);
      Cell.setNumFmt(ws, "A1", "#,##0.00");

      Cell.setValue(ws, "B1", 0.75);
      Cell.setNumFmt(ws, "B1", "0%");

      Cell.setValue(ws, "C1", new Date(2024, 11, 25));
      Cell.setNumFmt(ws, "C1", "yyyy-mm-dd");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "formats")!;
      expect(Cell.getNumFmt(ws2, "A1")).toBe("#,##0.00");
      expect(Cell.getNumFmt(ws2, "B1")).toBe("0%");
      expect(Cell.getNumFmt(ws2, "C1")).toBe("yyyy-mm-dd");
    });

    it("should preserve merged cells", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "merged");

      Cell.setValue(ws, "A1", "Merged Header");
      Worksheet.merge(ws, "A1:D1");

      Cell.setValue(ws, "A2", "Another Merge");
      Worksheet.merge(ws, "A2:B3");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "merged")!;
      // Check that merge info is preserved
      expect(Cell.getValue(ws2, "A1")).toBe("Merged Header");
      expect(Cell.getValue(ws2, "A2")).toBe("Another Merge");
      // B1, C1, D1 should be merge slaves
      expect(Cell.isMerged(ws2, "B1")).toBe(true);
      expect(Cell.isMerged(ws2, "C1")).toBe(true);
    });

    it("should preserve formulas", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "formulas");

      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "A3", { formula: "SUM(A1:A2)", result: 30 });
      Cell.setValue(ws, "B1", { formula: "A1*2", result: 20 });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "formulas")!;
      expect(Cell.getFormula(ws2, "A3")).toBe("SUM(A1:A2)");
      expect(Cell.getResult(ws2, "A3")).toBe(30);
      expect(Cell.getFormula(ws2, "B1")).toBe("A1*2");
    });

    it("should handle large data sets", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "large");

      // Create 1000 rows x 10 columns
      const rows = 1000;
      const cols = 10;
      for (let r = 1; r <= rows; r++) {
        for (let c = 1; c <= cols; c++) {
          Cell.setValue(ws, r, c, `R${r}C${c}`);
        }
      }

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "large")!;
      expect(Cell.getValue(ws2, 1, 1)).toBe("R1C1");
      expect(Cell.getValue(ws2, 500, 5)).toBe("R500C5");
      expect(Cell.getValue(ws2, 1000, 10)).toBe("R1000C10");
    });

    it("should preserve hyperlinks", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "links");

      Cell.setValue(ws, "A1", {
        text: "Google",
        hyperlink: "https://www.google.com"
      });
      Cell.setValue(ws, "A2", {
        text: "Email",
        hyperlink: "mailto:test@example.com"
      });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "links")!;
      expect(Cell.getText(ws2, "A1")).toBe("Google");
      expect(Cell.getHyperlink(ws2, "A1")).toBe("https://www.google.com");
      expect(Cell.getHyperlink(ws2, "A2")).toBe("mailto:test@example.com");
    });

    it("should preserve column widths and row heights", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "dimensions");

      Column.setWidth(ws, "A", 25);
      Column.setWidth(ws, "B", 50);
      Worksheet.getRow(ws, 1).height = 30;
      Worksheet.getRow(ws, 2).height = 40;

      Cell.setValue(ws, "A1", "Wide column");
      Cell.setValue(ws, "B1", "Wider column");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "dimensions")!;
      expect(Column.getWidth(ws2, "A")).toBe(25);
      expect(Column.getWidth(ws2, "B")).toBe(50);
      expect(Worksheet.getRow(ws2, 1).height).toBe(30);
      expect(Worksheet.getRow(ws2, 2).height).toBe(40);
    });

    it("should preserve data validation", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "validation");

      Cell.setValue(ws, "A1", "Yes");
      Cell.setValidation(ws, "A1", {
        type: "list",
        allowBlank: true,
        formulae: ['"Yes,No,Maybe"']
      });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "validation")!;
      expect(Cell.getValidation(ws2, "A1")).toBeTruthy();
      expect(Cell.getValidation(ws2, "A1")?.type).toBe("list");
    });

    it("should handle workbook with defined names", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "names");

      Cell.setValue(ws, "A1", 100);
      Cell.setName(ws, "A1", "MyValue");

      Cell.setValue(ws, "B1", { formula: "MyValue * 2", result: 200 });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      // Check that the value and formula are preserved
      const ws2 = Workbook.getWorksheet(wb2, "names")!;
      expect(Cell.getValue(ws2, "A1")).toBe(100);
      expect(Cell.getFormula(ws2, "B1")).toBe("MyValue * 2");
      expect(Cell.getResult(ws2, "B1")).toBe(200);
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
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "fallback-test");

        // Add various data types
        Cell.setValue(ws, "A1", "Hello, World!");
        Cell.setValue(ws, "A2", 12345);
        Cell.setValue(ws, "A3", new Date("2024-01-01"));
        Cell.setValue(ws, "A4", { formula: "A2*2", result: 24690 });

        // Write using JS fallback compression
        const buffer = await Workbook.toBuffer(wb);
        expect(buffer).toBeTruthy();
        expect(buffer.byteLength).toBeGreaterThan(0);

        // Read using JS fallback decompression
        const wb2 = Workbook.create();
        await Workbook.read(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "fallback-test")!;
        expect(ws2).toBeTruthy();
        expect(Cell.getValue(ws2!, "A1")).toBe("Hello, World!");
        expect(Cell.getValue(ws2!, "A2")).toBe(12345);
        expect(Cell.getFormula(ws2!, "A4")).toBe("A2*2");
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
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "large-data");

        // Create a larger dataset (500 rows)
        for (let i = 1; i <= 500; i++) {
          Cell.setValue(ws, `A${i}`, `Row ${i}`);
          Cell.setValue(ws, `B${i}`, i * 100);
          Cell.setValue(ws, `C${i}`, `Data ${i} with some repeated text`.repeat(3));
        }

        const buffer = await Workbook.toBuffer(wb);
        expect(buffer.byteLength).toBeGreaterThan(0);

        const wb2 = Workbook.create();
        await Workbook.read(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "large-data")!;
        expect(Cell.getValue(ws2, "A1")).toBe("Row 1");
        expect(Cell.getValue(ws2, "B500")).toBe(50000);
        expect(Cell.getValue(ws2, "A500")).toBe("Row 500");
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
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "styled");

        Cell.setValue(ws, "A1", "Bold Text");
        Cell.setFont(ws, "A1", { bold: true, size: 14 });
        Cell.setFill(ws, "A1", {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" }
        });

        Cell.setValue(ws, "B1", 1234.56);
        Cell.setNumFmt(ws, "B1", "$#,##0.00");

        const buffer = await Workbook.toBuffer(wb);

        const wb2 = Workbook.create();
        await Workbook.read(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "styled")!;
        expect(Cell.getFont(ws2, "A1")?.bold).toBe(true);
        expect(Cell.getNumFmt(ws2, "B1")).toBe("$#,##0.00");
      } finally {
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should read file created with native compression using fallback decompression", async () => {
      // First, create a file with native compression (if available)
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "native-created");
      Cell.setValue(ws, "A1", "Created with native compression");
      Cell.setValue(ws, "A2", 42);

      const buffer = await Workbook.toBuffer(wb);

      // Now disable native APIs and try to read
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      try {
        const wb2 = Workbook.create();
        await Workbook.read(wb2, buffer);

        const ws2 = Workbook.getWorksheet(wb2, "native-created")!;
        expect(Cell.getValue(ws2, "A1")).toBe("Created with native compression");
        expect(Cell.getValue(ws2, "A2")).toBe(42);
      } finally {
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }
    });

    it("should create file with fallback that native compression can read", async () => {
      // Disable native APIs
      const originalCompressionStream = globalThis.CompressionStream;
      const originalDecompressionStream = globalThis.DecompressionStream;

      globalThis.CompressionStream = undefined as any;
      globalThis.DecompressionStream = undefined as any;

      let buffer: ArrayBuffer;
      try {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "fallback-created");
        Cell.setValue(ws, "A1", "Created with JS fallback");
        Cell.setValue(ws, "A2", 123);

        buffer = await Workbook.toBuffer(wb);
      } finally {
        // Restore native APIs
        globalThis.CompressionStream = originalCompressionStream;
        globalThis.DecompressionStream = originalDecompressionStream;
      }

      // Now read with native compression restored
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "fallback-created")!;
      expect(Cell.getValue(ws2, "A1")).toBe("Created with JS fallback");
      expect(Cell.getValue(ws2, "A2")).toBe(123);
    });

    // Regression test: loading files with drawings via loadFromFiles path
    // Previously, _processDrawingEntry would fail because it tried to collect
    // data from an already-consumed text stream instead of using the provided rawData
    it("should load files with embedded images via buffer (loadFromFiles path)", async () => {
      // Create a workbook with an embedded image
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "with-image");
      Cell.setValue(ws, "A1", "Image Test");

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

      const imageId = Image.add(wb, {
        buffer: pngData,
        extension: "png"
      });
      Image.place(ws, imageId, "B2:D6");

      // Write to buffer
      const buffer = await Workbook.toBuffer(wb);

      // Load via xlsx.load() which uses loadFromFiles internally
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      // Verify data
      const ws2 = Workbook.getWorksheet(wb2, "with-image")!;
      expect(ws2).toBeTruthy();
      expect(Cell.getValue(ws2!, "A1")).toBe("Image Test");

      // Verify image was loaded
      const images = Image.list(ws2!);
      expect(images.length).toBe(1);
      expect(images[0].range.tl.nativeCol).toBe(1); // B column = 1
      expect(images[0].range.tl.nativeRow).toBe(1); // Row 2 = 1 (0-indexed)
    });
  });
});
