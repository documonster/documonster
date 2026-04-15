import { ZipArchive } from "@archive/zip";
import { PassThrough } from "@stream";
import { describe, it, expect } from "vitest";

import { Workbook, WorkbookReader } from "../../../index";

// =============================================================================
// Helpers
// =============================================================================

/** Build a minimal valid XLSX buffer for testing. */
async function buildMinimalXlsx(): Promise<Uint8Array> {
  const wb = new Workbook();
  wb.addWorksheet("Sheet1").getCell("A1").value = "hello";
  return wb.xlsx.writeBuffer();
}

/**
 * Build a dirty XLSX buffer whose shared strings contain invalid XML characters.
 *
 * Creates a normal XLSX via the Workbook API, then extracts the raw ZIP entries,
 * injects invalid bytes (0x7F, 0x01, etc.) into the shared-strings XML, and
 * re-zips the result.  This simulates real-world XLSX files from third-party
 * tools that embed non-XML-1.0 characters.
 */
async function buildDirtyXlsx(
  invalidChars: string = "\x7f",
  cellText: string = "hello"
): Promise<Uint8Array> {
  // 1. Build a clean XLSX
  const wb = new Workbook();
  wb.addWorksheet("Sheet1").getCell("A1").value = cellText;
  const cleanBuffer = await wb.xlsx.writeBuffer();

  // 2. Extract all ZIP entries
  const { extractAll } = await import("@archive/unzip/extract");
  const entries = await extractAll(cleanBuffer);

  // 3. Find and mutate the shared strings XML
  const ssiKey = "xl/sharedStrings.xml";
  const ssiEntry = entries.get(ssiKey);
  if (!ssiEntry) {
    throw new Error("no sharedStrings.xml found in generated XLSX");
  }
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let xml = decoder.decode(ssiEntry.data);
  // Inject invalid chars right before the cell text
  xml = xml.replace(cellText, invalidChars + cellText + invalidChars);
  ssiEntry.data = encoder.encode(xml);

  // 4. Re-zip all entries
  const archive = new ZipArchive({ level: 0, reproducible: true });
  for (const [name, entry] of entries) {
    archive.add(name, entry.data);
  }
  return archive.bytes();
}

/** Build a feature-rich XLSX buffer for round-trip testing. */
async function buildRichXlsx(): Promise<Uint8Array> {
  const wb = new Workbook();
  wb.creator = "TestAuthor";
  wb.created = new Date(2024, 0, 1);

  const ws1 = wb.addWorksheet("Data");
  ws1.columns = [
    { header: "ID", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "Value", key: "val", width: 15 }
  ];
  ws1.addRow({ id: 1, name: "Alice", val: 100 });
  ws1.addRow({ id: 2, name: "Bob", val: 200 });
  ws1.addRow({ id: 3, name: "Charlie", val: 300 });

  ws1.getCell("D1").value = new Date(2024, 5, 15);
  ws1.getCell("D1").numFmt = "yyyy-mm-dd";
  ws1.getCell("E1").value = true;
  ws1.getCell("F1").value = { formula: "SUM(C2:C4)", result: 600 };
  ws1.getCell("G1").value = { text: "Google", hyperlink: "https://google.com" };

  ws1.getCell("A1").font = { bold: true };
  ws1.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCCCCCC" } };

  ws1.mergeCells("H1:I2");
  ws1.getCell("H1").value = "merged";

  const ws2 = wb.addWorksheet("Hidden", { state: "hidden" });
  ws2.getCell("A1").value = "secret";

  return wb.xlsx.writeBuffer();
}

// =============================================================================
// Tests
// =============================================================================

describe("XLSX", () => {
  // ===========================================================================
  // writeBuffer / load Round-Trip
  // ===========================================================================

  describe("writeBuffer / load round-trip", () => {
    it("writes and loads a minimal workbook", async () => {
      const buffer = await buildMinimalXlsx();

      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);

      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("Sheet1")!.getCell("A1").value).toBe("hello");
    });

    it("preserves all value types through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      const ws = wb.getWorksheet("Data")!;
      expect(ws.getCell("A2").value).toBe(1); // number
      expect(ws.getCell("B2").value).toBe("Alice"); // string
      expect(ws.getCell("D1").value).toBeInstanceOf(Date); // date
      expect(ws.getCell("E1").value).toBe(true); // boolean

      // formula
      const f = ws.getCell("F1").value as { formula: string; result: number };
      expect(f.formula).toBe("SUM(C2:C4)");
      expect(f.result).toBe(600);

      // hyperlink
      const h = ws.getCell("G1").value as { text: string; hyperlink: string };
      expect(h.text).toBe("Google");
      expect(h.hyperlink).toBe("https://google.com");
    });

    it("preserves styles through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      const cell = wb.getWorksheet("Data")!.getCell("A1");
      expect(cell.font!.bold).toBe(true);

      const fill = cell.fill as { type: string; pattern: string; fgColor: { argb: string } };
      expect(fill.pattern).toBe("solid");
      expect(fill.fgColor.argb).toBe("FFCCCCCC");
    });

    it("preserves merged cells through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      const ws = wb.getWorksheet("Data")!;
      expect(ws.getCell("H1").value).toBe("merged");
      expect(ws.getCell("I2").isMerged).toBe(true);
    });

    it("preserves column definitions through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      const ws = wb.getWorksheet("Data")!;
      expect(ws.getColumn(1).width).toBe(10);
      expect(ws.getColumn(2).width).toBe(32);
      expect(ws.getColumn(3).width).toBe(15);
    });

    it("preserves metadata through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      expect(wb.creator).toBe("TestAuthor");
      expect(wb.created).toBeInstanceOf(Date);
    });

    it("preserves sheet state (hidden) through round-trip", async () => {
      const buffer = await buildRichXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);

      expect(wb.getWorksheet("Data")!.state).toBe("visible");
      expect(wb.getWorksheet("Hidden")!.state).toBe("hidden");
    });

    it("preserves multiple sheets with correct order", async () => {
      const wb1 = new Workbook();
      wb1.addWorksheet("First").getCell("A1").value = 1;
      wb1.addWorksheet("Second").getCell("A1").value = 2;
      wb1.addWorksheet("Third").getCell("A1").value = 3;

      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const names = wb2.worksheets.map(ws => ws.name);
      expect(names).toEqual(["First", "Second", "Third"]);
    });

    it("handles empty workbook (no worksheets)", async () => {
      const wb1 = new Workbook();
      const buffer = await wb1.xlsx.writeBuffer();

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.worksheets.length).toBe(0);
    });

    it("handles workbook with empty worksheet", async () => {
      const wb1 = new Workbook();
      wb1.addWorksheet("Empty");
      const buffer = await wb1.xlsx.writeBuffer();

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.worksheets.length).toBe(1);
      expect(wb2.getWorksheet("Empty")).toBeDefined();
    });
  });

  // ===========================================================================
  // Input Types for load()
  // ===========================================================================

  describe("load() input types", () => {
    it("accepts Uint8Array", async () => {
      const buffer = await buildMinimalXlsx();
      const wb = new Workbook();
      await wb.xlsx.load(buffer);
      expect(wb.worksheets.length).toBe(1);
    });

    it("accepts ArrayBuffer", async () => {
      const buffer = await buildMinimalXlsx();
      const arrayBuffer: ArrayBuffer = buffer.buffer.slice(
        buffer.byteOffset,
        buffer.byteOffset + buffer.byteLength
      ) as ArrayBuffer;

      const wb = new Workbook();
      await wb.xlsx.load(arrayBuffer);
      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("Sheet1")!.getCell("A1").value).toBe("hello");
    });

    it("accepts base64 string with option", async () => {
      const buffer = await buildMinimalXlsx();
      // Convert to base64
      const base64 = Buffer.from(buffer).toString("base64");

      const wb = new Workbook();
      await wb.xlsx.load(base64, { base64: true });
      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("Sheet1")!.getCell("A1").value).toBe("hello");
    });

    it("rejects invalid input (plain object)", async () => {
      const wb = new Workbook();
      await expect(wb.xlsx.load({} as any)).rejects.toThrow();
    });

    it("rejects null input", async () => {
      const wb = new Workbook();
      await expect(wb.xlsx.load(null as any)).rejects.toThrow();
    });

    it("rejects undefined input", async () => {
      const wb = new Workbook();
      await expect(wb.xlsx.load(undefined as any)).rejects.toThrow();
    });
  });

  // ===========================================================================
  // Corrupt / Invalid Input
  // ===========================================================================

  describe("corrupt input handling", () => {
    it("rejects empty Uint8Array", async () => {
      const wb = new Workbook();
      await expect(wb.xlsx.load(new Uint8Array(0))).rejects.toThrow();
    });

    it("rejects random bytes (not a ZIP file)", async () => {
      // Deterministic non-ZIP bytes (not starting with PK\x03\x04 magic)
      const garbage = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        garbage[i] = i;
      }

      const wb = new Workbook();
      await expect(wb.xlsx.load(garbage)).rejects.toThrow();
    });

    it("rejects a truncated ZIP file", async () => {
      const validBuffer = await buildMinimalXlsx();
      // Take only the first half
      const truncated = validBuffer.slice(0, Math.floor(validBuffer.length / 2));

      const wb = new Workbook();
      await expect(wb.xlsx.load(truncated)).rejects.toThrow();
    });

    it("rejects a JPEG file passed as XLSX", async () => {
      // JPEG magic bytes: FF D8 FF E0
      const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46]);

      const wb = new Workbook();
      await expect(wb.xlsx.load(jpeg)).rejects.toThrow();
    });

    it("rejects a plain text file", async () => {
      const text = new TextEncoder().encode("This is not an XLSX file");

      const wb = new Workbook();
      await expect(wb.xlsx.load(text)).rejects.toThrow();
    });
  });

  // ===========================================================================
  // write() / read() — Stream API
  // ===========================================================================

  describe("write(stream) / read(stream)", () => {
    it("write() pipes valid XLSX to a writable stream", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Stream").getCell("A1").value = "streamed";

      const output = new PassThrough();
      const chunks: Uint8Array[] = [];
      output.on("data", (chunk: Uint8Array) => chunks.push(chunk));

      await wb.xlsx.write(output);

      const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
      expect(totalLength).toBeGreaterThan(0);

      // Verify the stream output is a valid XLSX
      const buf = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        buf.set(chunk, offset);
        offset += chunk.length;
      }

      const wb2 = new Workbook();
      await wb2.xlsx.load(buf);
      expect(wb2.getWorksheet("Stream")!.getCell("A1").value).toBe("streamed");
    });

    it("read() loads workbook from a readable stream", async () => {
      const buffer = await buildMinimalXlsx();

      // Create a readable stream from the buffer
      const input = new PassThrough();
      input.end(buffer);

      const wb = new Workbook();
      await wb.xlsx.read(input);

      expect(wb.worksheets.length).toBe(1);
      expect(wb.getWorksheet("Sheet1")!.getCell("A1").value).toBe("hello");
    });

    it("read() from stream matches load() from buffer", async () => {
      const buffer = await buildRichXlsx();

      // Load via buffer
      const wb1 = new Workbook();
      await wb1.xlsx.load(buffer);

      // Load via stream
      const input = new PassThrough();
      input.end(buffer);
      const wb2 = new Workbook();
      await wb2.xlsx.read(input);

      // Same worksheets
      expect(wb1.worksheets.length).toBe(wb2.worksheets.length);
      expect(wb1.worksheets.map(ws => ws.name)).toEqual(wb2.worksheets.map(ws => ws.name));

      // Same values
      const ws1 = wb1.getWorksheet("Data")!;
      const ws2 = wb2.getWorksheet("Data")!;
      expect(ws1.getCell("A2").value).toEqual(ws2.getCell("A2").value);
      expect(ws1.getCell("B2").value).toEqual(ws2.getCell("B2").value);
    });
  });

  // ===========================================================================
  // Write Options
  // ===========================================================================

  describe("write options", () => {
    it("useSharedStrings option produces valid output", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "shared";
      ws.getCell("A2").value = "shared"; // same string

      const buffer = await wb.xlsx.writeBuffer({ useSharedStrings: true });
      expect(buffer.length).toBeGreaterThan(0);

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      expect(wb2.getWorksheet("Sheet1")!.getCell("A1").value).toBe("shared");
      expect(wb2.getWorksheet("Sheet1")!.getCell("A2").value).toBe("shared");
    });

    it("useStyles option produces valid output", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 42;
      ws.getCell("A1").font = { bold: true };

      const buffer = await wb.xlsx.writeBuffer({ useStyles: true });
      expect(buffer.length).toBeGreaterThan(0);

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);
      expect(wb2.getWorksheet("Sheet1")!.getCell("A1").font!.bold).toBe(true);
    });

    it("different compression levels produce valid output", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1").getCell("A1").value = "compress";

      const bufStore = await wb.xlsx.writeBuffer({ zip: { level: 0 } }); // STORE
      const bufDeflate = await wb.xlsx.writeBuffer({ zip: { level: 9 } }); // max compression

      // Level 0 (STORE) should be larger than level 9
      expect(bufStore.length).toBeGreaterThan(bufDeflate.length);

      // Both should be readable
      for (const buf of [bufStore, bufDeflate]) {
        const wb2 = new Workbook();
        await wb2.xlsx.load(buf);
        expect(wb2.getWorksheet("Sheet1")!.getCell("A1").value).toBe("compress");
      }
    });
  });

  // ===========================================================================
  // Double write / Double load
  // ===========================================================================

  describe("multiple operations", () => {
    it("writeBuffer() can be called multiple times on same workbook", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Sheet1").getCell("A1").value = "test";

      const buf1 = await wb.xlsx.writeBuffer();
      const buf2 = await wb.xlsx.writeBuffer();

      // Both buffers should be valid
      const wb1 = new Workbook();
      await wb1.xlsx.load(buf1);
      const wb2 = new Workbook();
      await wb2.xlsx.load(buf2);

      expect(wb1.getWorksheet("Sheet1")!.getCell("A1").value).toBe("test");
      expect(wb2.getWorksheet("Sheet1")!.getCell("A1").value).toBe("test");
    });

    it("load() replaces existing workbook content", async () => {
      const buf1 = await buildMinimalXlsx();

      const wb2 = new Workbook();
      wb2.addWorksheet("Original").getCell("A1").value = "original";

      // Load overwrites
      await wb2.xlsx.load(buf1);
      expect(wb2.getWorksheet("Original")).toBeUndefined();
      expect(wb2.getWorksheet("Sheet1")!.getCell("A1").value).toBe("hello");
    });
  });

  // ===========================================================================
  // Large Data
  // ===========================================================================

  describe("large data", () => {
    it("handles 10,000 rows via writeBuffer/load", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Big");
      for (let i = 1; i <= 10000; i++) {
        ws.addRow([i, `row-${i}`, i * 0.1]);
      }

      const buffer = await wb.xlsx.writeBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("Big")!;
      expect(ws2.getCell("A1").value).toBe(1);
      expect(ws2.getCell("A10000").value).toBe(10000);
      expect(ws2.getCell("B10000").value).toBe("row-10000");
    });
  });

  // ===========================================================================
  // Internal Hyperlinks (exceljs/exceljs#3027)
  // ===========================================================================

  describe("internal hyperlinks", () => {
    it("should round-trip internal hyperlink without duplication", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = {
        text: "Go to Sheet2",
        hyperlink: "#Sheet2!A1"
      };

      const buffer = await wb.xlsx.writeBuffer();

      // Inspect the raw XML to verify correct OOXML structure
      const { extractAll } = await import("@archive/unzip/extract");
      const entries = await extractAll(buffer);

      // Check sheet XML for hyperlink element
      const sheetXml = new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml")!.data);

      // Internal hyperlink should have location attribute, no r:id
      expect(sheetXml).toContain("location=");
      expect(sheetXml).toMatch(/hyperlink[^>]*location="Sheet2!A1"/);
      // Should NOT have r:id on internal hyperlink
      expect(sheetXml).not.toMatch(/hyperlink[^>]*r:id="[^"]*"[^>]*location=/);

      // There should be no sheet rels file for hyperlinks (only internal links)
      const sheetRels = entries.get("xl/worksheets/_rels/sheet1.xml.rels");
      if (sheetRels) {
        const relsXml = new TextDecoder().decode(sheetRels.data);
        // Should not contain hyperlink relationship for internal link
        expect(relsXml).not.toContain("hyperlink");
      }

      // Verify round-trip
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const cell = wb2.getWorksheet("Sheet1")!.getCell("A1");
      expect(cell.text).toBe("Go to Sheet2");
      // The hyperlink should be "#Sheet2!A1", not "#Sheet2!A1##Sheet2!A1"
      expect(cell.hyperlink).toBe("#Sheet2!A1");
    });
  });

  // ===========================================================================
  // Invalid XML Characters in XLSX (Regression)
  // ===========================================================================

  describe("invalid XML characters in XLSX", () => {
    it("reads XLSX with 0x7F (DEL) in shared string without crashing", async () => {
      const dirtyBuffer = await buildDirtyXlsx("\x7f");
      const wb = new Workbook();
      await wb.xlsx.load(dirtyBuffer);

      const ws = wb.getWorksheet("Sheet1")!;
      expect(ws).toBeDefined();
      // The invalid char should be stripped; "hello" should survive
      expect(ws.getCell("A1").text).toContain("hello");
    });

    it("reads XLSX with multiple control chars in shared string", async () => {
      const dirtyBuffer = await buildDirtyXlsx("\x01\x02\x03\x7f");
      const wb = new Workbook();
      await wb.xlsx.load(dirtyBuffer);

      const ws = wb.getWorksheet("Sheet1")!;
      expect(ws).toBeDefined();
      expect(ws.getCell("A1").text).toContain("hello");
    });

    it("reads dirty XLSX via streaming WorkbookReader", async () => {
      const dirtyBuffer = await buildDirtyXlsx("\x7f");
      const rows: string[] = [];

      const reader = new WorkbookReader(dirtyBuffer, { worksheets: "emit" });
      for await (const ws of reader) {
        for await (const row of ws) {
          const cell = row.getCell(1);
          if (cell.text) {
            rows.push(cell.text);
          }
        }
      }

      expect(rows.length).toBeGreaterThan(0);
      expect(rows[0]).toContain("hello");
    });

    it("reads XLSX with NUL bytes in shared string", async () => {
      const dirtyBuffer = await buildDirtyXlsx("\x00\x00\x00");
      const wb = new Workbook();
      await wb.xlsx.load(dirtyBuffer);

      const ws = wb.getWorksheet("Sheet1")!;
      expect(ws).toBeDefined();
      expect(ws.getCell("A1").text).toContain("hello");
    });
  });

  // ===========================================================================
  // Data Bar Conditional Formatting (exceljs/exceljs#3015)
  // ===========================================================================

  describe("data bar conditional formatting", () => {
    async function getSheetXml(buffer: Uint8Array): Promise<string> {
      const { extractAll } = await import("@archive/unzip/extract");
      const entries = await extractAll(buffer);
      return new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml")!.data);
    }

    function addDataBarWorkbook(ruleOverrides: Record<string, unknown> = {}): Promise<Uint8Array> {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      for (let i = 1; i <= 10; i++) {
        ws.getCell(`A${i}`).value = i * 10;
      }
      ws.addConditionalFormatting({
        ref: "A1:A10",
        rules: [{ type: "dataBar", priority: 1, ...ruleOverrides } as any]
      });
      return wb.xlsx.writeBuffer();
    }

    it("should produce valid data bar with default settings", async () => {
      const buf = await addDataBarWorkbook();
      const xml = await getSheetXml(buf);

      // Primary section must have <dataBar> with <cfvo> and <color>
      expect(xml).toMatch(/<dataBar>/);
      expect(xml).toMatch(/<cfvo type="min"/);
      expect(xml).toMatch(/<cfvo type="max"/);
      expect(xml).toMatch(/<color rgb="FF638EC6"/);

      // Primary section must have <x14:id> referencing ext section
      const primaryIdMatch = xml.match(/<x14:id>([^<]+)<\/x14:id>/);
      expect(primaryIdMatch).not.toBeNull();
      const x14Id = primaryIdMatch![1];
      expect(x14Id).toMatch(/^\{[0-9A-F-]+\}$/);

      // Ext section must have matching <x14:cfRule id="...">
      expect(xml).toContain(`<x14:cfRule type="dataBar" id="${x14Id}"`);
      expect(xml).toMatch(/<x14:dataBar/);
    });

    it("should produce valid data bar with explicit gradient=true", async () => {
      const buf = await addDataBarWorkbook({ gradient: true });
      const xml = await getSheetXml(buf);

      // Must still have ext section even when gradient is true
      const primaryIdMatch = xml.match(/<x14:id>([^<]+)<\/x14:id>/);
      expect(primaryIdMatch).not.toBeNull();
      const x14Id = primaryIdMatch![1];
      expect(xml).toContain(`<x14:cfRule type="dataBar" id="${x14Id}"`);
    });

    it("should produce valid data bar with gradient=false", async () => {
      const buf = await addDataBarWorkbook({ gradient: false });
      const xml = await getSheetXml(buf);

      const primaryIdMatch = xml.match(/<x14:id>([^<]+)<\/x14:id>/);
      expect(primaryIdMatch).not.toBeNull();
      const x14Id = primaryIdMatch![1];
      expect(xml).toContain(`<x14:cfRule type="dataBar" id="${x14Id}"`);

      // gradient=false should appear in ext section
      expect(xml).toMatch(/x14:dataBar[^>]*gradient="0"/);
    });

    it("should round-trip data bar without corruption", async () => {
      const buf1 = await addDataBarWorkbook();

      // Read it back
      const wb2 = new Workbook();
      await wb2.xlsx.load(buf1);

      // Verify conditional formatting survived
      const ws = wb2.getWorksheet("Sheet1")!;
      const cfs = ws.conditionalFormattings;
      expect(cfs.length).toBeGreaterThan(0);
      const dbRule = cfs[0].rules.find((r: any) => r.type === "dataBar");
      expect(dbRule).toBeDefined();

      // Write again
      const buf2 = await wb2.xlsx.writeBuffer();
      const xml = await getSheetXml(buf2);

      // Must still have matching primary + ext IDs after round-trip
      const primaryIdMatch = xml.match(/<x14:id>([^<]+)<\/x14:id>/);
      expect(primaryIdMatch).not.toBeNull();
      const x14Id = primaryIdMatch![1];
      expect(xml).toContain(`<x14:cfRule type="dataBar" id="${x14Id}"`);
    });
  });

  // ===========================================================================
  // Dynamic Array Formulas (exceljs/exceljs#2910)
  // ===========================================================================

  describe("dynamic array formulas", () => {
    async function getZipEntry(buffer: Uint8Array, path: string): Promise<string | undefined> {
      const { extractAll } = await import("@archive/unzip/extract");
      const entries = await extractAll(buffer);
      const entry = entries.get(path);
      return entry ? new TextDecoder().decode(entry.data) : undefined;
    }

    function createDynamicArrayWorkbook(): Promise<Uint8Array> {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      // Source data
      for (let i = 1; i <= 10; i++) {
        ws.getCell(`A${i}`).value = i;
        ws.getCell(`B${i}`).value = i > 5 ? 1 : 0;
      }
      // Dynamic array formula
      ws.getCell("D1").value = {
        formula: "_xlfn._xlws.FILTER(A1:A10,B1:B10=1)",
        shareType: "array",
        ref: "D1",
        result: 6,
        isDynamicArray: true
      };
      return wb.xlsx.writeBuffer();
    }

    it("should write cm attribute and metadata.xml for dynamic array formula", async () => {
      const buf = await createDynamicArrayWorkbook();

      // Check sheet XML has cm attribute
      const sheetXml = await getZipEntry(buf, "xl/worksheets/sheet1.xml");
      expect(sheetXml).toBeDefined();
      expect(sheetXml).toMatch(/<c r="D1"[^>]*cm="1"/);
      expect(sheetXml).toContain('t="array"');
      expect(sheetXml).toContain("_xlfn._xlws.FILTER");

      // Check metadata.xml exists
      const metadataXml = await getZipEntry(buf, "xl/metadata.xml");
      expect(metadataXml).toBeDefined();
      expect(metadataXml).toContain("XLDAPR");
      expect(metadataXml).toContain("xda:dynamicArrayProperties");
      expect(metadataXml).toContain('fDynamic="1"');

      // Check Content_Types.xml has metadata override
      const ctXml = await getZipEntry(buf, "[Content_Types].xml");
      expect(ctXml).toContain("sheetMetadata+xml");

      // Check workbook.xml.rels has metadata relationship
      const relsXml = await getZipEntry(buf, "xl/_rels/workbook.xml.rels");
      expect(relsXml).toContain("sheetMetadata");
      expect(relsXml).toContain("metadata.xml");
    });

    it("should not write metadata.xml when no dynamic array formulas exist", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { formula: "SUM(B1:B10)", result: 55 };
      const buf = await wb.xlsx.writeBuffer();

      const metadataXml = await getZipEntry(buf, "xl/metadata.xml");
      expect(metadataXml).toBeUndefined();

      const ctXml = await getZipEntry(buf, "[Content_Types].xml");
      expect(ctXml).not.toContain("sheetMetadata");
    });

    it("should round-trip dynamic array formula", async () => {
      const buf1 = await createDynamicArrayWorkbook();

      // Read back
      const wb2 = new Workbook();
      await wb2.xlsx.load(buf1);
      const ws2 = wb2.getWorksheet("Sheet1")!;
      const cellValue = ws2.getCell("D1").value as any;

      expect(cellValue).toBeDefined();
      expect(cellValue.formula).toBe("_xlfn._xlws.FILTER(A1:A10,B1:B10=1)");
      expect(cellValue.shareType).toBe("array");
      expect(cellValue.isDynamicArray).toBe(true);

      // Write again
      const buf2 = await wb2.xlsx.writeBuffer();

      // Verify second generation
      const sheetXml = await getZipEntry(buf2, "xl/worksheets/sheet1.xml");
      expect(sheetXml).toMatch(/<c r="D1"[^>]*cm="1"/);
      const metadataXml = await getZipEntry(buf2, "xl/metadata.xml");
      expect(metadataXml).toContain("XLDAPR");
    });

    it("should handle multiple dynamic array formulas in same workbook", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = {
        formula: "_xlfn._xlws.SORT(B1:B10)",
        shareType: "array",
        ref: "A1",
        result: 1,
        isDynamicArray: true
      };
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.UNIQUE(D1:D10)",
        shareType: "array",
        ref: "C1",
        result: "a",
        isDynamicArray: true
      };
      const buf = await wb.xlsx.writeBuffer();

      const sheetXml = (await getZipEntry(buf, "xl/worksheets/sheet1.xml"))!;
      // Both cells should have cm="1"
      expect(sheetXml).toMatch(/<c r="A1"[^>]*cm="1"/);
      expect(sheetXml).toMatch(/<c r="C1"[^>]*cm="1"/);

      // Only one metadata record needed
      const metadataXml = (await getZipEntry(buf, "xl/metadata.xml"))!;
      expect(metadataXml).toContain('<cellMetadata count="1"');
    });

    it("should handle mixed CSE array + dynamic array in same workbook", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      // Legacy CSE array formula
      ws.getCell("A1").value = {
        formula: "{ROW(1:3)}",
        shareType: "array",
        ref: "A1:A3",
        result: 1
      };
      // Dynamic array formula
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(B1:B10)",
        shareType: "array",
        ref: "C1",
        result: 1,
        isDynamicArray: true
      };
      const buf = await wb.xlsx.writeBuffer();

      const sheetXml = (await getZipEntry(buf, "xl/worksheets/sheet1.xml"))!;
      // Only the dynamic array cell should have cm
      expect(sheetXml).toMatch(/<c r="C1"[^>]*cm="1"/);
      // CSE cell should NOT have cm
      expect(sheetXml).not.toMatch(/<c r="A1"[^>]*cm="/);

      // Round-trip: read back and verify
      const wb2 = new Workbook();
      await wb2.xlsx.load(buf);
      const ws2 = wb2.getWorksheet("Sheet1")!;

      const cseVal = ws2.getCell("A1").value as any;
      expect(cseVal.formula).toBeDefined();
      expect(cseVal.isDynamicArray).toBeUndefined();

      const daVal = ws2.getCell("C1").value as any;
      expect(daVal.formula).toContain("SORT");
      expect(daVal.isDynamicArray).toBe(true);
    });

    it("should load Excel 365-style dynamic array XLSX built from raw ZIP parts", async () => {
      // Construct a minimal XLSX that mimics the exact structure Excel 365 produces
      // for a workbook with a FILTER() dynamic array formula in D1.
      // This tests our parser against externally-authored files, not just our own output.
      const archive = new ZipArchive({ level: 0, reproducible: true });
      const enc = new TextEncoder();

      // [Content_Types].xml
      archive.add(
        "[Content_Types].xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '  <Default Extension="xml" ContentType="application/xml"/>',
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
            '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
            '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
            '  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>',
            "</Types>"
          ].join("")
        )
      );

      // _rels/.rels
      archive.add(
        "_rels/.rels",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
            "</Relationships>"
          ].join("")
        )
      );

      // xl/_rels/workbook.xml.rels
      archive.add(
        "xl/_rels/workbook.xml.rels",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
            '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
            '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
            '  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>',
            "</Relationships>"
          ].join("")
        )
      );

      // xl/workbook.xml
      archive.add(
        "xl/workbook.xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            "  <sheets>",
            '    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>',
            "  </sheets>",
            "</workbook>"
          ].join("")
        )
      );

      // xl/styles.xml — minimal
      archive.add(
        "xl/styles.xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
            '  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
            '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
            '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
            '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
            "</styleSheet>"
          ].join("")
        )
      );

      // xl/sharedStrings.xml — empty
      archive.add(
        "xl/sharedStrings.xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>'
          ].join("")
        )
      );

      // xl/metadata.xml — exact Excel 365 structure
      archive.add(
        "xl/metadata.xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            '  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">',
            '  <metadataTypes count="1">',
            '    <metadataType name="XLDAPR" minSupportedVersion="120000"',
            '      copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1"',
            '      rowColShift="1" clearFormats="1" clearComments="1" assign="1"',
            '      coerce="1" adjust="1" cellMeta="1"/>',
            "  </metadataTypes>",
            '  <futureMetadata name="XLDAPR" count="1">',
            "    <bk>",
            "      <extLst>",
            '        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">',
            '          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
            "        </ext>",
            "      </extLst>",
            "    </bk>",
            "  </futureMetadata>",
            '  <cellMetadata count="1">',
            "    <bk>",
            '      <rc t="1" v="0"/>',
            "    </bk>",
            "  </cellMetadata>",
            "</metadata>"
          ].join("\n")
        )
      );

      // xl/worksheets/sheet1.xml — with cm="1" on D1
      archive.add(
        "xl/worksheets/sheet1.xml",
        enc.encode(
          [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '  <dimension ref="A1:D5"/>',
            "  <sheetData>",
            '    <row r="1">',
            '      <c r="A1"><v>10</v></c>',
            '      <c r="B1"><v>1</v></c>',
            '      <c r="D1" cm="1"><f t="array" ref="D1">_xlfn._xlws.FILTER(A1:A5,B1:B5=1)</f><v>10</v></c>',
            "    </row>",
            '    <row r="2">',
            '      <c r="A2"><v>20</v></c>',
            '      <c r="B2"><v>0</v></c>',
            "    </row>",
            '    <row r="3">',
            '      <c r="A3"><v>30</v></c>',
            '      <c r="B3"><v>1</v></c>',
            "    </row>",
            '    <row r="4">',
            '      <c r="A4"><v>40</v></c>',
            '      <c r="B4"><v>0</v></c>',
            "    </row>",
            '    <row r="5">',
            '      <c r="A5"><v>50</v></c>',
            '      <c r="B5"><v>1</v></c>',
            "    </row>",
            "  </sheetData>",
            "</worksheet>"
          ].join("\n")
        )
      );

      const xlsxBuffer = await archive.bytes();

      // ---- Test 1: Load and verify isDynamicArray ----
      const wb = new Workbook();
      await wb.xlsx.load(xlsxBuffer);
      const ws = wb.getWorksheet("Sheet1")!;

      // Regular cells should load normally
      expect(ws.getCell("A1").value).toBe(10);
      expect(ws.getCell("B1").value).toBe(1);

      // Dynamic array formula cell
      const d1 = ws.getCell("D1").value as any;
      expect(d1).toBeDefined();
      expect(d1.formula).toBe("_xlfn._xlws.FILTER(A1:A5,B1:B5=1)");
      expect(d1.shareType).toBe("array");
      expect(d1.isDynamicArray).toBe(true);
      expect(d1.result).toBe(10);

      // ---- Test 2: Round-trip — write back and verify structure ----
      const buf2 = await wb.xlsx.writeBuffer();
      const sheetXml = await getZipEntry(buf2, "xl/worksheets/sheet1.xml");
      expect(sheetXml).toMatch(/<c r="D1"[^>]*cm="1"/);

      const metadataXml = await getZipEntry(buf2, "xl/metadata.xml");
      expect(metadataXml).toBeDefined();
      expect(metadataXml).toContain("XLDAPR");

      // ---- Test 3: Streaming reader should also detect isDynamicArray ----
      const reader = new WorkbookReader(xlsxBuffer, {
        worksheets: "emit",
        sharedStrings: "cache"
      });

      let streamDaFound = false;
      for await (const wsReader of reader) {
        for await (const row of wsReader) {
          const cell = row.getCell(4);
          const val = cell.value as any;
          if (val && typeof val === "object" && val.formula) {
            streamDaFound = true;
            expect(val.formula).toContain("FILTER");
            expect(val.isDynamicArray).toBe(true);
          }
        }
      }
      expect(streamDaFound).toBe(true);
    });

    it("should not mark isDynamicArray when cm exists but metadata has no XLDAPR", async () => {
      // Construct a workbook with a dynamic array formula, then manually strip
      // the XLDAPR type from metadata.xml to simulate a non-XLDAPR cm reference.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = {
        formula: "_xlfn._xlws.FILTER(B1:B5,B1:B5>0)",
        shareType: "array",
        ref: "A1",
        result: 1,
        isDynamicArray: true
      };
      const buf = await wb.xlsx.writeBuffer();

      // Extract ZIP, replace metadata.xml with one that has NO XLDAPR type
      const { extractAll } = await import("@archive/unzip/extract");
      const { ZipArchive } = await import("@archive/zip");
      const entries = await extractAll(buf);

      const fakeMetadata = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '  <metadataTypes count="1">',
        '    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" cellMeta="1"/>',
        "  </metadataTypes>",
        '  <cellMetadata count="1">',
        "    <bk>",
        '      <rc t="1" v="0"/>',
        "    </bk>",
        "  </cellMetadata>",
        "</metadata>"
      ].join("\n");

      const encoder = new TextEncoder();
      entries.get("xl/metadata.xml")!.data = encoder.encode(fakeMetadata);

      const archive = new ZipArchive({ level: 0, reproducible: true });
      for (const [name, entry] of entries) {
        archive.add(name, entry.data);
      }
      const tamperedBuf = await archive.bytes();

      // Load the tampered file
      const wb2 = new Workbook();
      await wb2.xlsx.load(tamperedBuf);
      const cellVal = wb2.getWorksheet("Sheet1")!.getCell("A1").value as any;

      // cm was present but mapped to XLRICHVALUE, not XLDAPR — should NOT be isDynamicArray
      expect(cellVal.formula).toContain("FILTER");
      expect(cellVal.isDynamicArray).toBeUndefined();
    });
  });
});
