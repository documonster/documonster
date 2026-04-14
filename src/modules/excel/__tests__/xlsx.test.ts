import { describe, it, expect } from "vitest";
import { PassThrough } from "@stream";
import { Workbook, WorkbookReader } from "../../../index";
import { ZipArchive } from "@archive/zip";

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
});
