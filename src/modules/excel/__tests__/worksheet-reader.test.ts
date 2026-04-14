import type { Row } from "@excel/row";
import { describe, it, expect } from "vitest";

import { Workbook, WorkbookReader, ExcelStreamStateError } from "../../../index";

// =============================================================================
// Helpers
// =============================================================================

/**
 * Create a minimal XLSX buffer from a builder callback.
 * Avoids filesystem I/O — fast + cross-platform.
 */
async function buildXlsxBuffer(builder: (wb: Workbook) => void): Promise<Uint8Array> {
  const wb = new Workbook();
  builder(wb);
  return wb.xlsx.writeBuffer();
}

/** Result of streaming a single worksheet */
interface SheetResult {
  name: string;
  id: number | string;
  sheetNo: number;
  rows: Row[];
  dimensions: unknown;
  columns: unknown[];
}

/**
 * Read an XLSX buffer with the streaming WorkbookReader and collect all
 * worksheets + their rows. This consumes everything inside the reader's
 * iteration context (the only correct way to use the async iterator API).
 */
async function readAllSheets(
  buffer: Uint8Array,
  readerOptions?: Record<string, string>
): Promise<SheetResult[]> {
  const reader = new WorkbookReader(buffer, {
    worksheets: "emit",
    sharedStrings: "cache",
    styles: "cache",
    hyperlinks: "cache",
    entries: "emit",
    ...readerOptions
  });

  const sheets: SheetResult[] = [];
  for await (const ws of reader) {
    const rows: Row[] = [];
    for await (const row of ws) {
      rows.push(row as Row);
    }
    sheets.push({
      name: ws.name,
      id: ws.id,
      sheetNo: (ws as any).sheetNo,
      rows,
      dimensions: (ws as any).dimensions,
      columns: (ws as any).columns ?? []
    });
  }
  return sheets;
}

/**
 * Convenience: read a single-sheet buffer and return its rows + metadata.
 */
async function readFirstSheet(
  buffer: Uint8Array,
  readerOptions?: Record<string, string>
): Promise<SheetResult> {
  const sheets = await readAllSheets(buffer, readerOptions);
  if (sheets.length === 0) {
    throw new Error("No worksheet found in buffer");
  }
  return sheets[0];
}

// =============================================================================
// Tests
// =============================================================================

describe("WorksheetReader", () => {
  // ===========================================================================
  // Lifecycle
  // ===========================================================================

  describe("lifecycle", () => {
    it("destroy() throws ExcelStreamStateError", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = 1;
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        expect(() => ws.destroy()).toThrow(ExcelStreamStateError);
        // Must consume rows to avoid hanging
        for await (const _row of ws) {
          /* drain */
        }
      }
    });
  });

  // ===========================================================================
  // Properties
  // ===========================================================================

  describe("properties", () => {
    it("has correct name from workbook metadata", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("MySheet").getCell("A1").value = 1;
      });
      const sheet = await readFirstSheet(buffer);
      expect(sheet.name).toBe("MySheet");
    });

    it("exposes sheet id and sheetNo", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("First").getCell("A1").value = 1;
      });
      const sheet = await readFirstSheet(buffer);
      expect(sheet.id).toBeDefined();
      expect(sheet.sheetNo).toBeTypeOf("number");
    });
  });

  // ===========================================================================
  // Dimensions
  // ===========================================================================

  describe("dimensions", () => {
    it("reflects the parsed data extent", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 1;
        ws.getCell("C3").value = 2;
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(2);
      expect(sheet.dimensions).toBeDefined();
    });
  });

  // ===========================================================================
  // Columns
  // ===========================================================================

  describe("columns", () => {
    it("returns column definitions after parsing", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.columns = [
          { header: "ID", key: "id", width: 10 },
          { header: "Name", key: "name", width: 32 }
        ];
        ws.addRow({ id: 1, name: "Alice" });
      });
      const sheet = await readFirstSheet(buffer);

      // Columns should be populated from the <cols> XML element
      expect(sheet.columns.length).toBeGreaterThanOrEqual(2);
    });

    it("getColumn() by number creates columns on demand", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = 1;
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        // Getting column 5 should create columns 1-5
        const col = ws.getColumn(5);
        expect(col).toBeDefined();
        expect(ws.columns.length).toBeGreaterThanOrEqual(5);

        for await (const _row of ws) {
          /* drain */
        }
      }
    });

    it("getColumn() by letter resolves correctly", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = 1;
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        const colC = ws.getColumn("C");
        expect(colC).toBeDefined();
        expect(colC.number).toBe(3);

        for await (const _row of ws) {
          /* drain */
        }
      }
    });

    it("column key management works correctly", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = 1;
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        const col = ws.getColumn(1);
        ws.setColumnKey("myKey", col);
        expect(ws.getColumnKey("myKey")).toBe(col);

        // getColumn by key should resolve
        expect(ws.getColumn("myKey")).toBe(col);

        // eachColumnKey iterates
        const keys: string[] = [];
        ws.eachColumnKey((_column, key) => keys.push(key));
        expect(keys).toContain("myKey");

        // delete
        ws.deleteColumnKey("myKey");
        expect(ws.getColumnKey("myKey")).toBeUndefined();

        for await (const _row of ws) {
          /* drain */
        }
      }
    });
  });

  // ===========================================================================
  // Async Iteration — Cell Value Types
  // ===========================================================================

  describe("cell value types", () => {
    it("parses numbers", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 42;
        ws.getCell("B1").value = 3.14;
        ws.getCell("C1").value = -100;
        ws.getCell("D1").value = 0;
      });
      const sheet = await readFirstSheet(buffer);
      const row = sheet.rows[0];

      expect(row.getCell(1).value).toBe(42);
      expect(row.getCell(2).value).toBe(3.14);
      expect(row.getCell(3).value).toBe(-100);
      expect(row.getCell(4).value).toBe(0);
    });

    it("parses strings via shared strings", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "hello";
        ws.getCell("B1").value = "world";
        ws.getCell("A2").value = "hello"; // duplicate — same shared string
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe("hello");
      expect(sheet.rows[0].getCell(2).value).toBe("world");
      expect(sheet.rows[1].getCell(1).value).toBe("hello");
    });

    it("parses booleans", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = true;
        ws.getCell("B1").value = false;
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe(true);
      expect(sheet.rows[0].getCell(2).value).toBe(false);
    });

    it("parses dates when numFmt indicates a date format", async () => {
      const testDate = new Date(2024, 0, 15); // Jan 15, 2024
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = testDate;
        ws.getCell("A1").numFmt = "yyyy-mm-dd";
      });
      const sheet = await readFirstSheet(buffer);
      const cellValue = sheet.rows[0].getCell(1).value;

      expect(cellValue).toBeInstanceOf(Date);
      expect((cellValue as Date).getFullYear()).toBe(2024);
    });

    it("parses formulas with numeric results", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 10;
        ws.getCell("A2").value = { formula: "A1*2", result: 20 };
      });
      const sheet = await readFirstSheet(buffer);
      const val = sheet.rows[1].getCell(1).value as { formula: string; result: number };

      expect(val.formula).toBe("A1*2");
      expect(val.result).toBe(20);
    });

    it("parses formulas with string results", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = {
          formula: 'CONCATENATE("Hello", " ", "World")',
          result: "Hello World"
        };
      });
      const sheet = await readFirstSheet(buffer);
      const val = sheet.rows[0].getCell(1).value as { formula: string; result: string };

      expect(val.formula).toContain("CONCATENATE");
      expect(val.result).toBe("Hello World");
    });

    it("parses error values", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = { error: "#DIV/0!" as any };
      });
      const sheet = await readFirstSheet(buffer);
      const val = sheet.rows[0].getCell(1).value as { error: string };

      expect(val.error).toBe("#DIV/0!");
    });

    it("parses rich text via shared strings", async () => {
      const richText = {
        richText: [
          { font: { bold: true }, text: "Bold " },
          { font: { italic: true }, text: "Italic" }
        ]
      };
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = richText;
      });
      const sheet = await readFirstSheet(buffer);
      const val = sheet.rows[0].getCell(1).value as { richText: Array<{ text: string }> };

      expect(val).toHaveProperty("richText");
      expect(val.richText.length).toBe(2);
      expect(val.richText[0].text).toBe("Bold ");
      expect(val.richText[1].text).toBe("Italic");
    });
  });

  // ===========================================================================
  // Row Properties
  // ===========================================================================

  describe("row properties", () => {
    it("parses row height", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        const row = ws.getRow(1);
        row.height = 30;
        row.getCell(1).value = "tall row";
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].height).toBe(30);
    });

    it("parses row number correctly for sparse rows", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "row1";
        ws.getCell("A5").value = "row5"; // rows 2-4 are empty
      });
      const sheet = await readFirstSheet(buffer);

      // Only non-empty rows are emitted
      expect(sheet.rows.length).toBe(2);
      expect(sheet.rows[0].number).toBe(1);
      expect(sheet.rows[1].number).toBe(5);
    });
  });

  // ===========================================================================
  // Styles
  // ===========================================================================

  describe("styles", () => {
    it("parses cell styles when styles are cached", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 42;
        ws.getCell("A1").font = { bold: true, size: 14 };
        ws.getCell("A1").numFmt = "0.00%";
      });
      const sheet = await readFirstSheet(buffer);
      const cell = sheet.rows[0].getCell(1);

      expect(cell.value).toBe(42);
      expect(cell.font).toBeDefined();
      expect(cell.font!.bold).toBe(true);
      expect(cell.numFmt).toBe("0.00%");
    });
  });

  // ===========================================================================
  // Multi-column + Sparse Data
  // ===========================================================================

  describe("multi-column data", () => {
    it("correctly maps cells to columns across multiple rows", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.addRow([1, "Alice", true]);
        ws.addRow([2, "Bob", false]);
        ws.addRow([3, "Charlie", true]);
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(3);
      expect(sheet.rows[0].getCell(1).value).toBe(1);
      expect(sheet.rows[0].getCell(2).value).toBe("Alice");
      expect(sheet.rows[0].getCell(3).value).toBe(true);
      expect(sheet.rows[2].getCell(1).value).toBe(3);
      expect(sheet.rows[2].getCell(2).value).toBe("Charlie");
    });

    it("handles sparse rows (cells with gaps)", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 1;
        ws.getCell("D1").value = 4; // B1, C1 are empty
        ws.getCell("A2").value = "start";
        ws.getCell("F2").value = "end"; // B2-E2 are empty
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe(1);
      expect(sheet.rows[0].getCell(4).value).toBe(4);
      expect(sheet.rows[1].getCell(1).value).toBe("start");
      expect(sheet.rows[1].getCell(6).value).toBe("end");
    });
  });

  // ===========================================================================
  // Multiple Worksheets
  // ===========================================================================

  describe("multiple worksheets", () => {
    it("reads all worksheets in order with correct names and data", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("First").getCell("A1").value = "sheet1";
        wb.addWorksheet("Second").getCell("A1").value = "sheet2";
        wb.addWorksheet("Third").getCell("A1").value = "sheet3";
      });

      const sheets = await readAllSheets(buffer);

      expect(sheets.length).toBe(3);
      expect(sheets[0].name).toBe("First");
      expect(sheets[0].rows[0].getCell(1).value).toBe("sheet1");
      expect(sheets[1].name).toBe("Second");
      expect(sheets[1].rows[0].getCell(1).value).toBe("sheet2");
      expect(sheets[2].name).toBe("Third");
      expect(sheets[2].rows[0].getCell(1).value).toBe("sheet3");
    });
  });

  // ===========================================================================
  // read() Event-Based API
  // ===========================================================================

  describe("read() event API", () => {
    it("emits 'row' events and 'finished' at the end", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "one";
        ws.getCell("A2").value = "two";
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const result = await new Promise<{ rows: Row[]; finished: boolean }>((resolve, reject) => {
        const rows: Row[] = [];
        let finished = false;

        reader.on("worksheet", worksheet => {
          worksheet.on("row", (row: Row) => rows.push(row));
          worksheet.on("finished", () => {
            finished = true;
            resolve({ rows, finished });
          });
        });
        reader.on("error", reject);
        reader.read();
      });

      expect(result.finished).toBe(true);
      expect(result.rows.length).toBe(2);
      expect(result.rows[0].getCell(1).value).toBe("one");
      expect(result.rows[1].getCell(1).value).toBe("two");
    });
  });

  // ===========================================================================
  // Hyperlinks
  // ===========================================================================

  describe("hyperlinks", () => {
    it("caches hyperlinks when hyperlinks option is 'cache'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = {
          text: "Google",
          hyperlink: "https://www.google.com"
        };
        ws.getCell("A2").value = "no link";
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache",
        hyperlinks: "cache",
        entries: "emit"
      });

      for await (const ws of reader) {
        const rows: Row[] = [];
        for await (const row of ws) {
          rows.push(row as Row);
        }
        expect(rows.length).toBe(2);

        // hyperlinks should be cached on the WorksheetReader
        const hyperlinks = (ws as any).hyperlinks as
          | Record<string, { ref: string; rId: string }>
          | undefined;
        expect(hyperlinks).toBeDefined();
        expect(hyperlinks!["A1"]).toBeDefined();
        expect(hyperlinks!["A1"].rId).toBeTypeOf("string");
      }
    });
  });

  // ===========================================================================
  // Large Dataset
  // ===========================================================================

  describe("large dataset", () => {
    it("handles 1000 rows without issues", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        for (let i = 1; i <= 1000; i++) {
          ws.addRow([i, `name-${i}`, i * 1.5]);
        }
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(1000);
      expect(sheet.rows[0].getCell(1).value).toBe(1);
      expect(sheet.rows[999].getCell(1).value).toBe(1000);
      expect(sheet.rows[999].getCell(2).value).toBe("name-1000");
      expect(sheet.rows[999].getCell(3).value).toBe(1500);
    });
  });

  // ===========================================================================
  // Edge Cases
  // ===========================================================================

  describe("edge cases", () => {
    it("handles an empty worksheet", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Empty");
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows.length).toBe(0);
    });

    it("handles XML special characters in string values", async () => {
      const specialStr = "xml chars: & < > \" '";
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = specialStr;
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe(specialStr);
    });

    it("handles OOXML _xHHHH_ escape patterns in strings", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        wb.addWorksheet("Sheet1").getCell("A1").value = "_x000D_";
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe("_x000D_");
    });

    it("handles mixed value types in a single row", async () => {
      const now = new Date(2024, 5, 15);
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = 42;
        ws.getCell("B1").value = "text";
        ws.getCell("C1").value = true;
        ws.getCell("D1").value = now;
        ws.getCell("D1").numFmt = "yyyy-mm-dd";
      });
      const sheet = await readFirstSheet(buffer);
      const row = sheet.rows[0];

      expect(row.getCell(1).value).toBe(42);
      expect(row.getCell(2).value).toBe("text");
      expect(row.getCell(3).value).toBe(true);
      expect(row.getCell(4).value).toBeInstanceOf(Date);
    });

    it("handles unicode and CJK characters", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = wb.addWorksheet("Sheet1");
        ws.getCell("A1").value = "你好世界";
        ws.getCell("B1").value = "日本語テスト";
        ws.getCell("C1").value = "한국어";
        ws.getCell("D1").value = "Ñoño Año";
      });
      const sheet = await readFirstSheet(buffer);

      expect(sheet.rows[0].getCell(1).value).toBe("你好世界");
      expect(sheet.rows[0].getCell(2).value).toBe("日本語テスト");
      expect(sheet.rows[0].getCell(3).value).toBe("한국어");
      expect(sheet.rows[0].getCell(4).value).toBe("Ñoño Año");
    });
  });
});
