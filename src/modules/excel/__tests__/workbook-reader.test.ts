import { cellGetValue } from "@excel/cell";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { type RowData, rowCellCount } from "@excel/row";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import type { WorkbookData } from "@excel/workbook-core";
import { rowGetCell } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

// =============================================================================
// Helpers
// =============================================================================

/**
 * Create a minimal XLSX buffer from a builder callback.
 * Avoids filesystem I/O — fast + cross-platform.
 */
async function buildXlsxBuffer(builder: (wb: WorkbookData) => void): Promise<Uint8Array> {
  const wb = Workbook.create();
  builder(wb);
  return Workbook.toXlsxBuffer(wb);
}

// =============================================================================
// Tests
// =============================================================================

describe("WorkbookReader", () => {
  // ===========================================================================
  // Input Types
  // ===========================================================================

  describe("input types", () => {
    it("reads from Uint8Array", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "hello");
      });

      const reader = new WorkbookReader(buffer, { worksheets: "emit" });
      let seen = false;
      for await (const ws of reader) {
        seen = true;
        expect(ws.name).toBe("Sheet1");
        for await (const row of ws) {
          expect(cellGetValue(rowGetCell(row, 1))).toBe("hello");
        }
      }
      expect(seen).toBe(true);
    });

    it.skipIf(typeof ReadableStream === "undefined")(
      "reads from ReadableStream<Uint8Array>",
      async () => {
        const wb = Workbook.create();
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "hello");
        Cell.setValue(ws, "A2", 42);
        const data = await Workbook.toXlsxBuffer(wb);

        const webStream = new ReadableStream<Uint8Array>({
          start(controller) {
            const chunkSize = 64 * 1024;
            for (let i = 0; i < data.length; i += chunkSize) {
              controller.enqueue(data.slice(i, i + chunkSize));
            }
            controller.close();
          }
        });

        const reader = new WorkbookReader(webStream, { worksheets: "emit" });
        let seen = false;

        for await (const worksheet of reader) {
          seen = true;
          expect(worksheet.name).toBe("Sheet1");

          let rowCount = 0;
          for await (const row of worksheet) {
            rowCount++;
            if (row.number === 1) {
              expect(cellGetValue(rowGetCell(row, 1))).toBe("hello");
            }
            if (row.number === 2) {
              expect(cellGetValue(rowGetCell(row, 1))).toBe(42);
            }
          }
          expect(rowCount).toBeGreaterThan(0);
        }
        expect(seen).toBe(true);
      }
    );

    it("reads from ArrayBuffer", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", 99);
      });

      // Convert Uint8Array to ArrayBuffer
      const arrayBuffer: ArrayBuffer = buffer.buffer.slice(
        buffer.byteOffset,
        buffer.byteOffset + buffer.byteLength
      ) as ArrayBuffer;

      const reader = new WorkbookReader(arrayBuffer, { worksheets: "emit" });
      let found = false;
      for await (const ws of reader) {
        for await (const row of ws) {
          if (row.number === 1) {
            expect(cellGetValue(rowGetCell(row, 1))).toBe(99);
            found = true;
          }
        }
      }
      expect(found).toBe(true);
    });
  });

  // ===========================================================================
  // Async Iterator API
  // ===========================================================================

  describe("[Symbol.asyncIterator]", () => {
    it("yields WorksheetReader instances for each sheet", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Alpha"), "A1", 1);
        Cell.setValue(Workbook.addWorksheet(wb, "Beta"), "A1", 2);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const names: string[] = [];
      for await (const ws of reader) {
        names.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(names).toEqual(["Alpha", "Beta"]);
    });
  });

  // ===========================================================================
  // Event-Based read() API
  // ===========================================================================

  describe("read() event API", () => {
    it("emits worksheet and end events", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "test");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const result = await new Promise<{ wsNames: string[]; ended: boolean }>((resolve, reject) => {
        const wsNames: string[] = [];

        reader.on("worksheet", ws => {
          wsNames.push(ws.name);
          ws.on("row", () => {});
          ws.on("finished", () => {});
        });
        reader.on("end", () => resolve({ wsNames, ended: true }));
        reader.on("error", reject);
        reader.read();
      });

      expect(result.ended).toBe(true);
      expect(result.wsNames).toContain("Sheet1");
    });

    it("emits error for corrupted input", async () => {
      const garbage = new Uint8Array([0, 1, 2, 3, 4, 5]);
      const reader = new WorkbookReader(garbage, { worksheets: "emit" });

      const error = await new Promise<Error>((resolve, reject) => {
        reader.on("error", resolve);
        reader.on("end", () => reject(new Error("Expected error but got end")));
        reader.read();
      });

      expect(error).toBeInstanceOf(Error);
    });
  });

  // ===========================================================================
  // Reader Options
  // ===========================================================================

  describe("reader options", () => {
    it("worksheets: 'ignore' skips worksheet parsing", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet1"), "A1", "data");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "ignore",
        sharedStrings: "cache",
        styles: "cache"
      });

      const sheets: string[] = [];
      for await (const ws of reader) {
        sheets.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }

      // With worksheets: "ignore", no worksheet events should be emitted
      expect(sheets.length).toBe(0);
    });

    it("sharedStrings: 'cache' enables shared string resolution", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", "shared text");
        Cell.setValue(ws, "A2", "shared text"); // same string
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        const rows: RowData[] = [];
        for await (const row of ws) {
          rows.push(row as RowData);
        }
        expect(cellGetValue(rowGetCell(rows[0], 1))).toBe("shared text");
        expect(cellGetValue(rowGetCell(rows[1], 1))).toBe("shared text");
      }
    });

    it("styles: 'cache' enables date parsing", async () => {
      const testDate = new Date(2024, 0, 15);
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", testDate);
        Cell.setStyle(ws, "A1", { numFmt: "yyyy-mm-dd" });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        for await (const row of ws) {
          const val = cellGetValue(rowGetCell(row, 1));
          expect(val).toBeInstanceOf(Date);
        }
      }
    });
  });

  // ===========================================================================
  // Multiple Worksheets
  // ===========================================================================

  describe("multiple worksheets", () => {
    it("reads all worksheets with correct names and data", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet A"), "A1", 100);
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet B"), "A1", 200);
        Cell.setValue(Workbook.addWorksheet(wb, "Sheet C"), "A1", 300);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const results: Array<{ name: string; value: unknown }> = [];
      for await (const ws of reader) {
        for await (const row of ws) {
          results.push({ name: ws.name, value: cellGetValue(rowGetCell(row, 1)) });
        }
      }

      expect(results.length).toBe(3);
      expect(results[0]).toEqual({ name: "Sheet A", value: 100 });
      expect(results[1]).toEqual({ name: "Sheet B", value: 200 });
      expect(results[2]).toEqual({ name: "Sheet C", value: 300 });
    });

    it("reads worksheets that have different column structures", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws1 = Workbook.addWorksheet(wb, "Numbers");
        Worksheet.addRow(ws1, [1, 2, 3]);

        const ws2 = Workbook.addWorksheet(wb, "Strings");
        Worksheet.addRow(ws2, ["a", "b", "c", "d", "e"]);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const sheets: Array<{ name: string; colCount: number }> = [];
      for await (const ws of reader) {
        for await (const row of ws) {
          const cellCount = rowCellCount(row);
          sheets.push({ name: ws.name, colCount: cellCount });
        }
      }

      expect(sheets[0].name).toBe("Numbers");
      expect(sheets[0].colCount).toBe(3);
      expect(sheets[1].name).toBe("Strings");
      expect(sheets[1].colCount).toBe(5);
    });
  });

  // ===========================================================================
  // Hyperlinks
  // ===========================================================================

  describe("hyperlinks", () => {
    it("getHyperlinkTarget() resolves hyperlink rId to target URL", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", {
          text: "Google",
          hyperlink: "https://www.google.com"
        });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache",
        hyperlinks: "cache",
        entries: "emit"
      });

      const targets: string[] = [];
      for await (const ws of reader) {
        for await (const _row of ws) {
          /* drain rows first — hyperlinks are parsed from a separate zip entry */
        }

        const hyperlinks = (ws as any).hyperlinks as
          | Record<string, { ref: string; rId: string }>
          | undefined;
        expect(hyperlinks).toBeDefined();
        for (const ref in hyperlinks) {
          const target = reader.getHyperlinkTarget((ws as any).sheetNo, hyperlinks[ref].rId);
          expect(target).toBeTypeOf("string");
          targets.push(target!);
        }
      }

      expect(targets).toEqual(["https://www.google.com"]);
    });
  });

  // ===========================================================================
  // Edge Cases
  // ===========================================================================

  describe("edge cases", () => {
    it("handles workbook with empty worksheet", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Workbook.addWorksheet(wb, "Empty");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      for await (const ws of reader) {
        expect(ws.name).toBe("Empty");
        const rows: RowData[] = [];
        for await (const row of ws) {
          rows.push(row as RowData);
        }
        expect(rows.length).toBe(0);
      }
    });

    it("handles workbook with many worksheets", async () => {
      const sheetCount = 20;
      const buffer = await buildXlsxBuffer(wb => {
        for (let i = 1; i <= sheetCount; i++) {
          Cell.setValue(Workbook.addWorksheet(wb, `Sheet${i}`), "A1", i);
        }
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const names: string[] = [];
      for await (const ws of reader) {
        names.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(names.length).toBe(sheetCount);
    });

    it("handles workbook with large shared string table", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        for (let i = 1; i <= 500; i++) {
          Cell.setValue(ws, `A${i}`, `unique-string-${i}`);
        }
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      let rowCount = 0;
      for await (const ws of reader) {
        for await (const row of ws) {
          rowCount++;
          expect(cellGetValue(rowGetCell(row, 1))).toBe(`unique-string-${row.number}`);
        }
      }
      expect(rowCount).toBe(500);
    });
  });

  // ===========================================================================
  // Worksheet Name Resolution (GitHub exceljs/exceljs#3025)
  // ===========================================================================

  describe("worksheet name resolution", () => {
    it("preserves worksheet names with sharedStrings: 'cache'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Alpha"), "A1", 1);
        Cell.setValue(Workbook.addWorksheet(wb, "Beta"), "A1", 2);
        Cell.setValue(Workbook.addWorksheet(wb, "Gamma"), "A1", 3);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache",
        styles: "cache"
      });

      const names: string[] = [];
      for await (const ws of reader) {
        names.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(names).toEqual(["Alpha", "Beta", "Gamma"]);
    });

    it("preserves worksheet names with sharedStrings: 'ignore'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "First Sheet"), "A1", 1);
        Cell.setValue(Workbook.addWorksheet(wb, "Second Sheet"), "A1", 2);
        Cell.setValue(Workbook.addWorksheet(wb, "Third Sheet"), "A1", 3);
      });

      // This is the exact scenario from exceljs/exceljs#3025:
      // When sharedStrings is NOT "cache", the hasPrerequisites check
      // may pass before workbook.xml is parsed, causing worksheet names
      // to be default ("Sheet1", "Sheet2", ...) instead of the actual names.
      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "ignore",
        styles: "ignore"
      });

      const names: string[] = [];
      for await (const ws of reader) {
        names.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(names).toEqual(["First Sheet", "Second Sheet", "Third Sheet"]);
    });

    it("preserves worksheet names via event API with sharedStrings: 'ignore'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Report"), "A1", "data");
        Cell.setValue(Workbook.addWorksheet(wb, "Summary"), "A1", "summary");
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "ignore",
        styles: "ignore"
      });

      const names: string[] = [];
      await new Promise<void>((resolve, reject) => {
        reader.on("worksheet", ws => {
          names.push(ws.name);
          ws.on("row", () => {});
          ws.on("finished", () => {});
        });
        reader.on("end", () => resolve());
        reader.on("error", reject);
        reader.read();
      });

      expect(names).toEqual(["Report", "Summary"]);
    });

    it("preserves worksheet names with sharedStrings: 'emit'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Cell.setValue(Workbook.addWorksheet(wb, "Emit Sheet A"), "A1", 1);
        Cell.setValue(Workbook.addWorksheet(wb, "Emit Sheet B"), "A1", 2);
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "emit",
        styles: "ignore"
      });

      const names: string[] = [];
      for await (const ws of reader) {
        names.push(ws.name);
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(names).toEqual(["Emit Sheet A", "Emit Sheet B"]);
    });

    it("preserves worksheet state (hidden/veryHidden) with sharedStrings: 'ignore'", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        Workbook.addWorksheet(wb, "Visible");
        Workbook.addWorksheet(wb, "Hidden", { state: "hidden" });
        Workbook.addWorksheet(wb, "VeryHidden", { state: "veryHidden" });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "ignore",
        styles: "ignore"
      });

      const sheets: Array<{ name: string; state: string }> = [];
      for await (const ws of reader) {
        sheets.push({ name: ws.name, state: (ws as any).state ?? "visible" });
        for await (const _row of ws) {
          /* drain */
        }
      }
      expect(sheets).toEqual([
        { name: "Visible", state: "visible" },
        { name: "Hidden", state: "hidden" },
        { name: "VeryHidden", state: "veryHidden" }
      ]);
    });
  });

  // ===========================================================================
  // Dynamic Array Formulas via streaming reader (exceljs#2910)
  // ===========================================================================

  describe("dynamic array formulas via streaming reader", () => {
    it("should read isDynamicArray flag from dynamic array formula cells", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        for (let i = 1; i <= 5; i++) {
          Cell.setValue(ws, `A${i}`, i);
        }
        Cell.setValue(ws, "C1", {
          formula: "_xlfn._xlws.FILTER(A1:A5,A1:A5>2)",
          shareType: "array",
          ref: "C1",
          result: 3,
          isDynamicArray: true
        });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache"
      });

      let foundDynamicArray = false;
      for await (const ws of reader) {
        for await (const row of ws) {
          const cell = rowGetCell(row, 3);
          const val = cellGetValue(cell) as any;
          if (val && typeof val === "object" && val.formula && val.isDynamicArray) {
            foundDynamicArray = true;
            expect(val.formula).toContain("FILTER");
            expect(val.isDynamicArray).toBe(true);
          }
        }
      }
      expect(foundDynamicArray).toBe(true);
    });

    it("should not mark non-dynamic formulas as isDynamicArray", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        Cell.setValue(ws, "A1", { formula: "SUM(B1:B10)", result: 55 });
        Cell.setValue(ws, "A2", {
          formula: "{A1*2}",
          shareType: "array",
          ref: "A2",
          result: 110
        });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache"
      });

      for await (const ws of reader) {
        for await (const row of ws) {
          for (let col = 1; col <= 3; col++) {
            const val = cellGetValue(rowGetCell(row, col)) as any;
            if (val && typeof val === "object" && val.formula) {
              expect(val.isDynamicArray).toBeUndefined();
            }
          }
        }
      }
    });

    it("should handle mixed CSE array + dynamic array formulas in same workbook", async () => {
      const buffer = await buildXlsxBuffer(wb => {
        const ws = Workbook.addWorksheet(wb, "Sheet1");
        // Legacy CSE array formula
        Cell.setValue(ws, "A1", {
          formula: "{ROW(1:3)}",
          shareType: "array",
          ref: "A1:A3",
          result: 1
        });
        // Dynamic array formula
        Cell.setValue(ws, "C1", {
          formula: "_xlfn._xlws.SORT(B1:B10)",
          shareType: "array",
          ref: "C1",
          result: 1,
          isDynamicArray: true
        });
      });

      const reader = new WorkbookReader(buffer, {
        worksheets: "emit",
        sharedStrings: "cache"
      });

      let cseFormula: any = null;
      let dynamicFormula: any = null;
      for await (const ws of reader) {
        for await (const row of ws) {
          const a = cellGetValue(rowGetCell(row, 1)) as any;
          if (a && typeof a === "object" && a.formula) {
            cseFormula = a;
          }
          const c = cellGetValue(rowGetCell(row, 3)) as any;
          if (c && typeof c === "object" && c.formula) {
            dynamicFormula = c;
          }
        }
      }

      // CSE formula should NOT have isDynamicArray
      expect(cseFormula).toBeDefined();
      expect(cseFormula.isDynamicArray).toBeUndefined();

      // Dynamic array formula SHOULD have isDynamicArray
      expect(dynamicFormula).toBeDefined();
      expect(dynamicFormula.isDynamicArray).toBe(true);
    });
  });
});
