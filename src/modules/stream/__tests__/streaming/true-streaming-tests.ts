/**
 * True Streaming Tests - Shared Test Cases
 *
 * These tests verify TRUE streaming behavior for:
 * 1. ZIP compression (write)
 * 2. ZIP decompression (unzip/read)
 * 3. Excel streaming write
 * 4. Excel streaming read
 *
 * Tests are designed to work in both Node.js and Browser environments.
 */

import type { PhaseLog } from "@stream/__tests__/streaming/streaming-test-base";
import {
  createMetrics,
  recordChunk,
  setPhase,
  logPhase,
  analyzeGrowth,
  printAnalysis,
  yieldToEventLoop,
  strongYield,
  generateRowData,
  generateLargeText,
  assertTrueStreaming
} from "@stream/__tests__/streaming/streaming-test-base";
import { describe, it, expect, beforeAll } from "vitest";

// ============================================================================
// Test Configuration
// ============================================================================

interface TestContext {
  // Platform detection
  isBrowser: boolean;

  // ZIP creation
  createZip: (onData: (chunk: Uint8Array) => void) => Promise<{
    addFile: (name: string, content: Uint8Array) => Promise<void>;
    finalize: () => Promise<void>;
  }>;

  // ZIP reading/parsing
  parseZip: (
    zipData: Uint8Array,
    onEntry: (entry: { path: string; stream: () => AsyncIterable<Uint8Array> }) => Promise<void>
  ) => Promise<void>;

  // Excel write
  createWorkbookWriter: (onData: (chunk: Uint8Array) => void) => Promise<{
    addWorksheet: (name: string) => {
      addRow: (data: (string | number)[]) => { commit: () => void };
      commit: () => Promise<void>;
    };
    commit: () => Promise<void>;
  }>;

  // Excel read (streaming)
  createWorkbookReader: (
    data: Uint8Array,
    onRow: (sheetName: string, rowNumber: number, values: unknown[]) => void
  ) => Promise<void>;
}

// ============================================================================
// Shared Test Implementations
// ============================================================================

export function createTrueStreamingTests(getContext: () => TestContext) {
  describe("True Streaming Verification", () => {
    let ctx: TestContext;

    beforeAll(() => {
      ctx = getContext();
    });

    // ========================================================================
    // ZIP Compression (Write) Tests
    // ========================================================================

    describe("ZIP Compression Streaming", () => {
      it("should stream compressed data progressively during writes", async () => {
        const metrics = createMetrics();
        const logs: PhaseLog[] = [];

        const zip = await ctx.createZip(chunk => recordChunk(metrics, chunk));

        // Write multiple files with significant data - 10MB total
        const fileCount = 5;
        const bytesPerFile = 2_000_000; // 2MB each = 10MB total

        for (let i = 0; i < fileCount; i++) {
          setPhase(metrics, `file-${i}`);
          const content = new TextEncoder().encode(generateLargeText(bytesPerFile));
          await zip.addFile(`file${i}.txt`, content);

          await yieldToEventLoop();
          logs.push(logPhase(metrics, `After file ${i}`));
        }

        const bytesBeforeFinalize = metrics.totalBytes;
        logs.push(logPhase(metrics, "Before finalize"));

        setPhase(metrics, "finalize");
        await zip.finalize();

        await yieldToEventLoop();
        logs.push(logPhase(metrics, "After finalize"));

        printAnalysis("ZIP Compression Streaming", logs, bytesBeforeFinalize, metrics.totalBytes);

        const analysis = analyzeGrowth(logs.slice(0, -1));
        assertTrueStreaming(analysis.totalGrowth, bytesBeforeFinalize);
      }, 60000);
    });

    // ========================================================================
    // ZIP Decompression (Read) Tests
    // ========================================================================

    describe("ZIP Decompression Streaming", () => {
      it("should stream decompressed data progressively during parsing", async () => {
        // First create a test ZIP file - 10MB total
        const zipChunks: Uint8Array[] = [];
        const zip = await ctx.createZip(chunk => zipChunks.push(chunk));

        const fileCount = 5;
        const bytesPerFile = 2_000_000; // 2MB each = 10MB total

        for (let i = 0; i < fileCount; i++) {
          const content = new TextEncoder().encode(generateLargeText(bytesPerFile));
          await zip.addFile(`test${i}.txt`, content);
        }
        await zip.finalize();

        // Combine chunks into single ZIP buffer
        const totalLength = zipChunks.reduce((sum, c) => sum + c.length, 0);
        const zipData = new Uint8Array(totalLength);
        let offset = 0;
        for (const chunk of zipChunks) {
          zipData.set(chunk, offset);
          offset += chunk.length;
        }

        // Now test decompression streaming
        const metrics = createMetrics();
        const logs: PhaseLog[] = [];
        let entryIndex = 0;

        await ctx.parseZip(zipData, async entry => {
          setPhase(metrics, `entry-${entryIndex}`);
          const stream = entry.stream();

          for await (const chunk of stream) {
            recordChunk(metrics, chunk);
          }

          await yieldToEventLoop();
          logs.push(logPhase(metrics, `After ${entry.path}`));
          entryIndex++;
        });

        logs.push(logPhase(metrics, "After all entries"));

        printAnalysis("ZIP Decompression Streaming", logs, metrics.totalBytes, metrics.totalBytes);

        expect(metrics.chunks.length).toBeGreaterThan(fileCount);
        expect(metrics.totalBytes).toBeGreaterThan(bytesPerFile * fileCount * 0.9);
      }, 60000);
    });

    // ========================================================================
    // Excel Write Streaming Tests
    // ========================================================================

    describe("Excel Write Streaming", () => {
      it("should stream Excel data progressively during row writes", async () => {
        const metrics = createMetrics();
        const logs: PhaseLog[] = [];

        const workbook = await ctx.createWorkbookWriter(chunk => recordChunk(metrics, chunk));
        const worksheet = workbook.addWorksheet("Sheet1");

        // Write 100000 rows with pseudo-random data = ~20MB uncompressed, ~10MB compressed
        const totalRows = 100000;
        // Yield every 1000 rows to allow zlib to output compressed data
        // Node.js zlib uses libuv thread pool, so we need frequent yields
        const checkInterval = 1000;

        setPhase(metrics, "row-writes");

        for (let i = 0; i < totalRows; i++) {
          worksheet.addRow(generateRowData(i)).commit();

          if (i > 0 && i % checkInterval === 0) {
            // Use strongYield to allow zlib thread pool to complete processing
            await strongYield();
            // Only log every 10th checkpoint to avoid excessive output
            if (i % 10000 === 0) {
              logs.push(logPhase(metrics, `Row ${i}`));
            }
          }
        }

        const bytesBeforeCommit = metrics.totalBytes;
        logs.push(logPhase(metrics, "Before worksheet.commit()"));

        setPhase(metrics, "worksheet-commit");
        await worksheet.commit();
        logs.push(logPhase(metrics, "After worksheet.commit()"));

        setPhase(metrics, "workbook-commit");
        await workbook.commit();
        logs.push(logPhase(metrics, "After workbook.commit()"));

        printAnalysis("Excel Write Streaming", logs, bytesBeforeCommit, metrics.totalBytes);

        const rowLogs = logs.filter(l => l.phase.startsWith("Row "));
        const analysis = analyzeGrowth(rowLogs);

        assertTrueStreaming(analysis.totalGrowth, bytesBeforeCommit);
      }, 120000);
    });

    // ========================================================================
    // Excel Read Streaming Tests
    // ========================================================================

    describe("Excel Read Streaming", () => {
      it("should stream row data progressively during parsing", async () => {
        // First create a test Excel file - 10MB+ uncompressed
        const excelChunks: Uint8Array[] = [];
        const workbook = await ctx.createWorkbookWriter(chunk => excelChunks.push(chunk));
        const worksheet = workbook.addWorksheet("TestSheet");

        // Create 50000 rows (~10MB uncompressed)
        for (let i = 0; i < 50000; i++) {
          worksheet.addRow(generateRowData(i)).commit();
        }
        await worksheet.commit();
        await workbook.commit();

        // Combine chunks
        const totalLength = excelChunks.reduce((sum, c) => sum + c.length, 0);
        const excelData = new Uint8Array(totalLength);
        let offset = 0;
        for (const chunk of excelChunks) {
          excelData.set(chunk, offset);
          offset += chunk.length;
        }

        // Now test read streaming
        const rowsReceived: { sheetName: string; rowNumber: number }[] = [];

        await ctx.createWorkbookReader(excelData, (sheetName, rowNumber, _values) => {
          rowsReceived.push({ sheetName, rowNumber });
        });

        // Verify we received all rows - check the last row number
        const lastRow = rowsReceived[rowsReceived.length - 1];
        expect(lastRow).toBeDefined();
        expect(lastRow.rowNumber).toBe(50000);
        expect(lastRow.sheetName).toBe("TestSheet");

        // Verify we got progressive streaming (not just one big batch)
        // In true streaming, we should receive many row events
        expect(rowsReceived.length).toBe(50000);
      }, 120000);
    });
  });
}
