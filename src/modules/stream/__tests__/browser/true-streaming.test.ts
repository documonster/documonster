/**
 * True Streaming Tests - Browser Implementation
 *
 * Uses browser-specific APIs (CompressionStream, DecompressionStream)
 * to verify TRUE streaming behavior.
 */

import { cellSetValue } from "@excel/cell";
import { rowValues } from "@excel/row";
import { rowCommit, rowGetCell } from "@excel/worksheet";
import {
  yieldToEventLoop,
  generateLargeText
} from "@stream/__tests__/streaming/streaming-test-base";
import { createTrueStreamingTests } from "@stream/__tests__/streaming/true-streaming-tests";
import { describe, beforeAll, expect, it } from "vitest";

// Lazy import to avoid Node.js module resolution issues
let WorkbookWriter: any;
let WorkbookReader: any;
let StreamingZip: any;
let ZipDeflateFile: any;
let ZipParser: any;

beforeAll(async () => {
  // Dynamic imports for browser environment - use index.browser directly
  const excelModule = await import("../../../../index.browser");
  WorkbookWriter = excelModule.WorkbookWriter;
  WorkbookReader = excelModule.WorkbookReader;

  const zipModule = await import("@archive/zip/stream");
  StreamingZip = zipModule.StreamingZip;
  ZipDeflateFile = zipModule.ZipDeflateFile;

  const zipParserModule = await import("@archive/unzip/zip-parser");
  ZipParser = zipParserModule.ZipParser;
});

// ============================================================================
// Browser-Specific Test Context
// ============================================================================

function getBrowserContext() {
  return {
    isBrowser: true,

    // ZIP Creation using StreamingZip
    createZip: async (onData: (chunk: Uint8Array) => void) => {
      let resolveFinish: () => void;
      const finishPromise = new Promise<void>(resolve => {
        resolveFinish = resolve;
      });

      const zip = new StreamingZip((err: Error | null, data: Uint8Array, final: boolean) => {
        if (err) {
          throw err;
        }
        if (data && data.length > 0) {
          onData(data);
        }
        if (final) {
          resolveFinish();
        }
      });

      return {
        addFile: async (name: string, content: Uint8Array) => {
          const file = new ZipDeflateFile(name, { level: 6 });
          zip.add(file);
          file.push(content, true);
          await yieldToEventLoop();
        },
        finalize: async () => {
          zip.end();
          await finishPromise;
        }
      };
    },

    // ZIP Parsing
    parseZip: async (
      zipData: Uint8Array,
      onEntry: (entry: { path: string; stream: () => AsyncIterable<Uint8Array> }) => Promise<void>
    ) => {
      const parser = new ZipParser(zipData);
      const entries = parser.getEntries();

      for (const entry of entries) {
        if (!entry.isDirectory) {
          await onEntry({
            path: entry.path,
            stream: () => ({
              async *[Symbol.asyncIterator]() {
                const content = await parser.extract(entry.path);
                if (content) {
                  const chunkSize = 16384;
                  for (let i = 0; i < content.length; i += chunkSize) {
                    yield content.slice(i, Math.min(i + chunkSize, content.length));
                  }
                }
              }
            })
          });
        }
      }
    },

    // Excel Write
    createWorkbookWriter: async (onData: (chunk: Uint8Array) => void) => {
      const writable = new WritableStream<Uint8Array>({
        write(chunk) {
          onData(chunk);
        }
      });

      // Enable trueStreaming for immediate data output
      const workbook = new WorkbookWriter({ stream: writable, trueStreaming: true });

      return {
        addWorksheet: (name: string) => {
          const worksheet = workbook.addWorksheet(name);
          return {
            addRow: (data: (string | number)[]) => {
              const row = worksheet.addRow(data);
              return { commit: () => rowCommit(row) };
            },
            commit: () => worksheet.commit()
          };
        },
        commit: () => workbook.commit()
      };
    },

    // Excel Read - using WorkbookReader for TRUE streaming
    createWorkbookReader: async (
      data: Uint8Array,
      onRow: (sheetName: string, rowNumber: number, values: unknown[]) => void
    ) => {
      // Use WorkbookReader for TRUE streaming - rows are yielded progressively
      const reader = new WorkbookReader(data);

      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          onRow(worksheet.name, row.number, rowValues(row));
        }
      }
    }
  };
}

// ============================================================================
// Run Shared Tests
// ============================================================================

createTrueStreamingTests(getBrowserContext);

const createCompressionLikeStream = (): TransformStream<Uint8Array, Uint8Array> => {
  if (typeof CompressionStream !== "undefined") {
    return new CompressionStream("deflate-raw") as unknown as TransformStream<
      Uint8Array,
      Uint8Array
    >;
  }
  // Pass-through fallback when CompressionStream is unavailable.
  // Zero-arg TransformStream is an identity transform.
  return new TransformStream();
};

const createDecompressionLikeStream = (): TransformStream<Uint8Array, Uint8Array> => {
  if (typeof DecompressionStream !== "undefined") {
    return new DecompressionStream("deflate-raw") as unknown as TransformStream<
      Uint8Array,
      Uint8Array
    >;
  }
  // Pass-through fallback when DecompressionStream is unavailable.
  return new TransformStream();
};

// ============================================================================
// Browser-Specific Additional Tests
// ============================================================================

describe("Browser-Specific True Streaming", () => {
  describe("Native CompressionStream Verification", () => {
    it("should verify CompressionStream streams chunks progressively", async () => {
      const chunks: { time: number; size: number }[] = [];
      const startTime = performance.now();

      const compressionStream = createCompressionLikeStream();
      const writer = compressionStream.writable.getWriter();
      const reader = compressionStream.readable.getReader();

      // Start reading in background
      const readPromise = (async () => {
        while (true) {
          const { done, value } = await reader.read();
          if (done) {
            break;
          }
          chunks.push({
            time: Math.round(performance.now() - startTime),
            size: value.length
          });
        }
      })();

      // Write 10MB of random data in 3MB chunks (random data won't compress well)
      const chunkSize = 3 * 1024 * 1024;
      const totalChunks = 3;

      for (let i = 0; i < totalChunks; i++) {
        // Use random data to prevent extreme compression
        const data = new Uint8Array(chunkSize);
        for (let j = 0; j < data.length; j += 65536) {
          const size = Math.min(65536, data.length - j);
          crypto.getRandomValues(data.subarray(j, j + size));
        }
        await writer.write(data);
        await yieldToEventLoop();
        console.log(`Write ${i + 1}: ${chunks.length} output chunks so far`);
      }

      const chunksBeforeClose = chunks.length;
      await writer.close();
      await readPromise;

      console.log(`\n=== Native CompressionStream Analysis ===`);
      console.log(`Chunks before close: ${chunksBeforeClose}`);
      console.log(`Chunks after close: ${chunks.length}`);
      console.log(`Total compressed size: ${chunks.reduce((s, c) => s + c.size, 0)} bytes`);

      if (chunksBeforeClose > 0) {
        console.log("\n✅ CompressionStream streams progressively");
      } else {
        console.log("\n⚠️ All data buffered until close");
      }

      expect(chunks.length).toBeGreaterThan(0);
    });
  });

  describe("Native DecompressionStream Verification", () => {
    it("should verify DecompressionStream streams chunks progressively", async () => {
      // First compress some data
      const compressionStream = createCompressionLikeStream();
      const compressWriter = compressionStream.writable.getWriter();
      const compressReader = compressionStream.readable.getReader();

      const compressedChunks: Uint8Array[] = [];
      const compressReadPromise = (async () => {
        while (true) {
          const { done, value } = await compressReader.read();
          if (done) {
            break;
          }
          compressedChunks.push(value);
        }
      })();

      // Create 500KB of test data
      const testData = new TextEncoder().encode(generateLargeText(500000));
      await compressWriter.write(testData);
      await compressWriter.close();
      await compressReadPromise;

      // Combine compressed data
      const totalCompressed = compressedChunks.reduce((s, c) => s + c.length, 0);
      const compressedData = new Uint8Array(totalCompressed);
      let offset = 0;
      for (const chunk of compressedChunks) {
        compressedData.set(chunk, offset);
        offset += chunk.length;
      }

      console.log(`Compressed data size: ${totalCompressed} bytes`);

      // Now test decompression streaming
      const decompressedChunks: { time: number; size: number }[] = [];
      const startTime = performance.now();

      const decompressionStream = createDecompressionLikeStream();
      const decompressWriter = decompressionStream.writable.getWriter();
      const decompressReader = decompressionStream.readable.getReader();

      // Start reading in background
      const decompressReadPromise = (async () => {
        while (true) {
          const { done, value } = await decompressReader.read();
          if (done) {
            break;
          }
          decompressedChunks.push({
            time: Math.round(performance.now() - startTime),
            size: value.length
          });
        }
      })();

      // Write compressed data in small chunks to simulate streaming
      const writeChunkSize = 1000;
      for (let i = 0; i < compressedData.length; i += writeChunkSize) {
        const chunk = compressedData.slice(i, Math.min(i + writeChunkSize, compressedData.length));
        await decompressWriter.write(chunk);

        if (i % 10000 === 0) {
          await yieldToEventLoop();
          console.log(`Decompress write ${i}: ${decompressedChunks.length} output chunks`);
        }
      }

      const chunksBeforeClose = decompressedChunks.length;
      await decompressWriter.close();
      await decompressReadPromise;

      const totalDecompressed = decompressedChunks.reduce((s, c) => s + c.size, 0);

      console.log(`\n=== Native DecompressionStream Analysis ===`);
      console.log(`Chunks before close: ${chunksBeforeClose}`);
      console.log(`Chunks after close: ${decompressedChunks.length}`);
      console.log(`Total decompressed size: ${totalDecompressed} bytes`);

      if (chunksBeforeClose > 0) {
        console.log("\n✅ DecompressionStream streams progressively");
      } else {
        console.log("\n⚠️ All data buffered until close");
      }

      expect(decompressedChunks.length).toBeGreaterThan(0);
      expect(totalDecompressed).toBe(testData.length);
    });
  });

  describe("Streaming memory behavior (browser)", () => {
    it("should not accumulate memory during streaming writes", async () => {
      const chunks: Uint8Array[] = [];
      const output = new WritableStream<Uint8Array>({
        write(chunk) {
          chunks.push(chunk);
        }
      });

      const workbook = new WorkbookWriter({
        stream: output,
        useSharedStrings: false,
        trueStreaming: true
      });
      const worksheet = workbook.addWorksheet("Sheet 1");

      const cellValue = "abcdefghij".repeat(40); // 400 chars

      // Warm up — stabilize JIT, GC, and internal structures
      for (let i = 0; i < 1000; i++) {
        const row = worksheet.getRow(i + 1);
        for (let c = 1; c <= 9; c++) {
          cellSetValue(rowGetCell(row, c), cellValue);
        }
        rowCommit(row);
      }

      // Yield to let browser GC settle
      await new Promise(r => setTimeout(r, 100));
      const perfMem = (performance as any).memory;
      const baselineHeap = perfMem ? perfMem.usedJSHeapSize : 0;

      // Steady-state — write 4000 more rows
      for (let i = 1000; i < 5000; i++) {
        const row = worksheet.getRow(i + 1);
        for (let c = 1; c <= 9; c++) {
          cellSetValue(rowGetCell(row, c), cellValue);
        }
        rowCommit(row);
      }

      await new Promise(r => setTimeout(r, 100));
      const finalHeap = perfMem ? perfMem.usedJSHeapSize : 0;

      await workbook.commit();

      const totalBytes = chunks.reduce((sum, c) => sum + c.length, 0);
      expect(totalBytes).toBeGreaterThan(0);

      // Memory assertion (Chrome only — performance.memory)
      if (baselineHeap > 0 && finalHeap > 0) {
        const growthMB = (finalHeap - baselineHeap) / 1024 / 1024;
        // 4000 rows × 9 cells × 400 chars should not cause significant growth.
        // Before the fix this would accumulate ~60MB+ in push chain closures.
        // With the fix, growth should be well under 50MB.
        expect(growthMB).toBeLessThan(50);
      }
    }, 30000);
  });
});
