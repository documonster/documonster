/**
 * True Streaming Tests - Node.js Implementation
 *
 * Uses Node.js-specific APIs to verify TRUE streaming behavior.
 */

import { PassThrough, Readable } from "@stream";
import { createTrueStreamingTests } from "@stream/__tests__/streaming/true-streaming-tests";
import { beforeAll } from "vitest";

// Lazy imports
let WorkbookWriter: any;
let WorkbookReader: any;
let StreamingZip: any;
let ZipDeflateFile: any;
let ZipParser: any;

beforeAll(async () => {
  // Dynamic imports for Node.js environment
  const excelModule = await import("../../../index");
  WorkbookWriter = excelModule.WorkbookWriter;
  WorkbookReader = excelModule.WorkbookReader;

  const zipModule = await import("@archive/zip/stream");
  StreamingZip = zipModule.StreamingZip;
  ZipDeflateFile = zipModule.ZipDeflateFile;

  const zipParserModule = await import("@archive/unzip/zip-parser");
  ZipParser = zipParserModule.ZipParser;
});

// ============================================================================
// Node.js-Specific Test Context
// ============================================================================

function getNodeContext() {
  return {
    isBrowser: false,

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
          await file.push(content, true);
          // Wait for data to propagate through Node.js stream events
          await new Promise(resolve => setImmediate(resolve));
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
      const stream = new PassThrough();
      stream.on("data", (chunk: Buffer) => {
        onData(new Uint8Array(chunk));
      });

      // Enable trueStreaming for immediate data output
      const workbook = new WorkbookWriter({ stream, trueStreaming: true });

      return {
        addWorksheet: (name: string) => {
          const worksheet = workbook.addWorksheet(name);
          return {
            addRow: (data: (string | number)[]) => worksheet.addRow(data),
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
      // Convert Uint8Array to Readable stream since WorkbookReader expects string | Readable
      const stream = Readable.from(Buffer.from(data));
      const reader = new WorkbookReader(stream);

      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          onRow(worksheet.name, row.number, row.values);
        }
      }
    }
  };
}

// ============================================================================
// Run Shared Tests
// ============================================================================

createTrueStreamingTests(getNodeContext);
