/**
 * Streaming Zlib tests
 */

import { unzlibSync, isZlibData } from "@archive/compression/compress";
import { createZlibStream, createUnzlibStream } from "@archive/compression/streaming-compress";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect } from "vitest";

// Helper to convert Buffer to Uint8Array for comparison
function toUint8Array(data: Uint8Array | Buffer): Uint8Array {
  // Node.js Buffer is a subclass of Uint8Array but JSON-serializes differently
  // Create a plain Uint8Array copy for comparison
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}

describe("Streaming Zlib compression", () => {
  const testData = new TextEncoder().encode("Hello, Zlib streaming world!");

  describe("createZlibStream", () => {
    it("should compress data in streaming fashion", async () => {
      const stream = createZlibStream();
      const chunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        stream.on("data", chunk => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", reject);

        stream.write(testData);
        stream.end();
      });

      const compressed = concatUint8Arrays(chunks);
      expect(isZlibData(compressed)).toBe(true);

      // Verify by decompressing
      const decompressed = unzlibSync(compressed);
      expect(decompressed).toEqual(testData);
    });

    it("should handle multiple writes", async () => {
      const stream = createZlibStream();
      const chunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        stream.on("data", chunk => chunks.push(chunk));
        stream.on("end", () => resolve());
        stream.on("error", reject);

        stream.write(new TextEncoder().encode("Chunk 1 "));
        stream.write(new TextEncoder().encode("Chunk 2 "));
        stream.write(new TextEncoder().encode("Chunk 3"));
        stream.end();
      });

      const compressed = concatUint8Arrays(chunks);
      expect(isZlibData(compressed)).toBe(true);

      const decompressed = unzlibSync(compressed);
      expect(new TextDecoder().decode(decompressed)).toBe("Chunk 1 Chunk 2 Chunk 3");
    });

    it("should support different compression levels", async () => {
      const compressWithLevel = async (level: number): Promise<Uint8Array> => {
        const stream = createZlibStream({ level });
        const chunks: Uint8Array[] = [];

        await new Promise<void>((resolve, reject) => {
          stream.on("data", chunk => chunks.push(chunk));
          stream.on("end", () => resolve());
          stream.on("error", reject);

          stream.write(testData);
          stream.end();
        });

        return concatUint8Arrays(chunks);
      };

      const level1 = await compressWithLevel(1);
      const level9 = await compressWithLevel(9);

      expect(isZlibData(level1)).toBe(true);
      expect(isZlibData(level9)).toBe(true);

      // Both should decompress to same data
      expect(unzlibSync(level1)).toEqual(testData);
      expect(unzlibSync(level9)).toEqual(testData);
    });
  });

  describe("createUnzlibStream", () => {
    it("should decompress data in streaming fashion", async () => {
      // First compress the data
      const compressStream = createZlibStream();
      const compressedChunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        compressStream.on("data", chunk => compressedChunks.push(chunk));
        compressStream.on("end", () => resolve());
        compressStream.on("error", reject);

        compressStream.write(testData);
        compressStream.end();
      });

      const compressed = concatUint8Arrays(compressedChunks);

      // Now decompress
      const decompressStream = createUnzlibStream();
      const decompressedChunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        decompressStream.on("data", chunk => decompressedChunks.push(chunk));
        decompressStream.on("end", () => resolve());
        decompressStream.on("error", reject);

        decompressStream.write(compressed);
        decompressStream.end();
      });

      const decompressed = concatUint8Arrays(decompressedChunks);
      expect(toUint8Array(decompressed)).toEqual(testData);
    });

    it("should handle piped streams", async () => {
      const compressStream = createZlibStream();
      const decompressStream = createUnzlibStream();
      const chunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        decompressStream.on("data", chunk => chunks.push(chunk));
        decompressStream.on("end", () => resolve());
        decompressStream.on("error", reject);
        compressStream.on("error", reject);

        // Pipe compression output to decompression input
        compressStream.on("data", chunk => {
          decompressStream.write(chunk);
        });
        compressStream.on("end", () => {
          decompressStream.end();
        });

        compressStream.write(testData);
        compressStream.end();
      });

      const result = concatUint8Arrays(chunks);
      expect(toUint8Array(result)).toEqual(testData);
    });
  });

  describe("large data", () => {
    it("should handle large data correctly", async () => {
      const largeData = new Uint8Array(100000);
      for (let i = 0; i < largeData.length; i++) {
        largeData[i] = i % 256;
      }

      // Compress
      const compressStream = createZlibStream();
      const compressedChunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        compressStream.on("data", chunk => compressedChunks.push(chunk));
        compressStream.on("end", () => resolve());
        compressStream.on("error", reject);

        // Write in chunks
        const chunkSize = 16384;
        for (let i = 0; i < largeData.length; i += chunkSize) {
          compressStream.write(largeData.subarray(i, Math.min(i + chunkSize, largeData.length)));
        }
        compressStream.end();
      });

      const compressed = concatUint8Arrays(compressedChunks);
      expect(isZlibData(compressed)).toBe(true);
      expect(compressed.length).toBeLessThan(largeData.length);

      // Decompress
      const decompressStream = createUnzlibStream();
      const decompressedChunks: Uint8Array[] = [];

      await new Promise<void>((resolve, reject) => {
        decompressStream.on("data", chunk => decompressedChunks.push(chunk));
        decompressStream.on("end", () => resolve());
        decompressStream.on("error", reject);

        decompressStream.write(compressed);
        decompressStream.end();
      });

      const decompressed = concatUint8Arrays(decompressedChunks);
      expect(decompressed).toEqual(largeData);
    });
  });
});
