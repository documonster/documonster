import { isGzipData, GZIP_ID1, GZIP_ID2 } from "@archive/compression/compress";
import { createGzipStream, createGunzipStream } from "@archive/compression/streaming-compress";
import { describe, it, expect } from "vitest";

describe("streaming gzip", () => {
  describe("createGzipStream", () => {
    it("should compress data chunk by chunk", async () => {
      const gzipStream = createGzipStream();
      const chunks: Uint8Array[] = [];

      gzipStream.on("data", chunk => {
        chunks.push(chunk);
      });

      // Write data in chunks
      gzipStream.write(new Uint8Array([72, 101, 108, 108, 111])); // "Hello"
      gzipStream.write(new Uint8Array([44, 32])); // ", "
      gzipStream.write(new Uint8Array([87, 111, 114, 108, 100])); // "World"

      // Finalize
      await new Promise<void>(resolve => {
        gzipStream.on("end", resolve);
        gzipStream.end();
      });

      // Combine all chunks
      const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const compressed = new Uint8Array(totalLength);
      let offset = 0;
      for (const chunk of chunks) {
        compressed.set(chunk, offset);
        offset += chunk.length;
      }

      // Verify it's valid gzip
      expect(isGzipData(compressed)).toBe(true);
      expect(compressed[0]).toBe(GZIP_ID1);
      expect(compressed[1]).toBe(GZIP_ID2);

      // Decompress and verify
      const gunzipStream = createGunzipStream();
      const decompressedChunks: Uint8Array[] = [];

      gunzipStream.on("data", chunk => {
        decompressedChunks.push(chunk);
      });

      gunzipStream.write(compressed);
      await new Promise<void>(resolve => {
        gunzipStream.on("end", resolve);
        gunzipStream.end();
      });

      const decompressedLength = decompressedChunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const decompressed = new Uint8Array(decompressedLength);
      offset = 0;
      for (const chunk of decompressedChunks) {
        decompressed.set(chunk, offset);
        offset += chunk.length;
      }

      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World");
    });

    it("should respect compression level", async () => {
      const original = new TextEncoder().encode("Compress this text!".repeat(100));

      // Compress with level 1
      const gzip1 = createGzipStream({ level: 1 });
      const chunks1: Uint8Array[] = [];
      gzip1.on("data", chunk => chunks1.push(chunk));
      gzip1.write(original);
      await new Promise<void>(resolve => {
        gzip1.on("end", resolve);
        gzip1.end();
      });

      // Compress with level 9
      const gzip9 = createGzipStream({ level: 9 });
      const chunks9: Uint8Array[] = [];
      gzip9.on("data", chunk => chunks9.push(chunk));
      gzip9.write(original);
      await new Promise<void>(resolve => {
        gzip9.on("end", resolve);
        gzip9.end();
      });

      const size1 = chunks1.reduce((sum, chunk) => sum + chunk.length, 0);
      const size9 = chunks9.reduce((sum, chunk) => sum + chunk.length, 0);

      // Level 9 should produce smaller or equal output
      expect(size9).toBeLessThanOrEqual(size1);
    });

    it("should emit error on invalid gunzip input", async () => {
      const gunzipStream = createGunzipStream();
      let errorEmitted = false;

      gunzipStream.on("error", () => {
        errorEmitted = true;
      });

      // Write invalid data
      gunzipStream.write(new Uint8Array([0x00, 0x01, 0x02, 0x03]));

      // Try to finalize
      try {
        await new Promise<void>((resolve, reject) => {
          gunzipStream.on("end", resolve);
          gunzipStream.on("error", reject);
          gunzipStream.end();
        });
      } catch {
        errorEmitted = true;
      }

      expect(errorEmitted).toBe(true);
    });
  });

  describe("streaming roundtrip", () => {
    it("should handle large data streaming", async () => {
      // Generate 1MB of data
      const original = new Uint8Array(1024 * 1024);
      for (let i = 0; i < original.length; i++) {
        original[i] = i % 256;
      }

      const gzipStream = createGzipStream();
      const compressedChunks: Uint8Array[] = [];

      gzipStream.on("data", chunk => {
        compressedChunks.push(chunk);
      });

      // Write in 64KB chunks
      const chunkSize = 64 * 1024;
      for (let i = 0; i < original.length; i += chunkSize) {
        gzipStream.write(original.slice(i, i + chunkSize));
      }

      await new Promise<void>(resolve => {
        gzipStream.on("end", resolve);
        gzipStream.end();
      });

      // Combine compressed chunks
      const compressedLength = compressedChunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const compressed = new Uint8Array(compressedLength);
      let offset = 0;
      for (const chunk of compressedChunks) {
        compressed.set(chunk, offset);
        offset += chunk.length;
      }

      // Decompress
      const gunzipStream = createGunzipStream();
      const decompressedChunks: Uint8Array[] = [];

      gunzipStream.on("data", chunk => {
        decompressedChunks.push(chunk);
      });

      gunzipStream.write(compressed);
      await new Promise<void>(resolve => {
        gunzipStream.on("end", resolve);
        gunzipStream.end();
      });

      // Combine decompressed chunks
      const decompressedLength = decompressedChunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const decompressed = new Uint8Array(decompressedLength);
      offset = 0;
      for (const chunk of decompressedChunks) {
        decompressed.set(chunk, offset);
        offset += chunk.length;
      }

      expect(decompressed).toEqual(original);
    });

    it("should handle multiple small writes", async () => {
      const parts = ["Hello", " ", "World", "!", " ", "Testing", " ", "streaming"];
      const original = new TextEncoder().encode(parts.join(""));

      const gzipStream = createGzipStream();
      const compressedChunks: Uint8Array[] = [];

      gzipStream.on("data", chunk => {
        compressedChunks.push(chunk);
      });

      // Write each part separately
      for (const part of parts) {
        gzipStream.write(new TextEncoder().encode(part));
      }

      await new Promise<void>(resolve => {
        gzipStream.on("end", resolve);
        gzipStream.end();
      });

      // Combine and decompress
      const compressedLength = compressedChunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const compressed = new Uint8Array(compressedLength);
      let offset = 0;
      for (const chunk of compressedChunks) {
        compressed.set(chunk, offset);
        offset += chunk.length;
      }

      const gunzipStream = createGunzipStream();
      const decompressedChunks: Uint8Array[] = [];

      gunzipStream.on("data", chunk => {
        decompressedChunks.push(chunk);
      });

      gunzipStream.write(compressed);
      await new Promise<void>(resolve => {
        gunzipStream.on("end", resolve);
        gunzipStream.end();
      });

      const decompressedLength = decompressedChunks.reduce((sum, chunk) => sum + chunk.length, 0);
      const decompressed = new Uint8Array(decompressedLength);
      offset = 0;
      for (const chunk of decompressedChunks) {
        decompressed.set(chunk, offset);
        offset += chunk.length;
      }

      expect(decompressed).toEqual(original);
    });
  });
});
