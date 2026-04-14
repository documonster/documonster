import {
  gzip,
  gunzip,
  gzipSync,
  gunzipSync,
  isGzipData,
  GZIP_ID1,
  GZIP_ID2
} from "@archive/compression/compress";
import { describe, it, expect } from "vitest";

describe("gzip", () => {
  describe("isGzipData", () => {
    it("should detect valid gzip magic bytes", () => {
      const gzipHeader = new Uint8Array([GZIP_ID1, GZIP_ID2, 0x08, 0x00]);
      expect(isGzipData(gzipHeader)).toBe(true);
    });

    it("should reject non-gzip data", () => {
      const nonGzip = new Uint8Array([0x50, 0x4b, 0x03, 0x04]); // ZIP magic
      expect(isGzipData(nonGzip)).toBe(false);
    });

    it("should reject data that is too short", () => {
      const tooShort = new Uint8Array([GZIP_ID1]);
      expect(isGzipData(tooShort)).toBe(false);
    });
  });

  describe("async gzip/gunzip", () => {
    it("should compress and decompress simple text", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await gzip(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should produce valid gzip format", async () => {
      const original = new TextEncoder().encode("Test data");
      const compressed = await gzip(original);

      expect(isGzipData(compressed)).toBe(true);
      expect(compressed[0]).toBe(GZIP_ID1);
      expect(compressed[1]).toBe(GZIP_ID2);
      expect(compressed[2]).toBe(0x08); // DEFLATE method
    });

    it("should compress and decompress binary data", async () => {
      const original = new Uint8Array([0x00, 0x01, 0x02, 0xff, 0xfe, 0xfd]);
      const compressed = await gzip(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should compress and decompress large data", async () => {
      // 100KB of repeating data
      const original = new Uint8Array(100 * 1024);
      for (let i = 0; i < original.length; i++) {
        original[i] = i % 256;
      }

      const compressed = await gzip(original);
      expect(compressed.length).toBeLessThan(original.length);

      const decompressed = await gunzip(compressed);
      expect(decompressed).toEqual(original);
    });

    it("should compress with different levels", async () => {
      const original = new TextEncoder().encode("Hello, World!".repeat(100));

      const level1 = await gzip(original, { level: 1 });
      const level6 = await gzip(original, { level: 6 });
      const level9 = await gzip(original, { level: 9 });

      // All should decompress correctly
      expect(await gunzip(level1)).toEqual(original);
      expect(await gunzip(level6)).toEqual(original);
      expect(await gunzip(level9)).toEqual(original);
    });

    it("should handle level 0 (store)", async () => {
      const original = new TextEncoder().encode("Uncompressed data");
      const compressed = await gzip(original, { level: 0 });
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
      expect(isGzipData(compressed)).toBe(true);
    });

    it("should compress unicode text", async () => {
      const original = new TextEncoder().encode("你好世界 🌍 مرحبا");
      const compressed = await gzip(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
    });
  });

  describe("sync gzip/gunzip", () => {
    it("should compress and decompress simple text", () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = gzipSync(original);
      const decompressed = gunzipSync(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should produce valid gzip format", () => {
      const original = new TextEncoder().encode("Test data");
      const compressed = gzipSync(original);

      expect(isGzipData(compressed)).toBe(true);
    });

    it("should handle different compression levels", () => {
      const original = new TextEncoder().encode("Test compression levels");

      const level1 = gzipSync(original, { level: 1 });
      const level9 = gzipSync(original, { level: 9 });

      expect(gunzipSync(level1)).toEqual(original);
      expect(gunzipSync(level9)).toEqual(original);
    });
  });

  describe("error handling", () => {
    it("should throw on invalid gzip data", async () => {
      const invalidData = new Uint8Array([0x00, 0x01, 0x02, 0x03]);

      await expect(gunzip(invalidData)).rejects.toThrow();
    });

    it("should throw on truncated gzip data", async () => {
      const original = new TextEncoder().encode("Hello");
      const compressed = await gzip(original);
      const truncated = compressed.slice(0, compressed.length - 5);

      await expect(gunzip(truncated)).rejects.toThrow();
    });

    it("should throw on corrupted gzip data (sync)", () => {
      const invalidData = new Uint8Array([0x00, 0x01, 0x02, 0x03]);

      expect(() => gunzipSync(invalidData)).toThrow();
    });
  });

  describe("cross-compatibility", () => {
    it("async compressed data should be decompressible by sync", async () => {
      const original = new TextEncoder().encode("Cross-compat test");
      const compressed = await gzip(original);
      const decompressed = gunzipSync(compressed);

      expect(decompressed).toEqual(original);
    });

    it("sync compressed data should be decompressible by async", async () => {
      const original = new TextEncoder().encode("Cross-compat test");
      const compressed = gzipSync(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
    });
  });

  describe("edge cases", () => {
    it("should handle empty data", async () => {
      const original = new Uint8Array(0);
      const compressed = await gzip(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed.length).toBe(0);
    });

    it("should handle empty data (sync)", () => {
      const original = new Uint8Array(0);
      const compressed = gzipSync(original);
      const decompressed = gunzipSync(compressed);

      expect(decompressed.length).toBe(0);
    });

    it("should handle single byte data", async () => {
      const original = new Uint8Array([42]);
      const compressed = await gzip(original);
      const decompressed = await gunzip(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should throw on data with wrong magic number but correct size", async () => {
      // Create fake "gzip" data with valid size but wrong magic
      const fakeGzip = new Uint8Array(18);
      fakeGzip[0] = 0x1f;
      fakeGzip[1] = 0x00; // wrong magic byte 2

      await expect(gunzip(fakeGzip)).rejects.toThrow();
    });

    it("should throw on wrong compression method", async () => {
      const original = await gzip(new Uint8Array([1, 2, 3]));
      original[2] = 0x00; // change compression method from 8 (DEFLATE) to 0

      await expect(gunzip(original)).rejects.toThrow();
    });
  });
});
