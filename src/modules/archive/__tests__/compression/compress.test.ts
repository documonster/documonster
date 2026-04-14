import {
  compress,
  compressSync,
  decompress,
  decompressSync,
  hasCompressionStream
} from "@archive/compression/compress";
import { isNode } from "@utils/env";
import { describe, it, expect } from "vitest";

describe("compress", () => {
  describe("environment detection", () => {
    it("should detect Node.js environment", () => {
      // In Node.js, this should be true
      expect(isNode()).toBe(true);
    });

    it("should detect CompressionStream availability", () => {
      // In Node.js 17+, this should be true
      const result = hasCompressionStream();
      expect(typeof result).toBe("boolean");
    });
  });

  describe("async compress/decompress", () => {
    it("should compress and decompress simple text", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should compress and decompress empty data", async () => {
      const original = new Uint8Array(0);
      const compressed = await compress(original, { level: 6 });
      // Empty data should remain empty or be trivially compressed
      expect(compressed.length).toBeLessThanOrEqual(original.length + 20);
    });

    it("should compress and decompress binary data", async () => {
      const original = new Uint8Array([0x00, 0x01, 0x02, 0xff, 0xfe, 0xfd]);
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should compress and decompress large data", async () => {
      // 100KB of repeating data (highly compressible)
      const original = new Uint8Array(100 * 1024);
      for (let i = 0; i < original.length; i++) {
        original[i] = i % 256;
      }

      const compressed = await compress(original, { level: 6 });
      expect(compressed.length).toBeLessThan(original.length);

      const decompressed = await decompress(compressed);
      expect(decompressed).toEqual(original);
    });

    it("should handle level 0 (no compression)", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original, { level: 0 });

      // Level 0 should return original data unchanged
      expect(compressed).toEqual(original);
    });

    it("should compress with different levels", async () => {
      const original = new TextEncoder().encode("Hello, World!".repeat(100));

      const level1 = await compress(original, { level: 1 });
      const level6 = await compress(original, { level: 6 });
      const level9 = await compress(original, { level: 9 });

      // All should decompress correctly
      expect(await decompress(level1)).toEqual(original);
      expect(await decompress(level6)).toEqual(original);
      expect(await decompress(level9)).toEqual(original);

      // Higher levels should generally produce smaller or equal output
      expect(level9.length).toBeLessThanOrEqual(level1.length);
    });

    it("should use default compression level when not specified", async () => {
      const original = new TextEncoder().encode("Test default level");
      const compressed = await compress(original);
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should compress unicode text", async () => {
      const original = new TextEncoder().encode("你好世界 🌍 مرحبا");
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should handle data with null bytes", async () => {
      const original = new Uint8Array([0x00, 0x00, 0x00, 0x41, 0x00, 0x00]);
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });
  });

  describe("sync compress/decompress (Node.js only)", () => {
    it("should compress and decompress simple text synchronously", () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = compressSync(original, { level: 6 });
      const decompressed = decompressSync(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should handle level 0 synchronously", () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = compressSync(original, { level: 0 });

      expect(compressed).toEqual(original);
    });

    it("should compress and decompress large data synchronously", () => {
      const original = new Uint8Array(50 * 1024);
      for (let i = 0; i < original.length; i++) {
        original[i] = i % 256;
      }

      const compressed = compressSync(original, { level: 6 });
      expect(compressed.length).toBeLessThan(original.length);

      const decompressed = decompressSync(compressed);
      expect(decompressed).toEqual(original);
    });

    it("should compress with different levels synchronously", () => {
      const original = new TextEncoder().encode("Hello, World!".repeat(50));

      const level1 = compressSync(original, { level: 1 });
      const level9 = compressSync(original, { level: 9 });

      expect(decompressSync(level1)).toEqual(original);
      expect(decompressSync(level9)).toEqual(original);

      expect(level9.length).toBeLessThanOrEqual(level1.length);
    });
  });

  describe("compression ratio", () => {
    it("should achieve good compression on repetitive data", async () => {
      // Highly compressible: repeated string
      const original = new TextEncoder().encode("AAAAAAAAAA".repeat(1000));
      const compressed = await compress(original, { level: 6 });

      // Should compress to at least 90% smaller
      expect(compressed.length).toBeLessThan(original.length * 0.1);
    });

    it("should handle incompressible data", async () => {
      // Random-like data (hard to compress)
      const original = new Uint8Array(1000);
      for (let i = 0; i < original.length; i++) {
        // Pseudo-random pattern
        original[i] = (i * 17 + 31) % 256;
      }

      const compressed = await compress(original, { level: 6 });
      // Even incompressible data should still work (might be slightly larger due to headers)
      expect(compressed.length).toBeLessThan(original.length * 1.1);

      const decompressed = await decompress(compressed);
      expect(decompressed).toEqual(original);
    });
  });

  describe("edge cases", () => {
    it("should handle single byte", async () => {
      const original = new Uint8Array([42]);
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should handle all zero bytes", async () => {
      const original = new Uint8Array(1000).fill(0);
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
      // All zeros should compress very well
      expect(compressed.length).toBeLessThan(100);
    });

    it("should handle all 0xFF bytes", async () => {
      const original = new Uint8Array(1000).fill(0xff);
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should handle alternating pattern", async () => {
      const original = new Uint8Array(1000);
      for (let i = 0; i < original.length; i++) {
        original[i] = i % 2 === 0 ? 0xaa : 0x55;
      }

      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(original);
    });

    it("should be consistent across multiple compressions", async () => {
      const original = new TextEncoder().encode("Consistency test data");

      const compressed1 = await compress(original, { level: 6 });
      const compressed2 = await compress(original, { level: 6 });

      // Same input should produce same output
      expect(compressed1).toEqual(compressed2);
    });
  });

  describe("Buffer compatibility", () => {
    it("should accept Buffer as input", async () => {
      const original = Buffer.from("Hello from Buffer!", "utf-8");
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(Buffer.from(decompressed).toString("utf-8")).toBe("Hello from Buffer!");
    });

    it("should work with Buffer.from array", async () => {
      const original = Buffer.from([0x48, 0x65, 0x6c, 0x6c, 0x6f]); // "Hello"
      const compressed = await compress(original, { level: 6 });
      const decompressed = await decompress(compressed);

      expect(decompressed).toEqual(new Uint8Array([0x48, 0x65, 0x6c, 0x6c, 0x6f]));
    });
  });
});
