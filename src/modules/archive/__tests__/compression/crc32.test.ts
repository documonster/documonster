import { crc32, crc32Update, crc32Finalize } from "@archive/compression/crc32";
import { crc32UpdateByte } from "@archive/compression/crc32.base";
import { describe, it, expect } from "vitest";

describe("crc32", () => {
  describe("basic crc32 calculation", () => {
    it("should return 0 for empty data", () => {
      const data = new Uint8Array(0);
      expect(crc32(data)).toBe(0);
    });

    it("should calculate correct CRC for 'Hello, World!'", () => {
      const data = new TextEncoder().encode("Hello, World!");
      // Known CRC32 value for "Hello, World!"
      expect(crc32(data)).toBe(0xec4ac3d0);
    });

    it("should calculate correct CRC for single byte", () => {
      const data = new Uint8Array([0x00]);
      expect(crc32(data)).toBe(0xd202ef8d);
    });

    it("should calculate correct CRC for 0xFF byte", () => {
      const data = new Uint8Array([0xff]);
      expect(crc32(data)).toBe(0xff000000);
    });

    it("should calculate correct CRC for 'test'", () => {
      const data = new TextEncoder().encode("test");
      // Standard CRC32 for "test"
      expect(crc32(data)).toBe(0xd87f7e0c);
    });

    it("should calculate correct CRC for binary data", () => {
      const data = new Uint8Array([0x01, 0x02, 0x03, 0x04, 0x05]);
      expect(crc32(data)).toBe(0x470b99f4);
    });

    it("should handle large data (1MB)", () => {
      const data = new Uint8Array(1024 * 1024);
      // Fill with repeating pattern
      for (let i = 0; i < data.length; i++) {
        data[i] = i & 0xff;
      }
      // Should complete without error and return consistent result
      const result = crc32(data);
      expect(typeof result).toBe("number");
      expect(result).toBeGreaterThanOrEqual(0);
      expect(result).toBeLessThanOrEqual(0xffffffff);
      // Same data should produce same CRC
      expect(crc32(data)).toBe(result);
    });

    it("should return unsigned 32-bit integer", () => {
      const data = new TextEncoder().encode("test");
      const result = crc32(data);
      expect(result).toBeGreaterThanOrEqual(0);
      expect(result).toBeLessThanOrEqual(0xffffffff);
    });

    it("should work with Buffer (Node.js)", () => {
      const data = Buffer.from("Hello, World!", "utf-8");
      expect(crc32(data)).toBe(0xec4ac3d0);
    });
  });

  describe("incremental crc32 (crc32Update + crc32Finalize)", () => {
    it("should produce same result as single calculation", () => {
      const fullData = new TextEncoder().encode("Hello, World!");

      // Single calculation
      const singleResult = crc32(fullData);

      // Incremental calculation
      let crc = 0xffffffff;
      crc = crc32Update(crc, fullData);
      const incrementalResult = crc32Finalize(crc);

      expect(incrementalResult).toBe(singleResult);
    });

    it("should produce same result when data is split into chunks", () => {
      const part1 = new TextEncoder().encode("Hello, ");
      const part2 = new TextEncoder().encode("World!");
      const fullData = new TextEncoder().encode("Hello, World!");

      // Single calculation
      const singleResult = crc32(fullData);

      // Incremental calculation with chunks
      let crc = 0xffffffff;
      crc = crc32Update(crc, part1);
      crc = crc32Update(crc, part2);
      const incrementalResult = crc32Finalize(crc);

      expect(incrementalResult).toBe(singleResult);
    });

    it("should handle many small chunks", () => {
      const text = "abcdefghijklmnopqrstuvwxyz";
      const fullData = new TextEncoder().encode(text);
      const singleResult = crc32(fullData);

      // Calculate one character at a time
      let crc = 0xffffffff;
      for (let i = 0; i < text.length; i++) {
        const chunk = new TextEncoder().encode(text[i]);
        crc = crc32Update(crc, chunk);
      }
      const incrementalResult = crc32Finalize(crc);

      expect(incrementalResult).toBe(singleResult);
    });

    it("should handle empty chunks in the middle", () => {
      const part1 = new TextEncoder().encode("Hello");
      const empty = new Uint8Array(0);
      const part2 = new TextEncoder().encode("World");
      const fullData = new TextEncoder().encode("HelloWorld");

      const singleResult = crc32(fullData);

      let crc = 0xffffffff;
      crc = crc32Update(crc, part1);
      crc = crc32Update(crc, empty);
      crc = crc32Update(crc, part2);
      const incrementalResult = crc32Finalize(crc);

      expect(incrementalResult).toBe(singleResult);
    });
  });

  describe("edge cases", () => {
    it("should handle data with all zeros", () => {
      const data = new Uint8Array(100);
      const result = crc32(data);
      expect(typeof result).toBe("number");
      // Should be consistent
      expect(crc32(data)).toBe(result);
    });

    it("should handle data with all 0xFF", () => {
      const data = new Uint8Array(100).fill(0xff);
      const result = crc32(data);
      expect(typeof result).toBe("number");
      expect(crc32(data)).toBe(result);
    });

    it("should handle unicode strings", () => {
      const data = new TextEncoder().encode("你好世界 🌍");
      const result = crc32(data);
      expect(typeof result).toBe("number");
      expect(crc32(data)).toBe(result);
    });

    it("should handle newlines and special characters", () => {
      const data = new TextEncoder().encode("line1\nline2\r\nline3\ttab");
      const result = crc32(data);
      expect(typeof result).toBe("number");
      expect(crc32(data)).toBe(result);
    });

    it("should produce different CRCs for different data", () => {
      const data1 = new TextEncoder().encode("hello");
      const data2 = new TextEncoder().encode("Hello");
      const data3 = new TextEncoder().encode("hello ");

      const crc1 = crc32(data1);
      const crc2 = crc32(data2);
      const crc3 = crc32(data3);

      expect(crc1).not.toBe(crc2);
      expect(crc1).not.toBe(crc3);
      expect(crc2).not.toBe(crc3);
    });

    it("should be deterministic", () => {
      const data = new TextEncoder().encode("deterministic test");
      const results = Array.from({ length: 10 }, () => crc32(data));
      expect(new Set(results).size).toBe(1);
    });
  });

  describe("known test vectors", () => {
    // Test vectors from various CRC32 implementations
    const testVectors: [string, number][] = [
      ["", 0x00000000],
      ["a", 0xe8b7be43],
      ["abc", 0x352441c2],
      ["message digest", 0x20159d7f],
      ["abcdefghijklmnopqrstuvwxyz", 0x4c2750bd],
      ["ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", 0x1fc2e6d2],
      [
        "12345678901234567890123456789012345678901234567890123456789012345678901234567890",
        0x7ca94a72
      ]
    ];

    testVectors.forEach(([input, expected]) => {
      it(`should calculate correct CRC for "${input.substring(0, 20)}${input.length > 20 ? "..." : ""}"`, () => {
        const data = new TextEncoder().encode(input);
        expect(crc32(data)).toBe(expected);
      });
    });
  });

  describe("crc32UpdateByte", () => {
    it("should update CRC with a single byte (non-inverted form)", () => {
      // crc32UpdateByte uses the raw table lookup (non-inverted form)
      // This is used by ZipCrypto key derivation
      const initialCrc = 0x12345678;
      const byte = 0x42;

      const result = crc32UpdateByte(initialCrc, byte);

      // Result should be unsigned 32-bit
      expect(result).toBeGreaterThanOrEqual(0);
      expect(result).toBeLessThanOrEqual(0xffffffff);
      expect(result >>> 0).toBe(result);
    });

    it("should produce consistent results for same input", () => {
      const crc = 0x12345678;
      const byte = 0xab;

      const result1 = crc32UpdateByte(crc, byte);
      const result2 = crc32UpdateByte(crc, byte);

      expect(result1).toBe(result2);
    });

    it("should produce different results for different bytes", () => {
      const crc = 0x12345678;

      const result1 = crc32UpdateByte(crc, 0x00);
      const result2 = crc32UpdateByte(crc, 0xff);

      expect(result1).not.toBe(result2);
    });

    it("should match byte-by-byte update with crc32Update", () => {
      // Test that feeding bytes one at a time through crc32UpdateByte
      // produces results consistent with the bulk crc32Update
      const data = new TextEncoder().encode("test");

      // Using crc32Update (bulk)
      let bulkCrc = 0xffffffff;
      bulkCrc = crc32Update(bulkCrc, data);

      // Using crc32UpdateByte (one byte at a time)
      // Note: crc32UpdateByte uses the non-inverted form, so we need to
      // start with the XOR form
      let byteCrc = 0xffffffff;
      for (let i = 0; i < data.length; i++) {
        byteCrc = crc32UpdateByte(byteCrc, data[i]!);
      }

      // Both should produce the same intermediate state
      expect(byteCrc).toBe(bulkCrc);
    });
  });
});
