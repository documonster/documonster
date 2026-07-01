/**
 * Tests for archive byte utilities.
 */

import { EMPTY_UINT8ARRAY, indexOfUint8ArrayPattern } from "@archive/core/bytes";
import { describe, it, expect } from "vitest";

describe("bytes", () => {
  describe("EMPTY_UINT8ARRAY", () => {
    it("should be an empty Uint8Array", () => {
      expect(EMPTY_UINT8ARRAY).toBeInstanceOf(Uint8Array);
      expect(EMPTY_UINT8ARRAY.length).toBe(0);
    });

    it("should be reusable (same reference)", () => {
      const a = EMPTY_UINT8ARRAY;
      const b = EMPTY_UINT8ARRAY;
      expect(a).toBe(b);
    });
  });

  describe("indexOfUint8ArrayPattern", () => {
    describe("empty pattern", () => {
      it("should return 0 for empty pattern", () => {
        const buffer = Uint8Array.from([1, 2, 3]);
        const pattern = new Uint8Array(0);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });
    });

    describe("pattern longer than buffer", () => {
      it("should return -1 when pattern is longer than buffer", () => {
        const buffer = Uint8Array.from([1, 2]);
        const pattern = Uint8Array.from([1, 2, 3]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(-1);
      });
    });

    describe("1-byte pattern", () => {
      it("should find single byte at start", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });

      it("should find single byte in middle", () => {
        const buffer = Uint8Array.from([0x00, 0x50, 0x4b, 0x03]);
        const pattern = Uint8Array.from([0x50]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(1);
      });

      it("should return -1 when byte not found", () => {
        const buffer = Uint8Array.from([0x00, 0x01, 0x02]);
        const pattern = Uint8Array.from([0xff]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(-1);
      });
    });

    describe("2-byte pattern", () => {
      it("should find 2-byte pattern at start", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });

      it("should find 2-byte pattern in middle", () => {
        const buffer = Uint8Array.from([0x00, 0x50, 0x4b, 0x03]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(1);
      });

      it("should find 2-byte pattern at end", () => {
        const buffer = Uint8Array.from([0x00, 0x01, 0x50, 0x4b]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(2);
      });

      it("should return -1 when first byte matches but second does not", () => {
        const buffer = Uint8Array.from([0x50, 0x00, 0x50, 0x01]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(-1);
      });
    });

    describe("3-byte pattern", () => {
      it("should find 3-byte pattern", () => {
        const buffer = Uint8Array.from([0x00, 0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x03]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(1);
      });

      it("should return -1 when only partial match", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x00, 0x50, 0x4b, 0x01]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x03]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(-1);
      });
    });

    describe("4-byte pattern (ZIP signatures)", () => {
      it("should find local file header signature", () => {
        // ZIP local file header signature: 0x04034b50 (little-endian)
        const buffer = Uint8Array.from([0x00, 0x50, 0x4b, 0x03, 0x04, 0x00]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(1);
      });

      it("should find central directory signature", () => {
        // ZIP central directory signature: 0x02014b50
        const buffer = Uint8Array.from([0x50, 0x4b, 0x01, 0x02]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x01, 0x02]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });

      it("should find EOCD signature", () => {
        // End of central directory signature: 0x06054b50
        const buffer = Uint8Array.from([0x00, 0x00, 0x50, 0x4b, 0x05, 0x06, 0x00, 0x00]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x05, 0x06]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(2);
      });

      it("should return first occurrence when multiple matches exist", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });
    });

    describe("longer patterns (>4 bytes)", () => {
      it("should find 5-byte pattern", () => {
        const buffer = Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x05, 0x06]);
        const pattern = Uint8Array.from([0x02, 0x03, 0x04, 0x05, 0x06]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(1);
      });

      it("should find pattern in large buffer", () => {
        const buffer = new Uint8Array(1000);
        buffer.set([0xde, 0xad, 0xbe, 0xef, 0xca, 0xfe], 500);
        const pattern = Uint8Array.from([0xde, 0xad, 0xbe, 0xef, 0xca, 0xfe]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(500);
      });
    });

    describe("startIndex parameter", () => {
      it("should start searching from specified index", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x00, 0x50, 0x4b, 0x00]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern, 1)).toBe(3);
      });

      it("should return -1 if startIndex is beyond last possible match", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern, 4)).toBe(-1);
      });

      it("should handle negative startIndex as 0", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03]);
        const pattern = Uint8Array.from([0x50]);
        expect(indexOfUint8ArrayPattern(buffer, pattern, -5)).toBe(0);
      });

      it("should handle startIndex equal to buffer length", () => {
        const buffer = Uint8Array.from([0x50, 0x4b]);
        const pattern = Uint8Array.from([0x50]);
        expect(indexOfUint8ArrayPattern(buffer, pattern, 2)).toBe(-1);
      });
    });

    describe("edge cases", () => {
      it("should handle exact buffer match", () => {
        const buffer = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        const pattern = Uint8Array.from([0x50, 0x4b, 0x03, 0x04]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });

      it("should handle pattern at very end", () => {
        const buffer = Uint8Array.from([0x00, 0x00, 0x00, 0x50, 0x4b]);
        const pattern = Uint8Array.from([0x50, 0x4b]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(3);
      });

      it("should handle all same bytes", () => {
        const buffer = Uint8Array.from([0xff, 0xff, 0xff, 0xff]);
        const pattern = Uint8Array.from([0xff, 0xff]);
        expect(indexOfUint8ArrayPattern(buffer, pattern)).toBe(0);
      });
    });
  });
});
