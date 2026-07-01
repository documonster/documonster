/**
 * Tests for archive text encoding/decoding utilities.
 */

import { crc32 } from "@archive/compression/crc32";
import type { ZipStringCodec } from "@archive/core/text";
import {
  decodeCp437,
  encodeCp437,
  resolveZipStringCodec,
  encodeZipString,
  encodeZipStringWithCodec,
  decodeZipString,
  toArrayBuffer
} from "@archive/core/text";
import { describe, it, expect } from "vitest";

describe("text", () => {
  describe("decodeCp437", () => {
    it("should decode ASCII bytes unchanged", () => {
      const bytes = new TextEncoder().encode("Hello, World!");
      expect(decodeCp437(bytes)).toBe("Hello, World!");
    });

    it("should decode empty array", () => {
      expect(decodeCp437(new Uint8Array(0))).toBe("");
    });

    it("should decode high bytes to CP437 characters", () => {
      // 0x80 = Ç, 0x81 = ü, 0x82 = é
      const bytes = Uint8Array.from([0x80, 0x81, 0x82]);
      expect(decodeCp437(bytes)).toBe("Çüé");
    });

    it("should decode box-drawing characters", () => {
      // 0xB3 = │, 0xC4 = ─
      const bytes = Uint8Array.from([0xb3, 0xc4]);
      expect(decodeCp437(bytes)).toBe("│─");
    });

    it("should decode Greek letters", () => {
      // 0xE0 = α, 0xE1 = ß (German), 0xE2 = Γ, 0xE3 = π
      const bytes = Uint8Array.from([0xe0, 0xe1, 0xe2, 0xe3]);
      expect(decodeCp437(bytes)).toBe("αßΓπ");
    });

    it("should decode math symbols", () => {
      // 0xF0 = ≡, 0xF1 = ±, 0xFB = √
      const bytes = Uint8Array.from([0xf0, 0xf1, 0xfb]);
      expect(decodeCp437(bytes)).toBe("≡±√");
    });

    it("should decode mixed ASCII and high bytes", () => {
      // "caf" + 0x82 (é) = "café"
      const bytes = Uint8Array.from([0x63, 0x61, 0x66, 0x82]);
      expect(decodeCp437(bytes)).toBe("café");
    });

    it("should handle large ASCII-only input efficiently", () => {
      const text = "a".repeat(100000);
      const bytes = new TextEncoder().encode(text);
      expect(decodeCp437(bytes)).toBe(text);
    });
  });

  describe("encodeCp437", () => {
    it("should encode ASCII characters", () => {
      const result = encodeCp437("Hello");
      expect(result).toEqual(new TextEncoder().encode("Hello"));
    });

    it("should encode empty string", () => {
      expect(encodeCp437("")).toEqual(new Uint8Array(0));
    });

    it("should encode CP437 high characters", () => {
      const result = encodeCp437("Çüé");
      expect(result).toEqual(Uint8Array.from([0x80, 0x81, 0x82]));
    });

    it("should encode box-drawing characters", () => {
      const result = encodeCp437("│─");
      expect(result).toEqual(Uint8Array.from([0xb3, 0xc4]));
    });

    it("should replace unmappable characters with ?", () => {
      // Chinese character is not in CP437
      const result = encodeCp437("Hello中文");
      expect(result).toEqual(Uint8Array.from([0x48, 0x65, 0x6c, 0x6c, 0x6f, 0x3f, 0x3f]));
    });

    it("should roundtrip CP437 characters", () => {
      const original = "café │─ αßπ ≡±√";
      const encoded = encodeCp437(original);
      const decoded = decodeCp437(encoded);
      expect(decoded).toBe(original);
    });
  });

  describe("resolveZipStringCodec", () => {
    it("should return UTF-8 codec for undefined", () => {
      const codec = resolveZipStringCodec(undefined);
      expect(codec.name).toBe("utf-8");
      expect(codec.useUtf8Flag).toBe(true);
      expect(codec.useUnicodeExtraFields).toBe(false);
    });

    it("should return UTF-8 codec for 'utf-8'", () => {
      const codec = resolveZipStringCodec("utf-8");
      expect(codec.name).toBe("utf-8");
      expect(codec.useUtf8Flag).toBe(true);
    });

    it("should return CP437 codec for 'cp437'", () => {
      const codec = resolveZipStringCodec("cp437");
      expect(codec.name).toBe("cp437");
      expect(codec.useUtf8Flag).toBe(false);
      expect(codec.useUnicodeExtraFields).toBe(true);
    });

    it("should resolve custom codec with defaults", () => {
      const custom: ZipStringCodec = {
        encode: s => new TextEncoder().encode(s),
        decode: b => new TextDecoder().decode(b)
      };

      const resolved = resolveZipStringCodec(custom);
      expect(resolved.encode).toBe(custom.encode);
      expect(resolved.decode).toBe(custom.decode);
      expect(resolved.useUtf8Flag).toBe(false); // default when name !== "utf-8"
      expect(resolved.useUnicodeExtraFields).toBe(true);
    });

    it("should respect custom codec useUtf8Flag", () => {
      const custom: ZipStringCodec = {
        name: "custom",
        encode: s => new TextEncoder().encode(s),
        decode: b => new TextDecoder().decode(b),
        useUtf8Flag: true
      };

      const resolved = resolveZipStringCodec(custom);
      expect(resolved.useUtf8Flag).toBe(true);
      expect(resolved.useUnicodeExtraFields).toBe(false);
    });

    it("should cache resolved custom codecs", () => {
      const custom: ZipStringCodec = {
        encode: s => new TextEncoder().encode(s),
        decode: b => new TextDecoder().decode(b)
      };

      const first = resolveZipStringCodec(custom);
      const second = resolveZipStringCodec(custom);
      expect(first).toBe(second);
    });

    it("should infer useUtf8Flag from name=utf-8", () => {
      const custom: ZipStringCodec = {
        name: "utf-8",
        encode: s => new TextEncoder().encode(s),
        decode: b => new TextDecoder().decode(b)
      };

      const resolved = resolveZipStringCodec(custom);
      expect(resolved.useUtf8Flag).toBe(true);
    });
  });

  describe("encodeZipString", () => {
    it("should return empty array for undefined", () => {
      const result = encodeZipString(undefined);
      expect(result.length).toBe(0);
    });

    it("should return empty array for empty string", () => {
      const result = encodeZipString("");
      expect(result.length).toBe(0);
    });

    it("should encode with default UTF-8", () => {
      const result = encodeZipString("Hello");
      expect(result).toEqual(new TextEncoder().encode("Hello"));
    });

    it("should encode with CP437", () => {
      const result = encodeZipString("café", "cp437");
      expect(result).toEqual(encodeCp437("café"));
    });

    it("should return shared empty buffer for multiple calls", () => {
      const a = encodeZipString(undefined);
      const b = encodeZipString("");
      expect(a).toBe(b);
    });
  });

  describe("encodeZipStringWithCodec", () => {
    it("should use provided codec", () => {
      const codec = resolveZipStringCodec("cp437");
      const result = encodeZipStringWithCodec("Çüé", codec);
      expect(result).toEqual(Uint8Array.from([0x80, 0x81, 0x82]));
    });

    it("should return empty for undefined value", () => {
      const codec = resolveZipStringCodec("utf-8");
      const result = encodeZipStringWithCodec(undefined, codec);
      expect(result.length).toBe(0);
    });
  });

  describe("decodeZipString", () => {
    it("should return empty string for empty bytes", () => {
      expect(decodeZipString(new Uint8Array(0), 0)).toBe("");
    });

    it("should decode as UTF-8 when UTF-8 flag is set", () => {
      const bytes = new TextEncoder().encode("Hello 世界");
      const FLAG_UTF8 = 0x0800;
      expect(decodeZipString(bytes, FLAG_UTF8)).toBe("Hello 世界");
    });

    it("should decode as CP437 when UTF-8 flag is not set", () => {
      const bytes = Uint8Array.from([0x63, 0x61, 0x66, 0x82]); // "caf" + é
      expect(decodeZipString(bytes, 0)).toBe("café");
    });

    it("should use Unicode extra field when CRC32 matches", () => {
      const originalBytes = new TextEncoder().encode("test");
      const unicodeValue = "test-unicode";
      const unicodeInfo = {
        version: 1,
        originalCrc32: crc32(originalBytes),
        unicodeValue
      };

      expect(decodeZipString(originalBytes, 0, unicodeInfo)).toBe(unicodeValue);
    });

    it("should ignore Unicode extra field when CRC32 does not match", () => {
      const originalBytes = Uint8Array.from([0x63, 0x61, 0x66, 0x82]);
      const unicodeInfo = {
        version: 1,
        originalCrc32: 0x12345678, // wrong CRC
        unicodeValue: "should-not-use"
      };

      expect(decodeZipString(originalBytes, 0, unicodeInfo)).toBe("café");
    });

    it("should ignore Unicode extra field when version is not 1", () => {
      const originalBytes = new TextEncoder().encode("test");
      const unicodeInfo = {
        version: 2,
        originalCrc32: crc32(originalBytes),
        unicodeValue: "should-not-use"
      };

      expect(decodeZipString(originalBytes, 0, unicodeInfo)).toBe("test");
    });

    it("should use fallback decoder when provided", () => {
      const bytes = Uint8Array.from([0x48, 0x69]); // "Hi"
      const fallback = {
        decode: () => "custom-decoded"
      };

      expect(decodeZipString(bytes, 0, undefined, fallback)).toBe("custom-decoded");
    });

    it("should prioritize UTF-8 flag over Unicode extra field", () => {
      const bytes = new TextEncoder().encode("UTF-8 value");
      const FLAG_UTF8 = 0x0800;
      const unicodeInfo = {
        version: 1,
        originalCrc32: crc32(bytes),
        unicodeValue: "extra-field-value"
      };

      expect(decodeZipString(bytes, FLAG_UTF8, unicodeInfo)).toBe("UTF-8 value");
    });
  });

  describe("toArrayBuffer", () => {
    it("should return underlying buffer for zero-offset full-length view", () => {
      const buffer = new ArrayBuffer(10);
      const view = new Uint8Array(buffer);
      const result = toArrayBuffer(view);
      expect(result).toBe(buffer);
    });

    it("should copy when view has non-zero offset", () => {
      const buffer = new ArrayBuffer(10);
      const view = new Uint8Array(buffer, 2, 5);
      const result = toArrayBuffer(view);
      expect(result).not.toBe(buffer);
      expect(result.byteLength).toBe(5);
    });

    it("should copy when view does not cover full buffer", () => {
      const buffer = new ArrayBuffer(10);
      const view = new Uint8Array(buffer, 0, 5);
      const result = toArrayBuffer(view);
      expect(result).not.toBe(buffer);
      expect(result.byteLength).toBe(5);
    });

    it("should handle empty view", () => {
      const view = new Uint8Array(0);
      const result = toArrayBuffer(view);
      expect(result.byteLength).toBe(0);
    });
  });
});
