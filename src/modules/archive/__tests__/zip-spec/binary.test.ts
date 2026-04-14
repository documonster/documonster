/**
 * Tests for binary reader/writer utilities.
 */

import {
  BinaryReader,
  writeUint32LE,
  readUint32LE,
  parseFormatted,
  parseFormattedTyped
} from "@archive/zip-spec/binary";
import { describe, it, expect } from "vitest";

describe("binary", () => {
  describe("writeUint32LE", () => {
    it("should write a 32-bit unsigned integer in little-endian", () => {
      const result = writeUint32LE(0x12345678);
      expect(result).toEqual(Uint8Array.from([0x78, 0x56, 0x34, 0x12]));
    });

    it("should handle zero", () => {
      const result = writeUint32LE(0);
      expect(result).toEqual(Uint8Array.from([0, 0, 0, 0]));
    });

    it("should handle max uint32", () => {
      const result = writeUint32LE(0xffffffff);
      expect(result).toEqual(Uint8Array.from([0xff, 0xff, 0xff, 0xff]));
    });

    it("should handle negative numbers by wrapping", () => {
      // -1 >>> 0 = 0xFFFFFFFF
      const result = writeUint32LE(-1);
      expect(result).toEqual(Uint8Array.from([0xff, 0xff, 0xff, 0xff]));
    });
  });

  describe("readUint32LE", () => {
    it("should read a 32-bit unsigned integer in little-endian", () => {
      const data = Uint8Array.from([0x78, 0x56, 0x34, 0x12]);
      expect(readUint32LE(data, 0)).toBe(0x12345678);
    });

    it("should read from specified offset", () => {
      const data = Uint8Array.from([0x00, 0x00, 0x78, 0x56, 0x34, 0x12]);
      expect(readUint32LE(data, 2)).toBe(0x12345678);
    });

    it("should handle zero", () => {
      const data = Uint8Array.from([0, 0, 0, 0]);
      expect(readUint32LE(data, 0)).toBe(0);
    });

    it("should handle max uint32", () => {
      const data = Uint8Array.from([0xff, 0xff, 0xff, 0xff]);
      expect(readUint32LE(data, 0)).toBe(0xffffffff);
    });
  });

  describe("BinaryReader", () => {
    describe("construction and position", () => {
      it("should start at position 0 by default", () => {
        const reader = new BinaryReader(new Uint8Array(10));
        expect(reader.position).toBe(0);
      });

      it("should start at specified offset", () => {
        const reader = new BinaryReader(new Uint8Array(10), 5);
        expect(reader.position).toBe(5);
      });

      it("should allow setting position", () => {
        const reader = new BinaryReader(new Uint8Array(10));
        reader.position = 7;
        expect(reader.position).toBe(7);
      });

      it("should report remaining bytes", () => {
        const reader = new BinaryReader(new Uint8Array(10), 3);
        expect(reader.remaining).toBe(7);
      });
    });

    describe("readUint8", () => {
      it("should read a single byte and advance position", () => {
        const data = Uint8Array.from([0x12, 0x34, 0x56]);
        const reader = new BinaryReader(data);

        expect(reader.readUint8()).toBe(0x12);
        expect(reader.position).toBe(1);
        expect(reader.readUint8()).toBe(0x34);
        expect(reader.position).toBe(2);
      });
    });

    describe("readUint16", () => {
      it("should read a 16-bit LE integer and advance position", () => {
        const data = Uint8Array.from([0x34, 0x12, 0x78, 0x56]);
        const reader = new BinaryReader(data);

        expect(reader.readUint16()).toBe(0x1234);
        expect(reader.position).toBe(2);
        expect(reader.readUint16()).toBe(0x5678);
        expect(reader.position).toBe(4);
      });
    });

    describe("readUint32", () => {
      it("should read a 32-bit LE integer and advance position", () => {
        const data = Uint8Array.from([0x78, 0x56, 0x34, 0x12]);
        const reader = new BinaryReader(data);

        expect(reader.readUint32()).toBe(0x12345678);
        expect(reader.position).toBe(4);
      });
    });

    describe("readBigUint64", () => {
      it("should read a 64-bit LE integer and advance position", () => {
        const data = Uint8Array.from([0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01]);
        const reader = new BinaryReader(data);

        expect(reader.readBigUint64()).toBe(0x0100000000000001n);
        expect(reader.position).toBe(8);
      });

      it("should handle large values", () => {
        // 0xFFFFFFFFFFFFFFFF
        const data = Uint8Array.from([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff]);
        const reader = new BinaryReader(data);

        expect(reader.readBigUint64()).toBe(0xffffffffffffffffn);
      });
    });

    describe("readBytes", () => {
      it("should read specified number of bytes and advance position", () => {
        const data = Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x05]);
        const reader = new BinaryReader(data);

        const bytes = reader.readBytes(3);
        expect(bytes).toEqual(Uint8Array.from([0x01, 0x02, 0x03]));
        expect(reader.position).toBe(3);
      });

      it("should return a subarray (view, not copy)", () => {
        const data = Uint8Array.from([0x01, 0x02, 0x03]);
        const reader = new BinaryReader(data);

        const bytes = reader.readBytes(2);
        expect(bytes.buffer).toBe(data.buffer);
      });
    });

    describe("readString", () => {
      it("should read UTF-8 string by default", () => {
        const data = new TextEncoder().encode("Hello");
        const reader = new BinaryReader(data);

        expect(reader.readString(5)).toBe("Hello");
        expect(reader.position).toBe(5);
      });

      it("should read CP437 string when utf8=false", () => {
        // "caf" + 0x82 (é in CP437)
        const data = Uint8Array.from([0x63, 0x61, 0x66, 0x82]);
        const reader = new BinaryReader(data);

        expect(reader.readString(4, false)).toBe("café");
      });
    });

    describe("skip", () => {
      it("should advance position by specified amount", () => {
        const reader = new BinaryReader(new Uint8Array(20));
        reader.skip(5);
        expect(reader.position).toBe(5);
        reader.skip(3);
        expect(reader.position).toBe(8);
      });
    });

    describe("slice", () => {
      it("should return a subarray without changing position", () => {
        const data = Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x05]);
        const reader = new BinaryReader(data);

        const slice = reader.slice(1, 4);
        expect(slice).toEqual(Uint8Array.from([0x02, 0x03, 0x04]));
        expect(reader.position).toBe(0); // unchanged
      });
    });

    describe("peekUint32", () => {
      it("should read uint32 at offset without changing position", () => {
        const data = Uint8Array.from([0x01, 0x02, 0x03, 0x04, 0x78, 0x56, 0x34, 0x12]);
        const reader = new BinaryReader(data);

        expect(reader.peekUint32(4)).toBe(0x12345678);
        expect(reader.position).toBe(0); // unchanged
      });
    });

    describe("complex reading", () => {
      it("should correctly read a mock ZIP local file header", () => {
        // Simplified local file header structure
        const header = new Uint8Array(30);
        const view = new DataView(header.buffer);

        view.setUint32(0, 0x04034b50, true); // signature
        view.setUint16(4, 20, true); // version needed
        view.setUint16(6, 0, true); // flags
        view.setUint16(8, 8, true); // compression method (deflate)
        view.setUint16(10, 0x5678, true); // mod time
        view.setUint16(12, 0x1234, true); // mod date
        view.setUint32(14, 0xaabbccdd, true); // crc32
        view.setUint32(18, 100, true); // compressed size
        view.setUint32(22, 200, true); // uncompressed size
        view.setUint16(26, 8, true); // filename length
        view.setUint16(28, 0, true); // extra field length

        const reader = new BinaryReader(header);

        expect(reader.readUint32()).toBe(0x04034b50); // signature
        expect(reader.readUint16()).toBe(20); // version needed
        expect(reader.readUint16()).toBe(0); // flags
        expect(reader.readUint16()).toBe(8); // compression
        expect(reader.readUint16()).toBe(0x5678); // mod time
        expect(reader.readUint16()).toBe(0x1234); // mod date
        expect(reader.readUint32()).toBe(0xaabbccdd); // crc32
        expect(reader.readUint32()).toBe(100); // compressed size
        expect(reader.readUint32()).toBe(200); // uncompressed size
        expect(reader.readUint16()).toBe(8); // filename length
        expect(reader.readUint16()).toBe(0); // extra length
        expect(reader.position).toBe(30);
      });
    });
  });

  describe("parseFormatted", () => {
    it("should parse sequential fields according to format", () => {
      const data = Uint8Array.from([0x12, 0x34, 0x12, 0x78, 0x56, 0x34, 0x12]);
      const format: [string, number][] = [
        ["byte1", 1],
        ["word", 2],
        ["dword", 4]
      ];

      const result = parseFormatted(data, format);

      expect(result.byte1).toBe(0x12);
      expect(result.word).toBe(0x1234);
      expect(result.dword).toBe(0x12345678);
    });

    it("should return null for fields beyond buffer length", () => {
      const data = Uint8Array.from([0x12, 0x34]);
      const format: [string, number][] = [
        ["a", 1],
        ["b", 1],
        ["c", 1] // beyond buffer
      ];

      const result = parseFormatted(data, format);

      expect(result.a).toBe(0x12);
      expect(result.b).toBe(0x34);
      expect(result.c).toBeNull();
    });

    it("should handle 8-byte fields as Number", () => {
      const data = Uint8Array.from([0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01]);
      const format: [string, number][] = [["qword", 8]];

      const result = parseFormatted(data, format);

      // Little-endian: low 4 bytes = 0x00000001, high 4 bytes = 0x01000000
      // Result = high * 0x100000000 + low
      expect(result.qword).toBe(0x01000000 * 0x100000000 + 1);
    });

    it("should handle empty format", () => {
      const data = Uint8Array.from([0x12, 0x34]);
      const result = parseFormatted(data, []);
      expect(result).toEqual({});
    });

    it("should throw for unsupported size", () => {
      const data = Uint8Array.from([0x12, 0x34, 0x56]);
      const format: [string, number][] = [["invalid", 3]];

      expect(() => parseFormatted(data, format)).toThrow("Unsupported UInt LE size!");
    });
  });

  describe("parseFormattedTyped", () => {
    interface TestFormat {
      signature: number;
      version: number;
      flags: number;
    }

    it("should return typed result", () => {
      const data = Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x00, 0x00]);
      const format: [string, number][] = [
        ["signature", 4],
        ["version", 2],
        ["flags", 2]
      ];

      const result = parseFormattedTyped<TestFormat>(data, format);

      expect(result.signature).toBe(0x04034b50);
      expect(result.version).toBe(20);
      expect(result.flags).toBe(0);
    });
  });
});
