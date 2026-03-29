import { describe, it, expect } from "vitest";
import { parseFormatted as parse } from "@archive/zip-spec/binary";

const buf = Buffer.from([
  0x62, 0x75, 0x66, 0x68, 0x65, 0x72, 0xff, 0xae, 0x00, 0x11, 0x99, 0xd7, 0x7b, 0x13, 0x35
]);

describe("parse-buffer", () => {
  describe("parse", () => {
    it("should parse little endian values for increasing byte size", () => {
      const result = parse(buf, [
        ["key1", 1],
        ["key2", 2],
        ["key3", 4],
        ["key4", 8]
      ]);
      expect(result).toEqual({
        key1: 98,
        key2: 26229,
        key3: 4285687144,
        key4: 3824536674483896300
      });
    });

    it("should parse little endian values for decreasing byte size", () => {
      const result = parse(buf, [
        ["key1", 8],
        ["key2", 4],
        ["key3", 2],
        ["key4", 1]
      ]);
      expect(result).toEqual({
        key1: 12609923261529487000,
        key2: 3617132800,
        key3: 4987,
        key4: 53
      });
    });

    it("should parse little endian values with null keys due to small buffer", () => {
      const result = parse(buf, [
        ["key1", 8],
        ["key2", 8],
        ["key3", 8],
        ["key4", 8]
      ]);
      expect(result).toEqual({
        key1: 12609923261529487000,
        key2: null,
        key3: null,
        key4: null
      });
    });

    it("should parse 1-byte unsigned integers", () => {
      const buffer = Buffer.from([0x12, 0x34]);
      const result = parse(buffer, [
        ["a", 1],
        ["b", 1]
      ]);
      expect(result).toEqual({ a: 0x12, b: 0x34 });
    });

    it("should parse 2-byte unsigned little endian integers", () => {
      const buffer = Buffer.from([0x34, 0x12, 0x78, 0x56]);
      const result = parse(buffer, [
        ["a", 2],
        ["b", 2]
      ]);
      expect(result).toEqual({ a: 0x1234, b: 0x5678 });
    });

    it("should parse 4-byte unsigned little endian integers", () => {
      const buffer = Buffer.from([0x78, 0x56, 0x34, 0x12]);
      const result = parse(buffer, [["value", 4]]);
      expect(result).toEqual({ value: 0x12345678 });
    });

    it("should parse 8-byte unsigned little endian integers", () => {
      const buffer = Buffer.alloc(8);
      buffer.writeBigUInt64LE(BigInt("0x123456789ABCDEF0"), 0);
      const result = parse(buffer, [["value", 8]]);
      expect(result).toEqual({ value: Number(BigInt("0x123456789ABCDEF0")) });
    });

    it("should return null for incomplete buffer", () => {
      const buffer = Buffer.from([0x12, 0x34]);
      const result = parse(buffer, [
        ["a", 2],
        ["b", 4]
      ]);
      expect(result).toEqual({ a: 0x3412, b: null });
    });

    it("should handle mixed sizes", () => {
      const buffer = Buffer.from([0x01, 0x34, 0x12, 0x78, 0x56, 0x34, 0x12]);
      const result = parse(buffer, [
        ["byte", 1],
        ["short", 2],
        ["int", 4]
      ]);
      expect(result).toEqual({ byte: 0x01, short: 0x1234, int: 0x12345678 });
    });

    it("should handle empty format", () => {
      const buffer = Buffer.from([0x12, 0x34]);
      const result = parse(buffer, []);
      expect(result).toEqual({});
    });

    it("should throw error for unsupported size", () => {
      const buffer = Buffer.from([0x12, 0x34, 0x56]);
      expect(() => parse(buffer, [["value", 3]])).toThrow("Unsupported UInt LE size!");
    });
  });
});
