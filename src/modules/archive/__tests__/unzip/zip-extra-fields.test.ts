/**
 * Tests for ZIP extra field parsing: ZIP64, extended timestamp,
 * Unicode Path (0x7075), Unicode Comment (0x6375), CP437 decoding,
 * and decodeZipPath logic.
 */

import { crc32 } from "@archive/compression/crc32";
import { decodeCp437, decodeZipPath } from "@archive/shared/text";
import { parseExtraField, type ZipVars } from "@archive/unzip/parser-core";
import { FLAG_UTF8 } from "@archive/zip-spec/zip-records";
import { describe, it, expect } from "vitest";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function createUnicodePathExtraField(
  version: number,
  originalName: Uint8Array,
  unicodeName: string
): Uint8Array {
  const textEncoder = new TextEncoder();
  const unicodeBytes = textEncoder.encode(unicodeName);
  const crc = crc32(originalName);

  const extraField = new Uint8Array(4 + 1 + 4 + unicodeBytes.length);
  const view = new DataView(extraField.buffer);

  view.setUint16(0, 0x7075, true);
  view.setUint16(2, 1 + 4 + unicodeBytes.length, true);
  extraField[4] = version;
  view.setUint32(5, crc, true);
  extraField.set(unicodeBytes, 9);

  return extraField;
}

function createUnicodeCommentExtraField(
  version: number,
  originalComment: Uint8Array,
  unicodeComment: string
): Uint8Array {
  const textEncoder = new TextEncoder();
  const unicodeBytes = textEncoder.encode(unicodeComment);
  const crc = crc32(originalComment);

  const extraField = new Uint8Array(4 + 1 + 4 + unicodeBytes.length);
  const view = new DataView(extraField.buffer);

  view.setUint16(0, 0x6375, true);
  view.setUint16(2, 1 + 4 + unicodeBytes.length, true);
  extraField[4] = version;
  view.setUint32(5, crc, true);
  extraField.set(unicodeBytes, 9);

  return extraField;
}

// =============================================================================
// ZIP64 extra field
// =============================================================================

describe("parseExtraField: ZIP64", () => {
  it("should return empty object for empty extra field", () => {
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(Buffer.alloc(0), vars);
    expect(result).toEqual({});
  });

  it("should parse ZIP64 uncompressed size", () => {
    const extraField = Buffer.alloc(12);
    extraField.writeUInt16LE(0x0001, 0);
    extraField.writeUInt16LE(8, 2);
    extraField.writeBigUInt64LE(BigInt(0x100000000), 4);

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 0xffffffff };
    const result = parseExtraField(extraField, vars);
    expect(result.uncompressedSize).toBe(0x100000000);
    expect(result.uncompressedSize64).toBe(BigInt(0x100000000));
    expect(vars.uncompressedSize).toBe(0x100000000);
  });

  it("should parse ZIP64 compressed size", () => {
    const extraField = Buffer.alloc(12);
    extraField.writeUInt16LE(0x0001, 0);
    extraField.writeUInt16LE(8, 2);
    extraField.writeBigUInt64LE(BigInt(0x200000000), 4);

    const vars: ZipVars = { compressedSize: 0xffffffff, uncompressedSize: 100 };
    const result = parseExtraField(extraField, vars);
    expect(result.compressedSize).toBe(0x200000000);
    expect(result.compressedSize64).toBe(BigInt(0x200000000));
  });

  it("should parse ZIP64 with both sizes", () => {
    const extraField = Buffer.alloc(20);
    extraField.writeUInt16LE(0x0001, 0);
    extraField.writeUInt16LE(16, 2);
    extraField.writeBigUInt64LE(BigInt(0x100000000), 4);
    extraField.writeBigUInt64LE(BigInt(0x200000000), 12);

    const vars: ZipVars = { compressedSize: 0xffffffff, uncompressedSize: 0xffffffff };
    const result = parseExtraField(extraField, vars);
    expect(result.uncompressedSize).toBe(0x100000000);
    expect(result.compressedSize).toBe(0x200000000);
  });

  it("should skip non-ZIP64 headers", () => {
    const extraField = Buffer.alloc(20);
    extraField.writeUInt16LE(0x0007, 0);
    extraField.writeUInt16LE(4, 2);
    extraField.writeUInt32LE(0x12345678, 4);
    extraField.writeUInt16LE(0x0001, 8);
    extraField.writeUInt16LE(8, 10);
    extraField.writeBigUInt64LE(BigInt(0x300000000), 12);

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 0xffffffff };
    const result = parseExtraField(extraField, vars);
    expect(result.uncompressedSize).toBe(0x300000000);
  });

  it("should handle offset to local file header", () => {
    const extraField = Buffer.alloc(28);
    extraField.writeUInt16LE(0x0001, 0);
    extraField.writeUInt16LE(24, 2);
    extraField.writeBigUInt64LE(BigInt(0x100000000), 4);
    extraField.writeBigUInt64LE(BigInt(0x200000000), 12);
    extraField.writeBigUInt64LE(BigInt(0x300000000), 20);

    const vars: ZipVars = {
      compressedSize: 0xffffffff,
      uncompressedSize: 0xffffffff,
      offsetToLocalFileHeader: 0xffffffff
    };
    const result = parseExtraField(extraField, vars);
    expect(result.offsetToLocalFileHeader).toBe(0x300000000);
    expect(result.offsetToLocalFileHeader64).toBe(BigInt(0x300000000));
  });

  it("should expose BigInt values beyond JS safe integers", () => {
    const tooLarge = BigInt(Number.MAX_SAFE_INTEGER) + 1n;
    const extraField = Buffer.alloc(12);
    extraField.writeUInt16LE(0x0001, 0);
    extraField.writeUInt16LE(8, 2);
    extraField.writeBigUInt64LE(tooLarge, 4);

    const vars: ZipVars = { compressedSize: 123, uncompressedSize: 0xffffffff };
    const result = parseExtraField(extraField, vars);
    expect(result.uncompressedSize).toBeUndefined();
    expect(result.uncompressedSize64).toBe(tooLarge);
    expect(vars.uncompressedSize64).toBe(tooLarge);
  });
});

// =============================================================================
// Extended timestamp (0x5455)
// =============================================================================

describe("parseExtraField: extended timestamp", () => {
  it("should parse mtime from 0x5455 field", () => {
    const extraField = Buffer.alloc(9);
    extraField.writeUInt16LE(0x5455, 0);
    extraField.writeUInt16LE(5, 2);
    extraField.writeUInt8(0x01, 4);
    extraField.writeUInt32LE(123456789, 5);

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.mtimeUnixSeconds).toBe(123456789);
    expect(vars.compressedSize).toBe(100);
    expect(vars.uncompressedSize).toBe(200);
  });
});

// =============================================================================
// Unicode Path Extra Field (0x7075)
// =============================================================================

describe("parseExtraField: Unicode Path (0x7075)", () => {
  it("should parse valid Unicode Path extra field", () => {
    const originalName = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x65, 0x2e, 0x74, 0x78, 0x74]);
    const extraField = createUnicodePathExtraField(1, originalName, "Grüße.txt");

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodePath).toBeDefined();
    expect(result.unicodePath!.version).toBe(1);
    expect(result.unicodePath!.originalCrc32).toBe(crc32(originalName));
    expect(result.unicodePath!.unicodeValue).toBe("Grüße.txt");
  });

  it("should parse Unicode Path with Chinese characters", () => {
    const originalName = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodePathExtraField(1, originalName, "测试文件.txt");

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodePath!.unicodeValue).toBe("测试文件.txt");
  });

  it("should parse Unicode Path with Japanese characters", () => {
    const originalName = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodePathExtraField(1, originalName, "ファイル名.txt");

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodePath!.unicodeValue).toBe("ファイル名.txt");
  });

  it("should parse Unicode Path with emoji", () => {
    const originalName = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodePathExtraField(
      1,
      originalName,
      "\u{1F4C1}folder/\u{1F4C4}file.txt"
    );

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodePath!.unicodeValue).toBe("\u{1F4C1}folder/\u{1F4C4}file.txt");
  });

  it("should not reject unsupported version (parsed but not validated here)", () => {
    const originalName = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodePathExtraField(2, originalName, "test.txt");

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodePath).toBeDefined();
    expect(result.unicodePath!.version).toBe(2);
  });

  it("should handle Unicode Path alongside ZIP64 extra field", () => {
    const originalName = new Uint8Array([0x74, 0x65, 0x73, 0x74]);

    const zip64Field = new Uint8Array(12);
    const zip64View = new DataView(zip64Field.buffer);
    zip64View.setUint16(0, 0x0001, true);
    zip64View.setUint16(2, 8, true);
    zip64View.setBigUint64(4, BigInt(0x100000000), true);

    const unicodeField = createUnicodePathExtraField(1, originalName, "\u6D4B\u8BD5.txt");

    const combined = new Uint8Array(zip64Field.length + unicodeField.length);
    combined.set(zip64Field, 0);
    combined.set(unicodeField, zip64Field.length);

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 0xffffffff };
    const result = parseExtraField(combined, vars);
    expect(result.uncompressedSize).toBe(0x100000000);
    expect(result.unicodePath!.unicodeValue).toBe("\u6D4B\u8BD5.txt");
  });

  it("should skip malformed Unicode Path field (too short)", () => {
    const malformed = new Uint8Array(8);
    const view = new DataView(malformed.buffer);
    view.setUint16(0, 0x7075, true);
    view.setUint16(2, 3, true);
    malformed[4] = 1;

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(malformed, vars);
    expect(result.unicodePath).toBeUndefined();
  });
});

// =============================================================================
// Unicode Comment Extra Field (0x6375)
// =============================================================================

describe("parseExtraField: Unicode Comment (0x6375)", () => {
  it("should parse valid Unicode Comment extra field", () => {
    const originalComment = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodeCommentExtraField(
      1,
      originalComment,
      "\u8FD9\u662F\u6CE8\u91CA"
    );

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(extraField, vars);
    expect(result.unicodeComment).toBeDefined();
    expect(result.unicodeComment!.version).toBe(1);
    expect(result.unicodeComment!.originalCrc32).toBe(crc32(originalComment));
    expect(result.unicodeComment!.unicodeValue).toBe("\u8FD9\u662F\u6CE8\u91CA");
  });

  it("should parse both Unicode Path and Comment", () => {
    const originalName = new Uint8Array([0x6e, 0x61, 0x6d, 0x65]);
    const originalComment = new Uint8Array([0x63, 0x6f, 0x6d, 0x6d]);

    const pathField = createUnicodePathExtraField(1, originalName, "\u540D\u524D.txt");
    const commentField = createUnicodeCommentExtraField(
      1,
      originalComment,
      "\u30B3\u30E1\u30F3\u30C8"
    );

    const combined = new Uint8Array(pathField.length + commentField.length);
    combined.set(pathField, 0);
    combined.set(commentField, pathField.length);

    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const result = parseExtraField(combined, vars);
    expect(result.unicodePath!.unicodeValue).toBe("\u540D\u524D.txt");
    expect(result.unicodeComment!.unicodeValue).toBe("\u30B3\u30E1\u30F3\u30C8");
  });
});

// =============================================================================
// CP437 decoding
// =============================================================================

describe("decodeCp437", () => {
  it("should decode pure ASCII correctly", () => {
    expect(decodeCp437(new Uint8Array([72, 101, 108, 108, 111]))).toBe("Hello");
  });

  it("should decode empty buffer", () => {
    expect(decodeCp437(new Uint8Array(0))).toBe("");
  });

  it("should decode typical CP437 characters", () => {
    expect(decodeCp437(new Uint8Array([0x81, 0x82, 0x84, 0x94]))).toBe("\u00FC\u00E9\u00E4\u00F6");
  });

  it("should decode box-drawing characters", () => {
    expect(decodeCp437(new Uint8Array([0xda, 0xc4, 0xbf, 0xb3]))).toBe("\u250C\u2500\u2510\u2502");
  });

  it("should decode Greek letters", () => {
    expect(decodeCp437(new Uint8Array([0xe0, 0xe1, 0xe2, 0xe3]))).toBe("\u03B1\u00DF\u0393\u03C0");
  });

  it("should decode math symbols", () => {
    expect(decodeCp437(new Uint8Array([0xf1, 0xf6, 0xfb]))).toBe("\u00B1\u00F7\u221A");
  });

  it("should decode mixed ASCII and CP437", () => {
    expect(decodeCp437(new Uint8Array([0x63, 0x61, 0x66, 0x82]))).toBe("caf\u00E9");
  });

  it("should decode real-world German filename", () => {
    const germanName = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x65, 0x2e, 0x74, 0x78, 0x74]);
    expect(decodeCp437(germanName)).toBe("Gr\u00FC\u00DFe.txt");
  });

  it("should decode real-world French filename", () => {
    const frenchName = new Uint8Array([0x72, 0x82, 0x73, 0x75, 0x6d, 0x82, 0x2e, 0x64, 0x6f, 0x63]);
    expect(decodeCp437(frenchName)).toBe("r\u00E9sum\u00E9.doc");
  });

  it("should handle null bytes", () => {
    expect(decodeCp437(new Uint8Array([0x48, 0x00, 0x69]))).toBe("H\x00i");
  });

  it("should decode boundary characters", () => {
    expect(decodeCp437(new Uint8Array([0x80, 0xff]))).toBe("\u00C7\u00A0");
  });

  it("should decode path with directory separator", () => {
    const path = new Uint8Array([
      0x64, 0x6f, 0x6e, 0x6e, 0x82, 0x65, 0x73, 0x2f, 0x66, 0x69, 0x63, 0x68, 0x69, 0x65, 0x72,
      0x2e, 0x74, 0x78, 0x74
    ]);
    expect(decodeCp437(path)).toBe("donn\u00E9es/fichier.txt");
  });
});

// =============================================================================
// CRC32 validation for Unicode extra fields
// =============================================================================

describe("CRC32 for Unicode extra fields", () => {
  it("should compute consistent CRC32 for original filename", () => {
    const originalName = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x65]);
    const calculatedCrc = crc32(originalName);
    expect(typeof calculatedCrc).toBe("number");
    expect(calculatedCrc).not.toBe(0);

    const differentName = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x66]);
    expect(crc32(differentName)).not.toBe(calculatedCrc);
  });
});

// =============================================================================
// decodeZipPath
// =============================================================================

describe("decodeZipPath", () => {
  it("should use UTF-8 when UTF-8 flag is set", () => {
    const pathBuffer = new Uint8Array([0x63, 0x61, 0x66, 0xc3, 0xa9]);
    expect(decodeZipPath(pathBuffer, FLAG_UTF8, undefined)).toBe("caf\u00E9");
  });

  it("should use CP437 when no UTF-8 flag and no Unicode extra field", () => {
    const pathBuffer = new Uint8Array([0x63, 0x61, 0x66, 0x82]);
    expect(decodeZipPath(pathBuffer, 0, undefined)).toBe("caf\u00E9");
  });

  it("should prefer Unicode Path extra field when CRC32 matches", () => {
    const pathBuffer = new Uint8Array([0x74, 0x65, 0x73, 0x74]);
    const extraField = createUnicodePathExtraField(1, pathBuffer, "\u6D4B\u8BD5.txt");
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe("\u6D4B\u8BD5.txt");
  });

  it("should fall back to CP437 when Unicode Path CRC32 does not match", () => {
    const pathBuffer = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x65]);
    const differentBytes = new Uint8Array([0x6f, 0x74, 0x68, 0x65, 0x72]);
    const extraField = createUnicodePathExtraField(1, differentBytes, "unicode.txt");
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe("Gr\u00FC\u00DFe");
  });

  it("should fall back to CP437 when Unicode Path version is not 1", () => {
    const pathBuffer = new Uint8Array([0x47, 0x72, 0x81, 0xe1, 0x65]);
    const extraField = createUnicodePathExtraField(2, pathBuffer, "unicode.txt");
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe("Gr\u00FC\u00DFe");
  });

  it("should handle null flags gracefully", () => {
    const pathBuffer = new Uint8Array([0x63, 0x61, 0x66, 0x82]);
    expect(decodeZipPath(pathBuffer, null, undefined)).toBe("caf\u00E9");
  });

  it("should decode Chinese characters from Unicode Path extra field", () => {
    const pathBuffer = new Uint8Array([0x66, 0x69, 0x6c, 0x65]);
    const extraField = createUnicodePathExtraField(1, pathBuffer, "\u6587\u4EF6/\u6570\u636E.txt");
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe("\u6587\u4EF6/\u6570\u636E.txt");
  });

  it("should decode Japanese characters from Unicode Path extra field", () => {
    const pathBuffer = new Uint8Array([0x66, 0x69, 0x6c, 0x65]);
    const extraField = createUnicodePathExtraField(
      1,
      pathBuffer,
      "\u30D5\u30A9\u30EB\u30C0/\u30D5\u30A1\u30A4\u30EB.txt"
    );
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe(
      "\u30D5\u30A9\u30EB\u30C0/\u30D5\u30A1\u30A4\u30EB.txt"
    );
  });

  it("should decode emoji from Unicode Path extra field", () => {
    const pathBuffer = new Uint8Array([0x66, 0x69, 0x6c, 0x65]);
    const extraField = createUnicodePathExtraField(
      1,
      pathBuffer,
      "\u{1F4C1}folder/\u{1F4C4}document.txt"
    );
    const vars: ZipVars = { compressedSize: 100, uncompressedSize: 200 };
    const extra = parseExtraField(extraField, vars);

    expect(decodeZipPath(pathBuffer, 0, extra)).toBe("\u{1F4C1}folder/\u{1F4C4}document.txt");
  });
});
