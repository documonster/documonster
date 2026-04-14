/**
 * Zlib compression tests (RFC 1950)
 */

import {
  zlib,
  unzlib,
  zlibSync,
  unzlibSync,
  isZlibData,
  detectCompressionFormat,
  decompressAuto,
  decompressAutoSync,
  compress,
  gzip,
  ZLIB_CM_DEFLATE,
  ZLIB_CINFO_MAX,
  ZLIB_MIN_SIZE
} from "@archive/compression/compress";
import {
  adler32,
  getZlibHeader,
  buildZlibTrailer,
  parseZlibHeader,
  readZlibTrailer,
  verifyAdler32
} from "@archive/compression/compress.base";
import { describe, it, expect } from "vitest";

describe("Zlib compression (RFC 1950)", () => {
  const testData = new TextEncoder().encode("Hello, World! This is a test for Zlib compression.");
  const emptyData = new Uint8Array(0);
  const singleByte = new Uint8Array([42]);

  describe("constants", () => {
    it("should have correct ZLIB_CM_DEFLATE value", () => {
      expect(ZLIB_CM_DEFLATE).toBe(8);
    });

    it("should have correct ZLIB_CINFO_MAX value", () => {
      expect(ZLIB_CINFO_MAX).toBe(7);
    });

    it("should have correct ZLIB_MIN_SIZE value", () => {
      // 2 bytes header + 4 bytes Adler-32 trailer
      expect(ZLIB_MIN_SIZE).toBe(6);
    });
  });

  describe("isZlibData", () => {
    it("should return false for data that is too short", () => {
      expect(isZlibData(new Uint8Array([]))).toBe(false);
      expect(isZlibData(new Uint8Array([0x78]))).toBe(false);
    });

    it("should return true for valid zlib header (level 1)", () => {
      // 0x78 0x01 = CM=8, CINFO=7, FLEVEL=0, FDICT=0
      // (0x78 << 8) + 0x01 = 0x7801 = 30721, 30721 % 31 = 0
      expect(isZlibData(new Uint8Array([0x78, 0x01]))).toBe(true);
    });

    it("should return true for valid zlib header (level 6 default)", () => {
      // 0x78 0x9c = CM=8, CINFO=7, FLEVEL=2, FDICT=0
      // (0x78 << 8) + 0x9c = 0x789c = 30876, 30876 % 31 = 0
      expect(isZlibData(new Uint8Array([0x78, 0x9c]))).toBe(true);
    });

    it("should return true for valid zlib header (level 9)", () => {
      // 0x78 0xda = CM=8, CINFO=7, FLEVEL=3, FDICT=0
      // (0x78 << 8) + 0xda = 0x78da = 30938, 30938 % 31 = 0
      expect(isZlibData(new Uint8Array([0x78, 0xda]))).toBe(true);
    });

    it("should return false for invalid CM (not DEFLATE)", () => {
      // CM=7 instead of 8
      expect(isZlibData(new Uint8Array([0x77, 0x00]))).toBe(false);
    });

    it("should return false for invalid CINFO (> 7)", () => {
      // CINFO=8 (invalid)
      expect(isZlibData(new Uint8Array([0x88, 0x00]))).toBe(false);
    });

    it("should return false for GZIP data", () => {
      expect(isZlibData(new Uint8Array([0x1f, 0x8b]))).toBe(false);
    });

    it("should return false for raw DEFLATE data", () => {
      // Raw DEFLATE typically doesn't satisfy the zlib checksum
      expect(isZlibData(new Uint8Array([0x00, 0x00]))).toBe(false);
    });
  });

  describe("zlibSync / unzlibSync", () => {
    it("should compress and decompress data correctly", () => {
      const compressed = zlibSync(testData);
      expect(compressed).not.toEqual(testData);
      expect(isZlibData(compressed)).toBe(true);

      const decompressed = unzlibSync(compressed);
      expect(decompressed).toEqual(testData);
    });

    it("should handle empty data", () => {
      const compressed = zlibSync(emptyData);
      expect(isZlibData(compressed)).toBe(true);

      const decompressed = unzlibSync(compressed);
      expect(decompressed).toEqual(emptyData);
    });

    it("should handle single byte", () => {
      const compressed = zlibSync(singleByte);
      expect(isZlibData(compressed)).toBe(true);

      const decompressed = unzlibSync(compressed);
      expect(decompressed).toEqual(singleByte);
    });

    it("should compress with different levels", () => {
      const level1 = zlibSync(testData, { level: 1 });
      const level9 = zlibSync(testData, { level: 9 });

      // Both should be valid zlib
      expect(isZlibData(level1)).toBe(true);
      expect(isZlibData(level9)).toBe(true);

      // Level 9 should generally produce smaller output (or equal for small data)
      expect(level9.length).toBeLessThanOrEqual(level1.length);

      // Both should decompress correctly
      expect(unzlibSync(level1)).toEqual(testData);
      expect(unzlibSync(level9)).toEqual(testData);
    });
  });

  describe("zlib / unzlib (async)", () => {
    it("should compress and decompress data correctly", async () => {
      const compressed = await zlib(testData);
      expect(compressed).not.toEqual(testData);
      expect(isZlibData(compressed)).toBe(true);

      const decompressed = await unzlib(compressed);
      expect(decompressed).toEqual(testData);
    });

    it("should handle large data", async () => {
      const largeData = new Uint8Array(100000);
      for (let i = 0; i < largeData.length; i++) {
        largeData[i] = i % 256;
      }

      const compressed = await zlib(largeData);
      expect(isZlibData(compressed)).toBe(true);
      expect(compressed.length).toBeLessThan(largeData.length);

      const decompressed = await unzlib(compressed);
      expect(decompressed).toEqual(largeData);
    });
  });

  describe("error handling", () => {
    it("should throw on invalid zlib data (too small)", () => {
      expect(() => unzlibSync(new Uint8Array([0x78, 0x9c, 0x00]))).toThrow();
    });

    it("should throw on corrupted zlib data", () => {
      const compressed = zlibSync(testData);
      // Corrupt the Adler-32 checksum
      compressed[compressed.length - 1] ^= 0xff;
      expect(() => unzlibSync(compressed)).toThrow();
    });
  });
});

describe("detectCompressionFormat", () => {
  const testData = new TextEncoder().encode("Test data for format detection");

  it("should detect GZIP format", async () => {
    const gzipped = await gzip(testData);
    expect(detectCompressionFormat(gzipped)).toBe("gzip");
  });

  it("should detect Zlib format", () => {
    const zlibbed = zlibSync(testData);
    expect(detectCompressionFormat(zlibbed)).toBe("zlib");
  });

  it("should detect raw DEFLATE format", async () => {
    const deflated = await compress(testData);
    expect(detectCompressionFormat(deflated)).toBe("deflate-raw");
  });

  it("should default to deflate-raw for unknown data", () => {
    expect(detectCompressionFormat(new Uint8Array([0x00, 0x00, 0x00]))).toBe("deflate-raw");
  });
});

describe("decompressAuto / decompressAutoSync", () => {
  const testData = new TextEncoder().encode("Auto-detect decompression test data");

  describe("sync", () => {
    it("should auto-detect and decompress GZIP", async () => {
      const gzipped = await gzip(testData);
      const decompressed = decompressAutoSync(gzipped);
      expect(decompressed).toEqual(testData);
    });

    it("should auto-detect and decompress Zlib", () => {
      const zlibbed = zlibSync(testData);
      const decompressed = decompressAutoSync(zlibbed);
      expect(decompressed).toEqual(testData);
    });

    it("should auto-detect and decompress raw DEFLATE", async () => {
      const deflated = await compress(testData);
      const decompressed = decompressAutoSync(deflated);
      expect(decompressed).toEqual(testData);
    });
  });

  describe("async", () => {
    it("should auto-detect and decompress GZIP", async () => {
      const gzipped = await gzip(testData);
      const decompressed = await decompressAuto(gzipped);
      expect(decompressed).toEqual(testData);
    });

    it("should auto-detect and decompress Zlib", async () => {
      const zlibbed = await zlib(testData);
      const decompressed = await decompressAuto(zlibbed);
      expect(decompressed).toEqual(testData);
    });

    it("should auto-detect and decompress raw DEFLATE", async () => {
      const deflated = await compress(testData);
      const decompressed = await decompressAuto(deflated);
      expect(decompressed).toEqual(testData);
    });
  });
});

describe("Adler-32", () => {
  // Test vectors from https://www.ietf.org/rfc/rfc1950.txt
  it("should compute correct Adler-32 for Wikipedia example", async () => {
    // "Wikipedia" = Adler-32: 0x11E60398
    const data = new TextEncoder().encode("Wikipedia");
    const compressed = zlibSync(data);
    // Just verify it round-trips correctly (Adler-32 is verified internally)
    expect(unzlibSync(compressed)).toEqual(data);
  });

  it("should handle empty data", () => {
    const compressed = zlibSync(new Uint8Array(0));
    // Empty data has Adler-32 = 1
    expect(unzlibSync(compressed)).toEqual(new Uint8Array(0));
  });

  it("should handle all zeros", () => {
    const data = new Uint8Array(1000);
    const compressed = zlibSync(data);
    expect(unzlibSync(compressed)).toEqual(data);
  });

  it("should handle all 0xFF", () => {
    const data = new Uint8Array(1000).fill(0xff);
    const compressed = zlibSync(data);
    expect(unzlibSync(compressed)).toEqual(data);
  });
});

// =============================================================================
// Zlib Helper Functions (compress.base.ts)
// =============================================================================

describe("adler32", () => {
  it("should return 1 for empty data", () => {
    expect(adler32(new Uint8Array(0))).toBe(1);
  });

  it("should compute correct checksum for 'Wikipedia'", () => {
    // Known test vector: adler32("Wikipedia") = 0x11E60398
    const data = new TextEncoder().encode("Wikipedia");
    expect(adler32(data) >>> 0).toBe(0x11e60398);
  });

  it("should compute correct checksum for 'a'", () => {
    // adler32("a") = (1 + 97) | ((0 + 1 + 97) << 16) = 98 | (98 << 16) = 0x00620062
    const data = new TextEncoder().encode("a");
    expect(adler32(data) >>> 0).toBe(0x00620062);
  });

  it("should handle large data without overflow", () => {
    // Test the chunking logic (chunk size is 5552)
    const data = new Uint8Array(10000).fill(0xff);
    const result = adler32(data);
    expect(typeof result).toBe("number");
    expect(result >>> 0).toBeGreaterThan(0);
  });
});

describe("getZlibHeader", () => {
  it("should return correct header for level 0", () => {
    expect(getZlibHeader(0)).toEqual(new Uint8Array([0x78, 0x01]));
  });

  it("should return correct header for level 1", () => {
    expect(getZlibHeader(1)).toEqual(new Uint8Array([0x78, 0x01]));
  });

  it("should return correct header for level 6 (default)", () => {
    expect(getZlibHeader(6)).toEqual(new Uint8Array([0x78, 0x9c]));
  });

  it("should return correct header for level 9", () => {
    expect(getZlibHeader(9)).toEqual(new Uint8Array([0x78, 0xda]));
  });

  it("should clamp out-of-range levels", () => {
    expect(getZlibHeader(-1)).toEqual(new Uint8Array([0x78, 0x01])); // clamped to 0
    expect(getZlibHeader(10)).toEqual(new Uint8Array([0x78, 0xda])); // clamped to 9
  });

  it("should produce valid zlib headers (FCHECK)", () => {
    for (let level = 0; level <= 9; level++) {
      const header = getZlibHeader(level);
      const check = (header[0] << 8) | header[1];
      expect(check % 31).toBe(0);
    }
  });
});

describe("buildZlibTrailer", () => {
  it("should build big-endian Adler-32 trailer", () => {
    const trailer = buildZlibTrailer(0x11e60398);
    expect(trailer).toEqual(new Uint8Array([0x11, 0xe6, 0x03, 0x98]));
  });

  it("should handle zero checksum", () => {
    const trailer = buildZlibTrailer(0);
    expect(trailer).toEqual(new Uint8Array([0, 0, 0, 0]));
  });

  it("should handle max checksum", () => {
    const trailer = buildZlibTrailer(0xffffffff);
    expect(trailer).toEqual(new Uint8Array([0xff, 0xff, 0xff, 0xff]));
  });
});

describe("parseZlibHeader", () => {
  it("should return offset 2 for valid header", () => {
    const data = new Uint8Array([0x78, 0x9c, 0x00, 0x00, 0x00, 0x00]);
    expect(parseZlibHeader(data)).toBe(2);
  });

  it("should throw for data too small", () => {
    expect(() => parseZlibHeader(new Uint8Array([0x78]))).toThrow("too small");
  });

  it("should throw for invalid compression method", () => {
    const data = new Uint8Array([0x77, 0x00, 0x00, 0x00, 0x00, 0x00]); // CM=7
    expect(() => parseZlibHeader(data)).toThrow("compression method");
  });

  it("should throw for invalid CINFO", () => {
    const data = new Uint8Array([0x88, 0x00, 0x00, 0x00, 0x00, 0x00]); // CINFO=8
    expect(() => parseZlibHeader(data)).toThrow("CINFO");
  });

  it("should throw for invalid FCHECK", () => {
    const data = new Uint8Array([0x78, 0x00, 0x00, 0x00, 0x00, 0x00]); // invalid checksum
    expect(() => parseZlibHeader(data)).toThrow("checksum");
  });

  it("should throw for preset dictionary", () => {
    // FDICT bit set (0x20)
    const data = new Uint8Array([0x78, 0xbb, 0x00, 0x00, 0x00, 0x00]); // 0xbb has FDICT
    expect(() => parseZlibHeader(data)).toThrow("dictionary");
  });
});

describe("readZlibTrailer", () => {
  it("should read big-endian Adler-32 from end of data", () => {
    const data = new Uint8Array([0x00, 0x00, 0x11, 0xe6, 0x03, 0x98]);
    expect(readZlibTrailer(data) >>> 0).toBe(0x11e60398);
  });

  it("should handle data with header + payload + trailer", () => {
    const data = new Uint8Array([0x78, 0x9c, 0x01, 0x02, 0x03, 0xaa, 0xbb, 0xcc, 0xdd]);
    expect(readZlibTrailer(data) >>> 0).toBe(0xaabbccdd);
  });

  it("should return unsigned value for high-bit data", () => {
    // When first byte >= 0x80, ensure unsigned result
    const data = new Uint8Array([0x00, 0x00, 0xff, 0xff, 0xff, 0xff]);
    expect(readZlibTrailer(data)).toBe(0xffffffff);
    expect(readZlibTrailer(data)).toBeGreaterThan(0);
  });
});

describe("verifyAdler32", () => {
  it("should not throw for matching checksum", () => {
    const data = new TextEncoder().encode("Wikipedia");
    expect(() => verifyAdler32(data, 0x11e60398)).not.toThrow();
  });

  it("should throw for mismatched checksum", () => {
    const data = new TextEncoder().encode("Wikipedia");
    expect(() => verifyAdler32(data, 0x12345678)).toThrow("Adler-32 mismatch");
  });

  it("should handle unsigned comparison correctly", () => {
    const data = new Uint8Array(0);
    // adler32(empty) = 1, verify with signed vs unsigned
    expect(() => verifyAdler32(data, 1)).not.toThrow();
    expect(() => verifyAdler32(data, -1)).toThrow();
  });
});
