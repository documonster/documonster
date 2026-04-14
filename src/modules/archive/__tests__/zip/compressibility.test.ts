/**
 * Archive Compressibility Detection Unit Tests
 *
 * Tests for entropy-based incompressibility heuristics:
 * - isProbablyIncompressible: Detect high-entropy data
 * - isProbablyIncompressibleChunks: Chunk-based variant for streaming
 */

import {
  isProbablyIncompressible,
  isProbablyIncompressibleChunks
} from "@archive/zip/compressibility";
import { describe, it, expect } from "vitest";

// =============================================================================
// Test Data Generators
// =============================================================================

/**
 * Generate repeated pattern (low entropy)
 */
function generateRepeatedPattern(length: number, pattern: number[] = [0, 1, 2, 3]): Uint8Array {
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    bytes[i] = pattern[i % pattern.length]!;
  }
  return bytes;
}

/**
 * Generate text-like data (moderate entropy)
 */
function generateTextData(length: number): Uint8Array {
  const text = "The quick brown fox jumps over the lazy dog. ";
  const bytes = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    bytes[i] = text.charCodeAt(i % text.length);
  }
  return bytes;
}

/**
 * Generate all zeros (minimum entropy)
 */
function generateZeros(length: number): Uint8Array {
  return new Uint8Array(length);
}

// =============================================================================
// isProbablyIncompressible Tests
// =============================================================================

describe("isProbablyIncompressible", () => {
  describe("returns false for compressible data", () => {
    it("all zeros (minimum entropy)", () => {
      const data = generateZeros(64 * 1024);
      expect(isProbablyIncompressible(data)).toBe(false);
    });

    it("repeated pattern", () => {
      const data = generateRepeatedPattern(64 * 1024);
      expect(isProbablyIncompressible(data)).toBe(false);
    });

    it("text-like data", () => {
      const data = generateTextData(64 * 1024);
      expect(isProbablyIncompressible(data)).toBe(false);
    });

    it("ASCII range only", () => {
      const data = new Uint8Array(64 * 1024);
      for (let i = 0; i < data.length; i++) {
        data[i] = 32 + (i % 95); // Printable ASCII
      }
      expect(isProbablyIncompressible(data)).toBe(false);
    });
  });

  describe("returns true for incompressible data", () => {
    it("random bytes (high entropy)", () => {
      // Generate truly random data
      const crypto = globalThis.crypto;
      const data = new Uint8Array(64 * 1024);
      crypto.getRandomValues(data);

      expect(isProbablyIncompressible(data)).toBe(true);
    });

    it("all byte values equally distributed", () => {
      // Create data with maximum entropy - each byte value appears equally
      const data = new Uint8Array(256 * 256); // 65536 bytes
      for (let i = 0; i < data.length; i++) {
        data[i] = i % 256;
      }
      // Shuffle to break patterns
      for (let i = data.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [data[i], data[j]] = [data[j]!, data[i]!];
      }

      expect(isProbablyIncompressible(data)).toBe(true);
    });
  });

  describe("minimum decision bytes", () => {
    it("returns false for small data below threshold", () => {
      const crypto = globalThis.crypto;
      const smallData = new Uint8Array(1024); // < 16KB default minimum
      crypto.getRandomValues(smallData);

      expect(isProbablyIncompressible(smallData)).toBe(false);
    });

    it("respects custom minDecisionBytes", () => {
      const crypto = globalThis.crypto;
      // Need enough data to have 200+ unique bytes for the algorithm to proceed
      const data = new Uint8Array(8 * 1024);
      crypto.getRandomValues(data);

      // Default minDecisionBytes is 16KB, so 8KB should return false
      expect(isProbablyIncompressible(data)).toBe(false);

      // With lower threshold (8KB), should detect high entropy
      expect(isProbablyIncompressible(data, { minDecisionBytes: 8 * 1024 })).toBe(true);
    });
  });

  describe("sample bytes option", () => {
    it("only samples specified amount", () => {
      // First 32KB is random, rest is zeros
      const data = new Uint8Array(128 * 1024);
      const crypto = globalThis.crypto;
      crypto.getRandomValues(data.subarray(0, 32 * 1024));
      // Rest is zeros (default)

      // Sample only first 32KB (high entropy)
      expect(isProbablyIncompressible(data, { sampleBytes: 32 * 1024 })).toBe(true);

      // Sample all 128KB (mixed, but mostly zeros)
      expect(isProbablyIncompressible(data, { sampleBytes: 128 * 1024 })).toBe(false);
    });
  });

  describe("unique byte threshold", () => {
    it("returns false when fewer than 200 unique bytes", () => {
      // Data with only 100 unique byte values
      const data = new Uint8Array(64 * 1024);
      for (let i = 0; i < data.length; i++) {
        data[i] = i % 100;
      }
      expect(isProbablyIncompressible(data)).toBe(false);
    });
  });
});

// =============================================================================
// isProbablyIncompressibleChunks Tests
// =============================================================================

describe("isProbablyIncompressibleChunks", () => {
  describe("equivalent to isProbablyIncompressible for same data", () => {
    it("returns same result for compressible chunks", () => {
      const fullData = generateTextData(64 * 1024);

      // Split into chunks
      const chunks: Uint8Array[] = [];
      const chunkSize = 8 * 1024;
      for (let i = 0; i < fullData.length; i += chunkSize) {
        chunks.push(fullData.subarray(i, i + chunkSize));
      }

      expect(isProbablyIncompressibleChunks(chunks)).toBe(isProbablyIncompressible(fullData));
    });

    it("returns same result for incompressible chunks", () => {
      const fullData = new Uint8Array(64 * 1024);
      globalThis.crypto.getRandomValues(fullData);

      // Split into chunks
      const chunks: Uint8Array[] = [];
      const chunkSize = 8 * 1024;
      for (let i = 0; i < fullData.length; i += chunkSize) {
        chunks.push(fullData.subarray(i, i + chunkSize));
      }

      expect(isProbablyIncompressibleChunks(chunks)).toBe(isProbablyIncompressible(fullData));
    });
  });

  describe("handles various chunk configurations", () => {
    it("single large chunk", () => {
      const data = generateZeros(64 * 1024);
      expect(isProbablyIncompressibleChunks([data])).toBe(false);
    });

    it("many small chunks", () => {
      const chunks: Uint8Array[] = [];
      for (let i = 0; i < 64; i++) {
        chunks.push(generateZeros(1024));
      }
      expect(isProbablyIncompressibleChunks(chunks)).toBe(false);
    });

    it("empty chunks array", () => {
      expect(isProbablyIncompressibleChunks([])).toBe(false);
    });
  });

  describe("respects options", () => {
    it("respects minDecisionBytes", () => {
      // Need enough data to have 200+ unique bytes for the algorithm to proceed
      const randomChunk = new Uint8Array(8 * 1024);
      globalThis.crypto.getRandomValues(randomChunk);

      // Default minDecisionBytes is 16KB, so 8KB should return false
      expect(isProbablyIncompressibleChunks([randomChunk], { minDecisionBytes: 16 * 1024 })).toBe(
        false
      );
      // With lower threshold matching data size, should detect high entropy
      expect(isProbablyIncompressibleChunks([randomChunk], { minDecisionBytes: 8 * 1024 })).toBe(
        true
      );
    });

    it("respects sampleBytes", () => {
      // First chunk is random, second is zeros
      const randomChunk = new Uint8Array(32 * 1024);
      globalThis.crypto.getRandomValues(randomChunk);
      const zeroChunk = generateZeros(32 * 1024);

      // Sample only first chunk
      expect(
        isProbablyIncompressibleChunks([randomChunk, zeroChunk], { sampleBytes: 32 * 1024 })
      ).toBe(true);
    });
  });

  describe("iterable support", () => {
    it("works with generator function", () => {
      function* generateChunks(): Generator<Uint8Array> {
        for (let i = 0; i < 8; i++) {
          yield generateZeros(8 * 1024);
        }
      }

      expect(isProbablyIncompressibleChunks(generateChunks())).toBe(false);
    });
  });
});
