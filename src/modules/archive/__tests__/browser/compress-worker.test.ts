/**
 * Browser Compression with Worker Pool Tests
 *
 * Tests for the compression API with worker pool integration.
 * Verifies that useWorker option works correctly for both
 * compress() and decompress() functions.
 */

import { compress, decompress, hasWorkerSupport } from "@archive/compression/compress.browser";
import { terminateDefaultWorkerPool } from "@archive/compression/worker-pool/index.browser";
import { describe, it, expect, afterEach } from "vitest";

// Helper to create compressible data
function createCompressibleData(size: number): Uint8Array {
  const data = new Uint8Array(size);
  const pattern = new TextEncoder().encode("Hello World! This is a repeating pattern. ");
  for (let i = 0; i < size; i++) {
    data[i] = pattern[i % pattern.length];
  }
  return data;
}

describe("compress.browser with worker support", () => {
  afterEach(() => {
    terminateDefaultWorkerPool();
  });

  describe("compress with useWorker", () => {
    it("should compress data using worker pool", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original, { useWorker: true });

      expect(compressed).toBeInstanceOf(Uint8Array);
      expect(compressed.length).toBeGreaterThan(0);

      // Verify by decompressing
      const decompressed = await decompress(compressed);
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("should compress large data using worker pool", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(100 * 1024); // 100KB
      const compressed = await compress(original, { useWorker: true });

      // With highly compressible data, should achieve good compression
      // Note: for very small data, compression might add overhead
      expect(compressed.length).toBeLessThan(original.length * 0.5); // At least 50% compression

      // Verify round-trip
      const decompressed = await decompress(compressed);
      expect(decompressed).toEqual(original);
    });

    it("should support level option with worker", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(10 * 1024);

      // Different compression levels should produce same decompressed output
      const compressed6 = await compress(original, { useWorker: true, level: 6 });
      const decompressed = await decompress(compressed6);

      expect(decompressed).toEqual(original);
    });

    it("should support abort signal", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const controller = new AbortController();
      controller.abort();

      await expect(
        compress(createCompressibleData(1024), {
          useWorker: true,
          signal: controller.signal
        })
      ).rejects.toThrow(/aborted/i);
    });

    it("should fallback to main thread when useWorker is false", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original, { useWorker: false });

      expect(compressed).toBeInstanceOf(Uint8Array);

      const decompressed = await decompress(compressed);
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("should handle level 0 (no compression) even with useWorker", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const result = await compress(original, { useWorker: true, level: 0 });

      // Level 0 should return data as-is without going to worker
      expect(result).toEqual(original);
    });

    it("should support allowTransfer option for zero-copy", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(50 * 1024);
      // Copy the original data before transfer since the buffer will be detached
      const originalCopy = original.slice();
      const compressed = await compress(original, { useWorker: true, allowTransfer: true });

      expect(compressed).toBeInstanceOf(Uint8Array);
      expect(compressed.length).toBeGreaterThan(0);

      // Verify decompression using the copy
      const decompressed = await decompress(compressed);
      expect(decompressed).toEqual(originalCopy);
    });

    it("should respect autoWorkerThreshold option", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const smallData = createCompressibleData(500); // 500 bytes

      // With high threshold, should use main thread
      const result1 = await compress(smallData, { autoWorkerThreshold: 1024 });
      expect(result1).toBeInstanceOf(Uint8Array);

      // With low threshold, should use worker (but API is the same)
      const result2 = await compress(smallData, { autoWorkerThreshold: 100, useWorker: undefined });
      expect(result2).toBeInstanceOf(Uint8Array);
    });
  });

  describe("decompress with useWorker", () => {
    it("should decompress data using worker pool", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original);

      const decompressed = await decompress(compressed, { useWorker: true });
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("should decompress large data using worker pool", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(100 * 1024);
      const compressed = await compress(original);

      const decompressed = await decompress(compressed, { useWorker: true });
      expect(decompressed).toEqual(original);
    });

    it("should support abort signal for decompress", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await compress(original);

      const controller = new AbortController();
      controller.abort();

      await expect(
        decompress(compressed, { useWorker: true, signal: controller.signal })
      ).rejects.toThrow(/aborted/i);
    });
  });

  describe("round-trip combinations", () => {
    it("should work: compress(main) -> decompress(worker)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(10 * 1024);
      const compressed = await compress(original, { useWorker: false });
      const decompressed = await decompress(compressed, { useWorker: true });

      expect(decompressed).toEqual(original);
    });

    it("should work: compress(worker) -> decompress(main)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(10 * 1024);
      const compressed = await compress(original, { useWorker: true });
      const decompressed = await decompress(compressed, { useWorker: false });

      expect(decompressed).toEqual(original);
    });

    it("should work: compress(worker) -> decompress(worker)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(10 * 1024);
      const compressed = await compress(original, { useWorker: true });
      const decompressed = await decompress(compressed, { useWorker: true });

      expect(decompressed).toEqual(original);
    });
  });

  describe("concurrent operations", () => {
    it("should handle multiple concurrent compressions", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const tasks = Array.from({ length: 5 }, (_, i) =>
        compress(createCompressibleData(10 * 1024 * (i + 1)), { useWorker: true })
      );

      const results = await Promise.all(tasks);

      expect(results.length).toBe(5);
      results.forEach(result => {
        expect(result).toBeInstanceOf(Uint8Array);
        expect(result.length).toBeGreaterThan(0);
      });
    });

    it("should handle mixed compress and decompress operations", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original1 = createCompressibleData(5 * 1024);
      const original2 = createCompressibleData(10 * 1024);

      // Compress first
      const [compressed1, compressed2] = await Promise.all([
        compress(original1, { useWorker: true }),
        compress(original2, { useWorker: true })
      ]);

      // Decompress in parallel
      const [decompressed1, decompressed2] = await Promise.all([
        decompress(compressed1, { useWorker: true }),
        decompress(compressed2, { useWorker: true })
      ]);

      expect(decompressed1).toEqual(original1);
      expect(decompressed2).toEqual(original2);
    });
  });
});
