/**
 * Browser Streaming Compression with Worker Pool Tests
 *
 * Tests for the streaming compression API with worker pool integration.
 * Verifies that useWorker option works correctly for both
 * createDeflateStream() and createInflateStream() functions.
 */

import {
  createDeflateStream,
  createInflateStream,
  hasWorkerSupport
} from "@archive/compression/streaming-compress.browser";
import {
  WorkerPool,
  terminateDefaultWorkerPool
} from "@archive/compression/worker-pool/index.browser";
import { concatUint8Arrays } from "@utils/binary";
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

// Helper to collect stream output
function collectStreamOutput(stream: ReturnType<typeof createDeflateStream>): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];

    stream.on("data", (chunk: Uint8Array) => {
      chunks.push(chunk);
    });

    stream.on("end", () => {
      resolve(concatUint8Arrays(chunks));
    });

    stream.on("error", (err: Error) => {
      reject(err);
    });
  });
}

describe("streaming-compress.browser with worker support", () => {
  afterEach(() => {
    terminateDefaultWorkerPool();
  });

  describe("createDeflateStream with useWorker", () => {
    it("should compress data using worker-based stream", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const deflate = createDeflateStream({ useWorker: true });

      const outputPromise = collectStreamOutput(deflate);

      deflate.write(original);
      deflate.end();

      const compressed = await outputPromise;

      expect(compressed).toBeInstanceOf(Uint8Array);
      expect(compressed.length).toBeGreaterThan(0);
    });

    it("should compress large data using worker-based stream", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(100 * 1024); // 100KB
      const deflate = createDeflateStream({ useWorker: true, level: 6 });

      const outputPromise = collectStreamOutput(deflate);

      // Write in chunks
      const chunkSize = 16 * 1024;
      for (let i = 0; i < original.length; i += chunkSize) {
        deflate.write(original.slice(i, Math.min(i + chunkSize, original.length)));
      }
      deflate.end();

      const compressed = await outputPromise;

      // Should achieve compression
      expect(compressed.length).toBeLessThan(original.length);

      // Verify by decompressing
      const inflate = createInflateStream();
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(decompressed).toEqual(original);
    });

    it("should support write callback", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const deflate = createDeflateStream({ useWorker: true });
      const outputPromise = collectStreamOutput(deflate);

      await new Promise<void>((resolve, reject) => {
        deflate.write(new TextEncoder().encode("Hello"), err => {
          if (err) {
            reject(err);
          } else {
            resolve();
          }
        });
      });

      deflate.end();
      const compressed = await outputPromise;
      expect(compressed).toBeInstanceOf(Uint8Array);
    });

    it("should support end callback", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const deflate = createDeflateStream({ useWorker: true });
      deflate.write(new TextEncoder().encode("Hello"));

      await new Promise<void>((resolve, reject) => {
        deflate.on("end", resolve);
        deflate.on("error", reject);
        deflate.end();
      });
    });

    it("should emit error for write after end", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const deflate = createDeflateStream({ useWorker: true });
      deflate.end();

      await new Promise<void>(resolve => {
        deflate.write(new TextEncoder().encode("test"), err => {
          expect(err).toBeInstanceOf(Error);
          expect(err!.message).toContain("write after end");
          resolve();
        });
      });
    });

    it("should use custom worker pool when provided", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const customPool = new WorkerPool({ maxWorkers: 2 });

      try {
        const deflate = createDeflateStream({ useWorker: true, workerPool: customPool });
        const outputPromise = collectStreamOutput(deflate);

        deflate.write(new TextEncoder().encode("Hello, World!"));
        deflate.end();

        const compressed = await outputPromise;
        expect(compressed.length).toBeGreaterThan(0);

        // Verify pool was used
        const stats = customPool.getStats();
        expect(stats.completedTasks).toBe(1);
      } finally {
        customPool.terminate();
      }
    });
  });

  describe("createInflateStream with useWorker", () => {
    it("should decompress data using worker-based stream", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // First compress some data
      const original = new TextEncoder().encode("Hello, World!");
      const deflate = createDeflateStream();
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      // Now decompress with worker
      const inflate = createInflateStream({ useWorker: true });
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("should decompress large data using worker-based stream", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(100 * 1024);

      // Compress
      const deflate = createDeflateStream();
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      // Decompress with worker
      const inflate = createInflateStream({ useWorker: true });
      const decompressPromise = collectStreamOutput(inflate);

      // Write in chunks
      const chunkSize = 8 * 1024;
      for (let i = 0; i < compressed.length; i += chunkSize) {
        inflate.write(compressed.slice(i, Math.min(i + chunkSize, compressed.length)));
      }
      inflate.end();

      const decompressed = await decompressPromise;
      expect(decompressed).toEqual(original);
    });

    it("should handle destroy correctly", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const inflate = createInflateStream({ useWorker: true });

      let errorEmitted = false;
      inflate.on("error", () => {
        errorEmitted = true;
      });

      inflate.write(new Uint8Array([1, 2, 3]));
      inflate.destroy(new Error("Test error"));

      // Give time for error to be emitted
      await new Promise(resolve => setTimeout(resolve, 10));
      expect(errorEmitted).toBe(true);
    });
  });

  describe("round-trip with worker streams", () => {
    it("should work: deflate(worker) -> inflate(worker)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(50 * 1024);

      // Compress with worker
      const deflate = createDeflateStream({ useWorker: true, level: 6 });
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      // Decompress with worker
      const inflate = createInflateStream({ useWorker: true });
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(decompressed).toEqual(original);
    });

    it("should work: deflate(main) -> inflate(worker)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(50 * 1024);

      // Compress on main thread
      const deflate = createDeflateStream({ useWorker: false });
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      // Decompress with worker
      const inflate = createInflateStream({ useWorker: true });
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(decompressed).toEqual(original);
    });

    it("should work: deflate(worker) -> inflate(main)", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = createCompressibleData(50 * 1024);

      // Compress with worker
      const deflate = createDeflateStream({ useWorker: true });
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      // Decompress on main thread
      const inflate = createInflateStream({ useWorker: false });
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(decompressed).toEqual(original);
    });
  });

  describe("fallback behavior", () => {
    it("should use native streaming when useWorker is false", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const deflate = createDeflateStream({ useWorker: false });
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      const inflate = createInflateStream({ useWorker: false });
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("should default to non-worker when useWorker not specified", async () => {
      const original = new TextEncoder().encode("Hello, World!");
      const deflate = createDeflateStream(); // No options
      const compressPromise = collectStreamOutput(deflate);
      deflate.write(original);
      deflate.end();
      const compressed = await compressPromise;

      const inflate = createInflateStream(); // No options
      const decompressPromise = collectStreamOutput(inflate);
      inflate.write(compressed);
      inflate.end();

      const decompressed = await decompressPromise;
      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });
  });
});
