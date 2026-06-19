/**
 * Worker Pool Browser Tests
 *
 * Comprehensive test suite for the browser worker pool implementation.
 * Tests cover:
 * - Basic compression/decompression operations
 * - Worker pool lifecycle management
 * - Task prioritization and scheduling
 * - Concurrent task processing
 * - Error handling and recovery
 * - Task cancellation via AbortSignal
 * - Worker idle timeout and cleanup
 */

import type { WorkerPoolOptions } from "@archive/compression/worker-pool/index.browser";
import {
  WorkerPool,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool,
  deflateWithPool,
  inflateWithPool,
  hasWorkerSupport
} from "@archive/compression/worker-pool/index.browser";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

// Helper to create test data
function createTestData(size: number, pattern = 0x42): Uint8Array {
  const data = new Uint8Array(size);
  for (let i = 0; i < size; i++) {
    data[i] = (pattern + i) % 256;
  }
  return data;
}

// Helper to create compressible data (repetitive pattern)
function createCompressibleData(size: number): Uint8Array {
  const data = new Uint8Array(size);
  const pattern = new TextEncoder().encode("Hello World! This is a repeating pattern. ");
  for (let i = 0; i < size; i++) {
    data[i] = pattern[i % pattern.length];
  }
  return data;
}

describe("WorkerPool", () => {
  let pool: WorkerPool | null = null;

  beforeEach(() => {
    // Clean up any existing default pool
    terminateDefaultWorkerPool();
  });

  afterEach(() => {
    // Clean up test pool
    if (pool) {
      pool.terminate();
      pool = null;
    }
    // Clean up default pool
    terminateDefaultWorkerPool();
  });

  describe("hasWorkerSupport", () => {
    it("should return true in browser environment with Worker support", () => {
      // In browser test environment, Worker should be available
      expect(hasWorkerSupport()).toBe(typeof Worker !== "undefined");
    });
  });

  describe("basic operations", () => {
    it("should compress and decompress small data correctly", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const original = new TextEncoder().encode("Hello, World!");

      // Compress
      const compressResult = await pool.execute("deflate", original);
      expect(compressResult.data).toBeInstanceOf(Uint8Array);
      expect(compressResult.data.length).toBeGreaterThan(0);
      expect(compressResult.duration).toBeGreaterThanOrEqual(0);

      // Decompress
      const decompressResult = await pool.execute("inflate", compressResult.data);
      expect(decompressResult.data).toBeInstanceOf(Uint8Array);

      // Verify round-trip
      expect(new TextDecoder().decode(decompressResult.data)).toBe("Hello, World!");
    });

    it("should compress and decompress larger data correctly", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const original = createCompressibleData(100 * 1024); // 100KB

      // Compress
      const compressResult = await pool.execute("deflate", original);
      expect(compressResult.data.length).toBeLessThan(original.length);

      // Decompress
      const decompressResult = await pool.execute("inflate", compressResult.data);

      // Verify round-trip
      expect(decompressResult.data.length).toBe(original.length);
      expect(decompressResult.data).toEqual(original);
    });

    it("should handle empty data", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const empty = new Uint8Array(0);

      // Compress
      const compressResult = await pool.execute("deflate", empty);
      expect(compressResult.data).toBeInstanceOf(Uint8Array);

      // Decompress
      const decompressResult = await pool.execute("inflate", compressResult.data);
      expect(decompressResult.data.length).toBe(0);
    });

    it("should support allowTransfer for zero-copy transfer", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const original = createCompressibleData(10 * 1024);
      // Copy the original data before transfer since the buffer will be detached
      const originalCopy = original.slice();

      // Compress with allowTransfer
      const result = await pool.execute("deflate", original, { allowTransfer: true });
      expect(result.data).toBeInstanceOf(Uint8Array);
      expect(result.data.length).toBeGreaterThan(0);

      // Decompress to verify using the copy
      const decompressed = await pool.execute("inflate", result.data);
      expect(decompressed.data).toEqual(originalCopy);
    });

    it("should not detach input when allowTransfer is disabled", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const original = createCompressibleData(8 * 1024);
      const beforeLength = original.byteLength;

      const compressed = await pool.execute("deflate", original, { allowTransfer: false });
      expect(compressed.data.length).toBeGreaterThan(0);

      expect(original.byteLength).toBe(beforeLength);
      expect(original.length).toBe(beforeLength);
      expect(original[0]).toBeGreaterThanOrEqual(0);
    });
  });

  describe("pool lifecycle", () => {
    it("should create pool with default options", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      expect(pool.isTerminated()).toBe(false);
    });

    it("should create pool with custom options", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const options: WorkerPoolOptions = {
        maxWorkers: 2,
        minWorkers: 1,
        idleTimeout: 5000
      };
      pool = new WorkerPool(options);
      expect(pool.isTerminated()).toBe(false);
    });

    it("should terminate pool correctly", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      pool.terminate();
      expect(pool.isTerminated()).toBe(true);
    });

    it("should reject new tasks after termination", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      pool.terminate();

      await expect(pool.execute("deflate", new Uint8Array(10))).rejects.toThrow(
        "Worker pool has been terminated"
      );
    });

    it("should handle multiple terminate calls gracefully", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      pool.terminate();
      pool.terminate(); // Should not throw
      expect(pool.isTerminated()).toBe(true);
    });
  });

  describe("pool statistics", () => {
    it("should report initial statistics", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool({ minWorkers: 0 });
      const stats = pool.getStats();

      expect(stats.totalWorkers).toBe(0);
      expect(stats.activeWorkers).toBe(0);
      expect(stats.idleWorkers).toBe(0);
      expect(stats.pendingTasks).toBe(0);
      expect(stats.completedTasks).toBe(0);
      expect(stats.failedTasks).toBe(0);
    });

    it("should update statistics after task completion", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool({ minWorkers: 0 });
      await pool.execute("deflate", createTestData(100));

      const stats = pool.getStats();
      expect(stats.completedTasks).toBe(1);
      expect(stats.failedTasks).toBe(0);
    });

    it("should pre-warm minimum workers", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool({ minWorkers: 2, maxWorkers: 4 });

      // Give workers time to initialize
      return new Promise<void>(resolve => {
        setTimeout(() => {
          const stats = pool!.getStats();
          expect(stats.totalWorkers).toBeGreaterThanOrEqual(2);
          resolve();
        }, 100);
      });
    });
  });

  describe("concurrent processing", () => {
    it("should process multiple tasks concurrently", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool({ maxWorkers: 4, minWorkers: 2 });

      const tasks = Array.from({ length: 8 }, (_, i) =>
        pool!.execute("deflate", createCompressibleData(10 * 1024), { level: 6 })
      );

      const results = await Promise.all(tasks);

      expect(results.length).toBe(8);
      results.forEach(result => {
        expect(result.data).toBeInstanceOf(Uint8Array);
        expect(result.data.length).toBeGreaterThan(0);
      });

      const stats = pool.getStats();
      expect(stats.completedTasks).toBe(8);
    });

    it("should scale up workers based on workload", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool({ maxWorkers: 4, minWorkers: 0 });

      // Start multiple tasks without awaiting
      const promises = Array.from({ length: 4 }, () =>
        pool!.execute("deflate", createCompressibleData(50 * 1024))
      );

      // Give workers time to start
      await new Promise(resolve => setTimeout(resolve, 50));

      const stats = pool.getStats();
      // Should have created workers to handle the load
      expect(stats.totalWorkers).toBeGreaterThan(0);

      // Wait for all tasks to complete
      await Promise.all(promises);
    });
  });

  describe("task prioritization", () => {
    it("should process high priority tasks first", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // Use a single worker to ensure tasks are queued
      pool = new WorkerPool({ maxWorkers: 1, minWorkers: 1 });

      const results: string[] = [];

      // Start a long-running task to block the worker
      const blocker = pool.execute("deflate", createCompressibleData(200 * 1024));

      // Queue tasks with different priorities
      const lowTask = pool
        .execute("deflate", createTestData(100), { priority: "low" })
        .then(() => results.push("low"));
      const highTask = pool
        .execute("deflate", createTestData(100), { priority: "high" })
        .then(() => results.push("high"));
      const normalTask = pool
        .execute("deflate", createTestData(100), { priority: "normal" })
        .then(() => results.push("normal"));

      await Promise.all([blocker, lowTask, highTask, normalTask]);

      // High priority should complete before normal, which should complete before low
      expect(results.indexOf("high")).toBeLessThan(results.indexOf("normal"));
      expect(results.indexOf("normal")).toBeLessThan(results.indexOf("low"));
    });
  });

  describe("task cancellation", () => {
    it("should cancel task with AbortSignal before execution", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // Use a single worker with a blocking task
      pool = new WorkerPool({ maxWorkers: 1, minWorkers: 1 });

      // Block the worker
      const blocker = pool.execute("deflate", createCompressibleData(100 * 1024));

      // Create an already-aborted controller
      const controller = new AbortController();
      controller.abort();

      // Queue a task with the aborted signal
      const taskPromise = pool.execute("deflate", createTestData(100), {
        signal: controller.signal
      });

      await expect(taskPromise).rejects.toThrow(/aborted/i);

      // Let the blocker complete
      await blocker;
    });

    it("should cancel task with AbortSignal during queue", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // Use a single worker
      pool = new WorkerPool({ maxWorkers: 1, minWorkers: 1 });

      // Block the worker with a large task that takes time
      const blocker = pool.execute("deflate", createCompressibleData(2 * 1024 * 1024));

      // Queue a task
      const controller = new AbortController();
      const taskPromise = pool.execute("deflate", createTestData(100), {
        signal: controller.signal
      });

      // Abort immediately while task is still queued (worker is busy with blocker)
      controller.abort();

      await expect(taskPromise).rejects.toThrow(/aborted/i);

      await blocker;
    });
  });

  describe("error handling", () => {
    it("should handle invalid compressed data gracefully", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const invalidData = new Uint8Array([0xff, 0xfe, 0xfd, 0xfc]); // Invalid deflate

      await expect(pool.execute("inflate", invalidData)).rejects.toThrow();

      // Pool should still be functional after error
      const validData = new TextEncoder().encode("test");
      const result = await pool.execute("deflate", validData);
      expect(result.data).toBeInstanceOf(Uint8Array);
    });

    it("should track failed tasks in statistics", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();

      try {
        await pool.execute("inflate", new Uint8Array([0xff, 0xfe]));
      } catch {
        // Expected
      }

      const stats = pool.getStats();
      expect(stats.failedTasks).toBe(1);
    });
  });

  describe("worker idle timeout", () => {
    it("should terminate idle workers after timeout", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // Short idle timeout for testing
      pool = new WorkerPool({
        maxWorkers: 4,
        minWorkers: 0,
        idleTimeout: 100
      });

      // Execute a task to create a worker
      await pool.execute("deflate", createTestData(100));

      let stats = pool.getStats();
      expect(stats.totalWorkers).toBeGreaterThan(0);

      // Wait for idle timeout
      await new Promise(resolve => setTimeout(resolve, 200));

      stats = pool.getStats();
      expect(stats.totalWorkers).toBe(0);
    });

    it("should keep minimum workers alive", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      // Short idle timeout with minWorkers
      pool = new WorkerPool({
        maxWorkers: 4,
        minWorkers: 1,
        idleTimeout: 100
      });

      // Execute tasks to create multiple workers
      await Promise.all([
        pool.execute("deflate", createTestData(100)),
        pool.execute("deflate", createTestData(100)),
        pool.execute("deflate", createTestData(100))
      ]);

      // Wait for idle timeout
      await new Promise(resolve => setTimeout(resolve, 200));

      const stats = pool.getStats();
      // Should keep at least minWorkers alive
      expect(stats.totalWorkers).toBeGreaterThanOrEqual(1);
    });
  });

  describe("convenience functions", () => {
    it("deflateWithPool should compress data", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await deflateWithPool(original);

      expect(compressed).toBeInstanceOf(Uint8Array);
      expect(compressed.length).toBeGreaterThan(0);
    });

    it("inflateWithPool should decompress data", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const original = new TextEncoder().encode("Hello, World!");
      const compressed = await deflateWithPool(original);
      const decompressed = await inflateWithPool(compressed);

      expect(new TextDecoder().decode(decompressed)).toBe("Hello, World!");
    });

    it("getDefaultWorkerPool should return singleton", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const pool1 = getDefaultWorkerPool();
      const pool2 = getDefaultWorkerPool();

      expect(pool1).toBe(pool2);
    });

    it("terminateDefaultWorkerPool should terminate singleton", () => {
      if (!hasWorkerSupport()) {
        return;
      }

      const pool1 = getDefaultWorkerPool();
      terminateDefaultWorkerPool();

      expect(pool1.isTerminated()).toBe(true);

      // Getting default pool again should create a new one
      const pool2 = getDefaultWorkerPool();
      expect(pool2).not.toBe(pool1);
      expect(pool2.isTerminated()).toBe(false);
    });
  });

  describe("data integrity", () => {
    it("should preserve data integrity for binary data", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();

      // Create binary data with all byte values
      const original = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        original[i] = i;
      }

      const compressed = await pool.execute("deflate", original);
      const decompressed = await pool.execute("inflate", compressed.data);

      expect(decompressed.data).toEqual(original);
    });

    it("should preserve data integrity for repeated compression", async () => {
      if (!hasWorkerSupport()) {
        return;
      }

      pool = new WorkerPool();
      const original = createCompressibleData(10 * 1024);

      // Compress and decompress multiple times
      let data = original;
      for (let i = 0; i < 3; i++) {
        const compressed = await pool.execute("deflate", data);
        const decompressed = await pool.execute("inflate", compressed.data);
        data = decompressed.data;
      }

      expect(data).toEqual(original);
    });
  });
});
