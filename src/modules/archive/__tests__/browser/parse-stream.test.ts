/**
 * Browser Parse Stream Tests
 *
 * Tests for createParse streaming API in browser environment.
 * These tests verify that entry.buffer() works correctly.
 */

import { zip } from "@archive/create-archive";
import { createParse, type ZipEntry } from "@archive/unzip/stream.browser";
import { describe, it, expect } from "vitest";

describe("parse.browser - createParse streaming", () => {
  // Helper to create a test ZIP with known content
  function createTestZip(files: Array<{ name: string; content: string }>): Uint8Array {
    const encoder = new TextEncoder();
    const z = zip({ level: 6 });
    for (const f of files) {
      z.add(f.name, encoder.encode(f.content));
    }
    return z.bytesSync();
  }

  it("should parse entries and get buffer content", async () => {
    const testContent = "Hello, World!";
    const zipData = createTestZip([{ name: "test.txt", content: testContent }]);

    const parse = createParse();
    const results: Array<{ path: string; content: string }> = [];
    const bufferPromises: Promise<void>[] = [];

    return new Promise<void>((resolve, reject) => {
      parse.on("entry", (entry: ZipEntry) => {
        // Store promise for later await
        const promise = entry.buffer().then(buffer => {
          const content = new TextDecoder().decode(buffer);
          results.push({ path: entry.path, content });
        });
        bufferPromises.push(promise);
      });

      parse.on("close", async () => {
        try {
          // Wait for all buffer operations to complete
          await Promise.all(bufferPromises);
          expect(results.length).toBe(1);
          expect(results[0]!.path).toBe("test.txt");
          expect(results[0]!.content).toBe(testContent);
          resolve();
        } catch (err) {
          reject(err);
        }
      });

      parse.on("error", reject);

      // Feed data to the stream
      parse.end(zipData);
    });
  });

  it("should parse multiple entries", async () => {
    const files = [
      { name: "file1.txt", content: "Content 1" },
      { name: "file2.txt", content: "Content 2" },
      { name: "file3.txt", content: "Content 3" }
    ];
    const zipData = createTestZip(files);

    const parse = createParse();
    const results: Array<{ path: string; content: string }> = [];
    const pendingBuffers: Promise<void>[] = [];

    return new Promise<void>((resolve, reject) => {
      parse.on("entry", (entry: ZipEntry) => {
        const promise = entry.buffer().then(buffer => {
          const content = new TextDecoder().decode(buffer);
          results.push({ path: entry.path, content });
        });
        pendingBuffers.push(promise);
      });

      parse.on("close", async () => {
        try {
          // Wait for all buffers to be read
          await Promise.all(pendingBuffers);

          expect(results.length).toBe(3);
          // Sort by path for consistent comparison
          results.sort((a, b) => a.path.localeCompare(b.path));
          expect(results[0].content).toBe("Content 1");
          expect(results[1].content).toBe("Content 2");
          expect(results[2].content).toBe("Content 3");
          resolve();
        } catch (err) {
          reject(err);
        }
      });

      parse.on("error", reject);

      parse.end(zipData);
    });
  });

  it("should handle large files", async () => {
    // Create a 100KB file
    const largeContent = "x".repeat(100 * 1024);
    const zipData = createTestZip([{ name: "large.txt", content: largeContent }]);

    const parse = createParse();

    return new Promise<void>((resolve, reject) => {
      parse.on("entry", async (entry: ZipEntry) => {
        try {
          const buffer = await entry.buffer();
          const content = new TextDecoder().decode(buffer);
          expect(content.length).toBe(largeContent.length);
          expect(content).toBe(largeContent);
        } catch (err) {
          reject(err);
        }
      });

      parse.on("close", resolve);
      parse.on("error", reject);

      parse.end(zipData);
    });
  });

  it("should handle autodrain correctly", async () => {
    const zipData = createTestZip([
      { name: "read.txt", content: "Read this" },
      { name: "skip.txt", content: "Skip this" }
    ]);

    const parse = createParse();
    const results: string[] = [];

    return new Promise<void>((resolve, reject) => {
      parse.on("entry", async (entry: ZipEntry) => {
        if (entry.path === "read.txt") {
          const buffer = await entry.buffer();
          results.push(new TextDecoder().decode(buffer));
        } else {
          entry.autodrain();
        }
      });

      parse.on("close", () => {
        try {
          expect(results.length).toBe(1);
          expect(results[0]).toBe("Read this");
          resolve();
        } catch (err) {
          reject(err);
        }
      });

      parse.on("error", reject);

      parse.end(zipData);
    });
  });

  it("entry.buffer() should resolve even after close event", async () => {
    // This is the bug scenario: entry events fire, close fires,
    // but entry.buffer() never resolves
    const zipData = createTestZip([
      { name: "a.txt", content: "AAA" },
      { name: "b.txt", content: "BBB" }
    ]);

    const parse = createParse();
    const bufferPromises: Promise<Uint8Array>[] = [];

    return new Promise<void>((resolve, reject) => {
      // Set a timeout to detect hanging
      const timeout = setTimeout(() => {
        reject(new Error("Test timed out - entry.buffer() never resolved"));
      }, 5000);

      parse.on("entry", (entry: ZipEntry) => {
        // Collect all buffer promises
        bufferPromises.push(entry.buffer());
      });

      parse.on("close", async () => {
        try {
          // All buffer promises should resolve
          const buffers = await Promise.all(bufferPromises);
          clearTimeout(timeout);

          expect(buffers.length).toBe(2);
          expect(new TextDecoder().decode(buffers[0])).toBe("AAA");
          expect(new TextDecoder().decode(buffers[1])).toBe("BBB");
          resolve();
        } catch (err) {
          clearTimeout(timeout);
          reject(err);
        }
      });

      parse.on("error", err => {
        clearTimeout(timeout);
        reject(err);
      });

      parse.end(zipData);
    });
  });
});
