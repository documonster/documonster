import { DEFAULT_COMPRESS_THRESHOLD_BYTES } from "@archive/compression/compress.base";
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { Readable } from "@stream";
import { describe, it, expect } from "vitest";

describe("parse threshold optimization (browser)", () => {
  it("should have default threshold of 5MB", () => {
    expect(DEFAULT_COMPRESS_THRESHOLD_BYTES).toBe(8 * 1024 * 1024);
  });

  describe("small file sync decompression", () => {
    it("should correctly parse small compressed files (using sync path)", async () => {
      // Create a small zip with compressed content
      const smallContent = "Hello, World! This is a test file.";
      const smallContentBytes = new TextEncoder().encode(smallContent);

      const zipBuffer = zip({ level: 6 }).add("small.txt", smallContentBytes).bytesSync();

      // Create a readable stream from the buffer
      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true } });

      let extractedContent = "";

      for await (const entry of reader.entries()) {
        if (entry.path === "small.txt") {
          extractedContent = await entry.text("utf-8");
        } else {
          entry.discard();
        }
      }

      expect(extractedContent).toBe(smallContent);
    });

    it("should correctly parse uncompressed (STORE) files", async () => {
      // Create a zip with STORE (no compression)
      const content = "Uncompressed content for STORE mode test.";
      const contentBytes = new TextEncoder().encode(content);

      const zipBuffer = zip({ level: 0 }).add("store.txt", contentBytes).bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true } });

      let extractedContent = "";

      for await (const entry of reader.entries()) {
        if (entry.path === "store.txt") {
          extractedContent = await entry.text("utf-8");
        } else {
          entry.discard();
        }
      }

      expect(extractedContent).toBe(content);
    });

    it("should handle multiple small files in same archive", async () => {
      const files = [
        { name: "file1.txt", content: "Content of file 1" },
        { name: "file2.txt", content: "Content of file 2" },
        { name: "file3.txt", content: "Content of file 3" }
      ];

      const z = zip({ level: 6 });
      for (const f of files) {
        z.add(f.name, new TextEncoder().encode(f.content));
      }
      const zipBuffer = z.bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true } });

      const extracted: Record<string, string> = {};

      for await (const entry of reader.entries()) {
        extracted[entry.path] = await entry.text("utf-8");
      }

      for (const file of files) {
        expect(extracted[file.name]).toBe(file.content);
      }
    });

    it("should respect custom thresholdBytes option", async () => {
      // Create a file that's larger than our custom threshold
      const content = "x".repeat(1000); // 1KB
      const contentBytes = new TextEncoder().encode(content);

      const zipBuffer = zip({ level: 6 }).add("test.txt", contentBytes).bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      // Set threshold to 500 bytes - file should use streaming path
      const reader = unzip(readable, { parse: { forceStream: true, thresholdBytes: 500 } });

      let extractedContent = "";

      for await (const entry of reader.entries()) {
        if (entry.path === "test.txt") {
          extractedContent = await entry.text("utf-8");
        } else {
          entry.discard();
        }
      }

      // Content should still be extracted correctly regardless of path used
      expect(extractedContent).toBe(content);
    });

    it("should handle binary data correctly", async () => {
      // Create binary data (random bytes)
      const binaryData = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        binaryData[i] = i;
      }

      const zipBuffer = zip({ level: 6 }).add("binary.bin", binaryData).bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true } });

      let extractedData: Uint8Array | null = null;

      for await (const entry of reader.entries()) {
        if (entry.path === "binary.bin") {
          extractedData = await entry.bytes();
        } else {
          entry.discard();
        }
      }

      expect(extractedData).not.toBeNull();
      expect(extractedData!.length).toBe(256);
      for (let i = 0; i < 256; i++) {
        expect(extractedData![i]).toBe(i);
      }
    });
  });

  describe("threshold boundary conditions", () => {
    it("should handle empty files", async () => {
      const zipBuffer = zip({ level: 6 }).add("empty.txt", new Uint8Array(0)).bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true } });

      let extractedData: Uint8Array | null = null;

      for await (const entry of reader.entries()) {
        if (entry.path === "empty.txt") {
          extractedData = await entry.bytes();
        } else {
          entry.discard();
        }
      }

      expect(extractedData).not.toBeNull();
      expect(extractedData!.length).toBe(0);
    });

    it("should handle files at threshold boundary", async () => {
      // Test with a file exactly at threshold boundary
      const thresholdBytes = 1024; // 1KB for testing
      const content = "x".repeat(thresholdBytes);
      const contentBytes = new TextEncoder().encode(content);

      const zipBuffer = zip({ level: 6 }).add("boundary.txt", contentBytes).bytesSync();

      const readable = new Readable();
      readable.push(zipBuffer);
      readable.push(null);

      const reader = unzip(readable, { parse: { forceStream: true, thresholdBytes } });

      let extractedContent = "";

      for await (const entry of reader.entries()) {
        if (entry.path === "boundary.txt") {
          extractedContent = await entry.text("utf-8");
        } else {
          entry.discard();
        }
      }

      expect(extractedContent).toBe(content);
    });
  });
});
