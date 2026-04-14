/**
 * TAR Gzip Tests
 *
 * Tests for tar.gz / tgz archive creation and extraction.
 */

import { gzip, gunzip, gzipSync, gunzipSync } from "@archive/compression/compress";
import { TarGzArchive, targz, parseTarGz, untargz, TarArchive } from "@archive/tar";
import { describe, it, expect } from "vitest";

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

// Gzip magic number
const GZIP_MAGIC = [0x1f, 0x8b];

describe("TAR Gzip Module", () => {
  describe("TarGzArchive class", () => {
    it("should create a gzip-compressed tar archive", async () => {
      const archive = new TarGzArchive();
      archive.add("hello.txt", "Hello, TarGz!");

      const bytes = await archive.bytes();

      // Verify gzip magic number
      expect(bytes[0]).toBe(GZIP_MAGIC[0]);
      expect(bytes[1]).toBe(GZIP_MAGIC[1]);

      // Decompress and verify content
      const entries = await parseTarGz(bytes);
      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe("hello.txt");
      expect(await entries[0].text()).toBe("Hello, TarGz!");
    });

    it("should support compression levels", async () => {
      const content = "a".repeat(10000);

      const noCompress = new TarGzArchive({ level: 0 });
      noCompress.add("file.txt", content);
      const level0 = await noCompress.bytes();

      const maxCompress = new TarGzArchive({ level: 9 });
      maxCompress.add("file.txt", content);
      const level9 = await maxCompress.bytes();

      // Level 9 should be smaller than level 0 for compressible data
      expect(level9.length).toBeLessThan(level0.length);

      // Both should roundtrip correctly
      const entries0 = await parseTarGz(level0);
      const entries9 = await parseTarGz(level9);
      expect(await entries0[0].text()).toBe(content);
      expect(await entries9[0].text()).toBe(content);
    });

    it("should handle multiple files", async () => {
      const archive = new TarGzArchive();
      archive.add("file1.txt", "Content 1");
      archive.add("file2.txt", "Content 2");
      archive.addDirectory("dir");
      archive.add("dir/file3.txt", "Content 3");

      const bytes = await archive.bytes();
      const entries = await parseTarGz(bytes);

      expect(entries.length).toBe(4);
      expect(entries.map(e => e.info.path)).toEqual([
        "file1.txt",
        "file2.txt",
        "dir/",
        "dir/file3.txt"
      ]);
    });
  });

  describe("targz function", () => {
    it("should create tar.gz from Map", async () => {
      const files = new Map<string, string>([
        ["a.txt", "AAA"],
        ["b.txt", "BBB"]
      ]);

      const bytes = await targz(files);

      // Verify gzip format
      expect(bytes[0]).toBe(GZIP_MAGIC[0]);
      expect(bytes[1]).toBe(GZIP_MAGIC[1]);

      const entries = await parseTarGz(bytes);
      expect(entries.length).toBe(2);
    });

    it("should create tar.gz from array", async () => {
      const files = [
        { name: "x.txt", source: "XXX" },
        { name: "y.txt", source: textEncoder.encode("YYY") }
      ];

      const bytes = await targz(files);
      const entries = await parseTarGz(bytes);

      expect(entries.length).toBe(2);
      expect(await entries[0].text()).toBe("XXX");
      expect(await entries[1].text()).toBe("YYY");
    });
  });

  describe("parseTarGz", () => {
    it("should parse gzip-compressed tar", async () => {
      // Create a tar.gz
      const archive = new TarGzArchive();
      archive.add("test.txt", "Test content");
      const compressed = await archive.bytes();

      // Parse it
      const entries = await parseTarGz(compressed);

      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe("test.txt");
      expect(await entries[0].text()).toBe("Test content");
    });
  });

  describe("untargz", () => {
    it("should extract all entries", async () => {
      const archive = new TarGzArchive();
      archive.add("one.txt", "First");
      archive.add("two.txt", "Second");

      const bytes = await archive.bytes();
      const extracted = await untargz(bytes);

      expect(extracted.size).toBe(2);
      expect(textDecoder.decode(extracted.get("one.txt")!.data)).toBe("First");
      expect(textDecoder.decode(extracted.get("two.txt")!.data)).toBe("Second");
    });
  });

  describe("gzip/gunzip utilities", () => {
    it("gzip should compress data", async () => {
      const data = textEncoder.encode("Hello, World!");
      const compressed = await gzip(data);

      expect(compressed[0]).toBe(GZIP_MAGIC[0]);
      expect(compressed[1]).toBe(GZIP_MAGIC[1]);

      const decompressed = await gunzip(compressed);
      expect(textDecoder.decode(decompressed)).toBe("Hello, World!");
    });

    it("gzipSync/gunzipSync should work synchronously", () => {
      const data = textEncoder.encode("Sync compression test");
      const compressed = gzipSync(data);

      expect(compressed[0]).toBe(GZIP_MAGIC[0]);

      const decompressed = gunzipSync(compressed);
      expect(textDecoder.decode(decompressed)).toBe("Sync compression test");
    });

    it("should support different compression levels", async () => {
      const data = textEncoder.encode("a".repeat(10000));

      const level1 = await gzip(data, { level: 1 });
      const level9 = await gzip(data, { level: 9 });

      // Higher compression should produce smaller output
      expect(level9.length).toBeLessThan(level1.length);

      // Both should decompress correctly
      expect(await gunzip(level1)).toEqual(data);
      expect(await gunzip(level9)).toEqual(data);
    });
  });

  describe("gzip existing tar", () => {
    it("should compress an existing tar archive with gzip", async () => {
      // Create uncompressed tar
      const archive = new TarArchive();
      archive.add("file.txt", "Uncompressed TAR content");
      const tarBytes = await archive.bytes();

      // Compress with gzip
      const compressed = await gzip(tarBytes);

      expect(compressed[0]).toBe(GZIP_MAGIC[0]);
      expect(compressed[1]).toBe(GZIP_MAGIC[1]);

      // Verify roundtrip
      const entries = await parseTarGz(compressed);
      expect(entries.length).toBe(1);
      expect(await entries[0].text()).toBe("Uncompressed TAR content");
    });
  });

  describe("Roundtrip tests", () => {
    it("should roundtrip binary data through tar.gz", async () => {
      const binaryData = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        binaryData[i] = i;
      }

      const archive = new TarGzArchive();
      archive.add("binary.bin", binaryData);

      const compressed = await archive.bytes();
      const entries = await parseTarGz(compressed);

      expect(entries[0].info.size).toBe(256);
      expect(await entries[0].data()).toEqual(binaryData);
    });

    it("should roundtrip Unicode content through tar.gz", async () => {
      const content = "Hello 世界 🌍 مرحبا";

      const archive = new TarGzArchive();
      archive.add("unicode.txt", content);

      const compressed = await archive.bytes();
      const entries = await parseTarGz(compressed);

      expect(await entries[0].text()).toBe(content);
    });

    it("should roundtrip long filenames through tar.gz", async () => {
      const longPath = "very/deep/nested/directory/".repeat(5) + "file.txt";

      const archive = new TarGzArchive();
      archive.add(longPath, "Long path content");

      const compressed = await archive.bytes();
      const entries = await parseTarGz(compressed);

      expect(entries[0].info.path).toBe(longPath);
    });

    it("should preserve file metadata through tar.gz", async () => {
      const mtime = new Date("2023-01-15T10:30:00Z");
      const mode = 0o755;

      const archive = new TarGzArchive();
      archive.add("script.sh", "#!/bin/bash", { mode, mtime });

      const compressed = await archive.bytes();
      const entries = await parseTarGz(compressed);

      expect(entries[0].info.mode).toBe(mode);
      // TAR stores seconds precision
      const expectedTime = Math.floor(mtime.getTime() / 1000) * 1000;
      expect(entries[0].info.mtime.getTime()).toBe(expectedTime);
    });
  });

  describe("Compression efficiency", () => {
    it("should significantly compress repetitive data", async () => {
      const data = "hello world ".repeat(1000);
      const archive = new TarGzArchive({ level: 9 });
      archive.add("repetitive.txt", data);

      const compressed = await archive.bytes();
      const uncompressed = await new TarArchive().add("repetitive.txt", data).bytes();

      // Compressed should be significantly smaller
      expect(compressed.length).toBeLessThan(uncompressed.length / 5);
    });
  });
});
