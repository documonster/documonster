/**
 * TAR E2E Tests
 *
 * Comprehensive tests for TAR archive creation and parsing.
 * Tests roundtrip functionality, long filenames, symlinks, and streaming.
 */

import { TAR_BLOCK_SIZE, TAR_TYPE, TarArchive, tar, tarSync, parseTar, untar } from "@archive/tar";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect } from "vitest";

const textEncoder = new TextEncoder();
const textDecoder = new TextDecoder();

describe("TAR Module", () => {
  describe("TarArchive class", () => {
    it("should create empty tar archive", async () => {
      const archive = new TarArchive();
      const bytes = await archive.bytes();

      // Minimum tar is just end marker (2 blocks)
      expect(bytes.length).toBe(TAR_BLOCK_SIZE * 2);

      // Verify it's all zeros (end-of-archive marker)
      for (let i = 0; i < bytes.length; i++) {
        expect(bytes[i]).toBe(0);
      }
    });

    it("should add and retrieve a single file", async () => {
      const archive = new TarArchive();
      archive.add("hello.txt", "Hello, TAR!");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe("hello.txt");
      expect(entries[0].info.type).toBe(TAR_TYPE.FILE);
      expect(await entries[0].text()).toBe("Hello, TAR!");
    });

    it("should add multiple files", async () => {
      const archive = new TarArchive();
      archive.add("file1.txt", "Content 1");
      archive.add("file2.txt", "Content 2");
      archive.add("file3.txt", "Content 3");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(3);
      expect(entries.map(e => e.info.path)).toEqual(["file1.txt", "file2.txt", "file3.txt"]);
      expect(await entries[0].text()).toBe("Content 1");
      expect(await entries[1].text()).toBe("Content 2");
      expect(await entries[2].text()).toBe("Content 3");
    });

    it("should handle directories", async () => {
      const archive = new TarArchive();
      archive.addDirectory("mydir");
      archive.add("mydir/file.txt", "Hello");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(2);
      expect(entries[0].info.path).toBe("mydir/");
      expect(entries[0].info.type).toBe(TAR_TYPE.DIRECTORY);
      expect(entries[1].info.path).toBe("mydir/file.txt");
      expect(entries[1].info.type).toBe(TAR_TYPE.FILE);
    });

    it("should handle symlinks", async () => {
      const archive = new TarArchive();
      archive.add("original.txt", "Original content");
      archive.addSymlink("link.txt", "original.txt");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(2);
      expect(entries[1].info.path).toBe("link.txt");
      expect(entries[1].info.type).toBe(TAR_TYPE.SYMLINK);
      expect(entries[1].info.linkname).toBe("original.txt");
    });

    it("should preserve file modes", async () => {
      const archive = new TarArchive();
      archive.add("executable.sh", "#!/bin/bash\necho hello", { mode: 0o755 });
      archive.add("readonly.txt", "readonly content", { mode: 0o444 });

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries[0].info.mode).toBe(0o755);
      expect(entries[1].info.mode).toBe(0o444);
    });

    it("should preserve modification time", async () => {
      const mtime = new Date("2023-06-15T12:30:00Z");
      const archive = new TarArchive();
      archive.add("file.txt", "content", { mtime });

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      // TAR only stores seconds precision
      const expectedTime = Math.floor(mtime.getTime() / 1000) * 1000;
      expect(entries[0].info.mtime.getTime()).toBe(expectedTime);
    });

    it("should support method chaining", () => {
      const archive = new TarArchive();
      const result = archive
        .add("a.txt", "A")
        .add("b.txt", "B")
        .addDirectory("dir")
        .addSymlink("c.txt", "a.txt");

      expect(result).toBe(archive);
    });

    it("should throw when adding after streaming", async () => {
      const archive = new TarArchive();
      archive.add("file.txt", "content");

      // Start streaming
      const stream = archive.stream();
      for await (const _ of stream) {
        // Consume first chunk
        break;
      }

      // Should throw when adding more
      expect(() => archive.add("another.txt", "data")).toThrow();
    });
  });

  describe("tarSync function", () => {
    it("should create tar from Map", () => {
      const entries = new Map<string, string>([
        ["a.txt", "Content A"],
        ["b.txt", "Content B"]
      ]);

      const bytes = tarSync(entries);
      const parsed = parseTar(bytes);

      expect(parsed.length).toBe(2);
    });

    it("should create tar from array", () => {
      const entries = [
        { name: "file1.txt", source: "Hello" },
        { name: "file2.txt", source: textEncoder.encode("World") }
      ];

      const bytes = tarSync(entries);
      const parsed = parseTar(bytes);

      expect(parsed.length).toBe(2);
    });
  });

  describe("tar async function", () => {
    it("should create tar from Map", async () => {
      const entries = new Map<string, Uint8Array>([
        ["x.txt", textEncoder.encode("X")],
        ["y.txt", textEncoder.encode("Y")]
      ]);

      const bytes = await tar(entries);
      const parsed = parseTar(bytes);

      expect(parsed.length).toBe(2);
    });
  });

  describe("Long filename support (GNU extension)", () => {
    it("should handle filenames up to 100 chars", async () => {
      const filename = "a".repeat(99) + ".txt";
      const archive = new TarArchive();
      archive.add(filename, "short filename content");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe(filename);
    });

    it("should handle long filenames (> 100 chars) using GNU extension", async () => {
      const filename = "very/long/path/".repeat(10) + "file.txt";
      expect(filename.length).toBeGreaterThan(100);

      const archive = new TarArchive();
      archive.add(filename, "long filename content");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe(filename);
      expect(await entries[0].text()).toBe("long filename content");
    });

    it("should handle very long filenames (> 256 chars)", async () => {
      const longDir = "subdir/".repeat(50);
      const filename = longDir + "file.txt";
      expect(filename.length).toBeGreaterThan(256);

      const archive = new TarArchive();
      archive.add(filename, "very long path");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(1);
      expect(entries[0].info.path).toBe(filename);
    });
  });

  describe("Binary data", () => {
    it("should handle binary data correctly", async () => {
      const binaryData = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        binaryData[i] = i;
      }

      const archive = new TarArchive();
      archive.add("binary.bin", binaryData);

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(1);
      expect(entries[0].info.size).toBe(256);
      const data = await entries[0].data();
      expect(data).toEqual(binaryData);
    });

    it("should handle large files", async () => {
      const largeData = new Uint8Array(1024 * 100); // 100KB
      for (let i = 0; i < largeData.length; i++) {
        largeData[i] = i % 256;
      }

      const archive = new TarArchive();
      archive.add("large.bin", largeData);

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries[0].info.size).toBe(largeData.length);
      const data = await entries[0].data();
      expect(data).toEqual(largeData);
    });

    it("should handle empty files", async () => {
      const archive = new TarArchive();
      archive.add("empty.txt", new Uint8Array(0));

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries[0].info.size).toBe(0);
      const data = await entries[0].data();
      expect(data.length).toBe(0);
    });
  });

  describe("Padding and block alignment", () => {
    it("should correctly pad files to 512-byte blocks", async () => {
      // Test various sizes to verify padding
      const testSizes = [1, 100, 511, 512, 513, 1000, 1024];

      for (const size of testSizes) {
        const data = new Uint8Array(size);
        data.fill(0x42); // Fill with 'B'

        const archive = new TarArchive();
        archive.add("test.bin", data);

        const bytes = await archive.bytes();

        // Total size should be block-aligned
        expect(bytes.length % TAR_BLOCK_SIZE).toBe(0);

        // Parse and verify data integrity
        const entries = parseTar(bytes);
        expect(entries[0].info.size).toBe(size);
        const parsed = await entries[0].data();
        expect(parsed).toEqual(data);
      }
    });
  });

  describe("Unicode support", () => {
    it("should handle Unicode filenames", async () => {
      const archive = new TarArchive();
      archive.add("文件.txt", "中文内容");
      archive.add("ファイル.txt", "日本語コンテンツ");
      archive.add("🎉emoji🎊.txt", "emoji content 🚀");

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(3);
      expect(entries[0].info.path).toBe("文件.txt");
      expect(entries[1].info.path).toBe("ファイル.txt");
      expect(entries[2].info.path).toBe("🎉emoji🎊.txt");

      expect(await entries[0].text()).toBe("中文内容");
      expect(await entries[1].text()).toBe("日本語コンテンツ");
      expect(await entries[2].text()).toBe("emoji content 🚀");
    });

    it("should handle Unicode content", async () => {
      const content = "Hello 世界 🌍 مرحبا العالم שלום עולם";
      const archive = new TarArchive();
      archive.add("unicode.txt", content);

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(await entries[0].text()).toBe(content);
    });
  });

  describe("Streaming output", () => {
    it("should produce same result via stream() and bytes()", async () => {
      const archive1 = new TarArchive();
      archive1.add("file.txt", "content");
      const direct = await archive1.bytes();

      const archive2 = new TarArchive();
      archive2.add("file.txt", "content");
      const chunks: Uint8Array[] = [];
      for await (const chunk of archive2.stream()) {
        chunks.push(chunk);
      }
      const streamed = concatUint8Arrays(chunks);

      expect(streamed).toEqual(direct);
    });

    it("should emit progress events", async () => {
      const progress: Array<{ entry?: string; count: number }> = [];

      const archive = new TarArchive({
        onProgress: p => progress.push({ entry: p.currentEntry?.name, count: p.entriesDone })
      });
      archive.add("a.txt", "A");
      archive.add("b.txt", "B");

      await archive.bytes();

      expect(progress.length).toBeGreaterThan(0);
      expect(progress[progress.length - 1].count).toBe(2);
    });
  });

  describe("untar convenience function", () => {
    it("should extract all files to a Map", async () => {
      const archive = new TarArchive();
      archive.add("one.txt", "First");
      archive.add("two.txt", "Second");
      archive.addDirectory("dir");
      archive.add("dir/three.txt", "Third");

      const bytes = await archive.bytes();
      const files = await untar(bytes);

      // untar includes all entries including directories
      expect(files.size).toBe(4);
      expect(textDecoder.decode(files.get("one.txt")!.data)).toBe("First");
      expect(textDecoder.decode(files.get("two.txt")!.data)).toBe("Second");
      expect(textDecoder.decode(files.get("dir/three.txt")!.data)).toBe("Third");
      expect(files.get("dir/")!.info.type).toBe(TAR_TYPE.DIRECTORY);
    });
  });

  // NOTE: Comprehensive roundtrip tests are in archive-roundtrip.test.ts

  describe("Error handling", () => {
    it("should throw on empty entry name", () => {
      const archive = new TarArchive();
      expect(() => archive.add("", "content")).toThrow();
    });

    it("should handle corrupted tar gracefully", () => {
      const corrupted = new Uint8Array(TAR_BLOCK_SIZE * 3);
      // Put some garbage that doesn't have valid checksum
      corrupted[0] = 0x50;
      corrupted[1] = 0x4b;

      expect(() => parseTar(corrupted)).toThrow();
    });

    it("should reject files exceeding maxFileSize", async () => {
      const archive = new TarArchive();
      archive.add("large.bin", new Uint8Array(1000));
      const bytes = await archive.bytes();

      // Parse with very small maxFileSize
      expect(() => parseTar(bytes, { maxFileSize: 100 })).toThrow(/exceeds maximum file size/);
    });
  });

  describe("Hard link support", () => {
    it("should create hard link entries", async () => {
      const archive = new TarArchive();
      archive.add("original.txt", "Original content");
      // Add hard link entry manually using type
      archive.add("hardlink.txt", new Uint8Array(0), {
        type: "1" as any, // TAR_TYPE.HARD_LINK = "1"
        linkname: "original.txt"
      });

      const bytes = await archive.bytes();
      const entries = parseTar(bytes);

      expect(entries.length).toBe(2);
      expect(entries[1].info.type).toBe("1");
      expect(entries[1].info.linkname).toBe("original.txt");
    });
  });

  describe("TarReader class", () => {
    it("should read entries with entries() iterator", async () => {
      const archive = new TarArchive();
      archive.add("a.txt", "A");
      archive.add("b.txt", "B");
      const bytes = await archive.bytes();

      const { TarReader } = await import("@archive/tar");
      const reader = new TarReader(bytes);
      const paths: string[] = [];

      for await (const entry of reader.entries()) {
        paths.push(entry.path);
      }

      expect(paths).toEqual(["a.txt", "b.txt"]);
    });

    it("should get entry by path", async () => {
      const archive = new TarArchive();
      archive.add("file.txt", "Hello");
      const bytes = await archive.bytes();

      const { TarReader } = await import("@archive/tar");
      const reader = new TarReader(bytes);

      const entry = await reader.get("file.txt");
      expect(entry).not.toBeNull();
      expect(await entry!.text()).toBe("Hello");
    });

    it("should return null for non-existent path", async () => {
      const archive = new TarArchive();
      archive.add("file.txt", "Hello");
      const bytes = await archive.bytes();

      const { TarReader } = await import("@archive/tar");
      const reader = new TarReader(bytes);

      const entry = await reader.get("nonexistent.txt");
      expect(entry).toBeNull();
    });

    it("should list all paths", async () => {
      const archive = new TarArchive();
      archive.add("x.txt", "X");
      archive.add("y.txt", "Y");
      archive.addDirectory("dir");
      const bytes = await archive.bytes();

      const { TarReader } = await import("@archive/tar");
      const reader = new TarReader(bytes);

      const paths = await reader.list();
      expect(paths).toEqual(["x.txt", "y.txt", "dir/"]);
    });

    it("should get bytes by path", async () => {
      const archive = new TarArchive();
      archive.add("data.bin", new Uint8Array([1, 2, 3]));
      const bytes = await archive.bytes();

      const { TarReader } = await import("@archive/tar");
      const reader = new TarReader(bytes);

      const data = await reader.bytes("data.bin");
      expect(data).toEqual(new Uint8Array([1, 2, 3]));
    });
  });
});

// ============================================================================
// Test unified API with format switching via zip()/unzip()
// ============================================================================
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";

describe("TAR via unified API (format switching)", () => {
  it("should create TAR archive using zip({ format: 'tar' })", async () => {
    const archive = zip({ format: "tar" });
    archive.add("hello.txt", "Hello from TAR!");
    archive.add("world.txt", "World content");

    const bytes = await archive.bytes();

    // Parse and verify
    const entries = parseTar(bytes);
    expect(entries.length).toBe(2);
    expect(entries[0].info.path).toBe("hello.txt");
    expect(await entries[0].text()).toBe("Hello from TAR!");
  });

  it("should read TAR archive using unzip(data, { format: 'tar' })", async () => {
    // Create TAR archive
    const archive = new TarArchive();
    archive.add("file1.txt", "Content 1");
    archive.add("file2.txt", "Content 2");
    const bytes = await archive.bytes();

    // Read using unified API
    const reader = unzip(bytes, { format: "tar" });
    const entryPaths: string[] = [];
    for await (const entry of reader.entries()) {
      entryPaths.push(entry.path);
    }

    expect(entryPaths).toEqual(["file1.txt", "file2.txt"]);
  });

  it("should have compatible API between ZIP and TAR", async () => {
    // Both should support the same method chain
    const zipArchive = zip();
    const tarArchive = zip({ format: "tar" });

    // Same API for adding entries
    zipArchive.add("test.txt", "content");
    tarArchive.add("test.txt", "content");

    // Same API for getting bytes
    const zipBytes = await zipArchive.bytes();
    const tarBytes = await tarArchive.bytes();

    expect(zipBytes).toBeInstanceOf(Uint8Array);
    expect(tarBytes).toBeInstanceOf(Uint8Array);

    // Different formats should produce different bytes
    expect(zipBytes).not.toEqual(tarBytes);
  });
});
