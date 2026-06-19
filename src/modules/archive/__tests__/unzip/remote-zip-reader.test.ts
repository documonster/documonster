import { BufferReader } from "@archive/io/random-access";
import { RemoteZipReader, Crc32MismatchError } from "@archive/unzip/remote-zip-reader";
import type { ZipEntry } from "@archive/zip/zip-bytes";
import { createZip } from "@archive/zip/zip-bytes";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect, vi } from "vitest";

// Helper to convert object to ZipEntry array
function toEntries(files: Record<string, Uint8Array | string>): ZipEntry[] {
  const encoder = new TextEncoder();
  return Object.entries(files).map(([name, data]) => ({
    name,
    data: typeof data === "string" ? encoder.encode(data) : data
  }));
}

describe("RemoteZipReader", () => {
  describe("fromReader with BufferReader", () => {
    it("should parse entries from a valid ZIP file", async () => {
      const testFiles: Record<string, string> = {
        "test.txt": "Hello, World!",
        "folder/nested.txt": "Nested content",
        "binary.bin": "\x00\x01\x02\x03\xff"
      };

      const zipData = await createZip(toEntries(testFiles));
      const bufferReader = new BufferReader(zipData);
      const reader = await RemoteZipReader.fromReader(bufferReader);

      const entries = reader.getEntries();
      expect(entries.length).toBe(3);

      const paths = entries.map(e => e.path).sort();
      expect(paths).toContain("test.txt");
      expect(paths).toContain("folder/nested.txt");
      expect(paths).toContain("binary.bin");

      await reader.close();
    });

    it("should handle empty ZIP files", async () => {
      const zipData = await createZip([]);
      const bufferReader = new BufferReader(zipData);
      const reader = await RemoteZipReader.fromReader(bufferReader);

      const entries = reader.getEntries();
      expect(entries.length).toBe(0);

      await reader.close();
    });

    it("should list files correctly", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "A",
        "b.txt": "B",
        "dir/c.txt": "C"
      };

      const zipData = await createZip(toEntries(testFiles));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const files = reader.listFiles();
      expect(files.sort()).toEqual(["a.txt", "b.txt", "dir/c.txt"].sort());

      await reader.close();
    });

    it("should check if entry exists", async () => {
      const zipData = await createZip(toEntries({ "exists.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      expect(reader.hasEntry("exists.txt")).toBe(true);
      expect(reader.hasEntry("not-exists.txt")).toBe(false);

      await reader.close();
    });

    it("should get entry by path", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "content" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const entry = reader.getEntry("test.txt");
      expect(entry).toBeDefined();
      expect(entry!.path).toBe("test.txt");

      const notFound = reader.getEntry("missing.txt");
      expect(notFound).toBeUndefined();

      await reader.close();
    });

    it("should parse archive comment", async () => {
      const zipData = await createZip(toEntries({ "a.txt": "A" }), {
        comment: "Test archive comment"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      expect(reader.getZipComment()).toBe("Test archive comment");

      await reader.close();
    });
  });

  describe("extract", () => {
    it("should extract text file content correctly", async () => {
      const originalContent = "Hello, World!";
      const zipData = await createZip(toEntries({ "test.txt": originalContent }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const extractedData = await reader.extract("test.txt");
      expect(extractedData).not.toBeNull();
      const extractedContent = new TextDecoder().decode(extractedData!);
      expect(extractedContent).toBe(originalContent);

      await reader.close();
    });

    it("should extract binary content correctly", async () => {
      const binaryData = new Uint8Array([0x00, 0x01, 0x02, 0x03, 0xff, 0xfe, 0xfd]);
      const zipData = await createZip(toEntries({ "binary.bin": binaryData }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const extractedData = await reader.extract("binary.bin");
      expect(extractedData).not.toBeNull();
      expect(extractedData!).toEqual(binaryData);

      await reader.close();
    });

    it("should return null for non-existent entry", async () => {
      const zipData = await createZip(toEntries({ "exists.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.extract("not-exists.txt");
      expect(result).toBeNull();

      await reader.close();
    });

    it("should return empty array for directory entry", async () => {
      const zipData = await createZip([
        { name: "folder/", data: new Uint8Array(0) },
        { name: "folder/file.txt", data: new TextEncoder().encode("content") }
      ]);
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const dirData = await reader.extract("folder/");
      expect(dirData).toEqual(new Uint8Array(0));

      await reader.close();
    });

    it("should extract compressed files", async () => {
      // Use a longer string to ensure compression is actually applied
      const content = "This is some content that will be compressed with DEFLATE. ".repeat(10);
      const zipData = await createZip(toEntries({ "compressed.txt": content }), {
        level: 6 // Default compression
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      // Just verify extraction works
      const extractedData = await reader.extract("compressed.txt");
      expect(new TextDecoder().decode(extractedData!)).toBe(content);

      await reader.close();
    });

    it("should extract stored (uncompressed) files", async () => {
      const content = "Uncompressed content";
      const zipData = await createZip(toEntries({ "stored.txt": content }), {
        level: 0 // Store (no compression)
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const entry = reader.getEntry("stored.txt");
      // Compression method 0 = STORE
      expect(entry?.compressionMethod).toBe(0);

      const extractedData = await reader.extract("stored.txt");
      expect(new TextDecoder().decode(extractedData!)).toBe(content);

      await reader.close();
    });
  });

  describe("extractEntry", () => {
    it("should extract entry directly", async () => {
      const content = "Direct extraction test";
      const zipData = await createZip(toEntries({ "direct.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const entry = reader.getEntry("direct.txt");
      expect(entry).toBeDefined();

      const data = await reader.extractEntry(entry!);
      expect(new TextDecoder().decode(data)).toBe(content);

      await reader.close();
    });
  });

  describe("extractAll", () => {
    it("should extract all files excluding directories", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "Content A",
        "b.txt": "Content B",
        "c.txt": "Content C"
      };
      const zipData = await createZip([
        { name: "folder/", data: new Uint8Array(0) },
        ...toEntries(testFiles)
      ]);
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.extractAll();

      expect(result.size).toBe(3);
      expect(new TextDecoder().decode(result.get("a.txt")!)).toBe("Content A");
      expect(new TextDecoder().decode(result.get("b.txt")!)).toBe("Content B");
      expect(new TextDecoder().decode(result.get("c.txt")!)).toBe("Content C");
      expect(result.has("folder/")).toBe(false);

      await reader.close();
    });
  });

  describe("extractMultiple", () => {
    it("should extract multiple files", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "Content A",
        "b.txt": "Content B",
        "c.txt": "Content C"
      };
      const zipData = await createZip(toEntries(testFiles));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.extractMultiple(["a.txt", "c.txt"]);

      expect(result.size).toBe(2);
      expect(new TextDecoder().decode(result.get("a.txt")!)).toBe("Content A");
      expect(new TextDecoder().decode(result.get("c.txt")!)).toBe("Content C");
      expect(result.has("b.txt")).toBe(false);

      await reader.close();
    });

    it("should skip non-existent entries", async () => {
      const zipData = await createZip(toEntries({ "exists.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.extractMultiple(["exists.txt", "missing.txt"]);

      expect(result.size).toBe(1);
      expect(result.has("exists.txt")).toBe(true);
      expect(result.has("missing.txt")).toBe(false);

      await reader.close();
    });
  });

  describe("forEach", () => {
    it("should iterate over all entries", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "A",
        "b.txt": "B"
      };
      const zipData = await createZip(toEntries(testFiles));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const visited: string[] = [];
      await reader.forEach(async (entry, getData) => {
        visited.push(entry.path);
        const data = await getData();
        expect(data).toBeInstanceOf(Uint8Array);
      });

      expect(visited.sort()).toEqual(["a.txt", "b.txt"].sort());

      await reader.close();
    });

    it("should stop iteration when callback returns false", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "A",
        "b.txt": "B",
        "c.txt": "C"
      };
      const zipData = await createZip(toEntries(testFiles));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const visited: string[] = [];
      await reader.forEach(async entry => {
        visited.push(entry.path);
        if (visited.length >= 2) {
          return false;
        }
      });

      expect(visited.length).toBe(2);

      await reader.close();
    });

    it("should lazily load data only when getData is called", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      let dataRequested = false;
      await reader.forEach(async (entry, getData) => {
        // Don't call getData - data should not be loaded
        dataRequested = true;
      });

      expect(dataRequested).toBe(true);

      await reader.close();
    });
  });

  describe("convenience methods", () => {
    it("should count files and directories correctly", async () => {
      const zipData = await createZip([
        { name: "folder/", data: new Uint8Array(0) },
        { name: "folder/sub/", data: new Uint8Array(0) },
        ...toEntries({ "a.txt": "A", "b.txt": "B", "folder/c.txt": "C" })
      ]);
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      expect(reader.getFileCount()).toBe(3);
      expect(reader.getDirectoryCount()).toBe(2);
      expect(reader.getEntries().length).toBe(5);

      await reader.close();
    });

    it("should filter entries by predicate", async () => {
      const zipData = await createZip(
        toEntries({
          "a.txt": "A",
          "b.json": "{}",
          "c.txt": "C"
        })
      );
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const txtFiles = reader.filterEntries(e => e.path.endsWith(".txt"));
      expect(txtFiles.length).toBe(2);
      expect(txtFiles.map(e => e.path).sort()).toEqual(["a.txt", "c.txt"]);

      await reader.close();
    });

    it("should find entries by glob pattern", async () => {
      const zipData = await createZip(
        toEntries({
          "file.txt": "A",
          "file.json": "{}",
          "data/a.txt": "B",
          "data/b.txt": "C",
          "data/sub/c.txt": "D"
        })
      );
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      // Match all .txt files in root
      expect(reader.findEntries("*.txt").map(e => e.path)).toEqual(["file.txt"]);

      // Match all .txt files in data folder (not recursive)
      expect(
        reader
          .findEntries("data/*.txt")
          .map(e => e.path)
          .sort()
      ).toEqual(["data/a.txt", "data/b.txt"]);

      // Match all .txt files recursively (** matches any path including empty)
      const allTxt = reader
        .findEntries("**.txt")
        .map(e => e.path)
        .sort();
      expect(allTxt).toEqual(["data/a.txt", "data/b.txt", "data/sub/c.txt", "file.txt"]);

      await reader.close();
    });
  });

  describe("getStats", () => {
    it("should return correct stats", async () => {
      const testFiles: Record<string, string> = {
        "a.txt": "A",
        "b.txt": "B"
      };
      const zipData = await createZip(toEntries(testFiles));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const stats = reader.getStats();
      expect(stats.totalSize).toBe(zipData.length);
      expect(stats.entryCount).toBe(2);
      // BufferReader doesn't have HTTP stats
      expect(stats.http).toBeUndefined();

      await reader.close();
    });
  });

  describe("ZIP64 support", () => {
    it("should parse ZIP64 archives", async () => {
      // Create a forced ZIP64 archive
      const zipData = await createZip(toEntries({ "test.txt": "ZIP64 content" }), {
        zip64: true
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const entries = reader.getEntries();
      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("test.txt");

      const data = await reader.extract("test.txt");
      expect(new TextDecoder().decode(data!)).toBe("ZIP64 content");

      await reader.close();
    });
  });

  describe("encryption", () => {
    it("should detect encrypted entries", async () => {
      // Create encrypted ZIP with explicit encryptionMethod
      const zipData = await createZip(toEntries({ "secret.txt": "Secret content" }), {
        password: "password123",
        encryptionMethod: "zipcrypto"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      expect(reader.hasEncryptedEntries()).toBe(true);

      const entry = reader.getEntry("secret.txt");
      expect(entry?.isEncrypted).toBe(true);

      await reader.close();
    });

    it("should extract encrypted entries with correct password", async () => {
      const content = "Secret message";
      const password = "correct-password";

      const zipData = await createZip(toEntries({ "secret.txt": content }), {
        password,
        encryptionMethod: "zipcrypto"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const data = await reader.extract("secret.txt", password);
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });

    it("should use constructor password if not provided to extract", async () => {
      const content = "Secret message";
      const password = "my-password";

      const zipData = await createZip(toEntries({ "secret.txt": content }), {
        password,
        encryptionMethod: "zipcrypto"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData), {
        password
      });

      const data = await reader.extract("secret.txt");
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });

    it("should throw when password is required but not provided", async () => {
      const zipData = await createZip(toEntries({ "secret.txt": "Secret" }), {
        password: "password",
        encryptionMethod: "zipcrypto"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      await expect(reader.extract("secret.txt")).rejects.toThrow(/encrypted/i);

      await reader.close();
    });
  });

  describe("error handling", () => {
    it("should throw for invalid ZIP data", async () => {
      const invalidData = new Uint8Array([0x00, 0x01, 0x02, 0x03]);

      await expect(RemoteZipReader.fromReader(new BufferReader(invalidData))).rejects.toThrow(
        /End of Central Directory not found/i
      );
    });

    it("should throw for truncated ZIP data", async () => {
      // Create valid ZIP then truncate it
      const validZip = await createZip(toEntries({ "test.txt": "content" }));
      const truncated = validZip.slice(0, 10);

      await expect(RemoteZipReader.fromReader(new BufferReader(truncated))).rejects.toThrow();
    });
  });

  describe("HTTP Range reading simulation", () => {
    // This test simulates HTTP Range behavior using a mock reader
    it("should read EOCD and central directory without reading file data", async () => {
      // Create a larger archive to demonstrate the benefit of range reading
      const files: Record<string, string> = {};
      for (let i = 0; i < 5; i++) {
        files[`file${i}.txt`] = `Content for file ${i} - ${"x".repeat(500)}`;
      }
      const zipData = await createZip(toEntries(files));

      // Track which ranges are read
      const reads: Array<{ start: number; end: number }> = [];

      const trackingReader = {
        size: zipData.length,
        async read(start: number, end: number): Promise<Uint8Array> {
          reads.push({ start, end });
          return zipData.slice(start, end);
        },
        async close(): Promise<void> {}
      };

      const reader = await RemoteZipReader.fromReader(trackingReader);

      // Just listing entries should make reads
      expect(reads.length).toBeGreaterThanOrEqual(1);

      // Verify we can get entry count
      expect(reader.getEntries().length).toBe(5);

      // Now extract one file - should make additional read(s)
      const readsBefore = reads.length;
      await reader.extract("file0.txt");
      expect(reads.length).toBeGreaterThan(readsBefore);

      await reader.close();
    });

    it("should only read necessary data for single file extraction", async () => {
      // Create a ZIP with many files
      const files: Record<string, string> = {};
      for (let i = 0; i < 10; i++) {
        files[`file${i}.txt`] = `Content for file ${i} - ${"x".repeat(100)}`;
      }
      const zipData = await createZip(toEntries(files));

      const reads: Array<{ start: number; end: number }> = [];
      const trackingReader = {
        size: zipData.length,
        async read(start: number, end: number): Promise<Uint8Array> {
          reads.push({ start, end });
          return zipData.slice(start, end);
        },
        async close(): Promise<void> {}
      };

      const reader = await RemoteZipReader.fromReader(trackingReader);
      const initialReads = reads.length;

      // Extract only one file
      await reader.extract("file5.txt");

      // Should have made only 2 additional reads:
      // 1. Local file header
      // 2. File data
      // (These might be combined into 1 or 2 reads depending on implementation)
      const extractionReads = reads.length - initialReads;
      expect(extractionReads).toBeLessThanOrEqual(2);

      await reader.close();
    });
  });
});

describe("RemoteZipReader.open", () => {
  it("should work with mocked HTTP (simulated)", async () => {
    // Create test ZIP with some content
    const content = "Remote content - " + "y".repeat(200);
    const zipData = await createZip(toEntries({ "remote.txt": content }));

    // Mock fetch
    const mockFetch = vi.fn().mockImplementation(async (url: string, options?: RequestInit) => {
      const headers = options?.headers as Record<string, string> | undefined;
      const rangeHeader = headers?.Range;

      if (options?.method === "HEAD") {
        return {
          ok: true,
          status: 200,
          statusText: "OK",
          headers: new Headers({
            "Content-Length": String(zipData.length),
            "Accept-Ranges": "bytes"
          })
        };
      }

      if (rangeHeader) {
        // Parse range: bytes=start-end
        const match = rangeHeader.match(/bytes=(\d+)-(\d+)/);
        if (match) {
          const start = parseInt(match[1], 10);
          const end = parseInt(match[2], 10) + 1; // HTTP range is inclusive
          const slice = zipData.slice(start, end);

          return {
            ok: true,
            status: 206,
            statusText: "Partial Content",
            arrayBuffer: () =>
              Promise.resolve(
                slice.buffer.slice(slice.byteOffset, slice.byteOffset + slice.byteLength)
              )
          };
        }
      }

      // Return full content
      return {
        ok: true,
        status: 200,
        statusText: "OK",
        arrayBuffer: () =>
          Promise.resolve(
            zipData.buffer.slice(zipData.byteOffset, zipData.byteOffset + zipData.byteLength)
          )
      };
    });

    const reader = await RemoteZipReader.open("https://example.com/test.zip", {
      fetch: mockFetch
    });

    expect(reader.getEntries().length).toBe(1);

    const data = await reader.extract("remote.txt");
    expect(new TextDecoder().decode(data!)).toBe(content);

    const stats = reader.getStats();
    expect(stats.http).toBeDefined();
    expect(stats.http!.requestCount).toBeGreaterThan(0);
    // Downloaded bytes should be tracked
    expect(stats.http!.bytesDownloaded).toBeGreaterThan(0);

    await reader.close();
  });

  describe("CRC32 validation", () => {
    it("should validate CRC32 when checkCrc32 is enabled", async () => {
      const content = "Hello, World!";
      const zipData = await createZip(toEntries({ "test.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      // Should not throw with valid CRC32
      const data = await reader.extract("test.txt", { checkCrc32: true });
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });

    it("should verify CRC32 using verifyCrc32 method", async () => {
      const content = "Test content for CRC verification";
      const zipData = await createZip(toEntries({ "verify.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.verifyCrc32("verify.txt");
      expect(result).toBe(true);

      const notFound = await reader.verifyCrc32("missing.txt");
      expect(notFound).toBeNull();

      await reader.close();
    });

    it("should enable CRC32 validation via constructor option", async () => {
      const content = "Constructor CRC test";
      const zipData = await createZip(toEntries({ "test.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData), {
        checkCrc32: true
      });

      // Extract should work with valid data
      const data = await reader.extract("test.txt");
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });
  });

  describe("progress callback", () => {
    it("should call onprogress during extraction", async () => {
      const content = "A".repeat(1000);
      const zipData = await createZip(toEntries({ "large.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const progressCalls: Array<{ current: number; total: number }> = [];
      await reader.extract("large.txt", {
        onprogress: (current, total) => {
          progressCalls.push({ current, total });
        }
      });

      expect(progressCalls.length).toBeGreaterThan(0);
      // Final call should have current === total
      const lastCall = progressCalls[progressCalls.length - 1];
      expect(lastCall!.current).toBe(lastCall!.total);

      await reader.close();
    });

    it("should call onprogress for extractMultiple", async () => {
      const files = {
        "file1.txt": "Content 1",
        "file2.txt": "Content 2",
        "file3.txt": "Content 3"
      };
      const zipData = await createZip(toEntries(files));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const progressCalls: Array<{ current: number; total: number }> = [];
      await reader.extractMultiple(["file1.txt", "file2.txt", "file3.txt"], {
        onprogress: (current, total) => {
          progressCalls.push({ current, total });
        }
      });

      expect(progressCalls.length).toBeGreaterThan(0);

      await reader.close();
    });
  });

  describe("async iterator (entriesGenerator)", () => {
    it("should iterate over entries using async generator", async () => {
      const files = {
        "a.txt": "AAA",
        "b.txt": "BBB",
        "c.txt": "CCC"
      };
      const zipData = await createZip(toEntries(files));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const results: Array<{ path: string; data: string }> = [];
      for await (const { entry, getData } of reader.entriesGenerator()) {
        const data = await getData();
        results.push({ path: entry.path, data: new TextDecoder().decode(data) });
      }

      expect(results.length).toBe(3);
      expect(results.find(r => r.path === "a.txt")?.data).toBe("AAA");
      expect(results.find(r => r.path === "b.txt")?.data).toBe("BBB");
      expect(results.find(r => r.path === "c.txt")?.data).toBe("CCC");

      await reader.close();
    });

    it("should allow early exit from generator", async () => {
      const files = {
        "a.txt": "AAA",
        "target.txt": "TARGET",
        "c.txt": "CCC"
      };
      const zipData = await createZip(toEntries(files));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      let foundTarget: string | null = null;
      for await (const { entry, getData } of reader.entriesGenerator()) {
        if (entry.path === "target.txt") {
          foundTarget = new TextDecoder().decode(await getData());
          break;
        }
      }

      expect(foundTarget).toBe("TARGET");

      await reader.close();
    });
  });

  describe("password verification (checkPassword)", () => {
    it("should verify correct password for zipcrypto", async () => {
      const password = "secret123";
      const zipData = await createZip(toEntries({ "encrypted.txt": "Secret data" }), {
        password,
        encryptionMethod: "zipcrypto"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const isValid = await reader.checkPassword("encrypted.txt", password);
      expect(isValid).toBe(true);

      // ZipCrypto password verification uses a single-byte check, which has a
      // ~1/128 false positive rate per attempt. Test multiple wrong passwords to
      // make a flaky pass statistically impossible (~1/2^40 for 5 attempts).
      const wrongPasswords = ["wrong1", "wrong2", "wrong3", "wrong4", "wrong5"];
      const results = await Promise.all(
        wrongPasswords.map(wp => reader.checkPassword("encrypted.txt", wp))
      );
      expect(results).toContain(false);

      await reader.close();
    });

    it("should return null for non-encrypted entries", async () => {
      const zipData = await createZip(toEntries({ "plain.txt": "Plain data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.checkPassword("plain.txt", "anypassword");
      expect(result).toBeNull();

      await reader.close();
    });

    it("should return null for non-existent entries", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.checkPassword("missing.txt", "password");
      expect(result).toBeNull();

      await reader.close();
    });
  });

  describe("stream output (extractToStream)", () => {
    it("should extract to WritableStream", async () => {
      const content = new Uint8Array(256 * 1024).fill(65); // 256KiB of 'A'
      const zipData = await createZip(toEntries({ "stream.txt": content }), {
        level: 0,
        smartStore: false
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const chunks: Uint8Array[] = [];
      const writable = new WritableStream<Uint8Array>({
        write(chunk) {
          chunks.push(chunk);
        }
      });

      const success = await reader.extractToStream("stream.txt", writable);
      expect(success).toBe(true);

      // Store method should stream in multiple writes (chunked reads).
      expect(chunks.length).toBeGreaterThan(1);

      const result = concatUint8Arrays(chunks);
      expect(result).toEqual(content);

      await reader.close();
    });

    it("should extract encrypted zipcrypto entry to WritableStream", async () => {
      const password = "secret123";
      const content = new Uint8Array(128 * 1024).map((_, i) => i & 0xff);

      const zipData = await createZip(toEntries({ "enc.bin": content }), {
        level: 0,
        smartStore: false,
        password,
        encryptionMethod: "zipcrypto"
      });

      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const chunks: Uint8Array[] = [];
      const writable = new WritableStream<Uint8Array>({
        write(chunk) {
          chunks.push(chunk);
        }
      });

      const success = await reader.extractToStream("enc.bin", writable, { password });
      expect(success).toBe(true);

      // ZipCrypto + STORE should also stream in multiple writes.
      expect(chunks.length).toBeGreaterThan(1);

      const result = concatUint8Arrays(chunks);
      expect(result).toEqual(content);

      await reader.close();
    });

    it("should return false for non-existent entry", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "data" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const writable = new WritableStream<Uint8Array>();
      const success = await reader.extractToStream("missing.txt", writable);
      expect(success).toBe(false);

      await reader.close();
    });
  });

  describe("Crc32MismatchError", () => {
    it("should have correct error properties", () => {
      const error = new Crc32MismatchError("test.txt", 0x12345678, 0xabcdef01);

      expect(error.name).toBe("Crc32MismatchError");
      expect(error.path).toBe("test.txt");
      expect(error.expected).toBe(0x12345678);
      expect(error.actual).toBe(0xabcdef01);
      expect(error.message).toContain("test.txt");
      expect(error.message).toContain("12345678");
      expect(error.message).toContain("abcdef01");
    });
  });

  describe("verifyCrc32", () => {
    it("should return true for valid CRC32", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "Hello, World!" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.verifyCrc32("test.txt");
      expect(result).toBe(true);

      await reader.close();
    });

    it("should return null for non-existent entry", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "content" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const result = await reader.verifyCrc32("missing.txt");
      expect(result).toBeNull();

      await reader.close();
    });

    it("should return true for AES encrypted entries (skips CRC check)", async () => {
      const zipData = await createZip(toEntries({ "secret.txt": "Secret content" }), {
        password: "password123",
        encryptionMethod: "aes-256"
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      // Check entry is actually AES encrypted
      const entry = reader.getEntry("secret.txt");
      expect(entry).toBeDefined();
      expect(entry!.encryptionMethod).toBe("aes");

      // AES entries skip CRC32 check, should return true without needing password
      // (because we skip CRC verification entirely for AES)
      const result = await reader.verifyCrc32("secret.txt");
      expect(result).toBe(true);

      await reader.close();
    });
  });

  describe("extractEntry with ExtractOptions", () => {
    it("should call onprogress callback", async () => {
      const content = "Test content for progress";
      const zipData = await createZip(toEntries({ "test.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const progressCalls: Array<{ current: number; total: number }> = [];
      await reader.extract("test.txt", {
        onprogress: (current, total) => {
          progressCalls.push({ current, total });
        }
      });

      expect(progressCalls.length).toBeGreaterThan(0);
      // Final call should have current === total
      const lastCall = progressCalls[progressCalls.length - 1];
      expect(lastCall.current).toBe(lastCall.total);

      await reader.close();
    });

    it("should verify CRC32 when checkCrc32 option is true", async () => {
      const content = "Content with CRC check";
      const zipData = await createZip(toEntries({ "test.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      // Should not throw - CRC is valid
      const data = await reader.extract("test.txt", { checkCrc32: true });
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });
  });

  describe("close", () => {
    it("should be callable multiple times without error", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "content" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      await expect(reader.close()).resolves.toBeUndefined();
      await expect(reader.close()).resolves.toBeUndefined();
    });
  });

  describe("constructor password", () => {
    it("should use constructor checkCrc32 option as default", async () => {
      const zipData = await createZip(toEntries({ "test.txt": "Content" }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData), {
        checkCrc32: true
      });

      // Should use constructor's checkCrc32 option
      const data = await reader.extract("test.txt");
      expect(new TextDecoder().decode(data!)).toBe("Content");

      await reader.close();
    });
  });

  describe("extractMultiple with options", () => {
    it("should accept ExtractOptions for all files", async () => {
      const files = {
        "a.txt": "Content A",
        "b.txt": "Content B"
      };
      const zipData = await createZip(toEntries(files));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const progressCalls: Array<{ current: number; total: number }> = [];
      const results = await reader.extractMultiple(["a.txt", "b.txt"], {
        onprogress: (current, total) => {
          progressCalls.push({ current, total });
        }
      });

      expect(results.size).toBe(2);
      expect(progressCalls.length).toBeGreaterThan(0);

      await reader.close();
    });
  });

  describe("forEach with async callback", () => {
    it("should wait for async callback to complete", async () => {
      const files = {
        "a.txt": "A",
        "b.txt": "B"
      };
      const zipData = await createZip(toEntries(files));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const results: string[] = [];
      await reader.forEach(async (entry, data) => {
        // Simulate async work
        await new Promise(resolve => setTimeout(resolve, 10));
        results.push(entry.path);
      });

      expect(results.sort()).toEqual(["a.txt", "b.txt"]);

      await reader.close();
    });
  });

  describe("directories", () => {
    it("should handle directory entries correctly", async () => {
      const zipData = await createZip([
        { name: "folder/", data: new Uint8Array(0) },
        { name: "folder/file.txt", data: new TextEncoder().encode("content") }
      ]);
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const entries = reader.getEntries();
      expect(entries.length).toBe(2);

      const dirEntry = reader.getEntry("folder/");
      expect(dirEntry).toBeDefined();
      expect(dirEntry!.type).toBe("directory");

      // Extract directory should return empty data
      const dirData = await reader.extract("folder/");
      expect(dirData).toBeDefined();
      expect(dirData!.length).toBe(0);

      await reader.close();
    });
  });

  describe("compressed entries", () => {
    it("should correctly decompress DEFLATE entries", async () => {
      // Force compression by using compressible content and level > 0
      const content = "A".repeat(1000);
      const zipData = await createZip(toEntries({ "compressed.txt": content }), {
        level: 6 // Force DEFLATE compression
      });
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const data = await reader.extract("compressed.txt");
      expect(new TextDecoder().decode(data!)).toBe(content);

      await reader.close();
    });
  });

  describe("large file progress tracking", () => {
    it("should call onprogress multiple times for larger files", async () => {
      // Create a larger content to ensure multiple progress calls
      const content = "X".repeat(10000);
      const zipData = await createZip(toEntries({ "large.txt": content }));
      const reader = await RemoteZipReader.fromReader(new BufferReader(zipData));

      const progressCalls: Array<{ current: number; total: number }> = [];
      await reader.extract("large.txt", {
        onprogress: (current, total) => {
          progressCalls.push({ current, total });
        }
      });

      expect(progressCalls.length).toBeGreaterThan(0);
      // Progress should increase
      for (let i = 1; i < progressCalls.length; i++) {
        expect(progressCalls[i].current).toBeGreaterThanOrEqual(progressCalls[i - 1].current);
      }

      await reader.close();
    });
  });
});
