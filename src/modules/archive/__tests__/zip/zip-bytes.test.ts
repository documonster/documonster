import { ZipParser } from "@archive/unzip/zip-parser";
import { parseZipExtraFields } from "@archive/zip-spec/zip-extra-fields";
import type { ZipEntry } from "@archive/zip/zip-bytes";
import { createZip, createZipSync } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

function parseEntries(zipData: Uint8Array) {
  const parser = new ZipParser(zipData);
  return { parser, entries: parser.getEntries() };
}

function extractFileSync(zipData: Uint8Array, fileName: string): Uint8Array | null {
  const parser = new ZipParser(zipData);
  return parser.extractSync(fileName);
}

describe("zip-bytes", () => {
  describe("createZip (async)", () => {
    it("should create a valid empty ZIP", async () => {
      const zip = await createZip([]);
      const { entries } = parseEntries(zip);

      expect(entries.length).toBe(0);
    });

    it("should create ZIP with single file", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "hello.txt", data: content }]);

      const { entries } = parseEntries(zip);
      expect(entries.length).toBe(1);
      expect(entries[0]!.path).toBe("hello.txt");
      expect(entries[0]!.uncompressedSize).toBe(content.length);

      // Extract and verify content
      const extracted = extractFileSync(zip, "hello.txt");
      expect(extracted).toEqual(content);
    });

    it("should create ZIP with multiple files", async () => {
      const file1 = new TextEncoder().encode("File 1 content");
      const file2 = new TextEncoder().encode("File 2 content with more data");
      const file3 = new TextEncoder().encode("File 3");

      const zip = await createZip([
        { name: "file1.txt", data: file1 },
        { name: "file2.txt", data: file2 },
        { name: "file3.txt", data: file3 }
      ]);

      const { entries } = parseEntries(zip);
      expect(entries.length).toBe(3);

      // Verify all files can be extracted
      expect(extractFileSync(zip, "file1.txt")).toEqual(file1);
      expect(extractFileSync(zip, "file2.txt")).toEqual(file2);
      expect(extractFileSync(zip, "file3.txt")).toEqual(file3);
    });

    it("should create ZIP with nested directories", async () => {
      const content = new TextEncoder().encode("Nested file content");
      const zip = await createZip([{ name: "folder/subfolder/deep/file.txt", data: content }]);

      const { entries } = parseEntries(zip);
      expect(entries[0]!.path).toBe("folder/subfolder/deep/file.txt");

      const extracted = extractFileSync(zip, "folder/subfolder/deep/file.txt");
      expect(extracted).toEqual(content);
    });

    it("should handle empty file", async () => {
      const empty = new Uint8Array(0);
      const zip = await createZip([{ name: "empty.txt", data: empty }]);

      const { entries } = parseEntries(zip);
      expect(entries[0]!.uncompressedSize).toBe(0);
      expect(entries[0]!.compressionMethod).toBe(0); // STORE for empty

      const extracted = extractFileSync(zip, "empty.txt");
      expect(extracted).toEqual(empty);
    });

    it("should handle large file (1MB)", async () => {
      const large = new Uint8Array(1024 * 1024);
      // Highly compressible payload.
      large.fill(0);

      const zip = await createZip([{ name: "large.bin", data: large }], { level: 1 });

      const { entries } = parseEntries(zip);
      expect(entries[0]!.uncompressedSize).toBe(large.length);
      // Compressed size should be smaller
      expect(entries[0]!.compressedSize).toBeLessThan(large.length);

      const extracted = extractFileSync(zip, "large.bin");
      expect(extracted).toEqual(large);
    });

    it("should smart STORE incompressible data", async () => {
      const large = new Uint8Array(1024 * 1024);
      for (let i = 0; i < large.length; i++) {
        large[i] = i % 256;
      }

      const zip = await createZip([{ name: "incompressible.bin", data: large }], { level: 6 });

      const { entries } = parseEntries(zip);
      expect(entries[0]!.compressionMethod).toBe(0); // STORE
      expect(entries[0]!.compressedSize).toBe(large.length);

      const extracted = extractFileSync(zip, "incompressible.bin");
      expect(extracted).toEqual(large);
    });

    it("should use STORE method for level 0", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "uncompressed.txt", data: content }], { level: 0 });

      const { entries } = parseEntries(zip);
      expect(entries[0]!.compressionMethod).toBe(0); // STORE
      expect(entries[0]!.compressedSize).toBe(content.length);
    });

    it("should honor per-entry level override", () => {
      const dataA = new TextEncoder().encode("aaaaabbbbbcccccdddddeeeee".repeat(100));
      const dataB = new Uint8Array(64 * 1024);
      for (let i = 0; i < dataB.length; i++) {
        dataB[i] = i & 0xff;
      }

      const zipData = createZipSync(
        [
          { name: "a.txt", data: dataA, level: 6 },
          { name: "b.bin", data: dataB, level: 0 }
        ],
        { level: 6 }
      );

      const { entries } = parseEntries(zipData);
      const a = entries.find(e => e.path === "a.txt")!;
      const b = entries.find(e => e.path === "b.bin")!;

      expect(a.compressionMethod).toBe(8);
      expect(b.compressionMethod).toBe(0);
    });
    it("should use DEFLATE method for level > 0", async () => {
      const content = new TextEncoder().encode("Hello, World!".repeat(100));
      const zip = await createZip([{ name: "compressed.txt", data: content }], { level: 6 });

      const { entries } = parseEntries(zip);
      expect(entries[0]!.compressionMethod).toBe(8); // DEFLATE
      expect(entries[0]!.compressedSize).toBeLessThan(content.length);
    });

    it("should handle unicode filenames", async () => {
      const content = new TextEncoder().encode("Unicode content");
      const zip = await createZip(
        [
          { name: "文件.txt", data: content },
          { name: "файл.txt", data: content },
          { name: "αρχείο.txt", data: content }
        ],
        { noSort: true }
      );

      const { entries } = parseEntries(zip);
      expect(entries[0]!.path).toBe("文件.txt");
      expect(entries[1]!.path).toBe("файл.txt");
      expect(entries[2]!.path).toBe("αρχείο.txt");
    });

    it("should handle filenames with spaces", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip(
        [
          { name: "file with spaces.txt", data: content },
          { name: "folder name/file name.txt", data: content }
        ],
        { noSort: true }
      );

      const { entries } = parseEntries(zip);
      expect(entries[0]!.path).toBe("file with spaces.txt");
      expect(entries[1]!.path).toBe("folder name/file name.txt");
    });

    it("should handle binary data", async () => {
      const binary = new Uint8Array([0x00, 0x01, 0xff, 0xfe, 0x7f, 0x80]);
      const zip = await createZip([{ name: "binary.bin", data: binary }]);

      const extracted = extractFileSync(zip, "binary.bin");
      expect(extracted).toEqual(binary);
    });

    it("should calculate correct CRC32", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "test.txt", data: content }]);

      // CRC32 of "Hello, World!" is 0xec4ac3d0
      const { entries } = parseEntries(zip);
      expect(entries[0]!.crc32).toBe(0xec4ac3d0);
    });

    it("should support file modification time", async () => {
      const content = new TextEncoder().encode("Timestamped content");
      const modTime = new Date(2023, 5, 15, 10, 30, 0); // June 15, 2023 10:30:00

      const zip = await createZip([{ name: "dated.txt", data: content, modTime }]);

      // Verify ZIP was created (detailed time verification would require more parsing)
      const { entries } = parseEntries(zip);
      expect(entries.length).toBe(1);
    });

    it("should support ZIP comment", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([{ name: "file.txt", data: content }], {
        comment: "This is a ZIP comment"
      });

      const { parser } = parseEntries(zip);
      expect(parser.getZipComment()).toBe("This is a ZIP comment");
    });

    it("should support file comments", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([
        { name: "file.txt", data: content, comment: "File comment here" }
      ]);

      const { entries } = parseEntries(zip);
      expect(entries[0]!.comment).toBe("File comment here");
    });

    it("should write unix permissions to external attributes", async () => {
      const content = new TextEncoder().encode("perm");
      const zip = await createZip([
        {
          name: "perm.txt",
          data: content,
          // Intentionally omit the file type bits to ensure writer fills them.
          mode: 0o644
        }
      ]);

      const parser = new ZipParser(zip);
      const entry = parser.getEntry("perm.txt");
      expect(entry).toBeDefined();
      expect(entry!.versionMadeBy).toBe((3 << 8) | 20);
      expect((entry!.externalAttributes >>> 16) & 0xffff).toBe(0o100644);
    });

    it("should mark directories in DOS attrs and write unix dir mode", async () => {
      const zip = await createZip([
        {
          name: "folder/",
          data: new Uint8Array(0),
          mode: 0o755
        }
      ]);

      const parser = new ZipParser(zip);
      const entry = parser.getEntry("folder/");
      expect(entry).toBeDefined();
      expect((entry!.externalAttributes & 0xff & 0x10) !== 0).toBe(true);
      expect((entry!.externalAttributes >>> 16) & 0xffff).toBe(0o040755);
    });

    it("should normalize paths when path options are provided", async () => {
      const content = new TextEncoder().encode("p");
      const zip = await createZip([{ name: "\\foo\\bar\\..\\baz.txt", data: content }], {
        path: { mode: "posix", prependSlash: true }
      });

      const { entries } = parseEntries(zip);
      expect(entries[0]!.path).toBe("/foo/baz.txt");

      const extracted = extractFileSync(zip, "/foo/baz.txt");
      expect(extracted).toEqual(content);
    });

    it("should reject traversal paths in safe mode", async () => {
      await expect(
        createZip([{ name: "../evil.txt", data: new TextEncoder().encode("x") }], {
          path: { mode: "safe" }
        })
      ).rejects.toThrow(/Unsafe ZIP path/);
    });

    it("should write NTFS timestamps when configured", async () => {
      const content = new TextEncoder().encode("t");
      const modTime = new Date(Date.UTC(2024, 0, 2, 3, 4, 5));

      const zip = await createZip([{ name: "t.txt", data: content, modTime }], {
        timestamps: "dos+utc+ntfs"
      });

      const parser = new ZipParser(zip);
      const entry = parser.getEntry("t.txt");
      expect(entry).toBeDefined();

      const extra = parseZipExtraFields(entry!.extraField ?? new Uint8Array(0), {
        uncompressedSize: entry!.uncompressedSize,
        compressedSize: entry!.compressedSize
      });
      expect(extra.mtimeUnixSeconds).toBe(Math.floor(modTime.getTime() / 1000));
      expect(extra.ntfsTimes).toBeDefined();
      const EPOCH_DIFF_100NS = 116444736000000000n;
      expect(extra.ntfsTimes!.mtime).toBe(BigInt(modTime.getTime()) * 10000n + EPOCH_DIFF_100NS);
    });
  });

  describe("createZipSync", () => {
    it("should create valid ZIP synchronously", () => {
      const content = new TextEncoder().encode("Sync content");
      const zip = createZipSync([{ name: "sync.txt", data: content }]);

      const { entries } = parseEntries(zip);
      expect(entries.length).toBe(1);
    });

    it("should reject traversal paths in safe mode (sync)", () => {
      expect(() =>
        createZipSync([{ name: "../evil.txt", data: new TextEncoder().encode("x") }], {
          path: { mode: "safe" }
        })
      ).toThrow(/Unsafe ZIP path/);
    });

    it("should produce same result as async for same input", async () => {
      const content = new TextEncoder().encode("Same content");
      const entries: ZipEntry[] = [{ name: "file.txt", data: content }];

      const asyncZip = await createZip(entries, { level: 6 });
      const syncZip = createZipSync(entries, { level: 6 });

      // Structure should be identical
      const asyncEntries = parseEntries(asyncZip).entries;
      const syncEntries = parseEntries(syncZip).entries;

      expect(asyncEntries.length).toBe(syncEntries.length);
    });

    it("should produce deterministic bytes in reproducible mode", () => {
      const content = new TextEncoder().encode("Deterministic content");
      const entries: ZipEntry[] = [{ name: "file.txt", data: content }];

      const zip1 = createZipSync(entries, { reproducible: true });
      const zip2 = createZipSync(entries, { reproducible: true });

      expect(zip1).toEqual(zip2);
    });
  });

  describe("edge cases", () => {
    it("should handle very long filenames", async () => {
      const longName = "a".repeat(200) + ".txt";
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([{ name: longName, data: content }]);

      const { entries } = parseEntries(zip);
      expect(entries[0]!.path).toBe(longName);
    });

    it("should handle special characters in filenames", async () => {
      const content = new TextEncoder().encode("Content");
      const specialNames = [
        "file-with-dashes.txt",
        "file_with_underscores.txt",
        "file.multiple.dots.txt",
        "file (with) parens.txt",
        "file [with] brackets.txt"
      ];

      const zip = await createZip(
        specialNames.map(name => ({ name, data: content })),
        {
          noSort: true
        }
      );

      const { entries } = parseEntries(zip);
      specialNames.forEach((name, i) => {
        expect(entries[i]!.path).toBe(name);
      });
    });

    it("should handle data with all bytes 0x00", async () => {
      const zeros = new Uint8Array(1000).fill(0);
      const zip = await createZip([{ name: "zeros.bin", data: zeros }]);

      const extracted = extractFileSync(zip, "zeros.bin");
      expect(extracted).toEqual(zeros);
    });

    it("should handle data with all bytes 0xFF", async () => {
      const ones = new Uint8Array(1000).fill(0xff);
      const zip = await createZip([{ name: "ones.bin", data: ones }]);

      const extracted = extractFileSync(zip, "ones.bin");
      expect(extracted).toEqual(ones);
    });

    it("should handle mixed empty and non-empty files", async () => {
      const zip = await createZip(
        [
          { name: "empty1.txt", data: new Uint8Array(0) },
          { name: "content.txt", data: new TextEncoder().encode("Has content") },
          { name: "empty2.txt", data: new Uint8Array(0) }
        ],
        { noSort: true }
      );

      const { entries } = parseEntries(zip);
      expect(entries[0]!.uncompressedSize).toBe(0);
      expect(entries[1]!.uncompressedSize).toBe(11);
      expect(entries[2]!.uncompressedSize).toBe(0);
    });

    it("should handle XML-like content (common in XLSX)", async () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>`;
      const content = new TextEncoder().encode(xml);
      const zip = await createZip([{ name: "xl/worksheets/sheet1.xml", data: content }]);

      const extracted = extractFileSync(zip, "xl/worksheets/sheet1.xml");
      expect(extracted).toEqual(content);
    });

    it("should handle many small files", async () => {
      const entries: ZipEntry[] = [];
      for (let i = 0; i < 100; i++) {
        entries.push({
          name: `file${i.toString().padStart(3, "0")}.txt`,
          data: new TextEncoder().encode(`Content of file ${i}`)
        });
      }

      const zip = await createZip(entries, { level: 1 });

      const { entries: parsed } = parseEntries(zip);
      expect(parsed.length).toBe(100);
    });

    it("should sort entries alphabetically by default", async () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = await createZip(entries);

      const { entries: parsed } = parseEntries(zip);

      expect(parsed[0]!.path).toBe("a.txt");
      expect(parsed[1]!.path).toBe("b.txt");
      expect(parsed[2]!.path).toBe("c.txt");
    });

    it("should preserve original order with noSort: true", async () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = await createZip(entries, { noSort: true });

      const { entries: parsed } = parseEntries(zip);

      expect(parsed[0]!.path).toBe("c.txt");
      expect(parsed[1]!.path).toBe("a.txt");
      expect(parsed[2]!.path).toBe("b.txt");
    });
  });

  describe("compression levels", () => {
    const testData = new TextEncoder().encode("Compressible content ".repeat(100));

    it("should support all compression levels (0-9)", async () => {
      for (let level = 0; level <= 9; level++) {
        const zip = await createZip([{ name: "file.txt", data: testData }], { level });
        const { entries } = parseEntries(zip);

        if (level === 0) {
          expect(entries[0]!.compressionMethod).toBe(0);
        } else {
          expect(entries[0]!.compressionMethod).toBe(8);
        }

        // Verify content can be extracted
        const extracted = extractFileSync(zip, "file.txt");
        expect(extracted).toEqual(testData);
      }
    });

    it("should use default level when not specified", async () => {
      const zip = await createZip([{ name: "file.txt", data: testData }]);
      const { entries } = parseEntries(zip);

      // Default level should compress
      expect(entries[0]!.compressionMethod).toBe(8);
      expect(entries[0]!.compressedSize).toBeLessThan(testData.length);
    });
  });
});
