import { describe, it, expect } from "vitest";
import { createZip, createZipSync, type ZipEntry } from "@archive/zip/zip-bytes";
import { decompress } from "@archive/compression/compress";

// Helper to read ZIP structure
function parseZipStructure(data: Uint8Array): {
  localHeaders: Array<{
    signature: number;
    fileName: string;
    compressedSize: number;
    uncompressedSize: number;
    crc32: number;
    compressionMethod: number;
  }>;
  centralDirectory: Array<{
    signature: number;
    fileName: string;
    offset: number;
  }>;
  endOfCentralDir: {
    signature: number;
    entryCount: number;
    centralDirSize: number;
    centralDirOffset: number;
  };
} {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const decoder = new TextDecoder();
  const localHeaders: Array<{
    signature: number;
    fileName: string;
    compressedSize: number;
    uncompressedSize: number;
    crc32: number;
    compressionMethod: number;
  }> = [];
  const centralDirectory: Array<{
    signature: number;
    fileName: string;
    offset: number;
  }> = [];
  let endOfCentralDir = {
    signature: 0,
    entryCount: 0,
    centralDirSize: 0,
    centralDirOffset: 0
  };

  let offset = 0;

  // Parse local file headers
  while (offset < data.length) {
    const sig = view.getUint32(offset, true);
    if (sig !== 0x04034b50) {
      break;
    }

    const fileNameLen = view.getUint16(offset + 26, true);
    const extraLen = view.getUint16(offset + 28, true);
    const fileName = decoder.decode(data.subarray(offset + 30, offset + 30 + fileNameLen));
    const compressedSize = view.getUint32(offset + 18, true);

    localHeaders.push({
      signature: sig,
      fileName,
      compressedSize: view.getUint32(offset + 18, true),
      uncompressedSize: view.getUint32(offset + 22, true),
      crc32: view.getUint32(offset + 14, true),
      compressionMethod: view.getUint16(offset + 8, true)
    });

    offset += 30 + fileNameLen + extraLen + compressedSize;
  }

  // Parse central directory
  while (offset < data.length) {
    const sig = view.getUint32(offset, true);
    if (sig !== 0x02014b50) {
      break;
    }

    const fileNameLen = view.getUint16(offset + 28, true);
    const extraLen = view.getUint16(offset + 30, true);
    const commentLen = view.getUint16(offset + 32, true);
    const fileName = decoder.decode(data.subarray(offset + 46, offset + 46 + fileNameLen));

    centralDirectory.push({
      signature: sig,
      fileName,
      offset: view.getUint32(offset + 42, true)
    });

    offset += 46 + fileNameLen + extraLen + commentLen;
  }

  // Parse end of central directory
  if (offset < data.length) {
    const sig = view.getUint32(offset, true);
    if (sig === 0x06054b50) {
      endOfCentralDir = {
        signature: sig,
        entryCount: view.getUint16(offset + 10, true),
        centralDirSize: view.getUint32(offset + 12, true),
        centralDirOffset: view.getUint32(offset + 16, true)
      };
    }
  }

  return { localHeaders, centralDirectory, endOfCentralDir };
}

// Helper to extract file content from ZIP
async function extractFile(zipData: Uint8Array, fileName: string): Promise<Uint8Array | null> {
  const structure = parseZipStructure(zipData);
  const view = new DataView(zipData.buffer, zipData.byteOffset, zipData.byteLength);

  for (let i = 0; i < structure.localHeaders.length; i++) {
    if (structure.localHeaders[i].fileName === fileName) {
      const header = structure.localHeaders[i];
      const cdEntry = structure.centralDirectory[i];
      const offset = cdEntry.offset;

      const fileNameLen = view.getUint16(offset + 26, true);
      const extraLen = view.getUint16(offset + 28, true);
      const dataStart = offset + 30 + fileNameLen + extraLen;
      const compressedData = zipData.subarray(dataStart, dataStart + header.compressedSize);

      if (header.compressionMethod === 0) {
        // STORE - no compression
        return compressedData;
      } else if (header.compressionMethod === 8) {
        // DEFLATE
        return decompress(compressedData);
      }
    }
  }
  return null;
}

describe("zip-bytes", () => {
  describe("createZip (async)", () => {
    it("should create a valid empty ZIP", async () => {
      const zip = await createZip([]);
      const structure = parseZipStructure(zip);

      expect(structure.localHeaders.length).toBe(0);
      expect(structure.centralDirectory.length).toBe(0);
      expect(structure.endOfCentralDir.signature).toBe(0x06054b50);
      expect(structure.endOfCentralDir.entryCount).toBe(0);
    });

    it("should create ZIP with single file", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "hello.txt", data: content }]);

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders.length).toBe(1);
      expect(structure.localHeaders[0].fileName).toBe("hello.txt");
      expect(structure.localHeaders[0].uncompressedSize).toBe(content.length);
      expect(structure.endOfCentralDir.entryCount).toBe(1);

      // Extract and verify content
      const extracted = await extractFile(zip, "hello.txt");
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

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders.length).toBe(3);
      expect(structure.centralDirectory.length).toBe(3);
      expect(structure.endOfCentralDir.entryCount).toBe(3);

      // Verify all files can be extracted
      expect(await extractFile(zip, "file1.txt")).toEqual(file1);
      expect(await extractFile(zip, "file2.txt")).toEqual(file2);
      expect(await extractFile(zip, "file3.txt")).toEqual(file3);
    });

    it("should create ZIP with nested directories", async () => {
      const content = new TextEncoder().encode("Nested file content");
      const zip = await createZip([{ name: "folder/subfolder/deep/file.txt", data: content }]);

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].fileName).toBe("folder/subfolder/deep/file.txt");

      const extracted = await extractFile(zip, "folder/subfolder/deep/file.txt");
      expect(extracted).toEqual(content);
    });

    it("should handle empty file", async () => {
      const empty = new Uint8Array(0);
      const zip = await createZip([{ name: "empty.txt", data: empty }]);

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].uncompressedSize).toBe(0);
      expect(structure.localHeaders[0].compressionMethod).toBe(0); // STORE for empty

      const extracted = await extractFile(zip, "empty.txt");
      expect(extracted).toEqual(empty);
    });

    it("should handle large file (1MB)", async () => {
      const large = new Uint8Array(1024 * 1024);
      // Highly compressible payload.
      large.fill(0);

      const zip = await createZip([{ name: "large.bin", data: large }], { level: 1 });

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].uncompressedSize).toBe(large.length);
      // Compressed size should be smaller
      expect(structure.localHeaders[0].compressedSize).toBeLessThan(large.length);

      const extracted = await extractFile(zip, "large.bin");
      expect(extracted).toEqual(large);
    });

    it("should smart STORE incompressible data", async () => {
      const large = new Uint8Array(1024 * 1024);
      for (let i = 0; i < large.length; i++) {
        large[i] = i % 256;
      }

      const zip = await createZip([{ name: "incompressible.bin", data: large }], { level: 6 });

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].compressionMethod).toBe(0); // STORE
      expect(structure.localHeaders[0].compressedSize).toBe(large.length);

      const extracted = await extractFile(zip, "incompressible.bin");
      expect(extracted).toEqual(large);
    });

    it("should use STORE method for level 0", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "uncompressed.txt", data: content }], { level: 0 });

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].compressionMethod).toBe(0); // STORE
      expect(structure.localHeaders[0].compressedSize).toBe(content.length);
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

      const structure = parseZipStructure(zipData);
      const a = structure.localHeaders.find(h => h.fileName === "a.txt")!;
      const b = structure.localHeaders.find(h => h.fileName === "b.bin")!;

      expect(a.compressionMethod).toBe(8);
      expect(b.compressionMethod).toBe(0);
    });
    it("should use DEFLATE method for level > 0", async () => {
      const content = new TextEncoder().encode("Hello, World!".repeat(100));
      const zip = await createZip([{ name: "compressed.txt", data: content }], { level: 6 });

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].compressionMethod).toBe(8); // DEFLATE
      expect(structure.localHeaders[0].compressedSize).toBeLessThan(content.length);
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

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].fileName).toBe("文件.txt");
      expect(structure.localHeaders[1].fileName).toBe("файл.txt");
      expect(structure.localHeaders[2].fileName).toBe("αρχείο.txt");
    });

    it("should handle filenames with spaces", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([
        { name: "file with spaces.txt", data: content },
        { name: "folder name/file name.txt", data: content }
      ]);

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].fileName).toBe("file with spaces.txt");
      expect(structure.localHeaders[1].fileName).toBe("folder name/file name.txt");
    });

    it("should handle binary data", async () => {
      const binary = new Uint8Array([0x00, 0x01, 0xff, 0xfe, 0x7f, 0x80]);
      const zip = await createZip([{ name: "binary.bin", data: binary }]);

      const extracted = await extractFile(zip, "binary.bin");
      expect(extracted).toEqual(binary);
    });

    it("should calculate correct CRC32", async () => {
      const content = new TextEncoder().encode("Hello, World!");
      const zip = await createZip([{ name: "test.txt", data: content }]);

      const structure = parseZipStructure(zip);
      // CRC32 of "Hello, World!" is 0xec4ac3d0
      expect(structure.localHeaders[0].crc32).toBe(0xec4ac3d0);
    });

    it("should support file modification time", async () => {
      const content = new TextEncoder().encode("Timestamped content");
      const modTime = new Date(2023, 5, 15, 10, 30, 0); // June 15, 2023 10:30:00

      const zip = await createZip([{ name: "dated.txt", data: content, modTime }]);

      // Verify ZIP was created (detailed time verification would require more parsing)
      const structure = parseZipStructure(zip);
      expect(structure.localHeaders.length).toBe(1);
    });

    it("should support ZIP comment", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([{ name: "file.txt", data: content }], {
        comment: "This is a ZIP comment"
      });

      // ZIP comment is at the end of EOCD
      const decoder = new TextDecoder();
      const zipString = decoder.decode(zip.subarray(zip.length - 50));
      expect(zipString).toContain("This is a ZIP comment");
    });

    it("should support file comments", async () => {
      const content = new TextEncoder().encode("Content");
      const zip = await createZip([
        { name: "file.txt", data: content, comment: "File comment here" }
      ]);

      // Comment is stored in central directory
      const decoder = new TextDecoder();
      const zipString = decoder.decode(zip);
      expect(zipString).toContain("File comment here");
    });
  });

  describe("createZipSync", () => {
    it("should create valid ZIP synchronously", () => {
      const content = new TextEncoder().encode("Sync content");
      const zip = createZipSync([{ name: "sync.txt", data: content }]);

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders.length).toBe(1);
      expect(structure.endOfCentralDir.entryCount).toBe(1);
    });

    it("should produce same result as async for same input", async () => {
      const content = new TextEncoder().encode("Same content");
      const entries: ZipEntry[] = [{ name: "file.txt", data: content }];

      const asyncZip = await createZip(entries, { level: 6 });
      const syncZip = createZipSync(entries, { level: 6 });

      // Structure should be identical
      const asyncStructure = parseZipStructure(asyncZip);
      const syncStructure = parseZipStructure(syncZip);

      expect(asyncStructure.localHeaders.length).toBe(syncStructure.localHeaders.length);
      expect(asyncStructure.endOfCentralDir.entryCount).toBe(
        syncStructure.endOfCentralDir.entryCount
      );
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

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].fileName).toBe(longName);
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
        { noSort: true }
      );

      const structure = parseZipStructure(zip);
      specialNames.forEach((name, i) => {
        expect(structure.localHeaders[i].fileName).toBe(name);
      });
    });

    it("should handle data with all bytes 0x00", async () => {
      const zeros = new Uint8Array(1000).fill(0);
      const zip = await createZip([{ name: "zeros.bin", data: zeros }]);

      const extracted = await extractFile(zip, "zeros.bin");
      expect(extracted).toEqual(zeros);
    });

    it("should handle data with all bytes 0xFF", async () => {
      const ones = new Uint8Array(1000).fill(0xff);
      const zip = await createZip([{ name: "ones.bin", data: ones }]);

      const extracted = await extractFile(zip, "ones.bin");
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

      const structure = parseZipStructure(zip);
      expect(structure.localHeaders[0].uncompressedSize).toBe(0);
      expect(structure.localHeaders[1].uncompressedSize).toBe(11);
      expect(structure.localHeaders[2].uncompressedSize).toBe(0);
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

      const extracted = await extractFile(zip, "xl/worksheets/sheet1.xml");
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
      const structure = parseZipStructure(zip);

      expect(structure.localHeaders.length).toBe(100);
      expect(structure.centralDirectory.length).toBe(100);
      expect(structure.endOfCentralDir.entryCount).toBe(100);
    });

    it("should sort entries alphabetically by default", async () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = await createZip(entries);
      const structure = parseZipStructure(zip);

      // Entries should be sorted alphabetically
      expect(structure.localHeaders[0].fileName).toBe("a.txt");
      expect(structure.localHeaders[1].fileName).toBe("b.txt");
      expect(structure.localHeaders[2].fileName).toBe("c.txt");
    });

    it("should preserve original order with noSort: true", async () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = await createZip(entries, { noSort: true });
      const structure = parseZipStructure(zip);

      // Order should be preserved
      expect(structure.localHeaders[0].fileName).toBe("c.txt");
      expect(structure.localHeaders[1].fileName).toBe("a.txt");
      expect(structure.localHeaders[2].fileName).toBe("b.txt");
    });

    it("should sort entries alphabetically by default (sync)", () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = createZipSync(entries);
      const structure = parseZipStructure(zip);

      // Entries should be sorted alphabetically
      expect(structure.localHeaders[0].fileName).toBe("a.txt");
      expect(structure.localHeaders[1].fileName).toBe("b.txt");
      expect(structure.localHeaders[2].fileName).toBe("c.txt");
    });

    it("should preserve original order with noSort: true (sync)", () => {
      const entries: ZipEntry[] = [
        { name: "c.txt", data: new TextEncoder().encode("C") },
        { name: "a.txt", data: new TextEncoder().encode("A") },
        { name: "b.txt", data: new TextEncoder().encode("B") }
      ];

      const zip = createZipSync(entries, { noSort: true });
      const structure = parseZipStructure(zip);

      // Order should be preserved
      expect(structure.localHeaders[0].fileName).toBe("c.txt");
      expect(structure.localHeaders[1].fileName).toBe("a.txt");
      expect(structure.localHeaders[2].fileName).toBe("b.txt");
    });
  });

  describe("compression levels", () => {
    const testData = new TextEncoder().encode("Compressible content ".repeat(100));

    it("should support all compression levels (0-9)", async () => {
      for (let level = 0; level <= 9; level++) {
        const zip = await createZip([{ name: "file.txt", data: testData }], { level });
        const structure = parseZipStructure(zip);

        if (level === 0) {
          expect(structure.localHeaders[0].compressionMethod).toBe(0);
        } else {
          expect(structure.localHeaders[0].compressionMethod).toBe(8);
        }

        // Verify content can be extracted
        const extracted = await extractFile(zip, "file.txt");
        expect(extracted).toEqual(testData);
      }
    });

    it("should use default level when not specified", async () => {
      const zip = await createZip([{ name: "file.txt", data: testData }]);
      const structure = parseZipStructure(zip);

      // Default level should compress
      expect(structure.localHeaders[0].compressionMethod).toBe(8);
      expect(structure.localHeaders[0].compressedSize).toBeLessThan(testData.length);
    });
  });
});
