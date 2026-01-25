import { describe, it, expect } from "vitest";
import { ZipParser } from "@archive/unzip/zip-parser";
import { createZip, type ZipEntry } from "@archive/zip/zip-bytes";

// Helper to convert object to ZipEntry array
function toEntries(files: Record<string, Uint8Array>): ZipEntry[] {
  return Object.entries(files).map(([name, data]) => ({ name, data }));
}

describe("ZipParser", () => {
  describe("getEntries", () => {
    it("should parse entries from a valid ZIP file", async () => {
      // Create a test ZIP file
      const testFiles: Record<string, Uint8Array> = {
        "test.txt": new TextEncoder().encode("Hello, World!"),
        "folder/nested.txt": new TextEncoder().encode("Nested content"),
        "binary.bin": new Uint8Array([0x00, 0x01, 0x02, 0x03, 0xff])
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);
      const entries = parser.getEntries();

      expect(entries.length).toBe(3);

      const paths = entries.map(e => e.path).sort();
      expect(paths).toContain("test.txt");
      expect(paths).toContain("folder/nested.txt");
      expect(paths).toContain("binary.bin");
    });

    it("should handle empty ZIP files", async () => {
      const zipData = await createZip([]);
      const parser = new ZipParser(zipData);
      const entries = parser.getEntries();
      expect(entries.length).toBe(0);
    });
  });

  describe("extract", () => {
    it("should extract file content correctly", async () => {
      const originalContent = "Hello, World!";
      const testFiles: Record<string, Uint8Array> = {
        "test.txt": new TextEncoder().encode(originalContent)
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);
      const extractedData = await parser.extract("test.txt");
      expect(extractedData).not.toBeNull();
      const extractedContent = new TextDecoder().decode(extractedData);

      expect(extractedContent).toBe(originalContent);
    });

    it("should extract binary content correctly", async () => {
      const binaryData = new Uint8Array([0x00, 0x01, 0x02, 0x03, 0xff, 0xfe, 0xfd]);
      const testFiles: Record<string, Uint8Array> = {
        "binary.bin": binaryData
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);
      const extractedData = await parser.extract("binary.bin");
      expect(extractedData).not.toBeNull();
      expect(extractedData).toEqual(binaryData);
    });
  });

  describe("ZipParser class", () => {
    it("should list all files", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "a.txt": new TextEncoder().encode("A"),
        "b.txt": new TextEncoder().encode("B"),
        "dir/c.txt": new TextEncoder().encode("C")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const files = parser.listFiles();
      expect(files.sort()).toEqual(["a.txt", "b.txt", "dir/c.txt"].sort());
    });

    it("should extract single file", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "target.txt": new TextEncoder().encode("Target content")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const content = await parser.extract("target.txt");
      expect(content).not.toBeNull();
      expect(new TextDecoder().decode(content!)).toBe("Target content");
    });

    it("should return null for non-existent file", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "exists.txt": new TextEncoder().encode("I exist")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const content = await parser.extract("does-not-exist.txt");
      expect(content).toBeNull();
    });

    it("should extract all files", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "file1.txt": new TextEncoder().encode("Content 1"),
        "file2.txt": new TextEncoder().encode("Content 2")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const allFiles = await parser.extractAll();

      expect(allFiles.size).toBe(2);
      expect(new TextDecoder().decode(allFiles.get("file1.txt")!)).toBe("Content 1");
      expect(new TextDecoder().decode(allFiles.get("file2.txt")!)).toBe("Content 2");
    });

    it("should support forEach iteration", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "a.txt": new TextEncoder().encode("A"),
        "b.txt": new TextEncoder().encode("B")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const visited: string[] = [];
      await parser.forEach(async (entry, getData) => {
        visited.push(entry.path);
        const data = await getData();
        expect(data.length).toBeGreaterThan(0);
      });

      expect(visited.sort()).toEqual(["a.txt", "b.txt"].sort());
    });

    it("should support early termination in forEach", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "a.txt": new TextEncoder().encode("A"),
        "b.txt": new TextEncoder().encode("B"),
        "c.txt": new TextEncoder().encode("C")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const visited: string[] = [];
      await parser.forEach(async entry => {
        visited.push(entry.path);
        return visited.length < 2; // Stop after 2 entries
      });

      expect(visited.length).toBe(2);
    });

    it("should check if entry exists", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "exists.txt": new TextEncoder().encode("I exist")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      expect(parser.hasEntry("exists.txt")).toBe(true);
      expect(parser.hasEntry("not-exists.txt")).toBe(false);
    });

    it("should get entry by path", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "file.txt": new TextEncoder().encode("Content")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      const entry = parser.getEntry("file.txt");
      expect(entry).not.toBeUndefined();
      expect(entry!.path).toBe("file.txt");
      expect(entry!.isDirectory).toBe(false);
    });

    it("should count children in a directory", async () => {
      const testFiles: Record<string, Uint8Array> = {
        "root.txt": new TextEncoder().encode("root"),
        "folder/": new Uint8Array(0),
        "folder/a.txt": new TextEncoder().encode("a"),
        "folder/b.txt": new TextEncoder().encode("b"),
        "folder/sub/": new Uint8Array(0),
        "folder/sub/c.txt": new TextEncoder().encode("c")
      };

      const zipData = await createZip(toEntries(testFiles));
      const parser = new ZipParser(zipData);

      // folder/ has 4 children: a.txt, b.txt, sub/, sub/c.txt
      expect(parser.getChildCount("folder/")).toBe(4);

      // folder/sub/ has 1 child: c.txt
      expect(parser.getChildCount("folder/sub/")).toBe(1);

      // Non-directory returns 0
      expect(parser.getChildCount("root.txt")).toBe(0);

      // Non-existent returns 0
      expect(parser.getChildCount("nonexistent/")).toBe(0);
    });
  });
});
