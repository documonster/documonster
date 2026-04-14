/**
 * Tests for the unified ArchiveFile class.
 */

import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { S_IFLNK, ZIP_OS_UNIX } from "@archive/zip-spec/zip-records";
import type { ZipEntry } from "@archive/zip/zip-bytes";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

import { ArchiveFile } from "../archive-file.js";

describe("ArchiveFile", () => {
  let testDir: string;

  beforeEach(async () => {
    testDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "archive-file-test-"));
  });

  afterEach(async () => {
    await fs.promises.rm(testDir, { recursive: true, force: true });
  });

  describe("ZIP format (default)", () => {
    it("should create empty archive with format zip by default", () => {
      const archive = new ArchiveFile();
      expect(archive.format).toBe("zip");
    });

    it("should add buffer and build archive", async () => {
      const archive = new ArchiveFile();
      archive.addBuffer(new TextEncoder().encode("Hello, World!"), "hello.txt");

      const buffer = await archive.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });

    it("should add text and build archive", async () => {
      const archive = new ArchiveFile();
      archive.addText("Hello, World!", "hello.txt");

      const buffer = await archive.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
    });

    it("should support method chaining", () => {
      const archive = new ArchiveFile();
      const result = archive.addText("content1", "file1.txt").addText("content2", "file2.txt");

      expect(result).toBe(archive);
    });

    it("should read entries from buffer", () => {
      const archive = new ArchiveFile();
      archive.addText("Hello", "hello.txt");
      archive.addText("World", "world.txt");

      const buffer = archive.toBufferSync();
      const reader = ArchiveFile.fromBuffer(buffer);

      const entries = reader.getEntriesSync();
      expect(entries.length).toBe(2);
      expect(entries.map(e => e.path).sort()).toEqual(["hello.txt", "world.txt"]);
    });

    it("should write and read from file", async () => {
      const archive = new ArchiveFile();
      archive.addText("Test content", "test.txt");

      const zipPath = path.join(testDir, "test.zip");
      await archive.writeToFile(zipPath);

      expect(fs.existsSync(zipPath)).toBe(true);

      const reader = await ArchiveFile.fromFile(zipPath);
      const entries = await reader.getEntries();
      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("test.txt");
    });
  });

  describe("TAR format", () => {
    it("should create TAR archive with format option", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(archive.format).toBe("tar");
    });

    it("should add buffer and build TAR archive", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addBuffer(new TextEncoder().encode("Hello, World!"), "hello.txt");

      const buffer = await archive.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });

    it("should add text and build TAR archive", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("Hello, World!", "hello.txt");

      const buffer = await archive.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
    });

    it("should create gzipped TAR archive", async () => {
      const archive = new ArchiveFile({ format: "tar", gzip: true });
      archive.addText("Hello, World!", "hello.txt");

      const buffer = await archive.toBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);

      // Check for gzip magic bytes
      expect(buffer[0]).toBe(0x1f);
      expect(buffer[1]).toBe(0x8b);
    });

    it("should write TAR to file", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("Test content", "test.txt");

      const tarPath = path.join(testDir, "test.tar");
      await archive.writeToFile(tarPath);

      expect(fs.existsSync(tarPath)).toBe(true);
    });

    it("should write gzipped TAR to file", async () => {
      const archive = new ArchiveFile({ format: "tar", gzip: true });
      archive.addText("Test content", "test.txt");

      const tarPath = path.join(testDir, "test.tar.gz");
      await archive.writeToFile(tarPath);

      expect(fs.existsSync(tarPath)).toBe(true);

      // Verify gzip header
      const data = await fs.promises.readFile(tarPath);
      expect(data[0]).toBe(0x1f);
      expect(data[1]).toBe(0x8b);
    });

    it("should read TAR from file", async () => {
      // Create a TAR file
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("Test content", "test.txt");
      const tarPath = path.join(testDir, "test.tar");
      await archive.writeToFile(tarPath);

      // Read it back
      const reader = await ArchiveFile.fromFile(tarPath, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("test.txt");
    });

    it("should read gzipped TAR from file", async () => {
      // Create a gzipped TAR file
      const archive = new ArchiveFile({ format: "tar", gzip: true });
      archive.addText("Test content", "test.txt");
      const tarPath = path.join(testDir, "test.tar.gz");
      await archive.writeToFile(tarPath);

      // Read it back (auto-detect gzip from extension)
      const reader = await ArchiveFile.fromFile(tarPath, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("test.txt");
    });

    it("should support synchronous build", () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("Hello", "hello.txt");

      const buffer = archive.toBufferSync();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });

  describe("addFile", () => {
    it("should add file from disk (ZIP)", async () => {
      // Create a test file
      const testFile = path.join(testDir, "source.txt");
      await fs.promises.writeFile(testFile, "File content");

      const archive = new ArchiveFile();
      archive.addFile(testFile);

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("source.txt");
    });

    it("should add file from disk (TAR)", async () => {
      // Create a test file
      const testFile = path.join(testDir, "source.txt");
      await fs.promises.writeFile(testFile, "File content");

      const archive = new ArchiveFile({ format: "tar" });
      archive.addFile(testFile);

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("source.txt");
    });
  });

  describe("addDirectory", () => {
    it("should add directory recursively (ZIP)", async () => {
      // Create test directory structure
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "file1.txt"), "content1");
      await fs.promises.writeFile(path.join(srcDir, "file2.txt"), "content2");

      const archive = new ArchiveFile();
      archive.addDirectory(srcDir);

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(2);
    });

    it("should add directory recursively (TAR)", async () => {
      // Create test directory structure
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "file1.txt"), "content1");
      await fs.promises.writeFile(path.join(srcDir, "file2.txt"), "content2");

      const archive = new ArchiveFile({ format: "tar" });
      archive.addDirectory(srcDir);

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(2);
    });
  });

  describe("addGlob", () => {
    it("should add files matching glob pattern (ZIP)", async () => {
      // Create test files
      await fs.promises.writeFile(path.join(testDir, "file1.txt"), "content1");
      await fs.promises.writeFile(path.join(testDir, "file2.txt"), "content2");
      await fs.promises.writeFile(path.join(testDir, "file3.md"), "content3");

      const archive = new ArchiveFile();
      archive.addGlob("*.txt", { cwd: testDir });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(2);
      expect(entries.every(e => e.path.endsWith(".txt"))).toBe(true);
    });

    it("should add files matching glob pattern (TAR)", async () => {
      // Create test files
      await fs.promises.writeFile(path.join(testDir, "file1.txt"), "content1");
      await fs.promises.writeFile(path.join(testDir, "file2.txt"), "content2");
      await fs.promises.writeFile(path.join(testDir, "file3.md"), "content3");

      const archive = new ArchiveFile({ format: "tar" });
      archive.addGlob("*.txt", { cwd: testDir });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(2);
      expect(entries.every(e => e.path.endsWith(".txt"))).toBe(true);
    });
  });

  describe("transform function", () => {
    it("should skip entries when transform returns false (ZIP)", async () => {
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "keep.txt"), "keep");
      await fs.promises.writeFile(path.join(srcDir, "skip.log"), "skip");

      const archive = new ArchiveFile();
      archive.addDirectory(srcDir, {
        transform: data => (data.name.endsWith(".log") ? false : data)
      });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toContain("keep.txt");
    });

    it("should rename entries via transform (ZIP)", async () => {
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "old.txt"), "content");

      const archive = new ArchiveFile();
      archive.addDirectory(srcDir, {
        transform: data => ({ ...data, name: data.name.replace("old", "new") })
      });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toContain("new.txt");
    });

    it("should skip entries when transform returns false (TAR)", async () => {
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "keep.txt"), "keep");
      await fs.promises.writeFile(path.join(srcDir, "skip.log"), "skip");

      const archive = new ArchiveFile({ format: "tar" });
      archive.addDirectory(srcDir, {
        transform: data => (data.name.endsWith(".log") ? false : data)
      });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toContain("keep.txt");
    });

    it("should rename entries via transform (TAR)", async () => {
      const srcDir = path.join(testDir, "src");
      await fs.promises.mkdir(srcDir, { recursive: true });
      await fs.promises.writeFile(path.join(srcDir, "old.txt"), "content");

      const archive = new ArchiveFile({ format: "tar" });
      archive.addDirectory(srcDir, {
        transform: data => ({ ...data, name: data.name.replace("old", "new") })
      });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer, { format: "tar" });
      const entries = await reader.getEntries();

      expect(entries.length).toBe(1);
      expect(entries[0].path).toContain("new.txt");
    });

    it("should work with addGlob transform", async () => {
      await fs.promises.writeFile(path.join(testDir, "a.txt"), "a");
      await fs.promises.writeFile(path.join(testDir, "b.txt"), "b");
      await fs.promises.writeFile(path.join(testDir, "c.log"), "c");

      const archive = new ArchiveFile();
      archive.addGlob("*.*", {
        cwd: testDir,
        transform: data => {
          if (data.name.endsWith(".log")) {
            return false;
          }
          return { ...data, prefix: "files/" };
        }
      });

      const buffer = await archive.toBuffer();
      const reader = ArchiveFile.fromBuffer(buffer);
      const entries = reader.getEntriesSync();

      expect(entries.length).toBe(2);
      expect(entries.every(e => e.path.startsWith("files/"))).toBe(true);
    });
  });

  describe("extraction", () => {
    it("should extract ZIP to directory", async () => {
      const archive = new ArchiveFile();
      archive.addText("content1", "file1.txt");
      archive.addText("content2", "subdir/file2.txt");

      const zipPath = path.join(testDir, "test.zip");
      await archive.writeToFile(zipPath);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted");
      await reader.extractTo(extractDir);

      expect(fs.existsSync(path.join(extractDir, "file1.txt"))).toBe(true);
      expect(fs.existsSync(path.join(extractDir, "subdir/file2.txt"))).toBe(true);

      const content1 = await fs.promises.readFile(path.join(extractDir, "file1.txt"), "utf-8");
      expect(content1).toBe("content1");
    });

    it("should emit onWarning for non-writable target (ZIP)", async () => {
      // Skip on Windows where chmod has no effect on directory write permissions
      if (process.platform === "win32") {
        return;
      }
      const archive = new ArchiveFile();
      archive.addText("content", "file.txt");

      const zipPath = path.join(testDir, "test.zip");
      await archive.writeToFile(zipPath);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-ro");
      await fs.promises.mkdir(extractDir, { recursive: true });
      await fs.promises.chmod(extractDir, 0o555);

      const warnings: any[] = [];
      await reader.extractTo(extractDir, {
        onWarning: w => warnings.push(w)
      });

      expect(warnings.length).toBeGreaterThan(0);
      expect(fs.existsSync(path.join(extractDir, "file.txt"))).toBe(false);
    });

    it("should extract TAR to directory", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("content1", "file1.txt");
      archive.addText("content2", "subdir/file2.txt");

      const tarPath = path.join(testDir, "test.tar");
      await archive.writeToFile(tarPath);

      const reader = await ArchiveFile.fromFile(tarPath, { format: "tar" });
      const extractDir = path.join(testDir, "extracted");
      await reader.extractTo(extractDir);

      expect(fs.existsSync(path.join(extractDir, "file1.txt"))).toBe(true);
      expect(fs.existsSync(path.join(extractDir, "subdir/file2.txt"))).toBe(true);

      const content1 = await fs.promises.readFile(path.join(extractDir, "file1.txt"), "utf-8");
      expect(content1).toBe("content1");
    });

    it("should report onProgress during TAR extraction", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("content1", "file1.txt");
      archive.addText("content2", "subdir/file2.txt");

      const tarPath = path.join(testDir, "test.tar");
      await archive.writeToFile(tarPath);

      const reader = await ArchiveFile.fromFile(tarPath, { format: "tar" });
      const extractDir = path.join(testDir, "extracted-progress");

      const progress: any[] = [];
      await reader.extractTo(extractDir, {
        onProgress: p => progress.push(p)
      });

      expect(progress.length).toBeGreaterThan(0);
      expect(progress.at(-1)?.extractedEntries).toBe(2);
      expect(fs.existsSync(path.join(extractDir, "subdir/file2.txt"))).toBe(true);
    });

    it("should emit onWarning for non-writable target (TAR)", async () => {
      // Skip on Windows where chmod has no effect on directory write permissions
      if (process.platform === "win32") {
        return;
      }
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("content", "file.txt");

      const tarPath = path.join(testDir, "test.tar");
      await archive.writeToFile(tarPath);

      const reader = await ArchiveFile.fromFile(tarPath, { format: "tar" });
      const extractDir = path.join(testDir, "extracted-ro-tar");
      await fs.promises.mkdir(extractDir, { recursive: true });
      await fs.promises.chmod(extractDir, 0o555);

      const warnings: any[] = [];
      await reader.extractTo(extractDir, {
        onWarning: w => warnings.push(w)
      });

      expect(warnings.length).toBeGreaterThan(0);
      expect(fs.existsSync(path.join(extractDir, "file.txt"))).toBe(false);
    });

    it("should preserve file permissions when extracting ZIP (Unix only)", async function () {
      // Skip on Windows where chmod has no effect
      if (process.platform === "win32") {
        return;
      }

      const archive = new ArchiveFile({ writePermissions: true });
      // Add file with executable permission
      archive.addText("#!/bin/bash\necho hello", "script.sh", { mode: 0o100755 });
      archive.addText("regular content", "file.txt", { mode: 0o100644 });

      const zipPath = path.join(testDir, "perms.zip");
      await archive.writeToFile(zipPath);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-perms");
      await reader.extractTo(extractDir, { preservePermissions: true });

      const scriptStats = await fs.promises.stat(path.join(extractDir, "script.sh"));
      const fileStats = await fs.promises.stat(path.join(extractDir, "file.txt"));

      // Check executable bit is set on script
      expect(scriptStats.mode & 0o111).toBeTruthy();
      // Check file is readable
      expect(fileStats.mode & 0o444).toBeTruthy();
    });

    it("should use default permissions when ZIP has no mode info", async function () {
      // Skip on Windows
      if (process.platform === "win32") {
        return;
      }

      // Create a ZIP without permission info using low-level API
      const { createZipSync } = await import("@archive/zip/zip-bytes");

      // Windows-style ZIP (versionMadeBy = 0, externalAttributes = 0 for files)
      const entries: ZipEntry[] = [
        {
          name: "file.txt",
          data: new TextEncoder().encode("content"),
          externalAttributes: 0, // No Unix mode info
          versionMadeBy: 0x0014 // DOS (0) + version 20
        },
        {
          name: "dir/",
          data: new Uint8Array(0),
          externalAttributes: 0x10, // DOS directory attribute
          versionMadeBy: 0x0014 // DOS
        }
      ];

      const zipBytes = createZipSync(entries);
      const zipPath = path.join(testDir, "no-perms.zip");
      await fs.promises.writeFile(zipPath, zipBytes);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-defaults");
      await reader.extractTo(extractDir, { preservePermissions: true });

      const fileStats = await fs.promises.stat(path.join(extractDir, "file.txt"));
      const dirStats = await fs.promises.stat(path.join(extractDir, "dir"));

      // Default file: 0o644 (rw-r--r--)
      expect(fileStats.mode & 0o777).toBe(0o644);
      // Default dir: 0o755 (rwxr-xr-x)
      expect(dirStats.mode & 0o777).toBe(0o755);
    });

    it("should skip symlinks when createSymlinks is false", async () => {
      // Create ZIP with a symlink entry manually using ZipParser
      const { createZipSync } = await import("@archive/zip/zip-bytes");

      // Create a valid symlink entry (Unix mode with S_IFLNK)
      const entries: ZipEntry[] = [
        { name: "target.txt", data: new TextEncoder().encode("target content") },
        {
          name: "link.txt",
          data: new TextEncoder().encode("target.txt"),
          externalAttributes: (S_IFLNK | 0o777) << 16, // Unix symlink mode
          versionMadeBy: (ZIP_OS_UNIX << 8) | 30 // Unix (3) + version 30
        }
      ];

      const zipBytes = createZipSync(entries);
      const zipPath = path.join(testDir, "symlink.zip");
      await fs.promises.writeFile(zipPath, zipBytes);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-no-symlinks");

      await reader.extractTo(extractDir, { createSymlinks: false });

      // Target file should exist
      expect(fs.existsSync(path.join(extractDir, "target.txt"))).toBe(true);
      // Symlink should NOT exist (was skipped)
      expect(fs.existsSync(path.join(extractDir, "link.txt"))).toBe(false);
    });

    it("should create symlinks when extracting (Unix only)", async function () {
      // Skip on Windows (symlinks require admin privileges)
      if (process.platform === "win32") {
        return;
      }

      const { createZipSync } = await import("@archive/zip/zip-bytes");

      const entries: ZipEntry[] = [
        { name: "target.txt", data: new TextEncoder().encode("target content") },
        {
          name: "link.txt",
          data: new TextEncoder().encode("target.txt"),
          externalAttributes: (S_IFLNK | 0o777) << 16,
          versionMadeBy: (ZIP_OS_UNIX << 8) | 30
        }
      ];

      const zipBytes = createZipSync(entries);
      const zipPath = path.join(testDir, "symlink-create.zip");
      await fs.promises.writeFile(zipPath, zipBytes);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-symlinks");

      await reader.extractTo(extractDir, { createSymlinks: true });

      // Target file should exist
      expect(fs.existsSync(path.join(extractDir, "target.txt"))).toBe(true);

      // Symlink should exist and be a symlink
      const linkPath = path.join(extractDir, "link.txt");
      const stats = await fs.promises.lstat(linkPath);
      expect(stats.isSymbolicLink()).toBe(true);

      // Symlink should point to the correct target
      const linkTarget = await fs.promises.readlink(linkPath);
      expect(linkTarget).toBe("target.txt");
    });

    it("should warn and skip symlinks pointing outside extraction directory", async function () {
      // Skip on Windows
      if (process.platform === "win32") {
        return;
      }

      const { createZipSync } = await import("@archive/zip/zip-bytes");

      const entries: ZipEntry[] = [
        {
          name: "evil-link.txt",
          data: new TextEncoder().encode("../../../etc/passwd"),
          externalAttributes: (S_IFLNK | 0o777) << 16,
          versionMadeBy: (ZIP_OS_UNIX << 8) | 30
        }
      ];

      const zipBytes = createZipSync(entries);
      const zipPath = path.join(testDir, "evil-symlink.zip");
      await fs.promises.writeFile(zipPath, zipBytes);

      const reader = await ArchiveFile.fromFile(zipPath);
      const extractDir = path.join(testDir, "extracted-evil");

      const warnings: any[] = [];
      await reader.extractTo(extractDir, {
        createSymlinks: true,
        onWarning: w => warnings.push(w)
      });

      // Should have emitted a warning
      expect(warnings.length).toBeGreaterThan(0);
      expect(warnings[0].message).toContain("outside extraction directory");

      // Evil symlink should NOT have been created
      expect(fs.existsSync(path.join(extractDir, "evil-link.txt"))).toBe(false);
    });
  });

  describe("ZIP-specific methods", () => {
    it("should throw for TAR when using has()", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(() => (archive as any).has("test.txt")).toThrow(
        "has() is only available for ZIP archives"
      );
    });

    it("should throw for TAR when using delete()", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(() => (archive as any).delete("test.txt")).toThrow(
        "delete() is only available for ZIP archives"
      );
    });

    it("should throw for TAR when using set()", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(() => (archive as any).set("test.txt", "content")).toThrow(
        "set() is only available for ZIP archives"
      );
    });

    it("should throw for TAR when using rename()", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(() => (archive as any).rename("old.txt", "new.txt")).toThrow(
        "rename() is only available for ZIP archives"
      );
    });

    it("should throw for TAR when using setPassword()", () => {
      const archive = new ArchiveFile({ format: "tar" });
      expect(() => (archive as any).setPassword("secret")).toThrow(
        "setPassword() is only available for ZIP archives"
      );
    });
  });

  describe("entryCount", () => {
    it("should return pending entry count for new ZIP archive", () => {
      const archive = new ArchiveFile();
      archive.addText("content1", "file1.txt");
      archive.addText("content2", "file2.txt");

      expect(archive.entryCount).toBe(2);
    });

    it("should return pending entry count for new TAR archive", () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("content1", "file1.txt");
      archive.addText("content2", "file2.txt");

      expect(archive.entryCount).toBe(2);
    });
  });

  describe("Streaming", () => {
    it("should stream ZIP archive chunks", async () => {
      const archive = new ArchiveFile();
      archive.addText("Hello, World!", "hello.txt");
      archive.addText("Goodbye, World!", "goodbye.txt");

      const chunks: Uint8Array[] = [];
      for await (const chunk of archive.stream()) {
        chunks.push(chunk);
      }

      expect(chunks.length).toBeGreaterThan(0);

      // Verify the output is valid ZIP
      const combined = new Uint8Array(chunks.reduce((sum, c) => sum + c.length, 0));
      let offset = 0;
      for (const chunk of chunks) {
        combined.set(chunk, offset);
        offset += chunk.length;
      }

      const reader = ArchiveFile.fromBuffer(combined);
      const entries = reader.getEntriesSync();
      expect(entries.length).toBe(2);
    });

    it("should stream TAR archive chunks", async () => {
      const archive = new ArchiveFile({ format: "tar" });
      archive.addText("Hello, World!", "hello.txt");
      archive.addText("Goodbye, World!", "goodbye.txt");

      const chunks: Uint8Array[] = [];
      for await (const chunk of archive.stream()) {
        chunks.push(chunk);
      }

      expect(chunks.length).toBeGreaterThan(0);

      // Verify the output is valid TAR
      const combined = new Uint8Array(chunks.reduce((sum, c) => sum + c.length, 0));
      let offset = 0;
      for (const chunk of chunks) {
        combined.set(chunk, offset);
        offset += chunk.length;
      }

      const reader = ArchiveFile.fromBuffer(combined, { format: "tar" });
      const entries = await reader.getEntries();
      expect(entries.length).toBe(2);
    });

    it("should report progress during streaming", async () => {
      const archive = new ArchiveFile();
      archive.addText("File 1 content", "file1.txt");
      archive.addText("File 2 content", "file2.txt");

      const progressUpdates: number[] = [];
      const op = archive.operation({
        onProgress: p => {
          progressUpdates.push(p.entriesDone);
        }
      });

      for await (const _ of op.iterable) {
        // Consume chunks
      }

      // Should have received progress updates
      expect(progressUpdates.length).toBeGreaterThan(0);
    });

    it("should streamToFile write directly to disk", async () => {
      const archive = new ArchiveFile();
      archive.addText("Hello, World!", "hello.txt");
      archive.addText("Goodbye, World!", "goodbye.txt");

      const zipPath = path.join(testDir, "streamed.zip");
      await archive.streamToFile(zipPath);

      expect(fs.existsSync(zipPath)).toBe(true);

      // Verify the file is valid
      const reader = await ArchiveFile.fromFile(zipPath);
      const entries = await reader.getEntries();
      expect(entries.length).toBe(2);
    });

    it("should pipeTo write to a WritableStream", async () => {
      const archive = new ArchiveFile();
      archive.addText("Hello, World!", "hello.txt");

      const zipPath = path.join(testDir, "piped.zip");
      const writeStream = fs.createWriteStream(zipPath);

      await archive.pipeTo(writeStream);

      expect(fs.existsSync(zipPath)).toBe(true);

      // Verify the file is valid
      const reader = await ArchiveFile.fromFile(zipPath);
      const entries = await reader.getEntries();
      expect(entries.length).toBe(1);
    });

    it("should stream with file inputs using createReadStream internally", async () => {
      // Create a test file
      const testFilePath = path.join(testDir, "input.txt");
      await fs.promises.writeFile(testFilePath, "This is test file content");

      const archive = new ArchiveFile();
      archive.addFile(testFilePath, { name: "streamed-input.txt" });

      const zipPath = path.join(testDir, "file-streamed.zip");
      await archive.streamToFile(zipPath);

      expect(fs.existsSync(zipPath)).toBe(true);

      // Verify the file is valid and contains the right content
      const reader = await ArchiveFile.fromFile(zipPath);
      const entries = await reader.getEntries();
      expect(entries.length).toBe(1);
      expect(entries[0].path).toBe("streamed-input.txt");

      const content = await reader.readAsText("streamed-input.txt");
      expect(content).toBe("This is test file content");
    });
  });
});
