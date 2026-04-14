/**
 * Tests for shared file system utilities.
 */

import * as fs from "node:fs";
import * as fsp from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";

import {
  // Glob matching
  globToRegex,
  matchGlob,
  matchGlobAny,
  createGlobMatcher,
  clearGlobCache,

  // Directory traversal
  traverseDirectory,
  traverseDirectorySync,

  // Glob file search
  glob,
  globSync,

  // File existence
  fileExists,
  fileExistsSync,

  // Directory creation
  ensureDir,
  ensureDirSync,

  // File stats
  safeStats,
  safeStatsSync,

  // Binary I/O
  readFileBytes,
  readFileBytesSync,
  writeFileBytes,
  writeFileBytesSync,

  // Text I/O
  readFileText,
  readFileTextSync,
  writeFileText,
  writeFileTextSync,

  // File time
  setFileTime,
  setFileTimeSync,

  // File operations
  remove,
  removeSync,
  copyFile,
  copyFileSync
} from "@utils/fs";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

describe("fs utilities", () => {
  let tempDir: string;

  beforeEach(async () => {
    tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "fs-utils-test-"));
  });

  afterEach(async () => {
    try {
      await fsp.rm(tempDir, { recursive: true, force: true });
    } catch {
      // Ignore cleanup errors
    }
  });

  describe("glob pattern matching", () => {
    describe("globToRegex", () => {
      it("should match simple wildcards", () => {
        const regex = globToRegex("*.txt");
        expect(regex.test("foo.txt")).toBe(true);
        expect(regex.test("bar.txt")).toBe(true);
        expect(regex.test("foo.js")).toBe(false);
        expect(regex.test("dir/foo.txt")).toBe(false);
      });

      it("should match ** for any path", () => {
        const regex = globToRegex("**/*.txt");
        expect(regex.test("foo.txt")).toBe(true);
        expect(regex.test("dir/foo.txt")).toBe(true);
        expect(regex.test("dir/sub/foo.txt")).toBe(true);
        expect(regex.test("foo.js")).toBe(false);
      });

      it("should match ? for single character", () => {
        const regex = globToRegex("file?.txt");
        expect(regex.test("file1.txt")).toBe(true);
        expect(regex.test("fileA.txt")).toBe(true);
        expect(regex.test("file.txt")).toBe(false);
        expect(regex.test("file12.txt")).toBe(false);
      });

      it("should match character classes [abc]", () => {
        const regex = globToRegex("file[123].txt");
        expect(regex.test("file1.txt")).toBe(true);
        expect(regex.test("file2.txt")).toBe(true);
        expect(regex.test("file3.txt")).toBe(true);
        expect(regex.test("file4.txt")).toBe(false);
      });

      it("should match brace expansion {a,b,c}", () => {
        const regex = globToRegex("file.{js,ts,json}");
        expect(regex.test("file.js")).toBe(true);
        expect(regex.test("file.ts")).toBe(true);
        expect(regex.test("file.json")).toBe(true);
        expect(regex.test("file.txt")).toBe(false);
      });

      it("should escape special regex characters", () => {
        const regex = globToRegex("file(1).txt");
        expect(regex.test("file(1).txt")).toBe(true);
        expect(regex.test("file1.txt")).toBe(false);
      });

      it("should handle dot files with dot option", () => {
        const regexNoDot = globToRegex("*.txt", { dot: false });
        const regexWithDot = globToRegex("*.txt", { dot: true });

        expect(regexNoDot.test(".hidden.txt")).toBe(false);
        expect(regexWithDot.test(".hidden.txt")).toBe(true);
      });
    });

    describe("matchGlob", () => {
      it("should match simple patterns", () => {
        expect(matchGlob("foo.txt", "*.txt")).toBe(true);
        expect(matchGlob("foo.js", "*.txt")).toBe(false);
      });

      it("should match recursive patterns", () => {
        expect(matchGlob("src/index.ts", "**/*.ts")).toBe(true);
        expect(matchGlob("src/utils/helpers.ts", "**/*.ts")).toBe(true);
        expect(matchGlob("index.ts", "**/*.ts")).toBe(true);
      });

      it("should normalize path separators", () => {
        expect(matchGlob("src\\index.ts", "**/*.ts")).toBe(true);
      });
    });

    describe("matchGlobAny", () => {
      it("should match any of multiple patterns", () => {
        expect(matchGlobAny("foo.ts", ["*.js", "*.ts"])).toBe(true);
        expect(matchGlobAny("foo.js", ["*.js", "*.ts"])).toBe(true);
        expect(matchGlobAny("foo.txt", ["*.js", "*.ts"])).toBe(false);
      });
    });
  });

  describe("directory traversal", () => {
    beforeEach(async () => {
      // Create test directory structure
      await fsp.mkdir(path.join(tempDir, "src", "utils"), { recursive: true });
      await fsp.writeFile(path.join(tempDir, "file1.txt"), "content1");
      await fsp.writeFile(path.join(tempDir, "src", "index.ts"), "content2");
      await fsp.writeFile(path.join(tempDir, "src", "utils", "helper.ts"), "content3");
    });

    describe("traverseDirectory", () => {
      it("should yield all files and directories recursively", async () => {
        const entries: string[] = [];
        for await (const entry of traverseDirectory(tempDir)) {
          entries.push(entry.relativePath);
        }

        expect(entries).toContain("file1.txt");
        expect(entries).toContain("src");
        expect(entries).toContain(path.join("src", "index.ts"));
        expect(entries).toContain(path.join("src", "utils"));
        expect(entries).toContain(path.join("src", "utils", "helper.ts"));
      });

      it("should respect recursive option", async () => {
        const entries: string[] = [];
        for await (const entry of traverseDirectory(tempDir, { recursive: false })) {
          entries.push(entry.relativePath);
        }

        expect(entries).toContain("file1.txt");
        expect(entries).toContain("src");
        expect(entries).not.toContain(path.join("src", "index.ts"));
      });

      it("should apply filter function", async () => {
        const entries: string[] = [];
        for await (const entry of traverseDirectory(tempDir, {
          filter: e => !e.isDirectory
        })) {
          entries.push(entry.relativePath);
        }

        expect(entries).toContain("file1.txt");
        expect(entries).not.toContain("src");
      });
    });

    describe("traverseDirectorySync", () => {
      it("should return all entries synchronously", () => {
        const entries = traverseDirectorySync(tempDir);

        const paths = entries.map(e => e.relativePath);
        expect(paths).toContain("file1.txt");
        expect(paths).toContain("src");
      });
    });
  });

  describe("glob file search", () => {
    beforeEach(async () => {
      await fsp.mkdir(path.join(tempDir, "src"), { recursive: true });
      await fsp.writeFile(path.join(tempDir, "readme.md"), "readme");
      await fsp.writeFile(path.join(tempDir, "index.ts"), "index");
      await fsp.writeFile(path.join(tempDir, "src", "utils.ts"), "utils");
      await fsp.writeFile(path.join(tempDir, "src", "types.ts"), "types");
    });

    describe("glob", () => {
      it("should find files matching pattern", async () => {
        const files: string[] = [];
        for await (const entry of glob("**/*.ts", { cwd: tempDir })) {
          files.push(entry.relativePath);
        }

        expect(files).toContain("index.ts");
        expect(files).toContain("src/utils.ts");
        expect(files).toContain("src/types.ts");
        expect(files).not.toContain("readme.md");
      });

      it("should apply ignore patterns", async () => {
        const files: string[] = [];
        for await (const entry of glob("**/*.ts", {
          cwd: tempDir,
          ignore: "**/utils.ts"
        })) {
          files.push(entry.relativePath);
        }

        expect(files).toContain("index.ts");
        expect(files).not.toContain("src/utils.ts");
      });
    });

    describe("globSync", () => {
      it("should find files synchronously", () => {
        const files = globSync("*.ts", { cwd: tempDir });

        expect(files.map(f => f.relativePath)).toContain("index.ts");
        expect(files.map(f => f.relativePath)).not.toContain("src/utils.ts");
      });
    });
  });

  describe("file existence", () => {
    beforeEach(async () => {
      await fsp.writeFile(path.join(tempDir, "exists.txt"), "content");
    });

    describe("fileExists", () => {
      it("should return true for existing file", async () => {
        expect(await fileExists(path.join(tempDir, "exists.txt"))).toBe(true);
      });

      it("should return false for non-existing file", async () => {
        expect(await fileExists(path.join(tempDir, "not-exists.txt"))).toBe(false);
      });
    });

    describe("fileExistsSync", () => {
      it("should return true for existing file", () => {
        expect(fileExistsSync(path.join(tempDir, "exists.txt"))).toBe(true);
      });

      it("should return false for non-existing file", () => {
        expect(fileExistsSync(path.join(tempDir, "not-exists.txt"))).toBe(false);
      });
    });
  });

  describe("directory creation", () => {
    describe("ensureDir", () => {
      it("should create nested directories", async () => {
        const dirPath = path.join(tempDir, "a", "b", "c");
        await ensureDir(dirPath);

        expect(fs.existsSync(dirPath)).toBe(true);
      });

      it("should not throw if directory exists", async () => {
        const dirPath = path.join(tempDir, "existing");
        await fsp.mkdir(dirPath);

        await expect(ensureDir(dirPath)).resolves.not.toThrow();
      });
    });

    describe("ensureDirSync", () => {
      it("should create directories synchronously", () => {
        const dirPath = path.join(tempDir, "sync", "dir");
        ensureDirSync(dirPath);

        expect(fs.existsSync(dirPath)).toBe(true);
      });
    });
  });

  describe("file stats", () => {
    beforeEach(async () => {
      await fsp.writeFile(path.join(tempDir, "file.txt"), "content");
    });

    describe("safeStats", () => {
      it("should return stats for existing file", async () => {
        const stats = await safeStats(path.join(tempDir, "file.txt"));
        expect(stats).not.toBeNull();
        expect(stats!.isFile()).toBe(true);
      });

      it("should return null for non-existing file", async () => {
        const stats = await safeStats(path.join(tempDir, "not-exists.txt"));
        expect(stats).toBeNull();
      });
    });

    describe("safeStatsSync", () => {
      it("should return stats synchronously", () => {
        const stats = safeStatsSync(path.join(tempDir, "file.txt"));
        expect(stats).not.toBeNull();
        expect(stats!.isFile()).toBe(true);
      });

      it("should return null for non-existing file", () => {
        const stats = safeStatsSync(path.join(tempDir, "not-exists.txt"));
        expect(stats).toBeNull();
      });
    });
  });

  describe("binary I/O", () => {
    describe("readFileBytes / writeFileBytes", () => {
      it("should read and write binary data", async () => {
        const filePath = path.join(tempDir, "binary.bin");
        const data = new Uint8Array([0x00, 0x01, 0x02, 0xff]);

        await writeFileBytes(filePath, data);
        const read = await readFileBytes(filePath);

        expect(read).toEqual(data);
      });
    });

    describe("readFileBytesSync / writeFileBytesSync", () => {
      it("should read and write binary data synchronously", () => {
        const filePath = path.join(tempDir, "binary-sync.bin");
        const data = new Uint8Array([0xaa, 0xbb, 0xcc]);

        writeFileBytesSync(filePath, data);
        const read = readFileBytesSync(filePath);

        expect(read).toEqual(data);
      });
    });
  });

  describe("text I/O", () => {
    describe("readFileText / writeFileText", () => {
      it("should read and write text data", async () => {
        const filePath = path.join(tempDir, "text.txt");
        const content = "Hello, World!";

        await writeFileText(filePath, content);
        const read = await readFileText(filePath);

        expect(read).toBe(content);
      });

      it("should support different encodings", async () => {
        const filePath = path.join(tempDir, "utf16.txt");
        const content = "你好世界";

        await writeFileText(filePath, content, "utf8");
        const read = await readFileText(filePath, "utf8");

        expect(read).toBe(content);
      });
    });

    describe("readFileTextSync / writeFileTextSync", () => {
      it("should read and write text synchronously", () => {
        const filePath = path.join(tempDir, "text-sync.txt");
        const content = "Sync content";

        writeFileTextSync(filePath, content);
        const read = readFileTextSync(filePath);

        expect(read).toBe(content);
      });
    });
  });

  describe("file time", () => {
    describe("setFileTime", () => {
      it("should set file modification time", async () => {
        const filePath = path.join(tempDir, "timed.txt");
        await fsp.writeFile(filePath, "content");

        const newTime = new Date("2024-01-15T10:00:00Z");
        await setFileTime(filePath, newTime);

        const stats = await fsp.stat(filePath);
        expect(stats.mtime.getTime()).toBe(newTime.getTime());
      });
    });

    describe("setFileTimeSync", () => {
      it("should set file time synchronously", async () => {
        const filePath = path.join(tempDir, "timed-sync.txt");
        await fsp.writeFile(filePath, "content");

        const newTime = new Date("2024-06-15T12:00:00Z");
        setFileTimeSync(filePath, newTime);

        const stats = fs.statSync(filePath);
        expect(stats.mtime.getTime()).toBe(newTime.getTime());
      });
    });
  });

  describe("file operations", () => {
    describe("remove", () => {
      it("should remove a file", async () => {
        const filePath = path.join(tempDir, "to-remove.txt");
        await fsp.writeFile(filePath, "content");

        await remove(filePath);

        expect(fs.existsSync(filePath)).toBe(false);
      });

      it("should remove a directory recursively", async () => {
        const dirPath = path.join(tempDir, "dir-to-remove");
        await fsp.mkdir(dirPath);
        await fsp.writeFile(path.join(dirPath, "file.txt"), "content");

        await remove(dirPath);

        expect(fs.existsSync(dirPath)).toBe(false);
      });

      it("should not throw for non-existing path", async () => {
        await expect(remove(path.join(tempDir, "not-exists"))).resolves.not.toThrow();
      });
    });

    describe("removeSync", () => {
      it("should remove synchronously", async () => {
        const filePath = path.join(tempDir, "remove-sync.txt");
        await fsp.writeFile(filePath, "content");

        removeSync(filePath);

        expect(fs.existsSync(filePath)).toBe(false);
      });

      it("should not throw for non-existing path", () => {
        expect(() => removeSync(path.join(tempDir, "not-exists"))).not.toThrow();
      });
    });

    describe("copyFile", () => {
      it("should copy a file", async () => {
        const src = path.join(tempDir, "source.txt");
        const dest = path.join(tempDir, "dest.txt");
        await fsp.writeFile(src, "source content");

        await copyFile(src, dest);

        expect(await fsp.readFile(dest, "utf8")).toBe("source content");
      });

      it("should create destination directory if needed", async () => {
        const src = path.join(tempDir, "src.txt");
        const dest = path.join(tempDir, "new-dir", "dest.txt");
        await fsp.writeFile(src, "content");

        await copyFile(src, dest);

        expect(fs.existsSync(dest)).toBe(true);
      });
    });

    describe("copyFileSync", () => {
      it("should copy synchronously", async () => {
        const src = path.join(tempDir, "src-sync.txt");
        const dest = path.join(tempDir, "dest-sync.txt");
        await fsp.writeFile(src, "sync content");

        copyFileSync(src, dest);

        expect(fs.readFileSync(dest, "utf8")).toBe("sync content");
      });

      it("should create destination directory if needed", async () => {
        const src = path.join(tempDir, "copy-src.txt");
        const dest = path.join(tempDir, "copy-dir", "nested", "dest.txt");
        await fsp.writeFile(src, "nested content");

        copyFileSync(src, dest);

        expect(fs.existsSync(dest)).toBe(true);
        expect(fs.readFileSync(dest, "utf8")).toBe("nested content");
      });
    });
  });

  describe("glob edge cases", () => {
    describe("createGlobMatcher", () => {
      it("should return false for empty pattern array", () => {
        const matcher = createGlobMatcher([]);
        expect(matcher("any/file.txt")).toBe(false);
      });

      it("should match multiple patterns efficiently", () => {
        const matcher = createGlobMatcher(["*.js", "*.ts", "*.json"]);
        expect(matcher("file.js")).toBe(true);
        expect(matcher("file.ts")).toBe(true);
        expect(matcher("file.json")).toBe(true);
        expect(matcher("file.txt")).toBe(false);
      });
    });

    describe("globToRegex caching", () => {
      it("should return cached regex on repeated calls", () => {
        clearGlobCache();
        const regex1 = globToRegex("**/*.ts");
        const regex2 = globToRegex("**/*.ts");
        // Same object reference means it was cached
        expect(regex1).toBe(regex2);
      });

      it("should cache different patterns separately", () => {
        clearGlobCache();
        const regexTs = globToRegex("**/*.ts");
        const regexJs = globToRegex("**/*.js");
        expect(regexTs).not.toBe(regexJs);
      });

      it("should cache dot option variations separately", () => {
        clearGlobCache();
        const regexNoDot = globToRegex("*.txt", { dot: false });
        const regexDot = globToRegex("*.txt", { dot: true });
        expect(regexNoDot).not.toBe(regexDot);
      });
    });

    describe("glob with non-existent cwd", () => {
      it("should return empty results when cwd does not exist", async () => {
        const files: string[] = [];
        for await (const entry of glob("**/*.ts", { cwd: "/non/existent/path" })) {
          files.push(entry.relativePath);
        }
        expect(files).toHaveLength(0);
      });

      it("globSync should return empty array when cwd does not exist", () => {
        const files = globSync("**/*.ts", { cwd: "/non/existent/path" });
        expect(files).toHaveLength(0);
      });
    });

    describe("matchGlobAny with empty patterns", () => {
      it("should return false for empty pattern array", () => {
        expect(matchGlobAny("file.txt", [])).toBe(false);
      });
    });

    describe("globToRegex special patterns", () => {
      it("should handle negated character classes [^abc]", () => {
        const regex = globToRegex("file[^0-9].txt");
        expect(regex.test("filea.txt")).toBe(true);
        expect(regex.test("file1.txt")).toBe(false);
      });

      it("should handle escaped brackets in character class", () => {
        const regex = globToRegex("file[\\]].txt");
        expect(regex.test("file].txt")).toBe(true);
      });

      it("should handle ** followed by file extension", () => {
        const regex = globToRegex("**.ts");
        expect(regex.test("file.ts")).toBe(true);
        expect(regex.test("src/file.ts")).toBe(true);
      });
    });
  });

  describe("traverseDirectory edge cases", () => {
    it("should handle empty directory", async () => {
      const emptyDir = path.join(tempDir, "empty");
      await fsp.mkdir(emptyDir);

      const entries: string[] = [];
      for await (const entry of traverseDirectory(emptyDir)) {
        entries.push(entry.relativePath);
      }
      expect(entries).toHaveLength(0);
    });

    it("should handle deeply nested directories", async () => {
      const deepPath = path.join(tempDir, "a", "b", "c", "d", "e");
      await fsp.mkdir(deepPath, { recursive: true });
      await fsp.writeFile(path.join(deepPath, "deep.txt"), "deep");

      const entries: string[] = [];
      for await (const entry of traverseDirectory(tempDir)) {
        entries.push(entry.relativePath);
      }

      expect(entries).toContain(path.join("a", "b", "c", "d", "e", "deep.txt"));
    });

    it("traverseDirectorySync should handle empty directory", () => {
      const emptyDir = path.join(tempDir, "empty-sync");
      fs.mkdirSync(emptyDir);

      const entries = traverseDirectorySync(emptyDir);
      expect(entries).toHaveLength(0);
    });
  });
});
