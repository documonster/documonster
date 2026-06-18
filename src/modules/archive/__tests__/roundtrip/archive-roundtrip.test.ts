/**
 * Archive Roundtrip Tests
 *
 * Comprehensive roundtrip tests for both ZIP and TAR formats.
 * Tests verify that compression followed by decompression produces identical data.
 *
 * Coverage:
 * - ZIP: bytes() and stream() creation -> unzip() extraction
 * - TAR: bytes() and stream() creation -> parseTar() / TarReader extraction
 * - Unified API: zip({ format: 'tar' }) -> unzip({ format: 'tar' })
 * - Edge cases: Unicode, binary, empty files, long paths, special characters
 * - Performance: many small files, large files
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { TarArchive, parseTar, TarReader, TAR_TYPE } from "@archive/tar/index.browser";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect } from "vitest";

// =============================================================================
// Types & Constants
// =============================================================================

interface TestFileEntry {
  path: string;
  content: string | Uint8Array | null;
  isDirectory?: boolean;
}

const textEncoder = new TextEncoder();

// =============================================================================
// Test Fixtures
// =============================================================================

/**
 * Generates deterministic pseudo-random bytes using a simple LCG.
 * Useful for creating incompressible test data.
 */
function generateRandomBytes(size: number, seed = 12345): Uint8Array {
  const data = new Uint8Array(size);
  let state = seed;
  for (let i = 0; i < size; i++) {
    state = (state * 1103515245 + 12345) & 0x7fffffff;
    data[i] = state & 0xff;
  }
  return data;
}

const FIXTURES: Record<string, TestFileEntry[]> = {
  // Basic cases
  "single-text-file": [{ path: "hello.txt", content: "Hello, World!" }],

  "multiple-text-files": [
    { path: "file1.txt", content: "Content of file 1" },
    { path: "file2.txt", content: "Content of file 2" },
    { path: "file3.txt", content: "Content of file 3" }
  ],

  "directory-structure": [
    { path: "root.txt", content: "Root file content" },
    { path: "src/", content: null, isDirectory: true },
    { path: "src/index.ts", content: 'export const main = () => console.log("Hello");' },
    { path: "src/utils/", content: null, isDirectory: true },
    { path: "src/utils/helper.ts", content: "export const helper = (x: number) => x * 2;" }
  ],

  // Binary data
  "binary-all-bytes": [
    {
      path: "binary.bin",
      content: (() => {
        const data = new Uint8Array(256);
        for (let i = 0; i < 256; i++) {
          data[i] = i;
        }
        return data;
      })()
    }
  ],

  "mixed-content": [
    { path: "readme.txt", content: "This is a readme file" },
    { path: "data.bin", content: new Uint8Array([0x00, 0x01, 0xff, 0xfe, 0x42, 0x00, 0xff]) },
    { path: "config.json", content: '{"name": "test", "version": "1.0.0"}' }
  ],

  // Unicode - both content and filenames
  "unicode-content": [
    { path: "chinese.txt", content: "中文内容 - Chinese content" },
    { path: "japanese.txt", content: "日本語コンテンツ - Japanese content" },
    { path: "emoji.txt", content: "Emoji: 🎉🚀🌍💻🔥" },
    { path: "mixed-scripts.txt", content: "Hello 世界 مرحبا שלום Привет" }
  ],

  "unicode-filenames": [
    { path: "文件.txt", content: "Chinese filename" },
    { path: "ファイル.txt", content: "Japanese filename" },
    { path: "🎉party🎊.txt", content: "Emoji filename" },
    { path: "данные.txt", content: "Russian filename" }
  ],

  // Edge cases
  "empty-file": [{ path: "empty.txt", content: "" }],

  "empty-archive": [],

  "only-directories": [
    { path: "dir1/", content: null, isDirectory: true },
    { path: "dir1/subdir/", content: null, isDirectory: true },
    { path: "dir2/", content: null, isDirectory: true }
  ],

  // Size edge cases
  "large-file-100kb": [{ path: "large.bin", content: generateRandomBytes(100 * 1024) }],

  "exact-block-sizes": [
    { path: "512bytes.bin", content: generateRandomBytes(512) }, // TAR block size
    { path: "1024bytes.bin", content: generateRandomBytes(1024) },
    { path: "513bytes.bin", content: generateRandomBytes(513) } // Just over block boundary
  ],

  // Complex realistic structure
  "complex-project": [
    { path: "README.md", content: "# Project\n\nThis is a test project." },
    { path: "package.json", content: '{"name": "test-pkg", "version": "0.1.0"}' },
    { path: ".gitignore", content: "node_modules/\ndist/\n*.log" },
    { path: "src/", content: null, isDirectory: true },
    { path: "src/index.ts", content: "export * from './lib';" },
    { path: "src/lib/", content: null, isDirectory: true },
    { path: "src/lib/core.ts", content: "export class Core { run() {} }" },
    { path: "tests/", content: null, isDirectory: true },
    { path: "tests/core.test.ts", content: "describe('Core', () => { it('works', () => {}) });" }
  ]
};

// =============================================================================
// Helper Functions
// =============================================================================

function contentToBytes(content: string | Uint8Array): Uint8Array {
  return typeof content === "string" ? textEncoder.encode(content) : content;
}

function getExpectedFiles(entries: TestFileEntry[]): Map<string, Uint8Array> {
  const map = new Map<string, Uint8Array>();
  for (const entry of entries) {
    if (entry.content !== null && !entry.isDirectory) {
      map.set(entry.path, contentToBytes(entry.content));
    }
  }
  return map;
}

function assertFilesEqual(
  actual: Map<string, Uint8Array>,
  expected: Map<string, Uint8Array>
): void {
  expect(actual.size, "File count mismatch").toBe(expected.size);
  for (const [path, expectedContent] of expected) {
    const actualContent = actual.get(path);
    expect(actualContent, `Missing file: ${path}`).toBeDefined();
    expect(actualContent, `Content mismatch: ${path}`).toEqual(expectedContent);
  }
}

async function collectStream(stream: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream) {
    chunks.push(chunk);
  }
  return concatUint8Arrays(chunks);
}

// =============================================================================
// Archive Adapters (Abstract ZIP/TAR differences)
// =============================================================================

interface ArchiveAdapter {
  create(entries: TestFileEntry[]): {
    bytes(): Promise<Uint8Array>;
    stream(): AsyncIterable<Uint8Array>;
  };
  extract(data: Uint8Array): Promise<Map<string, Uint8Array>>;
}

const zipAdapter: ArchiveAdapter = {
  create(entries) {
    const archive = zip();
    for (const entry of entries) {
      if (!entry.isDirectory && entry.content !== null) {
        archive.add(entry.path, contentToBytes(entry.content));
      }
    }
    return archive;
  },
  async extract(data) {
    const files = new Map<string, Uint8Array>();
    for await (const entry of unzip(data).entries()) {
      if (entry.type === "directory") {
        entry.discard();
      } else {
        files.set(entry.path, await entry.bytes());
      }
    }
    return files;
  }
};

const tarAdapter: ArchiveAdapter = {
  create(entries) {
    const archive = new TarArchive();
    for (const entry of entries) {
      if (entry.isDirectory) {
        archive.addDirectory(entry.path.replace(/\/$/, ""));
      } else if (entry.content !== null) {
        archive.add(entry.path, contentToBytes(entry.content));
      }
    }
    return archive;
  },
  async extract(data) {
    const files = new Map<string, Uint8Array>();
    const parsed = parseTar(data);
    for (const entry of parsed) {
      if (entry.info.type === TAR_TYPE.FILE || entry.info.type === TAR_TYPE.FILE_OLD) {
        files.set(entry.info.path, await entry.data());
      }
    }
    return files;
  }
};

const tarReaderAdapter: ArchiveAdapter = {
  ...tarAdapter,
  async extract(data) {
    const files = new Map<string, Uint8Array>();
    for await (const entry of new TarReader(data).entries()) {
      if (!entry.isDirectory) {
        files.set(entry.path, await entry.bytes());
      }
    }
    return files;
  }
};

const unifiedTarAdapter: ArchiveAdapter = {
  create(entries) {
    const archive = zip({ format: "tar" });
    for (const entry of entries) {
      if (!entry.isDirectory && entry.content !== null) {
        archive.add(entry.path, contentToBytes(entry.content));
      }
    }
    return archive;
  },
  async extract(data) {
    const files = new Map<string, Uint8Array>();
    for await (const entry of unzip(data, { format: "tar" }).entries()) {
      if (!entry.isDirectory) {
        files.set(entry.path, await entry.bytes());
      }
    }
    return files;
  }
};

// =============================================================================
// Parameterized Roundtrip Tests
// =============================================================================

function runRoundtripTests(name: string, adapter: ArchiveAdapter): void {
  describe(name, () => {
    describe("bytes() roundtrip", () => {
      for (const [fixtureName, entries] of Object.entries(FIXTURES)) {
        it(`should roundtrip: ${fixtureName}`, async () => {
          const archiveBytes = await adapter.create(entries).bytes();
          const extracted = await adapter.extract(archiveBytes);
          assertFilesEqual(extracted, getExpectedFiles(entries));
        });
      }
    });

    describe("stream() roundtrip", () => {
      for (const [fixtureName, entries] of Object.entries(FIXTURES)) {
        it(`should roundtrip via stream: ${fixtureName}`, async () => {
          const archiveBytes = await collectStream(adapter.create(entries).stream());
          const extracted = await adapter.extract(archiveBytes);
          assertFilesEqual(extracted, getExpectedFiles(entries));
        });
      }
    });

    it("bytes() and stream() should produce equivalent archives", async () => {
      const entries = FIXTURES["multiple-text-files"];
      const bytesArchive = await adapter.create(entries).bytes();
      const streamArchive = await collectStream(adapter.create(entries).stream());

      const bytesExtracted = await adapter.extract(bytesArchive);
      const streamExtracted = await adapter.extract(streamArchive);

      assertFilesEqual(bytesExtracted, streamExtracted);
    });
  });
}

// Run parameterized tests for each format
runRoundtripTests("ZIP Roundtrip", zipAdapter);
runRoundtripTests("TAR Roundtrip (parseTar)", tarAdapter);
runRoundtripTests("TAR Roundtrip (TarReader)", tarReaderAdapter);
runRoundtripTests("Unified API TAR Roundtrip", unifiedTarAdapter);

// =============================================================================
// TAR-specific: bytes() and stream() should be byte-identical
// =============================================================================

describe("TAR byte-for-byte consistency", () => {
  it("bytes() and stream() should produce identical bytes", async () => {
    const entries = FIXTURES["multiple-text-files"];
    // Pin modTime so two separate TarArchive instances produce identical headers.
    const fixedTime = new Date("2024-01-01T00:00:00Z");
    const createArchive = () => {
      const archive = new TarArchive({ modTime: fixedTime });
      for (const entry of entries) {
        if (entry.isDirectory) {
          archive.addDirectory(entry.path.replace(/\/$/, ""));
        } else if (entry.content !== null) {
          archive.add(entry.path, contentToBytes(entry.content));
        }
      }
      return archive;
    };
    const bytesOutput = await createArchive().bytes();
    const streamOutput = await collectStream(createArchive().stream());
    expect(streamOutput).toEqual(bytesOutput);
  });
});

// =============================================================================
// Edge Case Tests
// =============================================================================

describe("Edge Cases", () => {
  describe("Special filenames", () => {
    const specialPaths = [
      "file with spaces.txt",
      "file-with-dashes.txt",
      "file_with_underscores.txt",
      "file.multiple.dots.txt",
      "file(with)parens.txt",
      "file[with]brackets.txt",
      "file@special#chars.txt",
      "UPPERCASE.TXT",
      "MixedCase.Txt"
    ];

    it("ZIP should handle special characters in filenames", async () => {
      const archive = zip();
      for (const path of specialPaths) {
        archive.add(path, textEncoder.encode(`content of ${path}`));
      }
      const extracted = await zipAdapter.extract(await archive.bytes());
      expect(extracted.size).toBe(specialPaths.length);
      for (const path of specialPaths) {
        expect(extracted.has(path), `Missing: ${path}`).toBe(true);
      }
    });

    it("TAR should handle special characters in filenames", async () => {
      const archive = new TarArchive();
      for (const path of specialPaths) {
        archive.add(path, `content of ${path}`);
      }
      const extracted = await tarAdapter.extract(await archive.bytes());
      expect(extracted.size).toBe(specialPaths.length);
      for (const path of specialPaths) {
        expect(extracted.has(path), `Missing: ${path}`).toBe(true);
      }
    });
  });

  describe("Path depth", () => {
    const deepPath = Array.from({ length: 20 }, (_, i) => `d${i}`).join("/") + "/file.txt";

    it("ZIP should handle deeply nested paths (20 levels)", async () => {
      const archive = zip();
      archive.add(deepPath, textEncoder.encode("deep"));
      const entry = await unzip(await archive.bytes()).get(deepPath);
      expect(entry).not.toBeNull();
      expect(new TextDecoder().decode(await entry!.bytes())).toBe("deep");
    });

    it("TAR should handle deeply nested paths (20 levels)", async () => {
      const archive = new TarArchive();
      archive.add(deepPath, "deep");
      const parsed = parseTar(await archive.bytes());
      expect(parsed[0].info.path).toBe(deepPath);
      expect(await parsed[0].text()).toBe("deep");
    });
  });

  describe("TAR long filename support", () => {
    it("should handle 100-155 char paths (ustar prefix)", async () => {
      const path = "dir/".repeat(25) + "file.txt"; // ~104 chars
      const archive = new TarArchive();
      archive.add(path, "content");
      const parsed = parseTar(await archive.bytes());
      expect(parsed[0].info.path).toBe(path);
    });

    it("should handle >155 char paths (GNU extension)", async () => {
      const path = "a".repeat(200) + ".txt";
      const archive = new TarArchive();
      archive.add(path, "content");
      const parsed = parseTar(await archive.bytes());
      expect(parsed[0].info.path).toBe(path);
    });

    it("should handle >500 char paths", async () => {
      const path = "very_long_directory_name/".repeat(25) + "file.txt";
      expect(path.length).toBeGreaterThan(500);
      const archive = new TarArchive();
      archive.add(path, "content");
      const parsed = parseTar(await archive.bytes());
      expect(parsed[0].info.path).toBe(path);
    });
  });

  describe("Binary content edge cases", () => {
    const cases: [string, Uint8Array][] = [
      ["null bytes at start", new Uint8Array([0x00, 0x00, 0x00, 0x41, 0x42])],
      ["null bytes at end", new Uint8Array([0x41, 0x42, 0x00, 0x00, 0x00])],
      ["all 0xFF", new Uint8Array(100).fill(0xff)],
      ["all 0x00", new Uint8Array(100).fill(0x00)],
      [
        "alternating bytes",
        (() => {
          const d = new Uint8Array(100);
          for (let i = 0; i < 100; i++) {
            d[i] = i % 2 === 0 ? 0x00 : 0xff;
          }
          return d;
        })()
      ],
      ["random incompressible", generateRandomBytes(1000)]
    ];

    for (const [name, content] of cases) {
      it(`ZIP should handle: ${name}`, async () => {
        const archive = zip();
        archive.add("test.bin", content);
        const extracted = await zipAdapter.extract(await archive.bytes());
        expect(extracted.get("test.bin")).toEqual(content);
      });

      it(`TAR should handle: ${name}`, async () => {
        const archive = new TarArchive();
        archive.add("test.bin", content);
        const extracted = await tarAdapter.extract(await archive.bytes());
        expect(extracted.get("test.bin")).toEqual(content);
      });
    }
  });

  describe("Whitespace content", () => {
    const whitespaceContent = "   \n\t\r\n   ";

    it("ZIP should preserve whitespace-only content", async () => {
      const archive = zip();
      archive.add("ws.txt", textEncoder.encode(whitespaceContent));
      const entry = await unzip(await archive.bytes()).get("ws.txt");
      expect(new TextDecoder().decode(await entry!.bytes())).toBe(whitespaceContent);
    });

    it("TAR should preserve whitespace-only content", async () => {
      const archive = new TarArchive();
      archive.add("ws.txt", whitespaceContent);
      const parsed = parseTar(await archive.bytes());
      expect(await parsed[0].text()).toBe(whitespaceContent);
    });
  });

  describe("Duplicate content optimization check", () => {
    it("should correctly store multiple files with identical content", async () => {
      const content = textEncoder.encode("identical content repeated");

      // ZIP
      const zipArchive = zip();
      zipArchive.add("a.txt", content);
      zipArchive.add("b.txt", content);
      zipArchive.add("c.txt", content);
      const zipExtracted = await zipAdapter.extract(await zipArchive.bytes());
      expect(zipExtracted.size).toBe(3);
      expect(zipExtracted.get("a.txt")).toEqual(content);
      expect(zipExtracted.get("b.txt")).toEqual(content);
      expect(zipExtracted.get("c.txt")).toEqual(content);

      // TAR
      const tarArchive = new TarArchive();
      tarArchive.add("a.txt", content);
      tarArchive.add("b.txt", content);
      tarArchive.add("c.txt", content);
      const tarExtracted = await tarAdapter.extract(await tarArchive.bytes());
      expect(tarExtracted.size).toBe(3);
      expect(tarExtracted.get("a.txt")).toEqual(content);
      expect(tarExtracted.get("b.txt")).toEqual(content);
      expect(tarExtracted.get("c.txt")).toEqual(content);
    });
  });
});

// =============================================================================
// Performance / Stress Tests
// =============================================================================

describe("Performance Tests", () => {
  it("should roundtrip 1000 small files via ZIP", async () => {
    const entries: TestFileEntry[] = [];
    for (let i = 0; i < 1000; i++) {
      entries.push({
        path: `file${i.toString().padStart(4, "0")}.txt`,
        content: `Content of file ${i}`
      });
    }
    const archiveBytes = await zipAdapter.create(entries).bytes();
    const extracted = await zipAdapter.extract(archiveBytes);
    assertFilesEqual(extracted, getExpectedFiles(entries));
  });

  it("should roundtrip 1000 small files via TAR", async () => {
    const entries: TestFileEntry[] = [];
    for (let i = 0; i < 1000; i++) {
      entries.push({
        path: `file${i.toString().padStart(4, "0")}.txt`,
        content: `Content of file ${i}`
      });
    }
    const archiveBytes = await tarAdapter.create(entries).bytes();
    const extracted = await tarAdapter.extract(archiveBytes);
    assertFilesEqual(extracted, getExpectedFiles(entries));
  });

  it("should roundtrip 1MB file via ZIP", async () => {
    const content = generateRandomBytes(1024 * 1024);
    const archive = zip();
    archive.add("large.bin", content);
    const extracted = await zipAdapter.extract(await archive.bytes());
    expect(extracted.get("large.bin")).toEqual(content);
  });

  it("should roundtrip 1MB file via TAR", async () => {
    const content = generateRandomBytes(1024 * 1024);
    const archive = new TarArchive();
    archive.add("large.bin", content);
    const extracted = await tarAdapter.extract(await archive.bytes());
    expect(extracted.get("large.bin")).toEqual(content);
  });

  it("should roundtrip highly compressible data via ZIP", async () => {
    // Data that compresses very well (repeated pattern)
    const content = new Uint8Array(100 * 1024).fill(0x41); // 100KB of 'A'
    const archive = zip();
    archive.add("compressible.txt", content);
    const archiveBytes = await archive.bytes();

    // Verify compression actually happened
    expect(archiveBytes.length).toBeLessThan(content.length / 10);

    const extracted = await zipAdapter.extract(archiveBytes);
    expect(extracted.get("compressible.txt")).toEqual(content);
  });
});

// =============================================================================
// Compression Level Tests (ZIP only)
// =============================================================================

describe("ZIP Compression Levels", () => {
  const content = textEncoder.encode("a".repeat(10000));

  it("level 0 (store) should not compress", async () => {
    const archive = zip({ level: 0 });
    archive.add("test.txt", content);
    const bytes = await archive.bytes();
    // Store mode: archive should be larger than content (headers + content)
    expect(bytes.length).toBeGreaterThan(content.length);
  });

  it("level 9 (max) should produce smaller output than level 1", async () => {
    const archive1 = zip({ level: 1 });
    archive1.add("test.txt", content);
    const bytes1 = await archive1.bytes();

    const archive9 = zip({ level: 9 });
    archive9.add("test.txt", content);
    const bytes9 = await archive9.bytes();

    expect(bytes9.length).toBeLessThanOrEqual(bytes1.length);

    // Both should extract correctly
    const extracted1 = await zipAdapter.extract(bytes1);
    const extracted9 = await zipAdapter.extract(bytes9);
    expect(extracted1.get("test.txt")).toEqual(content);
    expect(extracted9.get("test.txt")).toEqual(content);
  });
});
