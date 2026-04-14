/**
 * Comprehensive tests for the unzip streaming Parse class.
 *
 * Covers:
 * 1. Parse class basics — entry parsing, content, metadata, factory
 * 2. Consumption modes — on("entry"), for-await, data event, consistency
 * 3. Autodrain / entry lifecycle
 * 4. Error propagation
 * 5. Entry interruption / destroy
 * 6. Premature close prevention (Readable.from, slow consumer)
 * 7. Stress + stability
 */

import { createReadStream } from "fs";
import { join } from "path";
import { Readable } from "stream";

import { unzip, zip } from "@archive";
import { Parse, createParse, type ZipEntry } from "@archive/unzip/stream";
import { PassThrough } from "@stream";
import { describe, it, expect } from "vitest";

import { delay, chunkBytes, createDataDescriptorZip } from "./test-helpers";

const testFilePath = join(__dirname, "./data/formulas.xlsx");

// =============================================================================
// 1. Parse class basics
// =============================================================================

describe("parse: basics", () => {
  it("should parse a zip file and emit entries via for-await", async () => {
    const entries: string[] = [];
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      entries.push((entry as ZipEntry).path);
      (entry as ZipEntry).autodrain();
    }

    expect(entries.length).toBeGreaterThan(0);
    expect(entries).toContain("[Content_Types].xml");
    expect(entries.some(e => e.includes("xl/workbook.xml"))).toBe(true);
  });

  it("should parse file content correctly", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    let contentTypesContent = "";

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      if (zipEntry.path === "[Content_Types].xml") {
        const buffer = await zipEntry.buffer();
        contentTypesContent = new TextDecoder().decode(buffer);
      } else {
        zipEntry.autodrain();
      }
    }

    expect(contentTypesContent).toContain("<?xml");
    expect(contentTypesContent).toContain("ContentType");
  });

  it("should provide entry type (File or Directory)", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    let hasFile = false;
    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      if (zipEntry.type === "File") {
        hasFile = true;
      }
      zipEntry.autodrain();
    }

    expect(hasFile).toBe(true);
  });

  it("should provide entry vars with compression info", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      expect(zipEntry.vars).toBeDefined();
      expect(typeof zipEntry.vars.compressionMethod).toBe("number");
      expect(typeof zipEntry.vars.compressedSize).toBe("number");
      expect(typeof zipEntry.vars.uncompressedSize).toBe("number");
      zipEntry.autodrain();
      break;
    }
  });

  it("should set entry size after reading", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      if (zipEntry.type === "File" && zipEntry.vars.uncompressedSize! > 0) {
        const buffer = await zipEntry.buffer();
        expect(buffer.length).toBe(zipEntry.vars.uncompressedSize);
        break;
      } else {
        zipEntry.autodrain();
      }
    }
  });

  it("should provide entry props with flags", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      expect(zipEntry.props).toBeDefined();
      expect(zipEntry.props.path).toBe(zipEntry.path);
      expect(zipEntry.props.pathBuffer).toBeInstanceOf(Uint8Array);
      expect(typeof zipEntry.props.flags.isUnicode).toBe("boolean");
      zipEntry.autodrain();
      break;
    }
  });

  it("should provide lastModifiedDateTime in entry vars", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      expect(zipEntry.vars.lastModifiedDateTime).toBeInstanceOf(Date);
      zipEntry.autodrain();
      break;
    }
  });

  it("should parse archive with low highWaterMark (chunk boundary test)", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath, { highWaterMark: 3 });
    stream.pipe(parse);

    for await (const entry of parse) {
      (entry as ZipEntry).autodrain();
    }
  });
});

// =============================================================================
// 2. Consumption modes
// =============================================================================

describe("parse: consumption modes", () => {
  it("forceStream should emit data event instead of entry event", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);

    let dataEventEmitted = false;
    let entryEventEmitted = false;

    parse.on("data", (entry: ZipEntry) => {
      expect(entry).toBeInstanceOf(PassThrough);
      dataEventEmitted = true;
      entry.autodrain();
    });

    parse.on("entry", () => {
      entryEventEmitted = true;
    });

    stream.pipe(parse);
    await parse.promise();

    expect(dataEventEmitted).toBe(true);
    expect(entryEventEmitted).toBe(false);
  });

  it("on('entry') + promise() should work without for-await", async () => {
    const parse = new Parse();
    const entries: string[] = [];

    parse.on("entry", (entry: ZipEntry) => {
      entries.push(entry.path);
      entry.autodrain();
    });

    const stream = createReadStream(testFilePath);
    stream.pipe(parse);
    await parse.promise();

    expect(entries.length).toBeGreaterThan(0);
  });

  it("promise() should resolve when entries have been consumed via on('entry')", async () => {
    const parse = new Parse();
    const stream = createReadStream(testFilePath);
    let entryRead = false;

    parse.on("entry", (entry: ZipEntry) => {
      if (entry.path === "[Content_Types].xml") {
        entry.buffer().then(() => {
          entryRead = true;
        });
      } else {
        entry.autodrain();
      }
    });

    stream.pipe(parse);
    await parse.promise();
    expect(entryRead).toBe(true);
  });

  it(
    "on('entry') + promise() and for-await should yield same entry paths",
    { timeout: 10_000 },
    async () => {
      // Mode A: on("entry") + promise()
      const pathsA: string[] = [];
      const parseA = new Parse();
      parseA.on("entry", (entry: ZipEntry) => {
        pathsA.push(entry.path);
        entry.autodrain();
      });
      const streamA = createReadStream(testFilePath);
      streamA.pipe(parseA);
      await parseA.promise();

      // Mode B: for await (forceStream)
      const pathsB: string[] = [];
      const parseB = createParse({ forceStream: true });
      const streamB = createReadStream(testFilePath);
      streamB.pipe(parseB);
      for await (const entry of parseB) {
        pathsB.push((entry as ZipEntry).path);
        (entry as ZipEntry).autodrain();
      }

      expect(pathsA.length).toBeGreaterThan(0);
      expect(pathsA).toEqual(pathsB);
    }
  );

  it("unzip().entries() and on('entry') should yield same paths", { timeout: 10_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("file-1"))
      .add("b.txt", new TextEncoder().encode("file-2"))
      .add("c.txt", new TextEncoder().encode("file-3"))
      .bytes();

    // Mode A: unzip().entries()
    const pathsA: string[] = [];
    const readerA = unzip(chunkBytes(zipBytes, 128));
    for await (const entry of readerA.entries()) {
      pathsA.push(entry.path);
      entry.discard();
    }

    // Mode B: on("entry") + promise()
    const pathsB: string[] = [];
    const parseB = new Parse();
    parseB.on("entry", (entry: ZipEntry) => {
      pathsB.push(entry.path);
      entry.autodrain();
    });
    const readable = Readable.from(chunkBytes(zipBytes, 128));
    readable.pipe(parseB);
    await parseB.promise();

    expect(pathsA).toEqual(pathsB);
  });
});

// =============================================================================
// 3. Autodrain / entry lifecycle
// =============================================================================

describe("parse: autodrain", () => {
  it("immediate autodrain should not unzip content", async () => {
    const parse = new Parse();
    const stream = createReadStream(testFilePath);

    parse.on("entry", (entry: ZipEntry) => {
      entry.autodrain().on("finish", () => {
        expect(entry.__autodraining).toBe(true);
      });
    });

    stream.pipe(parse);
    await parse.promise();
  });

  it("autodrain().promise() should resolve", async () => {
    const parse = new Parse();
    const stream = createReadStream(testFilePath);

    parse.on("entry", (entry: ZipEntry) => {
      entry
        .autodrain()
        .promise()
        .then(() => {
          expect(entry.__autodraining).toBe(true);
        });
    });

    stream.pipe(parse);
    await parse.promise();
  });

  it("autodrain().promise() should work via for-await", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    for await (const entry of parse) {
      const zipEntry = entry as ZipEntry;
      await zipEntry.autodrain().promise();
      break;
    }
  });
});

// =============================================================================
// 4. Error propagation
// =============================================================================

describe("parse: error propagation", () => {
  it("invalid ZIP should reject promise()", async () => {
    const parse = new Parse();
    const nonArchive = join(__dirname, "../../../../../package.json");
    const stream = createReadStream(nonArchive);
    stream.pipe(parse);

    await expect(parse.promise()).rejects.toThrow(/invalid signature/);
  });

  it("invalid ZIP should throw in for-await iterator", async () => {
    const parse = createParse({ forceStream: true });
    const nonArchive = join(__dirname, "../../../../../package.json");
    const stream = createReadStream(nonArchive);
    stream.pipe(parse);

    await expect(async () => {
      for await (const entry of parse) {
        (entry as ZipEntry).autodrain();
      }
    }).rejects.toThrow(/invalid signature/);
  });

  it("user-emitted error should reject promise()", async () => {
    const parse = new Parse();
    const stream = createReadStream(testFilePath);

    parse.on("entry", () => {
      parse.emit("error", new Error("user error"));
    });

    stream.pipe(parse);
    await expect(parse.promise()).rejects.toThrow("user error");
  });

  it("user-emitted error should propagate to for-await iterator", async () => {
    const parse = createParse({ forceStream: true });
    const stream = createReadStream(testFilePath);
    stream.pipe(parse);

    let count = 0;
    await expect(async () => {
      for await (const entry of parse) {
        count++;
        if (count === 1) {
          parse.emit("error", new Error("mid-iteration error"));
        }
        (entry as ZipEntry).autodrain();
      }
    }).rejects.toThrow("mid-iteration error");
  });

  it("invalid ZIP via unzip() entries() should throw", async () => {
    const garbage = new Uint8Array(256).fill(0xff);

    await expect(async () => {
      const reader = unzip(garbage);
      for await (const entry of reader.entries()) {
        entry.discard();
      }
    }).rejects.toThrow();
  });
});

// =============================================================================
// 5. Entry interruption / destroy
// =============================================================================

describe("parse: entry interruption", () => {
  it("breaking from for-await over entries should not hang", { timeout: 10_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("1.txt", new TextEncoder().encode("one"))
      .add("2.txt", new TextEncoder().encode("two"))
      .add("3.txt", new TextEncoder().encode("three"))
      .bytes();

    const reader = unzip(chunkBytes(zipBytes, 128));
    let seen = 0;

    for await (const entry of reader.entries()) {
      seen++;
      entry.discard();
      if (seen === 1) {
        break;
      }
    }

    expect(seen).toBe(1);
  });

  it(
    "break from entry.stream() should let parser continue to next entry",
    { timeout: 10_000 },
    async () => {
      const zipBytes = await zip({ level: 0 })
        .add("first.bin", new Uint8Array(1024 * 32))
        .add("second.txt", new TextEncoder().encode("still here"))
        .bytes();

      const reader = unzip(chunkBytes(zipBytes, 4096));
      const seen: string[] = [];

      for await (const entry of reader.entries()) {
        seen.push(entry.path);

        if (entry.path.endsWith(".bin")) {
          let chunks = 0;
          for await (const _chunk of entry.stream()) {
            chunks++;
            if (chunks >= 1) {
              break;
            }
          }
        } else {
          const data = await entry.bytes();
          expect(new TextDecoder().decode(data)).toBe("still here");
        }
      }

      expect(seen.length).toBe(2);
      expect(seen).toContain("first.bin");
      expect(seen).toContain("second.txt");
    }
  );

  it(
    "alternating stream/discard/bytes consumption across entries",
    { timeout: 10_000 },
    async () => {
      const z = zip({ level: 0 });
      for (let i = 0; i < 6; i++) {
        z.add(`file-${i}.txt`, new TextEncoder().encode(`content-${i}`));
      }
      const zipBytes = await z.bytes();

      const reader = unzip(chunkBytes(zipBytes, 256));
      const results: string[] = [];
      let idx = 0;

      for await (const entry of reader.entries()) {
        const mode = idx % 3;
        if (mode === 0) {
          const chunks: Uint8Array[] = [];
          for await (const chunk of entry.stream()) {
            chunks.push(chunk);
          }
          const text = new TextDecoder().decode(
            chunks.length === 1
              ? chunks[0]
              : new Uint8Array(chunks.reduce((s, c) => s + c.length, 0))
          );
          results.push(text);
        } else if (mode === 1) {
          const data = await entry.bytes();
          results.push(new TextDecoder().decode(data));
        } else {
          entry.discard();
          results.push("discarded");
        }
        idx++;
      }

      expect(results).toEqual([
        "content-0",
        "content-1",
        "discarded",
        "content-3",
        "content-4",
        "discarded"
      ]);
    }
  );
});

// =============================================================================
// 6. Premature close prevention
// =============================================================================

describe("parse: premature close prevention", () => {
  it(
    "Readable.from(entry.stream()) should not throw Premature close",
    { timeout: 10_000 },
    async () => {
      const content = new Uint8Array(4096).fill(67);
      const zipBytes = await zip({ level: 0 }).add("data.bin", content).bytes();

      const reader = unzip(chunkBytes(zipBytes, 256));

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const readable = Readable.from(entry.stream());
        const chunks: Buffer[] = [];
        for await (const chunk of readable) {
          chunks.push(Buffer.from(chunk));
        }
        const result = Buffer.concat(chunks);
        expect(new Uint8Array(result)).toEqual(content);
      }
    }
  );

  it(
    "Readable.from(entry.stream()) with multiple entries and slow consumer",
    { timeout: 15_000 },
    async () => {
      const z = zip({ level: 0 });
      for (let i = 0; i < 5; i++) {
        z.add(`r-${i}.bin`, new Uint8Array(2048).fill(i));
      }
      const zipBytes = await z.bytes();

      const reader = unzip(chunkBytes(zipBytes, 512));
      let count = 0;

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const readable = Readable.from(entry.stream());
        const chunks: Buffer[] = [];
        for await (const chunk of readable) {
          await delay(1);
          chunks.push(Buffer.from(chunk));
        }
        const result = Buffer.concat(chunks);
        expect(result.length).toBe(2048);
        expect(result[0]).toBe(count);
        count++;
      }

      expect(count).toBe(5);
    }
  );

  it(
    "data descriptor ZIP: Readable.from(entry.stream()) should work",
    { timeout: 15_000 },
    async () => {
      const content1 = new TextEncoder().encode("dd-file-1");
      const content2 = new TextEncoder().encode("dd-file-2-longer");

      const zipBytes = await createDataDescriptorZip([
        { name: "dd1.txt", data: content1 },
        { name: "dd2.txt", data: content2 }
      ]);

      const reader = unzip(chunkBytes(zipBytes, 32));
      const results = new Map<string, Uint8Array>();

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const readable = Readable.from(entry.stream());
        const chunks: Buffer[] = [];
        for await (const chunk of readable) {
          chunks.push(Buffer.from(chunk));
        }
        results.set(entry.path, new Uint8Array(Buffer.concat(chunks)));
      }

      expect(results.size).toBe(2);
      expect(results.get("dd1.txt")).toEqual(content1);
      expect(results.get("dd2.txt")).toEqual(content2);
    }
  );
});

// =============================================================================
// 7. Stress + stability
// =============================================================================

describe("parse: stress + stability", () => {
  it("100 small files via entry.stream()", { timeout: 20_000 }, async () => {
    const count = 100;
    const z = zip({ level: 0 });
    for (let i = 0; i < count; i++) {
      z.add(`s-${i}.txt`, new TextEncoder().encode(`val-${i}`));
    }
    const zipBytes = await z.bytes();

    const reader = unzip(chunkBytes(zipBytes, 64));
    let seen = 0;

    for await (const entry of reader.entries()) {
      if (entry.type === "directory") {
        entry.discard();
        continue;
      }
      for await (const _chunk of entry.stream()) {
        // drain
      }
      seen++;
    }

    expect(seen).toBe(count);
  });

  it("100 small files via entry.bytes()", { timeout: 20_000 }, async () => {
    const count = 100;
    const z = zip({ level: 0 });
    const expected = new Map<string, string>();
    for (let i = 0; i < count; i++) {
      const name = `b-${String(i).padStart(3, "0")}.txt`;
      const val = `val-${i}`;
      z.add(name, new TextEncoder().encode(val));
      expected.set(name, val);
    }
    const zipBytes = await z.bytes();

    const reader = unzip(chunkBytes(zipBytes, 64));
    let seen = 0;

    for await (const entry of reader.entries()) {
      if (entry.type === "directory") {
        entry.discard();
        continue;
      }
      const data = await entry.bytes();
      const text = new TextDecoder().decode(data);
      expect(expected.get(entry.path)).toBe(text);
      seen++;
    }

    expect(seen).toBe(count);
  });

  it(
    "data descriptor ZIP: 100 files via Readable.from(entry.stream())",
    { timeout: 30_000 },
    async () => {
      const count = 100;
      const expected = new Map<string, string>();
      const entries = Array.from({ length: count }, (_, i) => {
        const name = `dd-${String(i).padStart(3, "0")}.txt`;
        const val = `dd-val-${i}`;
        expected.set(name, val);
        return { name, data: new TextEncoder().encode(val) };
      });

      const zipBytes = await createDataDescriptorZip(entries);
      const reader = unzip(chunkBytes(zipBytes, 512));
      let seen = 0;

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const readable = Readable.from(entry.stream());
        const chunks: Buffer[] = [];
        for await (const chunk of readable) {
          chunks.push(Buffer.from(chunk));
        }
        const text = Buffer.concat(chunks).toString();
        expect(expected.get(entry.path)).toBe(text);
        seen++;
      }

      expect(seen).toBe(count);
    }
  );

  it(
    "repeat with Readable.from 10 times to detect intermittent races",
    { timeout: 30_000 },
    async () => {
      const content1 = new Uint8Array(4096).fill(65);
      const content2 = new Uint8Array(4096).fill(66);
      const zipBytes = await zip({ level: 0 })
        .add("a.bin", content1)
        .add("b.bin", content2)
        .bytes();

      for (let round = 0; round < 10; round++) {
        const reader = unzip(chunkBytes(zipBytes, 128));
        const results = new Map<string, Uint8Array>();

        for await (const entry of reader.entries()) {
          if (entry.type === "directory") {
            entry.discard();
            continue;
          }
          const readable = Readable.from(entry.stream());
          const chunks: Buffer[] = [];
          for await (const chunk of readable) {
            chunks.push(Buffer.from(chunk));
          }
          results.set(entry.path, new Uint8Array(Buffer.concat(chunks)));
        }

        expect(results.size).toBe(2);
        expect(results.get("a.bin")).toEqual(content1);
        expect(results.get("b.bin")).toEqual(content2);
      }
    }
  );

  it("repeat with delayed consumption 10 times", { timeout: 30_000 }, async () => {
    const content = new TextEncoder().encode("stable-content");
    const zipBytes = await zip({ level: 0 }).add("x.txt", content).add("y.txt", content).bytes();

    for (let round = 0; round < 10; round++) {
      const reader = unzip(chunkBytes(zipBytes, 64));
      let count = 0;

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const chunks: Uint8Array[] = [];
        for await (const chunk of entry.stream()) {
          await delay(1);
          chunks.push(chunk);
        }
        count++;
      }

      expect(count).toBe(2);
    }
  });
});

// =============================================================================
// 8. createParse factory
// =============================================================================

describe("parse: createParse factory", () => {
  it("should create a Parse instance", () => {
    const parse = createParse();
    expect(parse).toBeInstanceOf(Parse);
  });

  it("should pass options to Parse", () => {
    const parse = createParse({ verbose: false, forceStream: true });
    expect(parse).toBeInstanceOf(Parse);
  });
});
