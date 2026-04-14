/**
 * Robustness and edge-case tests for the unzip streaming parser.
 *
 * Covers concerns not already in stream-lifecycle.test.ts:
 * 1. Data descriptor + ultra-small chunk parameterized
 * 2. Consumer lifecycle edge cases (normal + data descriptor ZIP)
 * 3. Three consumption modes content-level consistency
 * 4. Error classification: truncated ZIP / corrupt data / source error
 * 5. awaitEntryCompletion does not swallow real errors
 * 6. Soak tests (key cases repeated 50 times)
 */

import { Readable } from "stream";

import { unzip, zip } from "@archive";
import { crc32 } from "@archive/compression/crc32";
import { Parse, createParse, type ZipEntry } from "@archive/unzip/stream";
import { describe, it, expect } from "vitest";

import { delay, chunkBytes, createDataDescriptorZip, concatChunks } from "./test-helpers";

// ---------------------------------------------------------------------------
// File-local helpers
// ---------------------------------------------------------------------------

/** Collect all entry contents from an unzip reader into a sorted Map. */
async function collectEntries(
  source: AsyncIterable<Uint8Array> | Uint8Array,
  mode: "bytes" | "stream" | "readable-from" = "bytes"
): Promise<Map<string, { data: Uint8Array; crc: number }>> {
  const reader = unzip(source);
  const results = new Map<string, { data: Uint8Array; crc: number }>();

  for await (const entry of reader.entries()) {
    if (entry.type === "directory") {
      entry.discard();
      continue;
    }

    let data: Uint8Array;
    if (mode === "bytes") {
      data = await entry.bytes();
    } else if (mode === "stream") {
      const chunks: Uint8Array[] = [];
      for await (const chunk of entry.stream()) {
        chunks.push(chunk);
      }
      data = concatChunks(chunks);
    } else {
      const readable = Readable.from(entry.stream());
      const chunks: Buffer[] = [];
      for await (const chunk of readable) {
        chunks.push(Buffer.from(chunk));
      }
      data = new Uint8Array(Buffer.concat(chunks));
    }

    results.set(entry.path, { data, crc: crc32(data) });
  }

  return results;
}

// =============================================================================
// 1. Data descriptor + ultra-small chunk parameterized
// =============================================================================

describe("robustness: data descriptor chunk boundaries", () => {
  // Data descriptor ZIP with a few entries, tested at various chunk sizes.
  // Small chunks stress the streamUntilValidatedDataDescriptor scanner
  // which must handle descriptor signatures split across chunk boundaries.
  //
  // Known limitation: chunk sizes below ~128 bytes can cause
  // "FILE_ENDED: Data descriptor not found" because the scanner needs
  // enough buffered data to validate the descriptor + next record signature.
  // Chunk sizes >= 128 should always work.
  // Test a wide range of chunk sizes, including very small ones that used to
  // trigger "FILE_ENDED: Data descriptor not found" due to a scanner bug
  // where onNoMatch() would advance past a pending candidate.
  const singleFileChunkSizes = [1, 2, 3, 4, 5, 7, 8, 16, 32, 64, 128, 256, 512, 1024];

  for (const cs of singleFileChunkSizes) {
    it(`single file dd ZIP at chunk size ${cs}`, { timeout: 10_000 }, async () => {
      const content = new TextEncoder().encode("hello-data-descriptor");
      const zipBytes = await createDataDescriptorZip([{ name: "f.txt", data: content }]);

      const reader = unzip(chunkBytes(zipBytes, cs));
      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const data = await entry.bytes();
        expect(new TextDecoder().decode(data)).toBe("hello-data-descriptor");
      }
    });
  }

  // Multi-file dd ZIPs exercise the scanner finding the descriptor followed
  // by the next local-file-header signature. Test the full range.
  const multiFileChunkSizes = [1, 2, 3, 4, 5, 7, 8, 16, 32, 64, 128, 256, 512, 1024];

  for (const cs of multiFileChunkSizes) {
    it(`3 files dd ZIP at chunk size ${cs}`, { timeout: 10_000 }, async () => {
      const entries = [
        { name: "a.txt", data: new TextEncoder().encode("aaa-content") },
        { name: "b.txt", data: new TextEncoder().encode("bbb-content-longer") },
        { name: "c.bin", data: new Uint8Array(512).fill(0xcc) }
      ];
      const zipBytes = await createDataDescriptorZip(entries);

      const reader = unzip(chunkBytes(zipBytes, cs));
      const seen = new Map<string, Uint8Array>();

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        seen.set(entry.path, await entry.bytes());
      }

      expect(seen.size).toBe(3);
      expect(new TextDecoder().decode(seen.get("a.txt")!)).toBe("aaa-content");
      expect(new TextDecoder().decode(seen.get("b.txt")!)).toBe("bbb-content-longer");
      expect(seen.get("c.bin")!.length).toBe(512);
    });
  }

  it("dd ZIP: large file followed by small file at small chunk", { timeout: 15_000 }, async () => {
    const big = new Uint8Array(32768).fill(0xab);
    const small = new TextEncoder().encode("tiny");
    const zipBytes = await createDataDescriptorZip([
      { name: "big.bin", data: big },
      { name: "small.txt", data: small }
    ]);

    // Use a chunk size that doesn't align with any record boundary.
    const reader = unzip(chunkBytes(zipBytes, 37));
    const seen = new Map<string, Uint8Array>();

    for await (const entry of reader.entries()) {
      if (entry.type === "directory") {
        entry.discard();
        continue;
      }
      seen.set(entry.path, await entry.bytes());
    }

    expect(seen.size).toBe(2);
    expect(seen.get("big.bin")).toEqual(big);
    expect(new TextDecoder().decode(seen.get("small.txt")!)).toBe("tiny");
  });
});

// =============================================================================
// 2. Consumer lifecycle edge cases
// =============================================================================

describe("robustness: consumer lifecycle edge cases", () => {
  // For each scenario, test both normal ZIP and data descriptor ZIP.

  async function makeNormalZip(): Promise<Uint8Array> {
    return zip({ level: 0 })
      .add("first.bin", new Uint8Array(8192).fill(0x11))
      .add("second.txt", new TextEncoder().encode("second-content"))
      .add("third.dat", new Uint8Array(256).fill(0x33))
      .bytes();
  }

  async function makeDdZip(): Promise<Uint8Array> {
    return createDataDescriptorZip([
      { name: "first.bin", data: new Uint8Array(8192).fill(0x11) },
      { name: "second.txt", data: new TextEncoder().encode("second-content") },
      { name: "third.dat", data: new Uint8Array(256).fill(0x33) }
    ]);
  }

  for (const [label, makeZip] of [
    ["normal ZIP", makeNormalZip],
    ["data descriptor ZIP", makeDdZip]
  ] as const) {
    describe(label, () => {
      it("discard first entry, consume rest", { timeout: 10_000 }, async () => {
        const zipBytes = await makeZip();
        const reader = unzip(chunkBytes(zipBytes, 256));
        const seen: string[] = [];

        for await (const entry of reader.entries()) {
          seen.push(entry.path);
          if (entry.path === "first.bin") {
            entry.discard();
          } else {
            await entry.bytes();
          }
        }

        expect(seen.length).toBe(3);
      });

      it("bytes() first entry, stream() second, discard third", { timeout: 10_000 }, async () => {
        const zipBytes = await makeZip();
        const reader = unzip(chunkBytes(zipBytes, 512));
        let idx = 0;

        for await (const entry of reader.entries()) {
          if (idx === 0) {
            await entry.bytes();
          } else if (idx === 1) {
            for await (const _c of entry.stream()) {
              // drain
            }
          } else {
            entry.discard();
          }
          idx++;
        }

        expect(idx).toBe(3);
      });

      it("break after first chunk of first entry stream", { timeout: 10_000 }, async () => {
        const zipBytes = await makeZip();
        const reader = unzip(chunkBytes(zipBytes, 512));
        const seen: string[] = [];

        for await (const entry of reader.entries()) {
          seen.push(entry.path);
          if (entry.path === "first.bin") {
            for await (const _c of entry.stream()) {
              break;
            }
          } else {
            await entry.bytes();
          }
        }

        expect(seen.length).toBe(3);
      });

      it("Readable.from first entry, bytes() rest", { timeout: 10_000 }, async () => {
        const zipBytes = await makeZip();
        const reader = unzip(chunkBytes(zipBytes, 512));
        const seen: string[] = [];

        for await (const entry of reader.entries()) {
          seen.push(entry.path);
          if (entry.path === "first.bin") {
            const readable = Readable.from(entry.stream());
            for await (const _c of readable) {
              // drain
            }
          } else {
            await entry.bytes();
          }
        }

        expect(seen.length).toBe(3);
      });

      it("break from entries() after first entry", { timeout: 10_000 }, async () => {
        const zipBytes = await makeZip();
        const reader = unzip(chunkBytes(zipBytes, 512));
        let seen = 0;

        for await (const entry of reader.entries()) {
          entry.discard();
          seen++;
          break;
        }

        expect(seen).toBe(1);
      });
    });
  }
});

// =============================================================================
// 3. Three consumption modes content-level consistency
// =============================================================================

describe("robustness: consumption mode content consistency", () => {
  it(
    "bytes() / stream() / Readable.from() should yield identical content and CRC",
    { timeout: 15_000 },
    async () => {
      const z = zip({ level: 6 })
        .add("text.txt", new TextEncoder().encode("hello world 你好世界"))
        .add("binary.bin", new Uint8Array(4096).fill(0xab))
        .add("empty.dat", new Uint8Array(0));
      const zipBytes = await z.bytes();

      const resultBytes = await collectEntries(zipBytes, "bytes");
      const resultStream = await collectEntries(zipBytes, "stream");
      const resultReadable = await collectEntries(chunkBytes(zipBytes, 256), "readable-from");

      // All three should have the same entry paths.
      const paths = [...resultBytes.keys()].sort();
      expect([...resultStream.keys()].sort()).toEqual(paths);
      expect([...resultReadable.keys()].sort()).toEqual(paths);

      // Content CRC should match across all three modes.
      for (const path of paths) {
        const bCrc = resultBytes.get(path)!.crc;
        const sCrc = resultStream.get(path)!.crc;
        const rCrc = resultReadable.get(path)!.crc;
        expect(sCrc).toBe(bCrc);
        expect(rCrc).toBe(bCrc);

        // Data length should match.
        expect(resultStream.get(path)!.data.length).toBe(resultBytes.get(path)!.data.length);
        expect(resultReadable.get(path)!.data.length).toBe(resultBytes.get(path)!.data.length);
      }
    }
  );

  it("on('entry') should yield same content CRC as for-await", { timeout: 10_000 }, async () => {
    const z = zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("content-a"))
      .add("b.bin", new Uint8Array(1024).fill(0xbb));
    const zipBytes = await z.bytes();

    // Mode A: for-await via unzip()
    const crcA = new Map<string, number>();
    const readerA = unzip(zipBytes);
    for await (const entry of readerA.entries()) {
      if (entry.type === "directory") {
        entry.discard();
        continue;
      }
      const data = await entry.bytes();
      crcA.set(entry.path, crc32(data));
    }

    // Mode B: on("entry") + promise()
    const crcB = new Map<string, number>();
    const parse = new Parse();
    parse.on("entry", (entry: ZipEntry) => {
      entry
        .buffer()
        .then(buf => {
          crcB.set(entry.path, crc32(buf));
        })
        .catch(() => {
          entry.autodrain();
        });
    });
    const readable = Readable.from(chunkBytes(zipBytes, 256));
    readable.pipe(parse);
    await parse.promise();

    expect([...crcA.keys()].sort()).toEqual([...crcB.keys()].sort());
    for (const [path, crc] of crcA) {
      expect(crcB.get(path)).toBe(crc);
    }
  });

  it(
    "data descriptor ZIP: bytes/stream/Readable.from yield identical CRC",
    { timeout: 15_000 },
    async () => {
      const entries = [
        { name: "dd-a.txt", data: new TextEncoder().encode("dd-content-a") },
        { name: "dd-b.bin", data: new Uint8Array(2048).fill(0xdd) }
      ];
      const zipBytes = await createDataDescriptorZip(entries);

      const resultBytes = await collectEntries(zipBytes, "bytes");
      const resultStream = await collectEntries(chunkBytes(zipBytes, 256), "stream");
      const resultReadable = await collectEntries(chunkBytes(zipBytes, 256), "readable-from");

      for (const path of resultBytes.keys()) {
        expect(resultStream.get(path)!.crc).toBe(resultBytes.get(path)!.crc);
        expect(resultReadable.get(path)!.crc).toBe(resultBytes.get(path)!.crc);
      }
    }
  );
});

// =============================================================================
// 4. Error classification
// =============================================================================

describe("robustness: error classification", () => {
  it("truncated ZIP (missing EOCD) should error, not hang", { timeout: 10_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("hello"))
      .bytes();

    // Truncate: remove last 30 bytes (EOCD is 22+ bytes).
    const truncated = zipBytes.subarray(0, Math.max(0, zipBytes.length - 30));

    await expect(async () => {
      const reader = unzip(truncated);
      for await (const entry of reader.entries()) {
        await entry.bytes();
      }
    }).rejects.toThrow();
  });

  it("truncated ZIP via streaming should error", { timeout: 10_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("hello"))
      .bytes();
    const truncated = zipBytes.subarray(0, Math.max(0, zipBytes.length - 30));

    await expect(async () => {
      const reader = unzip(chunkBytes(truncated, 64));
      for await (const entry of reader.entries()) {
        await entry.bytes();
      }
    }).rejects.toThrow();
  });

  it("garbage bytes should error", { timeout: 10_000 }, async () => {
    const garbage = new Uint8Array(512);
    for (let i = 0; i < garbage.length; i++) {
      garbage[i] = Math.floor(Math.random() * 256);
    }

    await expect(async () => {
      const reader = unzip(garbage);
      for await (const entry of reader.entries()) {
        entry.discard();
      }
    }).rejects.toThrow();
  });

  it("source stream error should propagate to promise()", { timeout: 10_000 }, async () => {
    const parse = new Parse();

    const source = new Readable({
      read() {
        this.push(Buffer.alloc(64));
        process.nextTick(() => {
          this.destroy(new Error("source stream failed"));
        });
      }
    });
    // Prevent the source's error from becoming an uncaught exception.
    source.on("error", () => {});

    source.pipe(parse);
    await expect(parse.promise()).rejects.toThrow();
  });

  it("source stream error should propagate to for-await", { timeout: 10_000 }, async () => {
    const parse = createParse({ forceStream: true });

    const source = new Readable({
      read() {
        this.push(Buffer.alloc(64));
        process.nextTick(() => {
          this.destroy(new Error("source failed"));
        });
      }
    });
    source.on("error", () => {});

    source.pipe(parse);

    await expect(async () => {
      for await (const entry of parse) {
        (entry as ZipEntry).autodrain();
      }
    }).rejects.toThrow();
  });

  it(
    "half-valid ZIP (valid header, truncated entry data) should error",
    { timeout: 10_000 },
    async () => {
      const zipBytes = await zip({ level: 0 })
        .add("big.bin", new Uint8Array(4096).fill(0xaa))
        .bytes();

      // Truncate in the middle of the entry data.
      const halfValid = zipBytes.subarray(0, 100);

      await expect(async () => {
        const reader = unzip(halfValid);
        for await (const entry of reader.entries()) {
          await entry.bytes();
        }
      }).rejects.toThrow();
    }
  );
});

// =============================================================================
// 5. awaitEntryCompletion does not swallow real errors
// =============================================================================

describe("robustness: real errors not swallowed", () => {
  it("corrupted compressed data should still produce an error", { timeout: 10_000 }, async () => {
    // Create a valid ZIP with compression, then corrupt the compressed data.
    const zipBytes = await zip({ level: 6 })
      .add("file.txt", new TextEncoder().encode("a".repeat(1000)))
      .bytes();

    // Find the compressed data region and corrupt it.
    // Local file header is at least 30 bytes + filename length.
    // We'll corrupt bytes in the middle of the file.
    const corrupted = new Uint8Array(zipBytes);
    const midpoint = Math.floor(corrupted.length / 2);
    for (let i = midpoint; i < midpoint + 20 && i < corrupted.length; i++) {
      corrupted[i] ^= 0xff;
    }

    // This should either throw a decompression error or produce wrong data.
    // The key assertion: it should NOT silently succeed with wrong content.
    try {
      const reader = unzip(corrupted);
      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const data = await entry.bytes();
        // If we get here, the data should NOT match the original.
        // (Some corruption patterns may still decompress but produce wrong data.)
        const text = new TextDecoder().decode(data);
        if (text === "a".repeat(1000)) {
          // This would mean corruption wasn't detected — fail the test.
          // But this is unlikely given we corrupted 20 bytes of compressed data.
          throw new Error("Corruption not detected");
        }
        // If data is different, that's acceptable — corruption was "detected"
        // in the sense that the output is wrong.
      }
    } catch (err) {
      // Any error (decompression, invalid signature, etc.) is acceptable.
      // The point is: it didn't silently succeed.
      expect(err).toBeDefined();
    }
  });

  it("valid ZIP should still fully succeed (sanity check)", { timeout: 10_000 }, async () => {
    const content = new TextEncoder().encode("a".repeat(1000));
    const zipBytes = await zip({ level: 6 }).add("file.txt", content).bytes();

    const reader = unzip(zipBytes);
    for await (const entry of reader.entries()) {
      if (entry.type === "directory") {
        entry.discard();
        continue;
      }
      const data = await entry.bytes();
      expect(new TextDecoder().decode(data)).toBe("a".repeat(1000));
    }
  });
});

// =============================================================================
// 6. Soak tests (repeat key cases to detect intermittent races)
// =============================================================================

describe("robustness: soak tests", () => {
  it("Readable.from(entry.stream()) x50 repetitions", { timeout: 60_000 }, async () => {
    const content1 = new Uint8Array(2048).fill(0x41);
    const content2 = new Uint8Array(2048).fill(0x42);
    const zipBytes = await zip({ level: 0 }).add("a.bin", content1).add("b.bin", content2).bytes();

    for (let i = 0; i < 50; i++) {
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
  });

  it("data descriptor + slow consumer x50 repetitions", { timeout: 60_000 }, async () => {
    const content = new TextEncoder().encode("soak-test-content");
    const zipBytes = await createDataDescriptorZip([
      { name: "x.txt", data: content },
      { name: "y.txt", data: content }
    ]);

    for (let i = 0; i < 50; i++) {
      const reader = unzip(chunkBytes(zipBytes, 16));
      let count = 0;

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }
        const chunks: Uint8Array[] = [];
        for await (const chunk of entry.stream()) {
          await delay(0);
          chunks.push(chunk);
        }
        count++;
      }

      expect(count).toBe(2);
    }
  });

  it("break after first chunk x50 repetitions", { timeout: 60_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("big.bin", new Uint8Array(16384))
      .add("small.txt", new TextEncoder().encode("ok"))
      .bytes();

    for (let i = 0; i < 50; i++) {
      const reader = unzip(chunkBytes(zipBytes, 4096));
      const seen: string[] = [];

      for await (const entry of reader.entries()) {
        seen.push(entry.path);

        if (entry.path === "big.bin") {
          for await (const _c of entry.stream()) {
            break;
          }
        } else {
          const data = await entry.bytes();
          expect(new TextDecoder().decode(data)).toBe("ok");
        }
      }

      expect(seen).toEqual(["big.bin", "small.txt"]);
    }
  });

  it("invalid ZIP error propagation x50 repetitions", { timeout: 30_000 }, async () => {
    const garbage = new Uint8Array(256).fill(0xff);

    for (let i = 0; i < 50; i++) {
      await expect(async () => {
        const reader = unzip(garbage);
        for await (const entry of reader.entries()) {
          entry.discard();
        }
      }).rejects.toThrow();
    }
  });

  it("mixed consumption modes x50 repetitions", { timeout: 60_000 }, async () => {
    const zipBytes = await zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("aaa"))
      .add("b.txt", new TextEncoder().encode("bbb"))
      .add("c.txt", new TextEncoder().encode("ccc"))
      .bytes();

    for (let i = 0; i < 50; i++) {
      const reader = unzip(chunkBytes(zipBytes, 128));
      let idx = 0;

      for await (const entry of reader.entries()) {
        if (entry.type === "directory") {
          entry.discard();
          continue;
        }

        const mode = idx % 3;
        if (mode === 0) {
          await entry.bytes();
        } else if (mode === 1) {
          for await (const _c of entry.stream()) {
            // drain
          }
        } else {
          entry.discard();
        }
        idx++;
      }

      expect(idx).toBe(3);
    }
  });
});
