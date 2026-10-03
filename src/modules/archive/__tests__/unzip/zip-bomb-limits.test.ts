/**
 * Decompression-bomb limits for in-memory extraction: the inflater must stop at
 * the entry's declared size instead of materialising the full output first.
 */

import {
  compressSync,
  decompress,
  decompressAuto,
  decompressSync,
  gunzip,
  gunzipSync,
  gzipSync,
  unzlib,
  unzlibSync,
  zlibSync
} from "@archive/compression/compress";
import { decompressWithStream } from "@archive/compression/compress.base";
import { inflateRaw } from "@archive/compression/deflate-fallback";
import { ArchiveLimitError, EntrySizeMismatchError, FileTooLargeError } from "@archive/core/errors";
import { ZipParser } from "@archive/unzip/zip-parser";
import { createZipSync } from "@archive/zip/zip-bytes";
import { describe, expect, it } from "vitest";

const BOMB_SIZE = 10 * 1024 * 1024;
const LIED_SIZE = 10;

function readU32(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(offset, true);
}

function writeU32(data: Uint8Array, offset: number, value: number): void {
  new DataView(data.buffer, data.byteOffset, data.byteLength).setUint32(offset, value, true);
}

/** Rewrite every declared uncompressed size (local header + central directory). */
function lieAboutSize(zip: Uint8Array, size: number): Uint8Array {
  const out = zip.slice();
  for (let i = 0; i + 4 <= out.length; i++) {
    const sig = readU32(out, i);
    if (sig === 0x04034b50) {
      writeU32(out, i + 22, size);
    } else if (sig === 0x02014b50) {
      writeU32(out, i + 24, size);
    }
  }
  return out;
}

function buildBombZip(): Uint8Array {
  const zip = createZipSync([{ name: "bomb.bin", data: new Uint8Array(BOMB_SIZE) }], {
    level: 9
  });
  return lieAboutSize(zip, LIED_SIZE);
}

describe("decompression output bound", () => {
  const payload = compressSync(new Uint8Array(BOMB_SIZE), { level: 9 });

  it("Node zlib stops at maxOutputLength (async and sync)", async () => {
    await expect(decompress(payload, { maxOutputLength: 1024 })).rejects.toBeInstanceOf(
      ArchiveLimitError
    );
    expect(() => decompressSync(payload, { maxOutputLength: 1024 })).toThrow(ArchiveLimitError);
  });

  it("honours a zero bound", () => {
    expect(() => decompressSync(payload, { maxOutputLength: 0 })).toThrow(ArchiveLimitError);
  });

  it("pure-JS fallback stops at maxOutputLength", () => {
    expect(() => inflateRaw(payload, 1024)).toThrow(ArchiveLimitError);
  });

  it("DecompressionStream path stops at maxOutputLength", async () => {
    await expect(decompressWithStream(payload, 1024)).rejects.toBeInstanceOf(ArchiveLimitError);
  });

  it("gzip and zlib honour the bound on every entry point", async () => {
    const raw = new Uint8Array(BOMB_SIZE);
    const gz = gzipSync(raw);
    const zl = zlibSync(raw);
    const opts = { maxOutputLength: 1024 };
    await expect(gunzip(gz, opts)).rejects.toBeInstanceOf(ArchiveLimitError);
    expect(() => gunzipSync(gz, opts)).toThrow(ArchiveLimitError);
    await expect(unzlib(zl, opts)).rejects.toBeInstanceOf(ArchiveLimitError);
    expect(() => unzlibSync(zl, opts)).toThrow(ArchiveLimitError);
    await expect(decompressAuto(gz, opts)).rejects.toBeInstanceOf(ArchiveLimitError);
    expect((await unzlib(zl, { maxOutputLength: BOMB_SIZE })).length).toBe(BOMB_SIZE);
  });

  it("an exact bound still succeeds", async () => {
    expect((await decompress(payload, { maxOutputLength: BOMB_SIZE })).length).toBe(BOMB_SIZE);
    expect(inflateRaw(payload, BOMB_SIZE).length).toBe(BOMB_SIZE);
  });
});

describe("ZipParser bomb protection", () => {
  const bomb = buildBombZip();

  it("rejects an entry that inflates past its declared size (async)", async () => {
    const parser = new ZipParser(bomb);
    const err = await parser.extractAll().catch((e: unknown) => e);
    expect(err).toBeInstanceOf(EntrySizeMismatchError);
    expect((err as EntrySizeMismatchError).isZipBomb()).toBe(true);
    expect((err as EntrySizeMismatchError).actual).toBeUndefined();
    expect((err as Error).message).toMatch(`exceeds its declared ${LIED_SIZE} bytes`);
    expect((err as Error).cause).toBeInstanceOf(ArchiveLimitError);
  });

  it("rejects an entry that inflates past its declared size (sync)", () => {
    expect(() => new ZipParser(bomb).extractAllSync()).toThrow(EntrySizeMismatchError);
  });

  it("rejects a declared size above maxEntrySize before inflating", async () => {
    const zip = createZipSync([{ name: "a.bin", data: new Uint8Array(2048) }]);
    await expect(new ZipParser(zip, { maxEntrySize: 1024 }).extract("a.bin")).rejects.toThrow(
      FileTooLargeError
    );
  });

  it("rejects a Central Directory count the buffer cannot hold", () => {
    const zip = createZipSync([{ name: "a.txt", data: new Uint8Array([1]) }]);
    const tampered = zip.slice();
    const eocd = tampered.length - 22;
    expect(readU32(tampered, eocd)).toBe(0x06054b50);
    // Total entries (both the on-disk and overall counts) → 0xfffe.
    tampered[eocd + 8] = 0xfe;
    tampered[eocd + 9] = 0xff;
    tampered[eocd + 10] = 0xfe;
    tampered[eocd + 11] = 0xff;
    expect(() => new ZipParser(tampered)).toThrow(/more than its/);
  });

  it("enforces maxTotalUncompressedSize", async () => {
    const zip = createZipSync([
      { name: "a.bin", data: new Uint8Array(600) },
      { name: "b.bin", data: new Uint8Array(600) }
    ]);
    const limited = new ZipParser(zip, { maxTotalUncompressedSize: 1000 });
    await expect(limited.extractAll()).rejects.toBeInstanceOf(ArchiveLimitError);
    expect(() => limited.extractAllSync()).toThrow(ArchiveLimitError);
    await expect(
      limited.forEach(async (_entry, getData) => {
        await getData();
      })
    ).rejects.toBeInstanceOf(ArchiveLimitError);
    expect((await new ZipParser(zip).extractAll()).size).toBe(2);
  });
});

describe("maxEntries", () => {
  function manyEntries(count: number): Uint8Array {
    return createZipSync(
      Array.from({ length: count }, (_, i) => ({ name: `f${i}.txt`, data: new Uint8Array(0) })),
      { level: 0 }
    );
  }

  it("rejects a central directory declaring more entries than allowed", () => {
    const zip = manyEntries(50);
    let err: unknown;
    try {
      new ZipParser(zip, { maxEntries: 49 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(ArchiveLimitError);
    expect((err as ArchiveLimitError).limit).toBe("maxEntries");
    expect((err as ArchiveLimitError).allowed).toBe(49);
  });

  it("accepts a count at the bound, and is unbounded by default", () => {
    expect(new ZipParser(manyEntries(50), { maxEntries: 50 }).getEntries()).toHaveLength(50);
    expect(new ZipParser(manyEntries(50)).getEntries()).toHaveLength(50);
  });
});

describe("limit option validation", () => {
  const deflated = compressSync(new TextEncoder().encode("hello"));
  const gz = gzipSync(new TextEncoder().encode("hello"));
  const zl = zlibSync(new TextEncoder().encode("hello"));

  it.each([Number.NaN, -1])(
    "rejects maxOutputLength %s on every inflating entry point",
    async bad => {
      const options = { maxOutputLength: bad };
      expect(() => decompressSync(deflated, options)).toThrow(RangeError);
      expect(() => gunzipSync(gz, options)).toThrow(RangeError);
      expect(() => unzlibSync(zl, options)).toThrow(RangeError);
      expect(() => inflateRaw(deflated, bad)).toThrow(RangeError);
      await expect(decompress(deflated, options)).rejects.toThrow(RangeError);
      await expect(gunzip(gz, options)).rejects.toThrow(RangeError);
      await expect(unzlib(zl, options)).rejects.toThrow(RangeError);
    }
  );

  it("accepts Infinity as unbounded", async () => {
    const options = { maxOutputLength: Infinity };
    expect(new TextDecoder().decode(decompressSync(deflated, options))).toBe("hello");
    expect(new TextDecoder().decode(await gunzip(gz, options))).toBe("hello");
    expect(new TextDecoder().decode(inflateRaw(deflated, Infinity))).toBe("hello");
  });

  it.each([Number.NaN, -1])("rejects ZipParser size limits of %s", bad => {
    const zip = createZipSync([{ name: "a.txt", data: new TextEncoder().encode("a") }]);
    expect(() => new ZipParser(zip, { maxEntrySize: bad })).toThrow(RangeError);
    expect(() => new ZipParser(zip, { maxTotalUncompressedSize: bad })).toThrow(RangeError);
    expect(() => new ZipParser(zip, { maxEntrySize: Infinity })).not.toThrow();
  });
});
