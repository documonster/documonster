/**
 * Damage in the bytes an archive reads is named for what it is, whichever inflater met it: compressed data that does
 * not decompress is an `ArchiveDecompressionError`, a ZIP record that is not one a `ZipParseError`. A document reader
 * tells its caller "this file is damaged" from those classes alone, so nothing else may wear them.
 */
import {
  Archive,
  ArchiveDecompressionError,
  ArchiveError,
  ArchiveLimitExceededError,
  Compression,
  ZipParseError
} from "@archive";
import { isArchiveDataError } from "@archive/core/read-failure";
import { describe, expect, it } from "vitest";

const TEXT = new TextEncoder().encode(
  "hello hello hello hello hello hello hello hello ".repeat(20)
);

/** `bytes` with a byte past the header flipped, so the compressed stream no longer decodes. */
function flipped(bytes: Uint8Array, at: number): Uint8Array {
  const out = bytes.slice();
  out[at]! ^= 0xff;
  return out;
}

async function failure(run: () => unknown): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  throw new Error("expected a failure");
}

describe("ArchiveDecompressionError", () => {
  it("is what each codec reports for data that does not decompress", async () => {
    const deflated = await Compression.deflate(TEXT);
    const gzipped = await Compression.gzip(TEXT);
    const zlibbed = await Compression.zlib(TEXT);
    const cases: [string, () => unknown][] = [
      ["inflate", () => Compression.inflate(flipped(deflated, 0))],
      ["inflateSync", () => Compression.inflateSync(flipped(deflated, 0))],
      ["gunzip, header", () => Compression.gunzip(flipped(gzipped, 0))],
      ["gunzip, checksum", () => Compression.gunzip(flipped(gzipped, gzipped.length - 5))],
      ["gunzipSync, truncated", () => Compression.gunzipSync(gzipped.slice(0, 12))],
      ["unzlib, header", () => Compression.unzlib(flipped(zlibbed, 1))],
      ["unzlibSync, checksum", () => Compression.unzlibSync(flipped(zlibbed, zlibbed.length - 1))]
    ];
    for (const [name, run] of cases) {
      const error = await failure(run);
      expect(error, name).toBeInstanceOf(ArchiveDecompressionError);
      expect(error, name).toBeInstanceOf(ArchiveError);
      expect(isArchiveDataError(error), name).toBe(true);
    }
  });

  it("is what a streaming decompressor fails with, on every platform", async () => {
    const bad = new Uint8Array([0xff, 0xff, 0xff, 0xff, 1, 2, 3, 4, 5, 6, 7, 8]);
    const streams: [string, () => ReturnType<typeof Compression.createInflateStream>][] = [
      ["createInflateStream", () => Compression.createInflateStream()],
      ["createGunzipStream", () => Compression.createGunzipStream()],
      ["createUnzlibStream", () => Compression.createUnzlibStream()]
    ];
    for (const [name, make] of streams) {
      const stream = make();
      const error = await new Promise<unknown>(resolve => {
        stream.on("error", resolve);
        stream.on("data", () => {});
        stream.on("end", () => resolve(undefined));
        stream.write(bad);
        stream.end();
      });
      expect(error, name).toBeInstanceOf(ArchiveDecompressionError);
    }
  });

  // Data that asks for what no inflater here implements is reported as data that does not decompress, on every path —
  // the native inflaters cannot tell it from damage, so the archive does not try to. A zlib header with FDICT set
  // (0x78 0xbb is valid), and a gzip header whose method is not DEFLATE.
  it("is what data asking for an unimplemented feature reports", async () => {
    const presetDictionary = new Uint8Array([0x78, 0xbb, 0, 0, 0, 1, 0x4b, 0x04, 0, 0, 0, 0, 0, 0]);
    const gzipped = await Compression.gzip(TEXT);
    const otherMethod = gzipped.slice();
    otherMethod[2] = 7;
    const cases: [string, () => unknown][] = [
      ["unzlib", () => Compression.unzlib(presetDictionary)],
      ["unzlibSync", () => Compression.unzlibSync(presetDictionary)],
      ["gunzip", () => Compression.gunzip(otherMethod)],
      ["gunzipSync", () => Compression.gunzipSync(otherMethod)]
    ];
    for (const [name, run] of cases) {
      expect(await failure(run), name).toBeInstanceOf(ArchiveDecompressionError);
    }
    for (const [name, make, data] of [
      ["createUnzlibStream", Compression.createUnzlibStream, presetDictionary],
      ["createGunzipStream", Compression.createGunzipStream, otherMethod]
    ] as const) {
      const stream = make();
      const error = await new Promise<unknown>(resolve => {
        stream.on("error", resolve);
        stream.on("data", () => {});
        stream.on("end", () => resolve(undefined));
        stream.write(data);
        stream.end();
      });
      expect(error, name).toBeInstanceOf(ArchiveDecompressionError);
    }
  });

  it("is what a ZIP entry whose compressed data is broken reports", async () => {
    const zip = await Archive.packZip([["a.txt", TEXT]], { level: 9 });
    const nameLength = zip[26]! | (zip[27]! << 8);
    const extraLength = zip[28]! | (zip[29]! << 8);
    const broken = flipped(zip, 30 + nameLength + extraLength + 1);
    const error = await failure(() => Archive.unpackZip(broken));
    expect(isArchiveDataError(error)).toBe(true);
  });

  it("is what every tar.gz reader reports for a damaged gzip stream", async () => {
    const tgz = await Archive.packTarGz(new Map([["a.txt", TEXT]]));
    const broken = flipped(tgz, tgz.length - 5);
    const readers: [string, () => unknown][] = [
      ["parseTarGz", () => Archive.parseTarGz(broken)],
      ["unpackTarGz", () => Archive.unpackTarGz(broken)],
      [
        "parseTarGzStream",
        async () => {
          for await (const _entry of Archive.parseTarGzStream(broken)) {
            // read through
          }
        }
      ]
    ];
    for (const [name, read] of readers) {
      expect(await failure(read), name).toBeInstanceOf(ArchiveDecompressionError);
    }
  });

  // The class says the bytes are damaged, so a source that fails while it is read must not wear it.
  it("is not what a tar.gz reader reports for a source that fails", async () => {
    const tgz = await Archive.packTarGz(new Map([["a.txt", TEXT]]));
    const cause = new TypeError("network went away");
    async function* source(): AsyncGenerator<Uint8Array> {
      yield tgz.subarray(0, 20);
      throw cause;
    }
    const readers: [string, () => unknown][] = [
      ["parseTarGz", () => Archive.parseTarGz(source())],
      ["unpackTarGz", () => Archive.unpackTarGz(source())],
      [
        "parseTarGzStream",
        async () => {
          for await (const _entry of Archive.parseTarGzStream(source())) {
            // read through
          }
        }
      ]
    ];
    for (const [name, read] of readers) {
      expect(await failure(read), name).toBe(cause);
    }
  });
});

describe("a ZIP record that is not one", () => {
  it("is a ZipParseError, for the central directory and for a local header", async () => {
    const zip = await Archive.packZip([["a.txt", TEXT]], { level: 0 });
    // The central directory's first signature, and the first local header's.
    const central = zip.lastIndexOf(0x50, zip.length - 22 - 46 + 1);
    expect([...zip.slice(central, central + 4)]).toEqual([0x50, 0x4b, 0x01, 0x02]);
    expect(await failure(() => Archive.unpackZip(flipped(zip, central + 2)))).toBeInstanceOf(
      ZipParseError
    );
    expect(await failure(() => Archive.unpackZip(flipped(zip, 2)))).toBeInstanceOf(ZipParseError);
  });
});

describe("isArchiveDataError", () => {
  it("answers for what describes the bytes, and for nothing else", () => {
    expect(isArchiveDataError(new ArchiveDecompressionError("x"))).toBe(true);
    expect(isArchiveDataError(new ZipParseError("x"))).toBe(true);
    // A plain ArchiveError — a worker that failed, an entry read twice — says nothing about the bytes.
    expect(isArchiveDataError(new ArchiveError("Worker terminated unexpectedly"))).toBe(false);
    expect(isArchiveDataError(new ArchiveLimitExceededError("maxEntries", 1))).toBe(false);
    expect(isArchiveDataError(new Error("x"))).toBe(false);
  });
});
