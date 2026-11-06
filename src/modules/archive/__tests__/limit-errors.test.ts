/**
 * Every limit an archive reader enforces says which limit (`limit`, the option's name) and its
 * value (`allowed`). The per-entry limit is a `ArchiveFileTooLargeError`, a subclass of
 * `ArchiveLimitExceededError` that also names the entry (`path`), so catching the limit class
 * sees every limit.
 */
import {
  Archive,
  ArchiveError,
  ArchiveLimitExceededError,
  ArchiveFileTooLargeError,
  type ArchiveLimit
} from "@archive/index";
import { describe, expect, expectTypeOf, it } from "vitest";

const big = new Uint8Array(4096).fill(7);
const zip = Archive.packZipSync([["big.bin", big]], { level: 0 });
const deflated = Archive.packZipSync([["big.bin", big]], { level: 9 });
const tar = Archive.packTarSync([["big.bin", big]]);
const tgz = Archive.packTarGzSync([["big.bin", big]]);

async function failure(run: () => unknown): Promise<unknown> {
  try {
    await run();
    return undefined;
  } catch (error) {
    return error;
  }
}

function expectEntryLimit(error: unknown, allowed: number): void {
  expect(error).toBeInstanceOf(ArchiveFileTooLargeError);
  expect(error).toBeInstanceOf(ArchiveError);
  expect(error).toBeInstanceOf(ArchiveLimitExceededError);
  expect(error).toMatchObject({ limit: "maxEntrySize", allowed, path: "big.bin" });
  expect((error as Error).name).toBe("ArchiveFileTooLargeError");
}

describe("maxEntrySize is an ArchiveLimitExceededError naming the limit and its value", () => {
  it.each([
    ["unpackZip stored", () => Archive.unpackZip(zip, { maxEntrySize: 100 })],
    ["unpackZip deflated", () => Archive.unpackZip(deflated, { maxEntrySize: 100 })],
    ["unpackZipSync", () => Archive.unpackZipSync(zip, { maxEntrySize: 100 })],
    ["openZip bytes()", () => Archive.openZip(zip, { maxEntrySize: 100 }).bytes("big.bin")],
    [
      "openZip over a stream, entry.bytes()",
      async () => {
        const stream = new ReadableStream<Uint8Array>({
          start(c) {
            c.enqueue(zip);
            c.close();
          }
        });
        for await (const entry of Archive.openZip(stream, { maxEntrySize: 100 }).entries()) {
          await entry.bytes();
        }
      }
    ],
    ["parseTar", () => Archive.parseTar(tar, { maxEntrySize: 100 })],
    ["unpackTar", () => Archive.unpackTar(tar, { maxEntrySize: 100 })],
    ["unpackTarGz", () => Archive.unpackTarGz(tgz, { maxEntrySize: 100 })],
    [
      "parseTarStream",
      async () => {
        for await (const _ of Archive.parseTarStream(tar, { maxEntrySize: 100 })) {
          // read
        }
      }
    ],
    ["openTar list()", () => Archive.openTar(tar, { maxEntrySize: 100 }).list()]
  ])("%s", async (_name, run) => {
    expectEntryLimit(await failure(run), 100);
  });

  it("the constructor keeps its signature; `allowed` is NaN when the thrower does not give it", () => {
    const plain = new ArchiveFileTooLargeError("x.bin", "holds more than 10 bytes");
    expect(plain.message).toBe(
      'File "x.bin" is too large to extract into memory (holds more than 10 bytes)'
    );
    expect(plain.limit).toBe("maxEntrySize");
    expect(plain.allowed).toBeNaN();
    expect(new ArchiveFileTooLargeError("x.bin", "r", { allowed: 10 }).allowed).toBe(10);
  });

  it("reads like an ArchiveLimitExceededError: the same field names and meaning", () => {
    const limitError = new ArchiveLimitExceededError("maxEntries", 3);
    const sizeError = new ArchiveFileTooLargeError("x.bin", "r", { allowed: 10 });
    const describeLimit = (e: { limit: string; allowed: number }) => `${e.limit}=${e.allowed}`;
    expect(describeLimit(limitError)).toBe("maxEntries=3");
    expect(describeLimit(sizeError)).toBe("maxEntrySize=10");
    expectTypeOf(sizeError.limit).toEqualTypeOf<"maxEntrySize">();
    expectTypeOf(sizeError.allowed).toEqualTypeOf<number>();
  });

  it("ArchiveLimit names every limit, maxEntrySize included", () => {
    expectTypeOf<ArchiveLimit>().toEqualTypeOf<
      "maxOutputLength" | "maxTotalUncompressedSize" | "maxEntries" | "maxEntrySize"
    >();
    expectTypeOf(
      new ArchiveLimitExceededError("maxEntries", 1).limit
    ).toEqualTypeOf<ArchiveLimit>();
  });
});

describe("the internal extractAll helper", () => {
  it("charges maxTotalUncompressedSize as unpackZip does", async () => {
    const { extractAll } = await import("@archive/unzip/extract");
    await expect(Archive.unpackZip(zip, { maxTotalUncompressedSize: 100 })).rejects.toBeInstanceOf(
      ArchiveLimitExceededError
    );
    await expect(extractAll(zip, { maxTotalUncompressedSize: 100 })).rejects.toMatchObject({
      limit: "maxTotalUncompressedSize",
      allowed: 100
    });
  });
});

describe("maxTotalUncompressedSize while an entry inflates", () => {
  it("stops an entry that would cross the total, reporting the total", async () => {
    const two = Archive.packZipSync(
      [
        ["a.bin", new Uint8Array(3000).fill(1)],
        ["b.bin", new Uint8Array(3000).fill(2)]
      ],
      { level: 9 }
    );
    const { extractAll } = await import("@archive/unzip/extract");
    const error = await extractAll(two, { maxTotalUncompressedSize: 4000 }).catch(
      (e: unknown) => e
    );
    expect(error).toBeInstanceOf(ArchiveLimitExceededError);
    expect(error).not.toBeInstanceOf(ArchiveFileTooLargeError);
    expect(error).toMatchObject({ limit: "maxTotalUncompressedSize", allowed: 4000 });
  });
});

describe("a total limit never widens the per-entry default", () => {
  it("still refuses an entry declaring more than 512 MiB when the total allows 1 GiB", async () => {
    const { createZipSync } = await import("@archive/zip/zip-bytes");
    const { ZipParser } = await import("@archive/unzip/zip-parser");
    const bytes = createZipSync([{ name: "x.bin", data: new Uint8Array([1]) }], { level: 0 });
    // Declare a 600 MiB entry in the central directory; nothing that large is allocated.
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let i = 0; i < bytes.length - 28; i++) {
      if (view.getUint32(i, true) === 0x02014b50) {
        view.setUint32(i + 24, 600 * 1024 * 1024, true);
        break;
      }
    }
    const parser = new ZipParser(bytes, { maxTotalUncompressedSize: 1024 * 1024 * 1024 });
    const error = await parser
      .forEach(async (_entry, getData) => {
        await getData();
      })
      .catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ArchiveFileTooLargeError);
    expect(error).toMatchObject({ limit: "maxEntrySize", allowed: 512 * 1024 * 1024 });
  });
});
