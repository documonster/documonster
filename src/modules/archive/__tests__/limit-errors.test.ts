/**
 * Every limit an archive reader enforces says which limit (`limit`, the option's name) and its
 * value (`allowed`). The per-entry limit is a `ArchiveFileTooLargeError`, which said both only in its
 * message; it now carries the same `limit` / `allowed` fields as `ArchiveLimitError` (plus the
 * entry's `path`), while staying outside that class so the published `ArchiveLimit` union is
 * unchanged.
 */
import {
  Archive,
  ArchiveError,
  ArchiveLimitError,
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
  expect(error).not.toBeInstanceOf(ArchiveLimitError);
  expect(error).toMatchObject({ limit: "maxEntrySize", allowed, path: "big.bin" });
  expect((error as Error).name).toBe("ArchiveFileTooLargeError");
}

describe("maxEntrySize is an ArchiveLimitError naming the limit and its value", () => {
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

  it("reads like an ArchiveLimitError: the same field names and meaning", () => {
    const limitError = new ArchiveLimitError("maxEntries", 3);
    const sizeError = new ArchiveFileTooLargeError("x.bin", "r", { allowed: 10 });
    const describeLimit = (e: { limit: string; allowed: number }) => `${e.limit}=${e.allowed}`;
    expect(describeLimit(limitError)).toBe("maxEntries=3");
    expect(describeLimit(sizeError)).toBe("maxEntrySize=10");
    expectTypeOf(sizeError.limit).toEqualTypeOf<"maxEntrySize">();
    expectTypeOf(sizeError.allowed).toEqualTypeOf<number>();
  });

  it("ArchiveLimit keeps its three published members", () => {
    expectTypeOf<ArchiveLimit>().toEqualTypeOf<
      "maxOutputLength" | "maxTotalUncompressedSize" | "maxEntries"
    >();
    expectTypeOf(new ArchiveLimitError("maxEntries", 1).limit).toEqualTypeOf<ArchiveLimit>();
  });
});
