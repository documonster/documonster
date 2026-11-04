/**
 * The `Compression` namespace checks its arguments as the archive builders do and reports
 * damaged input as an `ArchiveError`, on every platform, rather than the codec's own error.
 */
import { ArchiveError, Compression } from "@archive";
import { describe, expect, it } from "vitest";

const bytes = new TextEncoder().encode("hello hello hello");

const compressors = [
  ["deflate", Compression.deflate],
  ["gzip", Compression.gzip],
  ["zlib", Compression.zlib]
] as const;
const compressorsSync = [
  ["deflateSync", Compression.deflateSync],
  ["gzipSync", Compression.gzipSync],
  ["zlibSync", Compression.zlibSync]
] as const;
const decompressors = [
  ["inflate", Compression.inflate],
  ["gunzip", Compression.gunzip],
  ["unzlib", Compression.unzlib],
  ["decompressAuto", Compression.decompressAuto]
] as const;
const decompressorsSync = [
  ["inflateSync", Compression.inflateSync],
  ["gunzipSync", Compression.gunzipSync],
  ["unzlibSync", Compression.unzlibSync],
  ["decompressAutoSync", Compression.decompressAutoSync]
] as const;

describe("Compression arguments", () => {
  it.each([42, -1, NaN, 1.5, "6"])("refuses level %s", async level => {
    for (const [, fn] of compressors) {
      await expect(fn(bytes, { level: level as never })).rejects.toBeInstanceOf(ArchiveError);
    }
    for (const [, fn] of compressorsSync) {
      expect(() => fn(bytes, { level: level as never })).toThrow(ArchiveError);
    }
    for (const create of [
      Compression.createDeflateStream,
      Compression.createGzipStream,
      Compression.createZlibStream
    ]) {
      expect(() => create({ level: level as never })).toThrow(ArchiveError);
    }
  });

  it.each([null, undefined, "text", [1, 2]])("refuses %s as data", async data => {
    for (const [, fn] of [...compressors, ...decompressors]) {
      await expect(fn(data as never)).rejects.toBeInstanceOf(ArchiveError);
    }
    for (const [, fn] of [...compressorsSync, ...decompressorsSync]) {
      expect(() => fn(data as never)).toThrow(ArchiveError);
    }
    expect(() => Compression.detectFormat(data as never)).toThrow(ArchiveError);
    expect(() => Compression.isGzipData(data as never)).toThrow(ArchiveError);
    expect(() => Compression.isZlibData(data as never)).toThrow(ArchiveError);
  });

  it("still accepts every level from 0 to 9", () => {
    for (let level = 0; level <= 9; level++) {
      expect(Compression.inflateSync(Compression.deflateSync(bytes, { level }))).toEqual(bytes);
    }
  });
});

describe("damaged input", () => {
  const garbage = new TextEncoder().encode("garbage!! not compressed at all");

  it.each(decompressors)("%s rejects with an ArchiveError", async (_name, fn) => {
    await expect(fn(garbage)).rejects.toBeInstanceOf(ArchiveError);
  });

  it.each(decompressorsSync)("%s throws an ArchiveError", (_name, fn) => {
    expect(() => fn(garbage)).toThrow(ArchiveError);
  });

  it("keeps the codec's message as the cause", () => {
    try {
      Compression.gunzipSync(garbage);
      expect.unreachable();
    } catch (error) {
      expect((error as Error).message).toMatch(/^Invalid gzip data/);
      expect((error as ArchiveError).cause).toBeInstanceOf(Error);
    }
  });
});
