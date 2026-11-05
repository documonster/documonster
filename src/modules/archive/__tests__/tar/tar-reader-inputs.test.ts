/**
 * The TAR readers accept the same sources and fail the same way. `unpackTar` and `openTar`
 * take any `ArchiveSource`, as `unpackTarGz` and `unpackZip` do; anything else is refused
 * with an `ArchiveError`. An archive cut short — inside an entry's data or inside a header —
 * is an error whether it arrives as bytes or as a stream, rather than an empty or shortened
 * result, and a gzip-compressed archive given to a plain TAR reader says so.
 */
import { Archive, ArchiveError, Compression } from "@archive";
import { describe, expect, it } from "vitest";

const tar = Archive.packTarSync([
  ["a.txt", "A"],
  ["big.bin", new Uint8Array(5000).fill(7)]
]);
const decoder = new TextDecoder();
const sources = (bytes: Uint8Array): Array<[string, () => unknown]> => [
  ["Uint8Array", () => bytes],
  ["ArrayBuffer", () => bytes.slice().buffer],
  ["Blob", () => new Blob([bytes as BlobPart])],
  ["ReadableStream", () => new Blob([bytes as BlobPart]).stream()],
  [
    "async iterable",
    () =>
      (async function* () {
        yield bytes;
      })()
  ]
];

describe("TAR reader sources", () => {
  it.each(sources(tar))("unpackTar reads a %s", async (_name, make) => {
    const map = await Archive.unpackTar(make() as never);
    expect(decoder.decode(map.get("a.txt")!.data)).toBe("A");
    expect(map.get("big.bin")!.data.length).toBe(5000);
  });

  it.each(sources(tar))("openTar reads a %s", async (_name, make) => {
    expect(await Archive.openTar(make() as never).list()).toEqual(["a.txt", "big.bin"]);
  });

  it("unpackTarSync reads an ArrayBuffer", () => {
    expect([...Archive.unpackTarSync(tar.slice().buffer as never).keys()]).toEqual([
      "a.txt",
      "big.bin"
    ]);
  });

  it.each([null, 42, {}, "not bytes"])("refuses %s", async source => {
    await expect(Archive.unpackTar(source as never)).rejects.toBeInstanceOf(ArchiveError);
    if (typeof source === "string") {
      // A string is a source (its UTF-8 bytes), just not a TAR: refused when it is read.
      await expect(Archive.openTar(source).list()).rejects.toBeInstanceOf(ArchiveError);
    } else {
      // Anything else is refused when the reader is opened.
      expect(() => Archive.openTar(source as never)).toThrow(ArchiveError);
    }
    expect(() => Archive.unpackTarSync(source as never)).toThrow(ArchiveError);
    expect(() => Archive.parseTar(source as never)).toThrow(ArchiveError);
  });

  it("parseTar reads bytes in memory and refuses the rest instead of returning no entries", () => {
    for (const [, make] of sources(tar).slice(0, 2)) {
      expect(Archive.parseTar(make() as never).map(entry => entry.info.path)).toEqual([
        "a.txt",
        "big.bin"
      ]);
    }
    expect(Archive.parseTar(new DataView(tar.buffer)).length).toBe(2);
    for (const [, make] of sources(tar).slice(2)) {
      expect(() => Archive.parseTar(make() as never)).toThrow(ArchiveError);
    }
  });

  it.each(sources(tar))(
    "parseTarStream reads a %s, as parseTarGzStream does",
    async (_name, make) => {
      const paths: string[] = [];
      for await (const entry of Archive.parseTarStream(make() as never)) {
        paths.push(entry.info.path);
        entry.discard();
      }
      expect(paths).toEqual(["a.txt", "big.bin"]);
    }
  );

  it("parseTarStream refuses what is not a source when it is called", () => {
    // Checked at the call, with its options, rather than on the first `next()`
    expect(() => Archive.parseTarStream(42 as never)).toThrow(ArchiveError);
  });
});

describe("a TAR cut short", () => {
  const inData = tar.subarray(0, 512 * 3 + 100); // inside big.bin's data
  const inHeader = tar.subarray(0, 512 * 2 + 100); // inside big.bin's header

  it.each([
    ["inside an entry's data", inData],
    ["inside a header", inHeader],
    ["before the first header ends", tar.subarray(0, 100)]
  ])("is an error %s, from bytes and from a stream", async (_name, bytes) => {
    expect(() => Archive.unpackTarSync(bytes)).toThrow(ArchiveError);
    expect(() => Archive.parseTar(bytes)).toThrow(ArchiveError);
    await expect(Archive.unpackTar(bytes)).rejects.toBeInstanceOf(ArchiveError);
    await expect(
      Archive.unpackTar(
        (async function* () {
          yield bytes;
        })()
      )
    ).rejects.toBeInstanceOf(ArchiveError);
    const tgz = Compression.gzipSync(bytes);
    await expect(Archive.unpackTarGz(tgz)).rejects.toBeInstanceOf(ArchiveError);
    expect(() => Archive.unpackTarGzSync(tgz)).toThrow(ArchiveError);
  });

  it("an archive that ends after a whole entry, with no end blocks, still reads", () => {
    const withoutEnd = tar.subarray(0, 1024); // a.txt's header and data block
    expect([...Archive.unpackTarSync(withoutEnd).keys()]).toEqual(["a.txt"]);
  });

  it("an empty input is an empty archive", async () => {
    expect(Archive.unpackTarSync(new Uint8Array(0)).size).toBe(0);
    expect((await Archive.unpackTar(new Uint8Array(0))).size).toBe(0);
  });
});

describe("a gzip-compressed TAR given to a TAR reader", () => {
  const tgz = Archive.packTarGzSync([["a.txt", "A"]]);

  it("is refused with a pointer to the tar.gz reader", async () => {
    expect(() => Archive.unpackTarSync(tgz)).toThrow(/gzip.*unpackTarGz/);
    expect(() => Archive.parseTar(tgz)).toThrow(/gzip/);
    await expect(Archive.unpackTar(tgz)).rejects.toThrow(/gzip/);
    await expect(Archive.openTar(tgz).list()).rejects.toThrow(/gzip/);
  });
});
