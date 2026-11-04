/**
 * Every archive reader checks its limit options where it is called, as `openZip` / `openTar`
 * do: a limit that is not a whole number from 0 (or `Infinity`) is an `ArchiveArgumentError`
 * naming the call and `options.<limit>`. A mistyped limit used to read the archive with no limit
 * at all, and a negative one was reported as the limit being hit.
 */
import { Archive, ArchiveArgumentError } from "@archive/index";
import { describe, expect, it } from "vitest";

const ENTRIES: Array<[string, string]> = [["a.txt", "A"]];
const zip = Archive.packZipSync(ENTRIES);
const tar = Archive.packTarSync(ENTRIES);
const tgz = Archive.packTarGzSync(ENTRIES);

type Reader = readonly [name: string, read: (options: unknown) => unknown, limits: string[]];

const ZIP_LIMITS = ["maxEntrySize", "maxEntries", "maxTotalUncompressedSize"];
const TAR_LIMITS = ["maxEntrySize", "maxFileSize"];

const drain = async (iterable: AsyncIterable<unknown>) => {
  for await (const _ of iterable) {
    // read it all
  }
};

const READERS: Reader[] = [
  ["Archive.unpackZip", o => Archive.unpackZip(zip, o as never), ZIP_LIMITS],
  ["Archive.unpackZipSync", o => Archive.unpackZipSync(zip, o as never), ZIP_LIMITS],
  ["Archive.unpackTar", o => Archive.unpackTar(tar, o as never), TAR_LIMITS],
  ["Archive.unpackTarSync", o => Archive.unpackTarSync(tar, o as never), TAR_LIMITS],
  ["Archive.parseTar", o => Archive.parseTar(tar, o as never), TAR_LIMITS],
  ["Archive.parseTarStream", o => drain(Archive.parseTarStream(tar, o as never)), TAR_LIMITS],
  ["Archive.unpackTarGz", o => Archive.unpackTarGz(tgz, o as never), TAR_LIMITS],
  ["Archive.unpackTarGzSync", o => Archive.unpackTarGzSync(tgz, o as never), TAR_LIMITS],
  ["Archive.parseTarGz", o => Archive.parseTarGz(tgz, o as never), TAR_LIMITS],
  ["Archive.parseTarGzStream", o => drain(Archive.parseTarGzStream(tgz, o as never)), TAR_LIMITS],
  ["Archive.openZip", o => Archive.openZip(zip, o as never), ["maxEntrySize", "maxEntries"]],
  ["Archive.openTar", o => Archive.openTar(tar, o as never), TAR_LIMITS],
  ["Archive.openTarGz", o => Archive.openTarGz(tgz, o as never), TAR_LIMITS]
];

/** The error `run` throws or rejects with, or `undefined` when it succeeds. */
async function failure(run: () => unknown): Promise<unknown> {
  try {
    await run();
    return undefined;
  } catch (error) {
    return error;
  }
}

describe("every archive reader checks its limit options", () => {
  for (const [name, read, limits] of READERS) {
    for (const limit of limits) {
      it.each(["x", -1, Number.NaN, 1.5, null])(`${name}: options.${limit} = %s`, async bad => {
        const error = await failure(() => read({ [limit]: bad }));
        expect(error).toBeInstanceOf(ArchiveArgumentError);
        expect((error as ArchiveArgumentError).field).toBe(`options.${limit}`);
        expect((error as Error).message.startsWith(`${name}: `)).toBe(true);
      });
    }

    it(`${name}: options that are not an object`, async () => {
      const error = await failure(() => read(5));
      expect(error).toBeInstanceOf(ArchiveArgumentError);
      expect((error as ArchiveArgumentError).field).toBe("options");
    });

    it(`${name}: Infinity and 0-based counts are limits`, async () => {
      expect(await failure(() => read(Object.fromEntries(limits.map(l => [l, Infinity]))))).toBe(
        undefined
      );
    });
  }

  it("a stream reader refuses its options when it is called, before it is iterated", () => {
    expect(() => Archive.parseTarStream(tar, { maxEntrySize: -1 })).toThrow(ArchiveArgumentError);
    expect(() => Archive.parseTarGzStream(tgz, { maxEntrySize: -1 })).toThrow(ArchiveArgumentError);
  });
});
