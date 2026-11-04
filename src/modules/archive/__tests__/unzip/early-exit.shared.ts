/**
 * Leaving a streaming archive iteration early — `break`, `return` or a `throw` from the loop
 * body — must settle at once and hand back what the iteration held: the caller's source,
 * the parser and its inflater.
 *
 * It used not to whenever the entry being left was larger than the parser's per-entry
 * buffer: the iteration's cleanup waited for the parser to finish, the parser waited for
 * someone to drain the entry, and nobody would. The shared suite runs in Node and in a
 * browser; the Node file adds file and Node streams.
 */

import { Archive, ArchiveError, Crc32MismatchError } from "@archive";
import { parseTarGzStream as deprecatedParseTarGzStream } from "@archive/tar/tar-gzip";
import { describe, expect, it } from "vitest";

import { createDataDescriptorZip } from "./test-helpers";

const encoder = new TextEncoder();

/** Larger than the parser buffers per entry, so an entry nobody drains blocks it. */
export const BIG = 600 * 1024;

export function noise(length: number, seed = 12345): Uint8Array {
  const out = new Uint8Array(length);
  let s = seed;
  for (let i = 0; i < length; i++) {
    s = (s * 1103515245 + 12345) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
}

function text(length: number): Uint8Array {
  return encoder
    .encode("the quick brown fox jumps over the lazy dog ".repeat(length / 44 + 1))
    .slice(0, length);
}

export interface Fixture {
  readonly names: readonly string[];
  readonly data: ReadonlyMap<string, Uint8Array>;
}

/** Small and large entries in alternation, so every break position has a large entry near it. */
export function fixture(): Fixture {
  const data = new Map<string, Uint8Array>([
    ["1-a.txt", encoder.encode("first")],
    ["2-big-noise.bin", noise(BIG)],
    ["3-c.txt", encoder.encode("middle")],
    ["4-big-text.txt", text(BIG)],
    ["5-z.txt", encoder.encode("last")]
  ]);
  return { names: [...data.keys()], data };
}

export const ZIP_KINDS: Record<string, (f: Fixture) => Promise<Uint8Array>> = {
  stored: f => Archive.packZip([...f.data], { level: 0 }),
  deflated: f => Archive.packZip([...f.data], { level: 6, smartStore: false }),
  "data descriptor": f =>
    createDataDescriptorZip([...f.data].map(([name, bytes]) => ({ name, data: bytes })))
};

/** A source handed to the reader, and how to tell that the reader has let go of it. */
export interface TrackedSource {
  readonly source: Parameters<typeof Archive.openZip>[0];
  /** Resolves once the source has been released: destroyed, cancelled or returned. */
  released(): Promise<void>;
}

export type SourceFactory = (bytes: Uint8Array) => TrackedSource;

function chunks(bytes: Uint8Array, size = 64 * 1024): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let i = 0; i < bytes.length; i += size) {
    out.push(bytes.slice(i, i + size));
  }
  return out;
}

function deferred(): { promise: Promise<void>; resolve: () => void } {
  let resolve!: () => void;
  const promise = new Promise<void>(r => (resolve = r));
  return { promise, resolve };
}

export const PORTABLE_SOURCES: Record<string, SourceFactory> = {
  "Web ReadableStream": bytes => {
    const parts = chunks(bytes);
    const done = deferred();
    let index = 0;
    const stream = new ReadableStream<Uint8Array>(
      {
        pull(controller) {
          if (index < parts.length) {
            controller.enqueue(parts[index++]!);
          } else {
            controller.close();
            done.resolve();
          }
        },
        cancel() {
          done.resolve();
        }
      },
      { highWaterMark: 0 }
    );
    return {
      source: stream,
      async released() {
        await done.promise;
        // Cancelled or read to the end — and handed back unlocked either way
        await waitUntil(() => !stream.locked);
      }
    };
  },
  "async generator": bytes => {
    const done = deferred();
    async function* generate(): AsyncGenerator<Uint8Array> {
      try {
        for (const part of chunks(bytes)) {
          await Promise.resolve();
          yield part;
        }
      } finally {
        done.resolve();
      }
    }
    return { source: generate(), released: () => done.promise };
  }
};

export function settleWithin<T>(promise: Promise<T>, ms = 10_000, what = "iteration"): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => {
      timer = setTimeout(() => reject(new Error(`${what} did not settle within ${ms} ms`)), ms);
    })
  ]).finally(() => clearTimeout(timer));
}

export async function waitUntil(test: () => boolean, ms = 10_000): Promise<void> {
  const start = Date.now();
  while (!test()) {
    if (Date.now() - start > ms) {
      throw new Error("condition not reached");
    }
    await new Promise(resolve => setTimeout(resolve, 5));
  }
}

async function readAll(iterable: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  for await (const part of iterable) {
    parts.push(part);
  }
  return Archive.concatUint8Arrays(parts);
}

export type LeaveMode = "unread" | "read" | "partly read" | "throw";

export const LEAVE_MODES: readonly LeaveMode[] = ["unread", "read", "partly read", "throw"];

/** Where to leave: the entry's index in {@link Fixture.names}. */
export const POSITIONS: Record<string, number> = { first: 0, middle: 1, last: 4 };

export interface EntryLike {
  readonly path: string;
  bytes(): Promise<Uint8Array>;
  stream(): AsyncIterable<Uint8Array>;
}

/**
 * Iterate `entries`, reading each entry before `stopAt`, and leave at `stopAt` in the given
 * way. Resolves with what was read, once the iteration has settled.
 */
export async function leaveAt(
  entries: AsyncIterable<EntryLike>,
  stopAt: string,
  mode: LeaveMode
): Promise<Map<string, Uint8Array>> {
  const read = new Map<string, Uint8Array>();
  const thrown = new Error("left by throwing");
  try {
    for await (const entry of entries) {
      if (entry.path !== stopAt) {
        read.set(entry.path, await entry.bytes());
        continue;
      }
      if (mode === "read") {
        read.set(entry.path, await entry.bytes());
      } else if (mode === "partly read") {
        for await (const _part of entry.stream()) {
          break;
        }
      } else if (mode === "throw") {
        throw thrown;
      }
      break;
    }
  } catch (error) {
    if (error !== thrown) {
      throw error;
    }
  }
  return read;
}

/**
 * Byte-for-byte equality, reported compactly. `toEqual` walks a large typed array element by
 * element and takes seconds on the entries here; a Node source also yields `Buffer`s, which
 * TAR hands back as views and `toEqual` would tell apart from a `Uint8Array`.
 */
export function expectBytes(
  actual: Uint8Array | undefined,
  expected: Uint8Array | undefined
): void {
  expect(actual).toBeDefined();
  expect(expected).toBeDefined();
  const a = actual!;
  const b = expected!;
  let firstDifference = a.length === b.length ? -1 : Math.min(a.length, b.length);
  for (let i = 0; firstDifference === -1 && i < a.length; i++) {
    if (a[i] !== b[i]) {
      firstDifference = i;
    }
  }
  expect({ length: a.length, firstDifference }).toEqual({ length: b.length, firstDifference: -1 });
}

export function expectRead(
  read: Map<string, Uint8Array>,
  f: Fixture,
  upTo: string,
  inclusive: boolean
) {
  const expected = f.names.slice(0, f.names.indexOf(upTo) + (inclusive ? 1 : 0));
  expect([...read.keys()]).toEqual(expected);
  for (const name of expected) {
    expectBytes(read.get(name), f.data.get(name));
  }
}

/** The ZIP streaming reader, left in every way at every position, over each source. */
export function defineZipEarlyExitSuite(sources: Record<string, SourceFactory>): void {
  const f = fixture();

  describe("Archive.openZip(stream).entries(): leaving the loop early", () => {
    for (const [kind, build] of Object.entries(ZIP_KINDS)) {
      for (const [sourceName, makeSource] of Object.entries(sources)) {
        for (const [where, index] of Object.entries(POSITIONS)) {
          for (const mode of LEAVE_MODES) {
            it(`${kind}, ${sourceName}: ${mode} at the ${where} entry`, async () => {
              const zip = await build(f);
              const tracked = makeSource(zip);
              const stopAt = f.names[index]!;
              const read = await settleWithin(
                leaveAt(Archive.openZip(tracked.source).entries(), stopAt, mode)
              );
              expectRead(read, f, stopAt, mode === "read");
              await settleWithin(tracked.released(), 10_000, "releasing the source");
            });
          }
        }
      }
    }
  });

  describe("Archive.openZip(stream).entries(): entries that are not read", () => {
    const sourceName = Object.keys(sources)[0]!;
    const makeSource = sources[sourceName]!;

    for (const [kind, build] of Object.entries(ZIP_KINDS)) {
      it(`${kind}: skipping every entry walks the whole archive (${sourceName})`, async () => {
        const tracked = makeSource(await build(f));
        const seen: string[] = [];
        await settleWithin(
          (async () => {
            for await (const entry of Archive.openZip(tracked.source).entries()) {
              seen.push(entry.path);
            }
          })()
        );
        expect(seen).toEqual(f.names);
        await settleWithin(tracked.released(), 10_000, "releasing the source");
      });

      it(`${kind}: a small entry moved past unread can still be read (${sourceName})`, async () => {
        const tracked = makeSource(await build(f));
        const kept: Archive.UnzipEntry[] = [];
        await settleWithin(
          (async () => {
            for await (const entry of Archive.openZip(tracked.source).entries()) {
              kept.push(entry);
            }
          })()
        );
        const byPath = new Map(kept.map(entry => [entry.path, entry]));
        for (const name of ["1-a.txt", "3-c.txt", "5-z.txt"]) {
          expectBytes(await byPath.get(name)!.bytes(), f.data.get(name));
          expectBytes(await readAll(byPath.get(name)!.stream()), f.data.get(name));
        }
        // A large one was skipped to let the archive move on, and says so instead of stalling
        await expect(settleWithin(byPath.get("2-big-noise.bin")!.bytes())).rejects.toThrow(
          /moved past/
        );
        await expect(settleWithin(readAll(byPath.get("4-big-text.txt")!.stream()))).rejects.toThrow(
          /moved past/
        );
      });

      it(`${kind}: an entry found, then read after leaving the loop (${sourceName})`, async () => {
        for (const name of ["1-a.txt", "3-c.txt", "5-z.txt"]) {
          const tracked = makeSource(await build(f));
          let found: Archive.UnzipEntry | undefined;
          for await (const entry of Archive.openZip(tracked.source).entries()) {
            if (entry.path === name) {
              found = entry;
              break;
            }
          }
          expectBytes(await settleWithin(found!.bytes()), f.data.get(name));
          await settleWithin(tracked.released(), 10_000, "releasing the source");
        }
      });

      it(`${kind}: a large entry found, then read after leaving the loop, says the loop ended (${sourceName})`, async () => {
        for (const name of ["2-big-noise.bin", "4-big-text.txt"]) {
          const tracked = makeSource(await build(f));
          let found: Archive.UnzipEntry | undefined;
          for await (const entry of Archive.openZip(tracked.source).entries()) {
            if (entry.path === name) {
              found = entry;
              break;
            }
          }
          // The caller left on this very entry: the reader did not move past it to another.
          const read = settleWithin(found!.bytes());
          await expect(read).rejects.toThrow(ArchiveError);
          await expect(read).rejects.toThrow(/iteration ended before it was read/);
          await expect(read).rejects.not.toThrow(/moved past/);
          await settleWithin(tracked.released(), 10_000, "releasing the source");
        }
      });

      it(`${kind}: a discarded entry says so when read (${sourceName})`, async () => {
        const tracked = makeSource(await build(f));
        let discarded: Archive.UnzipEntry | undefined;
        const seen: string[] = [];
        for await (const entry of Archive.openZip(tracked.source).entries()) {
          seen.push(entry.path);
          if (entry.path === "3-c.txt") {
            entry.discard();
            discarded = entry;
          } else {
            await entry.bytes();
          }
        }
        expect(seen).toEqual(f.names);
        await expect(settleWithin(discarded!.bytes())).rejects.toBeInstanceOf(ArchiveError);
      });
    }
  });

  describe("Archive.openZip(stream).entries(): a corrupt entry", () => {
    const sourceName = Object.keys(sources)[0]!;
    const makeSource = sources[sourceName]!;

    /** Flip the CRC-32 recorded for `name`, in its local header and its central record. */
    function corruptCrcOf(zip: Uint8Array, name: string): Uint8Array {
      const out = zip.slice();
      const view = new DataView(out.buffer);
      const nameBytes = encoder.encode(name);
      const matches = (at: number) => nameBytes.every((b, i) => out[at + i] === b);
      for (let i = 0; i + 46 <= out.length; i++) {
        const signature = view.getUint32(i, true);
        if (signature === 0x04034b50 && matches(i + 30)) {
          view.setUint32(i + 14, view.getUint32(i + 14, true) ^ 0x5a5a5a5a, true);
        } else if (signature === 0x02014b50 && matches(i + 46)) {
          view.setUint32(i + 16, view.getUint32(i + 16, true) ^ 0x5a5a5a5a, true);
        }
      }
      return out;
    }

    type Visit = (entry: Archive.UnzipEntry, seen: string[]) => Promise<void>;
    const visits: Record<string, Visit> = {
      "skipped unread": async () => {},
      discarded: async entry => entry.discard(),
      "read, failure caught": async (entry, seen) => {
        await expect(entry.bytes()).rejects.toBeInstanceOf(Crc32MismatchError);
        seen.push("caught");
      },
      "streamed, failure caught": async (entry, seen) => {
        await expect(readAll(entry.stream())).rejects.toBeInstanceOf(Crc32MismatchError);
        seen.push("caught");
      }
    };

    for (const kind of ["stored", "deflated"]) {
      for (const victim of ["3-c.txt", "4-big-text.txt"]) {
        for (const [visitName, visit] of Object.entries(visits)) {
          it(`${kind}, ${victim} ${visitName}: the iteration goes on, as over bytes`, async () => {
            const zip = corruptCrcOf(await ZIP_KINDS[kind]!(f), victim);
            const walk = async (source: Parameters<typeof Archive.openZip>[0]) => {
              const seen: string[] = [];
              for await (const entry of Archive.openZip(source).entries()) {
                if (entry.path === victim) {
                  await visit(entry, seen);
                } else {
                  expectBytes(await entry.bytes(), f.data.get(entry.path));
                }
                seen.push(entry.path);
              }
              return seen;
            };
            const overBytes = await walk(zip);
            const overStream = await settleWithin(walk(makeSource(zip).source));
            expect(overStream).toEqual(overBytes);
            expect(overStream.filter(p => p !== "caught")).toEqual(f.names);
          });
        }
      }

      it(`${kind}: a failure not caught leaves the loop, as over bytes`, async () => {
        const zip = corruptCrcOf(await ZIP_KINDS[kind]!(f), "3-c.txt");
        const tracked = makeSource(zip);
        const walk = async (source: Parameters<typeof Archive.openZip>[0]) => {
          for await (const entry of Archive.openZip(source).entries()) {
            await entry.bytes();
          }
        };
        await expect(walk(zip)).rejects.toBeInstanceOf(Crc32MismatchError);
        await expect(settleWithin(walk(tracked.source))).rejects.toBeInstanceOf(Crc32MismatchError);
        await settleWithin(tracked.released(), 10_000, "releasing the source");
      });
    }

    it("a skipped entry moved past unread reports its failure when it is read later", async () => {
      const zip = corruptCrcOf(await ZIP_KINDS.deflated!(f), "3-c.txt");
      const kept: Archive.UnzipEntry[] = [];
      for await (const entry of Archive.openZip(makeSource(zip).source).entries()) {
        kept.push(entry);
      }
      const victim = kept.find(entry => entry.path === "3-c.txt")!;
      await expect(victim.bytes()).rejects.toBeInstanceOf(Crc32MismatchError);
      expectBytes(await kept[0]!.bytes(), f.data.get("1-a.txt"));
    });
  });

  describe("RemoteZipReader.entriesGenerator(): leaving the loop early", () => {
    for (const [kind, build] of Object.entries(ZIP_KINDS)) {
      for (const [where, index] of Object.entries(POSITIONS)) {
        for (const mode of LEAVE_MODES) {
          it(`${kind}: ${mode} at the ${where} entry`, async () => {
            const reader = await Archive.RemoteZipReader.fromReader(
              Archive.createBufferReader(await build(f))
            );
            const entries: AsyncIterable<EntryLike> = {
              async *[Symbol.asyncIterator]() {
                for await (const { entry, getData } of reader.entriesGenerator()) {
                  yield {
                    path: entry.path,
                    bytes: getData,
                    async *stream() {
                      yield await getData();
                    }
                  };
                }
              }
            };
            const stopAt = f.names[index]!;
            const read = await settleWithin(leaveAt(entries, stopAt, mode));
            expectRead(read, f, stopAt, mode === "read");
          });
        }
      }
    }
  });
}

/** The TAR and tar.gz streaming readers: every one of them, left early at every position. */
export function defineTarEarlyExitSuite(sources: Record<string, SourceFactory>): void {
  const f = fixture();
  const tarEntries = () => [...f.data].map(([name, source]) => ({ name, source }));

  interface TarWay {
    build: () => Promise<Uint8Array>;
    entries: (source: TrackedSource["source"]) => AsyncIterable<EntryLike>;
  }

  function fromTarEntries(
    iterable: AsyncIterable<{ info: { path: string }; data(): Promise<Uint8Array> }>
  ): AsyncIterable<EntryLike> {
    return {
      async *[Symbol.asyncIterator]() {
        for await (const entry of iterable) {
          yield {
            path: entry.info.path,
            bytes: () => entry.data(),
            async *stream() {
              yield await entry.data();
            }
          };
        }
      }
    };
  }

  const ways: Record<string, TarWay> = {
    "openTar(stream).entries()": {
      build: () => Archive.packTar(tarEntries()),
      entries: source => Archive.openTar(source).entries()
    },
    "openTarGz(stream).entries()": {
      build: () => Archive.packTarGz(tarEntries()),
      entries: source => Archive.openTarGz(source).entries()
    },
    parseTarStream: {
      build: () => Archive.packTar(tarEntries()),
      entries: source => fromTarEntries(Archive.parseTarStream(source as AsyncIterable<Uint8Array>))
    },
    parseTarGzStream: {
      build: () => Archive.packTarGz(tarEntries()),
      entries: source => fromTarEntries(Archive.parseTarGzStream(source))
    },
    "parseTarGzStream (deprecated flat)": {
      build: () => Archive.packTarGz(tarEntries()),
      entries: source => fromTarEntries(deprecatedParseTarGzStream(source))
    }
  };

  describe("TAR streaming readers: leaving the loop early", () => {
    for (const [wayName, way] of Object.entries(ways)) {
      for (const [sourceName, makeSource] of Object.entries(sources)) {
        for (const [where, index] of Object.entries(POSITIONS)) {
          for (const mode of LEAVE_MODES) {
            it(`${wayName}, ${sourceName}: ${mode} at the ${where} entry`, async () => {
              const tracked = makeSource(await way.build());
              const stopAt = f.names[index]!;
              const read = await settleWithin(leaveAt(way.entries(tracked.source), stopAt, mode));
              expectRead(read, f, stopAt, mode === "read");
              await settleWithin(tracked.released(), 10_000, "releasing the source");
            });
          }
        }
      }
    }
  });
}
