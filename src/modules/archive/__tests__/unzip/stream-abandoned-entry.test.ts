/**
 * An entry stream must settle — and must never crash the process — however its consumer treats it.
 *
 * Two failures are covered. A CRC-32 mismatch is reported by the entry's own flush, after its bytes have been handed
 * on; a consumer that reads an entry without listening for `error` (or that has already moved on) used to receive it
 * as an unhandled `error` event, which Node turns into a process crash even when the caller catches the parse
 * failure. And an archive that ends inside an entry's data used to leave that entry open forever, so whoever was
 * reading it waited for an `end` that never came.
 */

import { Readable } from "stream";

import { Crc32MismatchError, ZipParseError } from "@archive/core/errors";
import { zip } from "@archive/create-archive";
import type { ZipEntry } from "@archive/unzip/stream";
import { createParse } from "@archive/unzip/stream";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

const encoder = new TextEncoder();

/** A one-entry archive with sizes in the local header, whose local-header CRC-32 is wrong. */
async function zipWithBadCrc(): Promise<Uint8Array> {
  const z = zip({ level: 6, modTime: new Date(0) });
  z.add("a.txt", encoder.encode("hello world ".repeat(200)));
  z.add("b.txt", encoder.encode("second"));
  const bytes = Uint8Array.from(await z.toBuffer());
  // Local file header: signature(4) version(2) flags(2) method(2) time(2) date(2) crc32(4)
  bytes[14] ^= 0xff;
  return bytes;
}

/** Incompressible bytes, so a truncation lands inside the entry's compressed data. */
function noise(length: number): Uint8Array {
  const out = new Uint8Array(length);
  let seed = 12345;
  for (let i = 0; i < length; i++) {
    seed = (seed * 1103515245 + 12345) >>> 0;
    out[i] = seed >>> 24;
  }
  return out;
}

function settleWithin<T>(promise: Promise<T>, ms: number): Promise<T> {
  return Promise.race([
    promise,
    new Promise<T>((_, reject) =>
      setTimeout(() => reject(new Error(`did not settle within ${ms} ms`)), ms)
    )
  ]);
}

let uncaught: unknown[] = [];
const onUncaught = (error: unknown): void => {
  uncaught.push(error);
};

beforeEach(() => {
  uncaught = [];
  process.on("uncaughtException", onUncaught);
});

afterEach(() => {
  process.off("uncaughtException", onUncaught);
});

describe("unzip stream: an entry whose consumer does not listen for errors", () => {
  it("reports a CRC-32 mismatch through the parser instead of crashing", async () => {
    const parse = createParse({ forceStream: true });
    Readable.from([await zipWithBadCrc()]).pipe(parse);

    const failure = parse.promise().then(
      () => undefined,
      (error: unknown) => error
    );
    const iterated = (async () => {
      for await (const entry of parse) {
        // Reads the bytes, but has no `error` listener — as a consumer that has already given up would not.
        (entry as ZipEntry).resume();
      }
    })();

    await expect(settleWithin(iterated, 5000)).rejects.toBeInstanceOf(Crc32MismatchError);
    expect(await failure).toBeInstanceOf(Crc32MismatchError);
    await new Promise(resolve => setTimeout(resolve, 50));
    expect(uncaught).toEqual([]);
  });

  it("surfaces the mismatch on the parser when the consumer has no error listener", async () => {
    const parse = createParse({ forceStream: true });
    Readable.from([await zipWithBadCrc()]).pipe(parse);
    const done = parse.promise();
    parse.on("entry", (entry: ZipEntry) => entry.resume());
    parse.resume();

    await expect(settleWithin(done, 5000)).rejects.toBeInstanceOf(Crc32MismatchError);
    await new Promise(resolve => setTimeout(resolve, 50));
    expect(uncaught).toEqual([]);
  });

  it("reports a truncation through the parser when the consumer has no error listener", async () => {
    const z = zip({ level: 1, modTime: new Date(0) });
    z.add("big.bin", noise(200_000));
    const bytes = await z.toBuffer();
    const parse = createParse({ forceStream: true, thresholdBytes: 0 });
    Readable.from([bytes.subarray(0, Math.floor(bytes.length / 2))]).pipe(parse);
    const done = parse.promise();
    parse.on("entry", (entry: ZipEntry) => entry.resume());
    parse.resume();

    await expect(settleWithin(done, 5000)).rejects.toBeInstanceOf(ZipParseError);
    await new Promise(resolve => setTimeout(resolve, 50));
    expect(uncaught).toEqual([]);
  });

  it("still reports the mismatch to a consumer that does listen", async () => {
    const parse = createParse({ forceStream: true });
    Readable.from([await zipWithBadCrc()]).pipe(parse);
    parse.promise().catch(() => {});

    let entryError: unknown;
    for await (const entry of parse) {
      try {
        await (entry as ZipEntry).buffer();
      } catch (error) {
        entryError = error;
        break;
      }
    }
    expect(entryError).toBeInstanceOf(Crc32MismatchError);
    await new Promise(resolve => setTimeout(resolve, 50));
    expect(uncaught).toEqual([]);
  });
});

describe("unzip stream: an archive that ends inside an entry", () => {
  for (const [label, thresholdBytes] of [
    ["in-memory inflate", undefined],
    ["streamed inflate", 0]
  ] as const) {
    it(`fails the entry being read (${label})`, async () => {
      const z = zip({ level: 1, modTime: new Date(0) });
      z.add("big.bin", noise(200_000));
      const bytes = await z.toBuffer();
      const cut = bytes.subarray(0, Math.floor(bytes.length / 2));

      const parse = createParse({ forceStream: true, thresholdBytes });
      parse.on("error", () => {});
      Readable.from([cut]).pipe(parse);

      const read = (async () => {
        for await (const entry of parse) {
          await (entry as ZipEntry).buffer();
        }
      })();

      const error = await settleWithin(read, 5000).then(
        () => undefined,
        (e: unknown) => e
      );
      expect(error).toBeInstanceOf(ZipParseError);
      await expect(settleWithin(parse.promise(), 5000)).rejects.toBeInstanceOf(ZipParseError);
      expect(uncaught).toEqual([]);
    });
  }

  // Sizes in a data descriptor after the data: the entry's end is found by scanning for the descriptor, and the scan
  // used to stop after flushing what it had, so input that had already finished was never looked at again.
  for (const level of [0, 1]) {
    it(`fails an entry whose size is in a data descriptor (level ${level})`, async () => {
      const z = zip({ level, modTime: new Date(0) });
      z.add("a.txt", encoder.encode("first"));
      z.add("big.bin", noise(200_000));
      const chunks: Uint8Array[] = [];
      for await (const chunk of z.stream()) {
        chunks.push(chunk);
      }
      const bytes = Buffer.concat(chunks);
      expect(bytes[6]! & 0x08).toBe(0x08);

      for (const fraction of [0.3, 0.6, 0.9]) {
        const parse = createParse({ forceStream: true });
        parse.on("error", () => {});
        Readable.from([bytes.subarray(0, Math.floor(bytes.length * fraction))]).pipe(parse);
        const read = (async () => {
          for await (const entry of parse) {
            await (entry as ZipEntry).buffer();
          }
        })();
        const error = await settleWithin(read, 5000).then(
          () => undefined,
          (e: unknown) => e
        );
        expect(error).toBeInstanceOf(ZipParseError);
      }
      expect(uncaught).toEqual([]);
    });
  }
});
