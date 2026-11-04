/**
 * Corrupt DEFLATE data is malformed input, and every unzip path reports it as an `ArchiveError` — the error the
 * pure-JS inflater already raised — rather than the bare `Error` Node's zlib produces, with that kept as the `cause`.
 */

import { Readable } from "stream";

import { ArchiveError } from "@archive/core/errors";
import { zip } from "@archive/create-archive";
import type { ZipEntry } from "@archive/unzip/stream";
import { createParse } from "@archive/unzip/stream";
import { ZipParser } from "@archive/unzip/zip-parser";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();

/** One deflated entry whose first block header names the reserved block type. */
async function corruptDeflate(options: { descriptor: boolean }): Promise<Uint8Array> {
  const z = zip({ level: 6, modTime: new Date(0) });
  z.add("a.txt", encoder.encode("hello world ".repeat(400)));
  z.add("b.txt", encoder.encode("second"));
  let bytes: Uint8Array;
  if (options.descriptor) {
    const chunks: Uint8Array[] = [];
    for await (const chunk of z.stream()) {
      chunks.push(chunk);
    }
    bytes = Uint8Array.from(Buffer.concat(chunks));
  } else {
    bytes = Uint8Array.from(await z.toBuffer());
  }
  const view = new DataView(bytes.buffer);
  expect(view.getUint16(8, true)).toBe(8);
  const dataStart = 30 + view.getUint16(26, true) + view.getUint16(28, true);
  bytes[dataStart] = 0xff;
  return bytes;
}

async function streamAll(bytes: Uint8Array, thresholdBytes?: number): Promise<unknown> {
  const parse = createParse({ forceStream: true, thresholdBytes, checkCrc32: false });
  parse.on("error", () => {});
  Readable.from([bytes]).pipe(parse);
  try {
    for await (const entry of parse) {
      await (entry as ZipEntry).buffer();
    }
  } catch (error) {
    return error;
  }
  return undefined;
}

function expectTypedInflateError(error: unknown): void {
  expect(error).toBeInstanceOf(ArchiveError);
  expect((error as Error).message).toMatch(/DEFLATE/);
  expect((error as Error).cause).toBeInstanceOf(Error);
}

describe("unzip: corrupt DEFLATE data", () => {
  it("is an ArchiveError from the buffer parser", async () => {
    const parser = new ZipParser(await corruptDeflate({ descriptor: false }), {
      checkCrc32: false
    });
    expectTypedInflateError(await parser.extract("a.txt").catch((e: unknown) => e));
    let syncError: unknown;
    try {
      parser.extractAllSync();
    } catch (error) {
      syncError = error;
    }
    expectTypedInflateError(syncError);
  });

  it("is an ArchiveError from the streaming parser (in-memory inflate)", async () => {
    expectTypedInflateError(await streamAll(await corruptDeflate({ descriptor: false })));
  });

  it("is an ArchiveError from the streaming parser (streamed inflate)", async () => {
    expectTypedInflateError(await streamAll(await corruptDeflate({ descriptor: false }), 0));
  });

  it("is an ArchiveError from the streaming parser (size in a data descriptor)", async () => {
    expectTypedInflateError(await streamAll(await corruptDeflate({ descriptor: true })));
  });
});
