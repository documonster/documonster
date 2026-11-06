/**
 * A ZIP whose header fields point past its data is malformed input, and is reported as a `ZipParseError` — not as the
 * `RangeError` the byte reader underneath raises, which a caller cannot tell apart from a bug in this library.
 */

import { ZipParseError } from "@archive/core/errors";
import { ZipParser } from "@archive/unzip/zip-parser";
import { createZip } from "@archive/zip/zip-bytes";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();

async function oneEntryZip(): Promise<{ bytes: Uint8Array; centralDirectory: number }> {
  const bytes = await createZip([{ name: "a.txt", data: encoder.encode("hello") }]);
  const centralDirectory = Buffer.from(bytes).indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  expect(centralDirectory).toBeGreaterThan(0);
  return { bytes: Uint8Array.from(bytes), centralDirectory };
}

describe("ZipParser on header fields that point past the data", () => {
  it("rejects a central-directory name longer than the archive", async () => {
    const { bytes, centralDirectory } = await oneEntryZip();
    // File name length, at offset 28 of a central-directory header.
    new DataView(bytes.buffer).setUint16(centralDirectory + 28, 0xfff0, true);
    expect(() => new ZipParser(bytes).getEntries()).toThrow(ZipParseError);
  });

  it("rejects a local-header offset outside the archive", async () => {
    const { bytes, centralDirectory } = await oneEntryZip();
    // Relative offset of the local header, at offset 42.
    new DataView(bytes.buffer).setUint32(centralDirectory + 42, 0x7fff_0000, true);
    const parser = new ZipParser(bytes);
    await expect(parser.extract("a.txt")).rejects.toBeInstanceOf(ZipParseError);
    expect(() => parser.extractAllSync()).toThrow(ZipParseError);
  });
});
