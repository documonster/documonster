/**
 * `Pdf.extract`'s limit option and inputs, held to the contracts the other modules' readers
 * follow: a bad limit is a `PdfArgumentError` naming `options.maxDecodedBytes`; a limit hit says
 * which limit fired and its value (`limit`, `allowed`, as `ArchiveLimitExceededError` does); and the async
 * reader takes a `Blob` — a browser `File` — as the archive and CSV readers do.
 */
import { PdfArgumentError, PdfLimitExceededError, PdfStructureError } from "@pdf/errors";
import { parseObject } from "@pdf/reader/pdf-parser";
import { PdfTokenizer } from "@pdf/reader/pdf-tokenizer";
import { decodeStreamFilters } from "@pdf/reader/stream-filters";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";

const enc = new TextEncoder();

async function onePage(text: string): Promise<Uint8Array> {
  const builder = new Pdf.Builder();
  builder.addPage().drawText(text, { x: 72, y: 750, fontSize: 12 });
  return builder.toBuffer();
}

function caught(fn: () => unknown): unknown {
  try {
    fn();
  } catch (err) {
    return err;
  }
  throw new Error("expected a throw");
}

describe("Pdf.extract({ maxDecodedBytes }) is checked as an argument", () => {
  it.each([Number.NaN, -1, "16", null])("refuses %s with PdfArgumentError", async bad => {
    const bytes = await onePage("x");
    const err = await Pdf.extract(bytes, { maxDecodedBytes: bad as never }).catch(e => e);
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect(err.field).toBe("options.maxDecodedBytes");
    expect(err.message).toMatch(/^Pdf\.extract: options\.maxDecodedBytes must be/);
  });

  it("accepts 0, a finite size and Infinity", async () => {
    const bytes = await onePage("x");
    await expect(Pdf.extract(bytes, { maxDecodedBytes: Infinity })).resolves.toBeDefined();
    await expect(Pdf.extract(bytes, { maxDecodedBytes: 1 << 20 })).resolves.toBeDefined();
    const zero = await Pdf.extract(bytes, { maxDecodedBytes: 0 }).catch(e => e);
    expect(zero).not.toBeInstanceOf(PdfArgumentError);
  });
});

describe("PdfLimitExceededError says which limit fired", () => {
  it("maxDecodedBytes: the option name and the value given", async () => {
    const bytes = await onePage("x".repeat(2000));
    const err = await Pdf.extract(bytes, { maxDecodedBytes: 16 }).catch(e => e);
    expect(err).toBeInstanceOf(PdfLimitExceededError);
    expect(err.limit).toBe("maxDecodedBytes");
    expect(err.allowed).toBe(16);
  });

  it("filterChain: the longest chain read", () => {
    const chain = Array.from({ length: 17 }, () => "ASCIIHexDecode");
    const dict = new Map<string, unknown>([["Filter", chain]]);
    const err = caught(() => decodeStreamFilters(enc.encode("00>"), dict as never));
    expect(err).toBeInstanceOf(PdfLimitExceededError);
    expect(err).toMatchObject({ limit: "filterChain", allowed: 16 });
  });

  it("depth: the deepest nesting read", () => {
    const deep = "[".repeat(1000) + "]".repeat(1000);
    const err = caught(() => parseObject(new PdfTokenizer(enc.encode(deep))));
    expect(err).toBeInstanceOf(PdfLimitExceededError);
    expect(err).toMatchObject({ limit: "depth", allowed: 512 });
  });

  it("is still a PdfStructureError, with the message it had", () => {
    const err = new PdfLimitExceededError("maxDecodedBytes", 8, "too big");
    expect(err).toBeInstanceOf(PdfStructureError);
    expect(err.message).toBe("too big");
    expect(err).toMatchObject({ limit: "maxDecodedBytes", allowed: 8 });
  });

  it("the deprecated message-only constructor keeps working", () => {
    const err = new PdfLimitExceededError("too big");
    expect(err.message).toBe("too big");
    expect(err.limit).toBeUndefined();
    expect(err.allowed).toBeUndefined();
  });
});

describe("Pdf.extract accepts a Blob", () => {
  it("reads a Blob as its bytes", async () => {
    const bytes = await onePage("FROM BLOB");
    const result = await Pdf.extract(new Blob([bytes as BlobPart]));
    expect(result.pages[0].text).toContain("FROM BLOB");
  });

  it("Pdf.verifySignatures reads a Blob too", async () => {
    const bytes = await onePage("x");
    await expect(Pdf.verifySignatures(new Blob([bytes as BlobPart]))).resolves.toEqual([]);
    const err = await Pdf.verifySignatures(42 as never).catch(e => e);
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect(err.message).toMatch(/^Pdf\.verifySignatures: data must be .*a Blob/);
  });

  it("still refuses what is neither bytes nor a Blob, naming the Blob", async () => {
    const err = await Pdf.extract("not a pdf" as never).catch(e => e);
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect(err.field).toBe("data");
    expect(err.message).toMatch(/Blob/);
  });
});
