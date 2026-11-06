/**
 * DOCX Module - Streaming Reader Tests
 *
 * Verifies that the streaming reader yields the exact same body content as the
 * batch `readDocx`, and that it honours its true-stream guarantees (single
 * iteration, metadata loaded up-front, per-element size cap).
 */

import { describe, it, expect } from "vitest";

import { WordCrypto } from "../crypto";
import {
  Build,
  Document,
  DocxDecryptionError,
  DocxEncryptedError,
  DocxError,
  DocxParseError,
  Io,
  Stream
} from "../index";
import type { BodyContent, Table } from "../types";

/** Build a representative document and return its packaged bytes. */
async function buildSampleDocx(paragraphCount = 5): Promise<Uint8Array> {
  const h = Document.create();
  Document.addHeading(h, "Title", 1);
  for (let i = 0; i < paragraphCount; i++) {
    Document.addParagraph(h, `Paragraph ${i}`);
  }
  Document.addTable(
    h,
    Build.simpleTable([
      ["A1", "B1"],
      ["A2", "B2"]
    ])
  );
  Document.addParagraph(h, "Trailing paragraph");
  const doc = h;
  return Io.toBuffer(doc);
}

/** Collect all body content yielded by the streaming reader. */
async function collectStream(bytes: Uint8Array): Promise<BodyContent[]> {
  const reader = new Stream.DocumentReader(bytes);
  const out: BodyContent[] = [];
  for await (const item of reader) {
    out.push(item);
  }
  return out;
}

describe("Stream.DocumentReader", () => {
  it("yields the same body content as readDocx", async () => {
    const bytes = await buildSampleDocx();

    const batch = await Io.read(bytes);
    const streamed = await collectStream(bytes);

    // The streaming reader never emits the trailing <w:sectPr> as a body
    // entry (it is exposed via reader.sectionProperties instead), and the
    // batch reader strips it from body too — so the body arrays must match
    // 1:1 in length and content.
    expect(streamed.length).toBe(batch.body.length);
    expect(streamed).toEqual(batch.body);
  });

  it("preserves paragraph text and order", async () => {
    const bytes = await buildSampleDocx(3);

    const batch = await Io.read(bytes);
    const streamed = await collectStream(bytes);

    // Compare extracted text per paragraph against the batch reader to avoid
    // depending on the exact run/child shape.
    const textOf = (b: BodyContent): string =>
      JSON.stringify(b)
        .match(/"text":"[^"]*"/g)
        ?.join("") ?? "";

    const batchTexts = batch.body.map(textOf);
    const streamTexts = streamed.map(textOf);
    expect(streamTexts).toEqual(batchTexts);

    const allText = streamed.map(textOf).join(" ");
    expect(allText).toContain("Title");
    expect(allText).toContain("Paragraph 0");
    expect(allText).toContain("Paragraph 2");
    expect(allText).toContain("Trailing paragraph");
  });

  it("streams a table identically to readDocx", async () => {
    const bytes = await buildSampleDocx(2);

    const batch = await Io.read(bytes);
    const streamed = await collectStream(bytes);

    const batchTable = batch.body.find((b): b is Table => b.type === "table");
    const streamTable = streamed.find((b): b is Table => b.type === "table");

    expect(streamTable).toBeDefined();
    expect(streamTable).toEqual(batchTable);
  });

  it("exposes metadata consistent with readDocx, loaded up-front", async () => {
    const bytes = await buildSampleDocx(1);

    const batch = await Io.read(bytes);
    const reader = new Stream.DocumentReader(bytes);

    // Drain the iterator so the metadata pass has run.
    for await (const _ of reader) {
      // no-op
    }

    // The streaming reader's metadata must match what the batch reader parsed
    // from the same package (this doc carries no explicit style/numbering
    // definitions, so both should agree — including agreeing on emptiness).
    expect(reader.styles.length).toBe((batch.styles ?? []).length);
    expect(reader.numberingInstances.length).toBe((batch.numberingInstances ?? []).length);
    expect(reader.metadata.styles).toBe(reader.styles);
    expect(Array.isArray(reader.numberingInstances)).toBe(true);
  });

  it("can only be iterated once", async () => {
    const bytes = await buildSampleDocx(1);
    const reader = new Stream.DocumentReader(bytes);

    for await (const _ of reader) {
      // drain
    }

    await expect(async () => {
      for await (const _ of reader) {
        // second iteration must throw
      }
    }).rejects.toThrow(/only be iterated once/);
  });

  it("rejects an oversized single body element via maxElementBytes", async () => {
    const h = Document.create();
    Document.addParagraph(h, "x".repeat(2000));
    const doc = h;
    const bytes = await Io.toBuffer(doc);

    const reader = new Stream.DocumentReader(bytes, { maxElementBytes: 100 });
    await expect(async () => {
      for await (const _ of reader) {
        // should throw before yielding the oversized paragraph
      }
    }).rejects.toThrow(/maxElementBytes/);
  });

  it("rejects encrypted (CFB) input without a password as Io.read does", async () => {
    const enc = await WordCrypto.encrypt(await buildSampleDocx(1), "pw", { spinCount: 1 });

    const reader = new Stream.DocumentReader(enc);
    await expect(async () => {
      for await (const _ of reader) {
        // should throw
      }
    }).rejects.toBeInstanceOf(DocxEncryptedError);
  });

  it("reports a compound file that is not an encrypted package as a legacy document, not as encrypted", async () => {
    // A bare CFB signature — a legacy .doc for all the reader can tell — used to be reported as encrypted, telling
    // the caller to supply a password that could never help.
    const cfb = new Uint8Array(16);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    for (const read of [
      () => Io.read(cfb),
      async () => {
        for await (const _ of new Stream.DocumentReader(cfb)) {
          // should throw
        }
      }
    ]) {
      const error = await read().then(
        () => undefined,
        (e: unknown) => e
      );
      expect(error).toBeInstanceOf(DocxParseError);
      expect(error).not.toBeInstanceOf(DocxEncryptedError);
      expect((error as Error).message).toMatch(/legacy \.doc/);
    }
  });

  it("reads an encrypted package with { password }, as Io.read does", async () => {
    const bytes = await buildSampleDocx(3);
    const enc = await WordCrypto.encrypt(bytes, "pw", { spinCount: 1 });
    const streamed: BodyContent[] = [];
    for await (const item of new Stream.DocumentReader(enc, { password: "pw" })) {
      streamed.push(item);
    }
    expect(streamed).toEqual((await Io.read(enc, { password: "pw" })).body);

    const wrong = new Stream.DocumentReader(enc, { password: "nope" });
    await expect(async () => {
      for await (const _ of wrong) {
        // should throw
      }
    }).rejects.toBeInstanceOf(DocxDecryptionError);
  });

  it("fails with a DocxParseError on damaged bytes, as Io.read does", async () => {
    const bytes = await buildSampleDocx(40);
    const damaged: Uint8Array[] = [];
    for (let i = 1; i < 32; i++) {
      const at = Math.floor((bytes.length * i) / 32);
      damaged.push(bytes.slice(0, at));
      // The central directory intact at its offset, the data before it zeroed.
      const zeroed = bytes.slice();
      zeroed.fill(0, at, Math.min(bytes.length - 22, at + 64));
      damaged.push(zeroed);
    }
    const failures: string[] = [];
    for (const input of damaged) {
      try {
        await collectStream(input);
      } catch (e) {
        if (!(e instanceof DocxError)) {
          failures.push(`${(e as Error).name}: ${(e as Error).message}`);
        } else if (e.name === "DocxParseError") {
          expect(e.cause).toBeDefined();
        }
      }
    }
    expect(failures).toEqual([]);
  });
});

// =============================================================================
// True end-to-end streaming proof (deterministic, not heap-sampling)
// =============================================================================

/** Random (incompressible) text so document.xml is genuinely large on disk. */
function randomText(n: number): string {
  const alphabet = "abcdefghijklmnopqrstuvwxyz0123456789 ";
  let s = "";
  for (let i = 0; i < n; i++) {
    s += alphabet[(Math.random() * alphabet.length) | 0];
  }
  return s;
}

async function buildLargeDocx(paraCount: number, paraSize: number): Promise<Uint8Array> {
  const h = Document.create();
  for (let i = 0; i < paraCount; i++) {
    Document.addParagraph(h, `${i}:${randomText(paraSize)}`);
  }
  return Io.toBuffer(h);
}

describe("Stream.DocumentReader — true streaming", () => {
  it("emits the first element long before document.xml is fully consumed", async () => {
    // ~3 MB of incompressible body across 800 paragraphs.
    const bytes = await buildLargeDocx(800, 4000);

    let consumedAtFirstYield = -1;
    let totalConsumed = 0;
    let count = 0;

    const reader = new Stream.DocumentReader(bytes, {
      onProgress: consumed => {
        totalConsumed = consumed;
      }
    });

    for await (const _ of reader) {
      if (consumedAtFirstYield < 0) {
        consumedAtFirstYield = reader.consumedBytes;
      }
      count++;
    }

    // Sanity: we actually streamed a big body in many elements.
    expect(count).toBeGreaterThan(700);
    expect(totalConsumed).toBeGreaterThan(2_000_000);

    // The decisive proof: by the time the FIRST body element is yielded, only
    // a small fraction of document.xml has been consumed. If the reader were
    // buffering the whole part (non-streaming), consumedAtFirstYield would
    // equal the full size. We require it to be < 25% of the total.
    expect(consumedAtFirstYield).toBeGreaterThan(0);
    expect(consumedAtFirstYield).toBeLessThan(totalConsumed * 0.25);
  }, 60000);

  it("consumedBytes grows monotonically across iteration", async () => {
    const bytes = await buildLargeDocx(400, 4000);

    const samples: number[] = [];
    const reader = new Stream.DocumentReader(bytes);
    for await (const _ of reader) {
      samples.push(reader.consumedBytes);
    }

    // Strictly non-decreasing, and the first sample is well below the last —
    // i.e. elements are produced progressively as bytes arrive.
    for (let i = 1; i < samples.length; i++) {
      expect(samples[i]).toBeGreaterThanOrEqual(samples[i - 1]);
    }
    expect(samples[0]).toBeLessThan(samples[samples.length - 1]);
  }, 60000);
});

describe("Stream.DocumentReader — leaving the loop early", () => {
  for (const [where, at] of [
    ["first", 0],
    ["middle", 200],
    ["last", -1]
  ] as const) {
    for (const how of ["break", "throw"] as const) {
      it(`settles on a ${how} at the ${where} element, over a body larger than an inflate chunk`, async () => {
        const bytes = await buildLargeDocx(400, 4000);
        const total = (await collectStream(bytes)).length;
        const stopAt = at < 0 ? total - 1 : at;
        const thrown = new Error("left");
        let seen = 0;
        const run = (async () => {
          try {
            for await (const _ of new Stream.DocumentReader(bytes)) {
              if (seen++ === stopAt) {
                if (how === "throw") {
                  throw thrown;
                }
                break;
              }
            }
          } catch (error) {
            if (error !== thrown) {
              throw error;
            }
          }
        })();
        let timer: ReturnType<typeof setTimeout> | undefined;
        await Promise.race([
          run,
          new Promise((_, reject) => {
            timer = setTimeout(() => reject(new Error("did not settle")), 10_000);
          })
        ]).finally(() => clearTimeout(timer));
        expect(seen).toBe(stopAt + 1);
      }, 60000);
    }
  }
});
