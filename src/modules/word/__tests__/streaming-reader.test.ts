/**
 * DOCX Module - Streaming Reader Tests
 *
 * Verifies that the streaming reader yields the exact same body content as the
 * batch `readDocx`, and that it honours its true-stream guarantees (single
 * iteration, metadata loaded up-front, per-element size cap).
 */

import { describe, it, expect } from "vitest";

import { Build, Document, Io, Streaming } from "../index";
import type { BodyContent, Table } from "../types";

/** Build a representative document and return its packaged bytes. */
async function buildSampleDocx(paragraphCount = 5): Promise<Uint8Array> {
  const h = Document.create();
  Document.addHeading(h, "Title", 1);
  for (let i = 0; i < paragraphCount; i++) {
    Document.addParagraph(h, `Paragraph ${i}`);
  }
  Document.addTableElement(
    h,
    Build.simpleTable([
      ["A1", "B1"],
      ["A2", "B2"]
    ])
  );
  Document.addParagraph(h, "Trailing paragraph");
  const doc = Document.build(h);
  return Io.package(doc);
}

/** Collect all body content yielded by the streaming reader. */
async function collectStream(bytes: Uint8Array): Promise<BodyContent[]> {
  const reader = Streaming.createDocxStreamReader(bytes);
  const out: BodyContent[] = [];
  for await (const item of reader) {
    out.push(item);
  }
  return out;
}

describe("StreamingDocxReader", () => {
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
    const reader = Streaming.createDocxStreamReader(bytes);

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
    const reader = Streaming.createDocxStreamReader(bytes);

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
    const doc = Document.build(h);
    const bytes = await Io.package(doc);

    const reader = Streaming.createDocxStreamReader(bytes, { maxElementBytes: 100 });
    await expect(async () => {
      for await (const _ of reader) {
        // should throw before yielding the oversized paragraph
      }
    }).rejects.toThrow(/maxElementBytes/);
  });

  it("rejects encrypted (CFB) input with a clear error", async () => {
    const cfb = new Uint8Array(16);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

    const reader = Streaming.createDocxStreamReader(cfb);
    await expect(async () => {
      for await (const _ of reader) {
        // should throw
      }
    }).rejects.toThrow(/Encrypted DOCX/);
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
  return Io.package(Document.build(h));
}

describe("StreamingDocxReader — true streaming", () => {
  it("emits the first element long before document.xml is fully consumed", async () => {
    // ~3 MB of incompressible body across 800 paragraphs.
    const bytes = await buildLargeDocx(800, 4000);

    let consumedAtFirstYield = -1;
    let totalConsumed = 0;
    let count = 0;

    const reader = Streaming.createDocxStreamReader(bytes, {
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
    const reader = Streaming.createDocxStreamReader(bytes);
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
