/**
 * Ids are allocated in O(1) amortised. Bookmark and drawing ids: the tree is
 * walked once per document to seed them, never per call — walking it per call
 * made a document of n bookmarks O(n²) to build. Footnote, endnote, comment
 * and image ids: a collection is read from where the last call stopped, never
 * again from the start.
 *
 * A wall-clock bound, so it runs under Node only (`.node.test.ts`): in the
 * shared browser run it measured the CI machine's load as much as the code.
 * The bound sits an order of magnitude above the measurement and an order of
 * magnitude below the per-call tree walk, so it still fails on the regression
 * it exists for.
 */
import { unzip } from "@archive/read-archive";
import { describe, it, expect } from "vitest";

import { Build, Document, Io } from "../../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

describe("Document id allocation cost", () => {
  it("builds 10k bookmarks and 10k drawings on a 50k-paragraph document quickly", async () => {
    const doc = Document.create();
    for (let i = 0; i < 50_000; i++) {
      Document.addParagraph(doc, `Paragraph ${i}`);
    }
    const started = performance.now();
    for (let i = 0; i < 10_000; i++) {
      const id = Document.allocateBookmarkId(doc);
      Document.addParagraph(
        doc,
        Build.paragraph([Build.bookmarkStart(id, `b${i}`), Build.text("x"), Build.bookmarkEnd(id)])
      );
      Document.addImage(doc, PNG, "png", 9525, 9525);
    }
    const elapsed = performance.now() - started;
    // Walking the tree per call took ~10 s at this size.
    expect(elapsed).toBeLessThan(1000);

    let xml = "";
    for await (const entry of unzip(await Io.toBuffer(doc)).entries()) {
      if (entry.path === "word/document.xml") {
        xml = new TextDecoder().decode(await entry.bytes());
      }
    }
    const starts = [...xml.matchAll(/<w:bookmarkStart [^>]*\bw:id="(\d+)"/g)].map(m => m[1]);
    const docPr = [...xml.matchAll(/<wp:docPr [^>]*\bid="(\d+)"/g)].map(m => m[1]);
    expect(starts).toHaveLength(10_000);
    expect(new Set(starts).size).toBe(starts.length);
    expect(docPr).toHaveLength(10_000);
    expect(new Set(docPr).size).toBe(docPr.length);
  });

  it("adds 100k footnotes, endnotes and comments in linear time", async () => {
    const doc = Document.create();
    const started = performance.now();
    for (let i = 0; i < 100_000; i++) {
      Document.addFootnote(doc, "f");
      Document.addEndnote(doc, "e");
      Document.addComment(doc, "me", "c");
    }
    const elapsed = performance.now() - started;
    // Measured ~100 ms; comparing every entry on every call took ~18 s per kind.
    expect(elapsed).toBeLessThan(3000);
    expect(new Set(doc.footnotes!.map(f => f.id)).size).toBe(100_000);

    // Entries appended directly to a large collection are seen at once.
    (doc.footnotes as { id: number; content: [] }[]).push({ id: 400_000, content: [] });
    expect(Document.addFootnote(doc, "g")).toBe(400_001);

    // So is an entry replaced, at any position, through a new array.
    const notes = doc.footnotes!;
    (doc as { footnotes?: unknown }).footnotes = [{ id: 500_000, content: [] }, ...notes.slice(1)];
    expect(Document.addFootnote(doc, "h")).toBe(500_001);

    // An entry replaced in place deep inside is not seen by the helper; the
    // write refuses the duplicate it produced rather than packaging it.
    const inPlace = doc.footnotes as { id: number; content: [] }[];
    inPlace[3] = { id: 500_002, content: [] };
    expect(Document.addFootnote(doc, "i")).toBe(500_002);
    await expect(Io.toBuffer(doc)).rejects.toMatchObject({
      name: "DocxDuplicateIdError",
      collection: "footnotes",
      id: 500_002
    });
  });

  it("adds 100k images in linear time", () => {
    const doc = Document.create();
    const started = performance.now();
    for (let i = 0; i < 100_000; i++) {
      Document.addImage(doc, PNG, "png", 9525, 9525);
    }
    const elapsed = performance.now() - started;
    // Measured ~60 ms; comparing every entry on every call is quadratic.
    expect(elapsed).toBeLessThan(3000);
    expect(new Set(doc.images!.map(img => img.rId)).size).toBe(100_000);
  });
});
