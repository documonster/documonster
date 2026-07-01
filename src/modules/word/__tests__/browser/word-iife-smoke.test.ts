import { loadIife } from "@test/browser/load-iife";
import { describe, it, expect, beforeAll } from "vitest";

/**
 * Smoke test for the shipped `documonster.word.iife.min.js` bundle.
 *
 * Loads the real IIFE artifact in a browser, asserts the `Documonster.Word`
 * namespace surface, and runs a minimal build → serialize round-trip so a
 * bundling/runtime regression (e.g. a Node-only API leaking into the bundle)
 * fails here rather than silently shipping.
 */
describe("Documonster.Word IIFE bundle", () => {
  let Word: any;

  beforeAll(async () => {
    Word = await loadIife("word", "Word");
  }, 60000);

  it("exposes the expected namespace members", () => {
    for (const member of ["Document", "Io", "Build", "Styles", "Units"]) {
      expect(Word[member], `Documonster.Word.${member}`).toBeTruthy();
    }
  });

  it("builds a document and serializes it to a non-empty docx buffer", async () => {
    const doc = Word.Document.create();
    Word.Document.addHeading(doc, "Browser Heading", 1);
    Word.Document.addParagraph(doc, "Hello from the Word IIFE bundle.");

    const buffer: Uint8Array = await Word.Io.toBuffer(doc);
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.byteLength).toBeGreaterThan(0);
    // A .docx is a ZIP — verify the local-file-header signature "PK\x03\x04".
    expect([buffer[0], buffer[1], buffer[2], buffer[3]]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it("round-trips a document through read()", async () => {
    const doc = Word.Document.create();
    Word.Document.addParagraph(doc, "Round trip");
    const buffer: Uint8Array = await Word.Io.toBuffer(doc);

    const reopened = await Word.Io.read(buffer);
    expect(reopened).toBeTruthy();
  });
});
