import { loadIife } from "@test/browser/load-iife";
import { describe, it, expect, beforeAll } from "vitest";

/**
 * Smoke test for the shipped `documonster.pdf.iife.min.js` bundle.
 *
 * Loads the real IIFE artifact, asserts the `Documonster.Pdf` surface, and
 * builds → reads back a minimal PDF in a browser so a bundling/runtime
 * regression in the shipped PDF bundle surfaces here.
 */
describe("Documonster.Pdf IIFE bundle", () => {
  let Pdf: any;

  beforeAll(async () => {
    // `Documonster.Pdf` is the module namespace; its `Pdf` member is the
    // surface that carries Builder / create / read (mirrors Csv.Csv, Xml.Xml).
    ({ Pdf } = await loadIife<{ Pdf: any }>("pdf", "Pdf"));
  }, 60000);

  it("exposes the expected namespace members", () => {
    for (const member of ["Builder", "create", "read", "PageSizes"]) {
      expect(Pdf[member], `Documonster.Pdf.${member}`).toBeTruthy();
    }
  });

  it("builds a one-page PDF with the correct header", async () => {
    const doc = new Pdf.Builder();
    const page = doc.addPage();
    page.drawText?.("Hello from the PDF IIFE bundle", { x: 72, y: 700 });

    const bytes: Uint8Array = await doc.build();
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.byteLength).toBeGreaterThan(0);
    const header = new TextDecoder().decode(bytes.subarray(0, 5));
    expect(header).toBe("%PDF-");
  });

  it("reads back the PDF it just built", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    const bytes: Uint8Array = await doc.build();

    const result = await Pdf.read(bytes);
    expect(result.pages.length).toBe(1);
  });
});
