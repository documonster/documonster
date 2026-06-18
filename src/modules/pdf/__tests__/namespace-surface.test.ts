import * as PdfMod from "@pdf/index";
/**
 * Public namespace-surface contract for `@cj-tech-master/excelts/pdf`.
 *
 * Verifies the `Pdf` namespace shape (writing/reading/building/editing/
 * conversion/signatures) and that a representative member produces output.
 */
import { describe, it, expect } from "vitest";

const { Pdf } = PdfMod;

describe("@cj-tech-master/excelts/pdf namespace surface", () => {
  it("exposes the expected members", () => {
    for (const m of [
      "create",
      "read",
      "fromExcel",
      "fromDocx",
      "fromChart",
      "verifySignature",
      "sign",
      "parseSvgPath"
    ]) {
      expect(typeof (Pdf as Record<string, unknown>)[m], `Pdf.${m}`).toBe("function");
    }
    expect(typeof Pdf.Builder).toBe("function"); // class
    expect(typeof Pdf.Editor).toBe("function"); // class
    expect(Pdf.PageSizes).toBeDefined();
  });

  it("Pdf.create produces a non-trivial PDF byte stream", async () => {
    const bytes = await Pdf.create([
      ["Product", "Revenue"],
      ["Widget", 1000]
    ]);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(100);
    // PDF magic header "%PDF"
    expect(bytes[0]).toBe(0x25);
    expect(bytes[1]).toBe(0x50);
  });

  it("error classes stay flat", () => {
    expect(typeof PdfMod.PdfError).toBe("function");
    expect(typeof PdfMod.isPdfError).toBe("function");
  });
});
