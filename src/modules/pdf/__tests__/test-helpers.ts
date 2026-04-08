/**
 * Shared test helpers for PDF tests.
 */
import { expect } from "vitest";

/**
 * Decode a PDF Uint8Array to string for assertion.
 */
export function pdfToString(pdf: Uint8Array): string {
  return new TextDecoder().decode(pdf);
}

/**
 * Verify basic PDF structure (header, xref, trailer, EOF).
 */
export function expectValidPdf(pdf: Uint8Array): void {
  const text = pdfToString(pdf);
  expect(text).toContain("%PDF-2.0");
  expect(text).toContain("xref");
  expect(text).toContain("trailer");
  expect(text).toContain("%%EOF");
  expect(text).toContain("/Catalog");
  expect(text).toContain("/Pages");
}
