/**
 * The PDF module's TrueType error contract survives the parser moving to Layer 0.
 *
 * `parseTtf` now lives in `@utils/font-ttf` and throws `FontParseError`, which
 * cannot extend `PdfFontError` — that is a `PdfError`, and moving it down would
 * drag the whole PDF error hierarchy to Layer 0. `@pdf/font/ttf-parser` translates
 * at the boundary instead, so everything published from `documonster/pdf` behaves
 * as it did.
 *
 * Worth pinning rather than assuming: `PdfFontError` is exported, `isPdfError()`
 * answers for it, and `pdf-exporter.ts` branches on it to tell a caller's bad font
 * from an internal failure. A silent change to any of those is a breaking change
 * to a published surface.
 */

import { PdfFontError, isPdfError } from "@pdf/errors";
import { countTtfFaces, parseTtf } from "@pdf/font/ttf-parser";
import { FontParseError } from "@utils/errors";
import { parseTtf as parseAtLayer0 } from "@utils/font-ttf";
import { describe, expect, it } from "vitest";

/** A `ttcf` header claiming fonts it does not contain. */
const TRUNCATED_TTC = new Uint8Array([0x74, 0x74, 0x63, 0x66]);

describe("PDF TrueType parsing boundary", () => {
  it("reports a bad font as PdfFontError, not the Layer 0 error", () => {
    expect(() => parseTtf(TRUNCATED_TTC)).toThrow(PdfFontError);
    expect(() => countTtfFaces(TRUNCATED_TTC)).toThrow(PdfFontError);
  });

  it("keeps the translated error inside the PdfError hierarchy", () => {
    let thrown: unknown;
    try {
      parseTtf(TRUNCATED_TTC);
    } catch (err) {
      thrown = err;
    }
    expect(isPdfError(thrown)).toBe(true);
    expect((thrown as Error).name).toBe("PdfFontError");
  });

  it("keeps the original failure as the cause rather than discarding it", () => {
    let thrown: unknown;
    try {
      parseTtf(TRUNCATED_TTC);
    } catch (err) {
      thrown = err;
    }
    // The diagnosis names the table or header at fault; losing it would make a
    // font bug unreadable from the PDF side.
    expect((thrown as Error).cause).toBeInstanceOf(FontParseError);
    expect((thrown as Error).message).toBe(((thrown as Error).cause as Error).message);
  });

  it("leaves the Layer 0 parser throwing its own error type", () => {
    // The translation belongs to the boundary, not to the parser: `draw` consumes
    // the same function and must not receive a PDF error.
    expect(() => parseAtLayer0(TRUNCATED_TTC)).toThrow(FontParseError);
    let thrown: unknown;
    try {
      parseAtLayer0(TRUNCATED_TTC);
    } catch (err) {
      thrown = err;
    }
    expect(isPdfError(thrown)).toBe(false);
  });
});
