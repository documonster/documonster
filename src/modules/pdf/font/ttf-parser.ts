/**
 * TrueType parsing for the PDF module — the layer boundary, not the parser.
 *
 * The parser itself is `@utils/font-ttf`, at Layer 0, because the glyph
 * rasteriser in `draw` needs the same tables and `draw` cannot import `pdf`. See
 * that module's header for why it moved.
 *
 * ## Why this file still exists
 *
 * `PdfFontError` is published from `documonster/pdf`, it is what
 * `isPdfError()` answers for, and `pdf-exporter.ts` branches on it to decide
 * whether a font failure is the caller's fault. A parser at Layer 0 cannot throw
 * it — `PdfFontError extends PdfError`, and moving it down would drag the whole
 * PDF error hierarchy with it. So the translation happens here, at the boundary:
 * everything inside `pdf` keeps importing `parseTtf` from this path and keeps
 * getting `PdfFontError`, exactly as before the parser moved.
 *
 * The alternative was to let `FontParseError` escape and change the seven tests
 * and two `instanceof` branches that name `PdfFontError`. That is a breaking
 * change to a published error type in return for deleting twenty lines, which is
 * the wrong trade.
 *
 * @module
 */

import { PdfFontError } from "@pdf/errors";
import { FontParseError } from "@utils/errors";
import { countTtfFaces as countFaces, parseTtf as parse } from "@utils/font-ttf";

export type { TableEntry, TtfFont } from "@utils/font-ttf";

/**
 * Re-throw a Layer 0 parse failure as this module's published error type.
 *
 * Anything that is not a `FontParseError` is passed through untouched: a bug in
 * the parser should surface as itself rather than be relabelled as bad input.
 */
function asPdfFontError(err: unknown): never {
  if (err instanceof FontParseError) {
    throw new PdfFontError(err.message, { cause: err });
  }
  throw err;
}

/**
 * Parse a TrueType font file.
 *
 * @param data - Raw .ttf, .otf or .ttc file bytes
 * @param collectionIndex - Face index within a `.ttc`; must be 0 otherwise
 * @returns Parsed font data
 * @throws {PdfFontError} If the font is invalid or unsupported
 */
export function parseTtf(data: Uint8Array, collectionIndex = 0): ReturnType<typeof parse> {
  try {
    return parse(data, collectionIndex);
  } catch (err) {
    asPdfFontError(err);
  }
}

/**
 * Count the faces inside a font file: the number of fonts in a TrueType
 * Collection, or `1` for a single-face `.ttf`.
 *
 * @param data - Raw font file bytes
 * @returns Face count (at least 1)
 * @throws {PdfFontError} If the collection header is invalid or truncated
 */
export function countTtfFaces(data: Uint8Array): number {
  try {
    return countFaces(data);
  } catch (err) {
    asPdfFontError(err);
  }
}
