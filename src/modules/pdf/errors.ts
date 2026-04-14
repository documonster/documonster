/**
 * PDF module error types.
 */

import { BaseError } from "@utils/errors";

/**
 * Base class for all PDF-related errors.
 */
export class PdfError extends BaseError {
  override name = "PdfError";
}

/**
 * Error thrown when PDF rendering fails (layout, drawing, content generation).
 */
export class PdfRenderError extends PdfError {
  override name = "PdfRenderError";
}

/**
 * Error thrown when font operations fail (missing glyph, unsupported font).
 */
export class PdfFontError extends PdfError {
  override name = "PdfFontError";
}

/**
 * Error thrown when the PDF file structure is invalid.
 */
export class PdfStructureError extends PdfError {
  override name = "PdfStructureError";
}

/**
 * Check if an error is a PdfError.
 */
export function isPdfError(err: unknown): err is PdfError {
  return err instanceof PdfError;
}
