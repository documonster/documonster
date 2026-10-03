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
 * Error thrown when a resource limit is hit while reading: object nesting
 * depth, decoded stream size (`maxDecodedBytes`) or filter-chain length.
 *
 * Unlike other structural problems, which the reader tolerates and reports
 * as page warnings, a limit hit is always a hard failure: continuing would
 * hand back a silently truncated result.
 */
export class PdfLimitExceededError extends PdfStructureError {
  override name = "PdfLimitExceededError";
}

/**
 * Error thrown when an edit would invalidate a digital signature already
 * present in the document, and the caller did not opt in with
 * `invalidateSignatures: true`.
 */
export class PdfSignatureInvalidationError extends PdfError {
  override name = "PdfSignatureInvalidationError";
}

/**
 * Check if an error is a PdfError.
 */
export function isPdfError(err: unknown): err is PdfError {
  return err instanceof PdfError;
}
