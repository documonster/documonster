/**
 * PDF module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

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
 * Error thrown when an encrypted document cannot be opened with the password given — none,
 * or the wrong one. A {@link PdfStructureError}, the class it has always been, so existing
 * `catch` blocks are unaffected; its own class is what lets a caller prompt for a password
 * instead of telling a wrong password apart from a corrupt file by its message.
 */
export class PdfPasswordError extends PdfStructureError {
  override name = "PdfPasswordError";
}

/**
 * Error thrown when a file uses something the PDF reader recognises but does not implement — an
 * encryption filter, revision or crypt-filter method it has no handler for. A
 * {@link PdfStructureError}, the class these were reported as before, so existing `catch` blocks
 * are unaffected; its own class tells "not supported" apart from "damaged".
 */
export class PdfUnsupportedFeatureError extends PdfStructureError {
  override name = "PdfUnsupportedFeatureError";
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

  /**
   * Which limit was hit: the `maxDecodedBytes` option, the object nesting `depth`, or the
   * length of a stream's `filterChain`. `undefined` only on an error built with the deprecated
   * message-only constructor; every error the reader throws carries it.
   */
  readonly limit: PdfLimit | undefined;

  /**
   * The value of that limit — the bytes `maxDecodedBytes` allows, the deepest nesting or the
   * longest filter chain read. `undefined` exactly when {@link limit} is.
   */
  readonly allowed: number | undefined;

  constructor(limit: PdfLimit, allowed: number, message: string, options?: BaseErrorOptions);
  /**
   * @deprecated Pass the limit and its value first:
   * `new PdfLimitExceededError(limit, allowed, message, options)`. This form leaves `limit`
   * and `allowed` undefined. Removed in 1.0.
   */
  constructor(message: string, options?: BaseErrorOptions);
  constructor(
    limitOrMessage: string,
    allowedOrOptions?: number | BaseErrorOptions,
    message?: string,
    options?: BaseErrorOptions
  ) {
    const legacy = typeof allowedOrOptions !== "number";
    super(
      legacy ? limitOrMessage : (message ?? ""),
      legacy ? (allowedOrOptions as BaseErrorOptions | undefined) : options
    );
    this.limit = legacy ? undefined : (limitOrMessage as PdfLimit);
    this.allowed = legacy ? undefined : (allowedOrOptions as number);
  }
}

/**
 * A resource limit the PDF reader enforces (see {@link PdfLimitExceededError.limit}):
 * `maxDecodedBytes` is the `Pdf.extract` option of that name, `depth` the object nesting
 * limit and `filterChain` the `/Filter` chain limit.
 */
export type PdfLimit = "maxDecodedBytes" | "depth" | "filterChain";

/**
 * Error thrown when an edit would invalidate a digital signature already
 * present in the document, and the caller did not opt in with
 * `invalidateSignatures: true`.
 */
export class PdfSignatureInvalidationError extends PdfError {
  override name = "PdfSignatureInvalidationError";
}

/**
 * Error thrown when a signature is to be filled into a file that holds no signature
 * placeholder — no `/ByteRange` with a `/Contents` hex string beside it. `Pdf.Builder#sign`
 * and `Pdf.Editor#sign` write the placeholder and fill it in one call;
 * `Pdf.Signature.buildDictPlaceholder` builds one by hand.
 *
 * A {@link PdfStructureError}, the class this has always been reported as.
 */
export class PdfSignaturePlaceholderMissingError extends PdfStructureError {
  override name = "PdfSignaturePlaceholderMissingError";
}

/**
 * Error thrown when `toBuffer({ incremental: "required" })` cannot append an update —
 * structural edits, recovered cross-reference data, or a page with no object number —
 * and the caller asked for an error rather than a full rewrite.
 */
export class PdfIncrementalUpdateError extends PdfError {
  override name = "PdfIncrementalUpdateError";
}

/**
 * Error thrown when an encrypted document cannot be written encrypted — its `/Encrypt`
 * dictionary is unreadable, so a full rewrite has no security handler to encrypt with.
 * Nothing is written; pass `{ encryption: "remove" }` to write the document decrypted.
 */
export class PdfEncryptionError extends PdfError {
  override name = "PdfEncryptionError";
}

/**
 * Error thrown when form field values are set on a dynamic XFA form — one whose catalog
 * says `/NeedsRendering true`, or whose XFA packet has no AcroForm field tree beside it.
 * Such a form is defined entirely by its XFA; the editor fills AcroForm fields, which an
 * XFA viewer would ignore, so it refuses rather than write values nobody will see.
 * Nothing is written.
 */
export class PdfXfaFormError extends PdfError {
  override name = "PdfXfaFormError";
}

/**
 * Error thrown when a builder, editor or converter is given a value that has no valid form
 * in the file — a font size that is `NaN`, a colour that is not `{ r, g, b }` in 0–1, a
 * page width of zero, an option outside its enumeration. Thrown by the call that received
 * the value, before it changes anything, with a message that starts with the call as the
 * caller wrote it (`Pdf.fromExcel: …`, `Pdf.PageBuilder#drawText: …`).
 */
export class PdfArgumentError extends PdfError {
  override name = "PdfArgumentError";
  /**
   * The argument or option at fault, as a path from the parameter name in the public
   * signature: `data`, `options.fontSize`, `options.watermark.pages[0]`, `metadata.title`.
   */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/**
 * Check if an error is a PdfError.
 */
export function isPdfError(err: unknown): err is PdfError {
  return err instanceof PdfError;
}
