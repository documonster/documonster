/**
 * DOCX module error types.
 *
 * All DOCX-related errors extend DocxError.
 */

import { BaseError, type BaseErrorOptions } from "@utils/errors";

// Re-export common utilities from base
export {
  AbortError,
  createAbortError,
  isAbortError,
  throwIfAborted,
  createLinkedAbortController,
  toError,
  suppressUnhandledRejection,
  errorToJSON,
  getErrorChain,
  getRootCause,
  type BaseErrorOptions
} from "@utils/errors";

/**
 * Base class for all DOCX-related errors.
 */
export class DocxError extends BaseError {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "DocxError";
  }
}

/**
 * Check if an error is a DOCX error.
 */
export function isDocxError(err: unknown): err is DocxError {
  return err instanceof DocxError;
}

/**
 * Error thrown when DOCX parsing fails.
 */
export class DocxParseError extends DocxError {
  override name = "DocxParseError";
}

/**
 * Error thrown when DOCX generation/writing fails.
 */
export class DocxWriteError extends DocxError {
  override name = "DocxWriteError";
}

/**
 * Error thrown when a required DOCX part is missing.
 */
export class DocxMissingPartError extends DocxParseError {
  override name = "DocxMissingPartError";

  constructor(partPath: string) {
    super(`Required DOCX part not found: ${partPath}`);
  }
}

/**
 * Error thrown when document structure is invalid.
 */
export class DocxInvalidStructureError extends DocxParseError {
  override name = "DocxInvalidStructureError";
}

/**
 * Error thrown when an unsupported feature is encountered during parsing.
 */
export class DocxUnsupportedFeatureError extends DocxError {
  override name = "DocxUnsupportedFeatureError";

  constructor(feature: string) {
    super(`Unsupported DOCX feature: ${feature}`);
  }
}
