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

/**
 * Error thrown when a DOCX file is encrypted (CFB format) and no password
 * was provided. Users should use `decryptDocx()` from "excelts/word/crypto"
 * to decrypt before calling `readDocx()`.
 */
export class DocxEncryptedError extends DocxError {
  override name = "DocxEncryptedError";

  constructor() {
    super(
      "The document is encrypted (password-protected). " +
        'Use decryptDocx(buffer, password) from "excelts/word/crypto" to decrypt it first.'
    );
  }
}

/**
 * Error thrown when an encrypted DOCX cannot be decrypted with the provided
 * password (wrong password) or when the encryption metadata is malformed.
 */
export class DocxDecryptionError extends DocxError {
  override name = "DocxDecryptionError";
}

/**
 * Error thrown when an input package exceeds a declared resource limit (e.g.
 * total package size, single part size, number of parts). Used to defend
 * against ZIP bombs and runaway memory usage.
 */
export class DocxLimitExceededError extends DocxParseError {
  override name = "DocxLimitExceededError";
  /** Limit category that was exceeded. */
  readonly limit: "packageSize" | "partSize" | "partCount";
  /** Configured maximum. */
  readonly maximum: number;
  /** Actual measured value (or the value that would have been reached). */
  readonly actual: number;

  constructor(
    limit: "packageSize" | "partSize" | "partCount",
    maximum: number,
    actual: number,
    detail?: string
  ) {
    super(
      `DOCX ${limit} limit exceeded: actual ${actual} > maximum ${maximum}` +
        (detail ? ` (${detail})` : "")
    );
    this.limit = limit;
    this.maximum = maximum;
    this.actual = actual;
  }
}
