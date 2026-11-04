/**
 * Archive module error types.
 *
 * All archive-related errors extend ArchiveError.
 */

import type { BaseErrorOptions } from "@utils/errors";
import { BaseError } from "@utils/errors";

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
 * Base class for all archive-related errors.
 */
export class ArchiveError extends BaseError {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "ArchiveError";
  }
}

/**
 * A wrong argument or option given to an archive function: the caller's mistake, not damage
 * in archive data (which stays an {@link ArchiveError}). It is an `ArchiveError`, so existing
 * `catch` blocks and `isArchiveError` still see it.
 */
export class ArchiveArgumentError extends ArchiveError {
  override name = "ArchiveArgumentError";
  /** The argument or option at fault, e.g. `level`, `source` or `entries[0]`. */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/**
 * The error a reader raises when it is asked to read a one-shot source (a `ReadableStream`,
 * a Node stream, an async iterable) a second time. Such a source is empty once read; reading it
 * again used to answer "no entries" for TAR and "not a ZIP file" for ZIP.
 */
export function oneShotSourceReadError(where: string): ArchiveError {
  return new ArchiveError(
    `${where}: the archive's source is a stream that an earlier read already consumed, and a ` +
      "stream cannot be read twice; open the archive again, or give the reader bytes or a Blob"
  );
}

/**
 * Check if an error is an ArchiveError.
 *
 * An {@link AbortError} is not one: cancellation is reported by the shared
 * `AbortError` every module throws, not by the archive module.
 */
export function isArchiveError(err: unknown): err is ArchiveError {
  return err instanceof ArchiveError;
}

// -----------------------------------------------------------------------------
// ZIP Parsing Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when ZIP parsing fails.
 */
export class ZipParseError extends ArchiveError {
  override name = "ZipParseError";
}

/**
 * Error thrown when an invalid ZIP signature is encountered.
 */
export class InvalidZipSignatureError extends ZipParseError {
  override name = "InvalidZipSignatureError";

  constructor(expected: string, actual: number, context?: string, options?: BaseErrorOptions) {
    const msg = context
      ? `Invalid ${context}: expected ${expected}, got 0x${actual.toString(16).padStart(8, "0")}`
      : `Invalid signature: expected ${expected}, got 0x${actual.toString(16).padStart(8, "0")}`;
    super(msg, options);
  }
}

/**
 * Error thrown when End of Central Directory is not found.
 */
export class EocdNotFoundError extends ZipParseError {
  override name = "EocdNotFoundError";

  constructor(options?: BaseErrorOptions) {
    super("Invalid ZIP file: End of Central Directory not found", options);
  }
}

/**
 * A native inflater's failure — one of Node's zlib `Z_*` errors, or a `DecompressionStream`'s — as the archive's own
 * error, with the original as its `cause`.
 *
 * Corrupt DEFLATE data is malformed input like any other, and the pure-JS inflater already reports it as an
 * `ArchiveError`; without this, the same damage surfaced as a bare `Error` or `TypeError` on the platforms that use a
 * native one. An error this library already typed (a limit, an abort, the pure-JS inflater's) passes through.
 */
export function toInflateError(error: unknown, path?: string): Error {
  if (error instanceof BaseError) {
    return error;
  }
  const cause = error instanceof Error ? error : new Error(String(error));
  const where = path === undefined ? "" : ` in "${path}"`;
  return new ArchiveError(`Invalid DEFLATE data${where}: ${cause.message}`, { cause });
}

/** Whether `error` came from Node's zlib binding, whose errors carry a `Z_*` code. */
export function isZlibError(error: unknown): boolean {
  const code = (error as { code?: unknown } | null)?.code;
  return typeof code === "string" && code.startsWith("Z_");
}

// -----------------------------------------------------------------------------
// CRC32 Validation Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when CRC32 validation fails.
 */
export class Crc32MismatchError extends ArchiveError {
  override name = "Crc32MismatchError";

  constructor(
    public readonly path: string,
    public readonly expected: number,
    public readonly actual: number,
    options?: BaseErrorOptions
  ) {
    super(
      `CRC32 mismatch for "${path}": expected 0x${expected.toString(16).padStart(8, "0")}, got 0x${actual.toString(16).padStart(8, "0")}`,
      options
    );
  }
}

// -----------------------------------------------------------------------------
// Entry Size Validation Errors
// -----------------------------------------------------------------------------

/**
 * Reason for entry size mismatch.
 * - `too-many-bytes`: ZIP bomb detected - actual size exceeds declared size
 * - `too-few-bytes`: Corruption detected - actual size is less than declared size
 */
export type EntrySizeMismatchReason = "too-many-bytes" | "too-few-bytes";

/**
 * Error thrown when the actual decompressed size doesn't match the declared size.
 * This is a security feature to detect ZIP bombs and corrupted archives.
 */
export class EntrySizeMismatchError extends ArchiveError {
  override name = "EntrySizeMismatchError";

  constructor(
    public readonly path: string,
    public readonly expected: number,
    /**
     * Bytes produced. `undefined` when the inflater was stopped at the
     * declared size, so no count beyond it exists.
     */
    public readonly actual: number | undefined,
    public readonly reason: EntrySizeMismatchReason,
    options?: BaseErrorOptions
  ) {
    const msg =
      reason === "too-few-bytes"
        ? `Entry "${path}" produced fewer bytes than declared: expected ${expected}, got ${actual}`
        : actual === undefined
          ? `Entry "${path}" exceeds its declared ${expected} bytes`
          : `Entry "${path}" produced more bytes than declared: expected ${expected}, got at least ${actual}`;
    super(msg, options);
  }

  /**
   * Check if this error indicates a potential ZIP bomb (too many bytes).
   */
  isZipBomb(): boolean {
    return this.reason === "too-many-bytes";
  }

  /**
   * Check if this error indicates data corruption (too few bytes).
   */
  isCorruption(): boolean {
    return this.reason === "too-few-bytes";
  }
}

// -----------------------------------------------------------------------------
// Limit Errors
// -----------------------------------------------------------------------------

/**
 * Which configured bound an {@link ArchiveLimitError} reports.
 *
 * - `maxOutputLength`: one decompression would produce more than its bound.
 *   Decompression stops as soon as the bound is crossed, so the full (possibly
 *   enormous) output is never materialised.
 * - `maxTotalUncompressedSize`: a bulk extraction would produce more in total.
 * - `maxEntries`: the central directory declares more entries than allowed;
 *   rejected before any entry record is allocated.
 */
export type ArchiveLimit = "maxOutputLength" | "maxTotalUncompressedSize" | "maxEntries";

/**
 * Error thrown when decompressing or extracting would exceed a configured limit. `limit` names
 * the option that set the bound and `allowed` is its value, so a caller can tell which limit was
 * hit without reading the message. The per-entry size (`maxEntrySize`) is reported by
 * {@link FileTooLargeError}, which carries the same `limit` / `allowed` fields.
 */
export class ArchiveLimitError extends ArchiveError {
  override name = "ArchiveLimitError";

  constructor(
    public readonly limit: ArchiveLimit,
    public readonly allowed: number,
    options?: BaseErrorOptions
  ) {
    super(
      limit === "maxOutputLength"
        ? `Decompressed output exceeds maxOutputLength (${allowed} bytes)`
        : limit === "maxEntries"
          ? `Archive entry count exceeds maxEntries (${allowed})`
          : `Archive uncompressed size exceeds maxTotalUncompressedSize (${allowed} bytes)`,
      options
    );
  }
}

// -----------------------------------------------------------------------------
// Encryption Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when decryption fails (wrong password or corrupted data).
 */
export class DecryptionError extends ArchiveError {
  override name = "DecryptionError";

  constructor(path: string, details?: string, options?: BaseErrorOptions) {
    super(
      details
        ? `Failed to decrypt "${path}": ${details}`
        : `Failed to decrypt "${path}": incorrect password or corrupted data`,
      options
    );
  }
}

/**
 * Error thrown when a password is required but not provided.
 */
export class PasswordRequiredError extends ArchiveError {
  override name = "PasswordRequiredError";

  constructor(path: string, options?: BaseErrorOptions) {
    super(`File "${path}" is encrypted. Please provide a password to extract.`, options);
  }
}

// -----------------------------------------------------------------------------
// HTTP / Network Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when the server doesn't support Range requests.
 */
export class RangeNotSupportedError extends ArchiveError {
  override name = "RangeNotSupportedError";

  constructor(url: string, options?: BaseErrorOptions) {
    super(`Server does not support Range requests for: ${url}`, options);
  }
}

/**
 * Error thrown when an HTTP request fails.
 */
export class HttpRangeError extends ArchiveError {
  override name = "HttpRangeError";

  constructor(
    public readonly url: string,
    public readonly status: number,
    public readonly statusText: string,
    options?: BaseErrorOptions
  ) {
    super(`HTTP ${status} ${statusText} for: ${url}`, options);
  }
}

// -----------------------------------------------------------------------------
// ZIP64 / Size Errors
// -----------------------------------------------------------------------------

/** Options for {@link FileTooLargeError}. */
export interface FileTooLargeErrorOptions extends BaseErrorOptions {
  /** The bound the entry exceeds, in bytes; becomes {@link FileTooLargeError.allowed}. */
  allowed?: number;
}

/**
 * Error thrown when an entry is too large to read into memory: the `maxEntrySize` limit. It
 * carries the same `limit` / `allowed` fields as {@link ArchiveLimitError}, read the same way,
 * plus the entry's `path`; it is not an `ArchiveLimitError`, so catch both to see every limit.
 */
export class FileTooLargeError extends ArchiveError {
  override name = "FileTooLargeError";

  /** The limit that was hit: always `"maxEntrySize"`. */
  readonly limit = "maxEntrySize" as const;

  /**
   * The bound the entry exceeds, in bytes: the configured `maxEntrySize`, or
   * `Number.MAX_SAFE_INTEGER` for a ZIP64 size or offset JavaScript cannot hold; `NaN` when the
   * code that threw did not say.
   */
  readonly allowed: number;

  /** The entry that is too large. */
  readonly path: string;

  constructor(path: string, reason: string, options?: FileTooLargeErrorOptions) {
    const { allowed = Number.NaN, ...errorOptions } = options ?? {};
    super(
      `File "${path}" is too large to extract into memory (${reason})`,
      options && errorOptions
    );
    this.allowed = allowed;
    this.path = path;
  }
}

// -----------------------------------------------------------------------------
// Unsupported Feature Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when an unsupported compression method is encountered.
 */
export class UnsupportedCompressionError extends ArchiveError {
  override name = "UnsupportedCompressionError";

  constructor(method: number, options?: BaseErrorOptions) {
    super(`Unsupported compression method: ${method}`, options);
  }
}

// -----------------------------------------------------------------------------
// Worker Errors
// -----------------------------------------------------------------------------

/**
 * A worker streaming session (`Compression.Workers.WorkerPool#openStream`) whose worker did
 * not start — it timed out, failed or died first. Nothing was sent to it, so the caller can
 * run the same stream elsewhere. The library's own streams fall back on their own and never
 * report it; a caller of `openStream` can tell it apart with `instanceof`.
 */
export class WorkerStreamStartError extends ArchiveError {
  override name = "WorkerStreamStartError";
}
