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
export class ZipInvalidSignatureError extends ZipParseError {
  override name = "ZipInvalidSignatureError";

  constructor(expected: string, actual: number, context?: string, options?: BaseErrorOptions) {
    const msg = context
      ? `Invalid ${context}: expected ${expected}, got 0x${actual.toString(16).padStart(8, "0")}`
      : `Invalid signature: expected ${expected}, got 0x${actual.toString(16).padStart(8, "0")}`;
    super(msg, options);
  }
}

/** @deprecated Use {@link ZipInvalidSignatureError}. Removed in 1.0. */
export const InvalidZipSignatureError = ZipInvalidSignatureError;
/** @deprecated Use {@link ZipInvalidSignatureError}. Removed in 1.0. */
export type InvalidZipSignatureError = ZipInvalidSignatureError;

/**
 * Error thrown when End of Central Directory is not found.
 */
export class ZipEocdNotFoundError extends ZipParseError {
  override name = "ZipEocdNotFoundError";

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
export class ZipCrc32MismatchError extends ArchiveError {
  override name = "ZipCrc32MismatchError";

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
export type ZipEntrySizeMismatchReason = "too-many-bytes" | "too-few-bytes";

/**
 * Error thrown when the actual decompressed size doesn't match the declared size.
 * This is a security feature to detect ZIP bombs and corrupted archives.
 */
export class ZipEntrySizeMismatchError extends ArchiveError {
  override name = "ZipEntrySizeMismatchError";

  constructor(
    public readonly path: string,
    public readonly expected: number,
    /**
     * Bytes produced. `undefined` when the inflater was stopped at the
     * declared size, so no count beyond it exists.
     */
    public readonly actual: number | undefined,
    public readonly reason: ZipEntrySizeMismatchReason,
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
 * Which configured bound an {@link ArchiveLimitExceededError} reports.
 *
 * - `maxOutputLength`: one decompression would produce more than its bound.
 *   Decompression stops as soon as the bound is crossed, so the full (possibly
 *   enormous) output is never materialised.
 * - `maxTotalUncompressedSize`: a bulk extraction would produce more in total.
 * - `maxEntries`: the central directory declares more entries than allowed;
 *   rejected before any entry record is allocated.
 * - `maxEntrySize`: one entry is larger than allowed — reported by the subclass
 *   {@link ArchiveFileTooLargeError}, which also names the entry.
 */
export type ArchiveLimit =
  | "maxOutputLength"
  | "maxTotalUncompressedSize"
  | "maxEntries"
  | "maxEntrySize";

/**
 * Error thrown when decompressing or extracting would exceed a configured limit. `limit` names
 * the option that set the bound and `allowed` is its value, so a caller can tell which limit was
 * hit without reading the message. The per-entry size (`maxEntrySize`) is reported by the
 * subclass {@link ArchiveFileTooLargeError}, so catching this class sees every limit.
 */
export class ArchiveLimitExceededError extends ArchiveError {
  override name = "ArchiveLimitExceededError";

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
          : limit === "maxEntrySize"
            ? `Archive entry exceeds maxEntrySize (${allowed} bytes)`
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
export class ZipDecryptionError extends ArchiveError {
  override name = "ZipDecryptionError";

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
export class ZipPasswordRequiredError extends ArchiveError {
  override name = "ZipPasswordRequiredError";

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
export class ArchiveRangeNotSupportedError extends ArchiveError {
  override name = "ArchiveRangeNotSupportedError";

  constructor(url: string, options?: BaseErrorOptions) {
    super(`Server does not support Range requests for: ${url}`, options);
  }
}

/**
 * Error thrown when an HTTP request fails.
 */
export class ArchiveHttpRangeError extends ArchiveError {
  override name = "ArchiveHttpRangeError";

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

/** Options for {@link ArchiveFileTooLargeError}. */
export interface ArchiveFileTooLargeErrorOptions extends BaseErrorOptions {
  /** The bound the entry exceeds, in bytes; becomes {@link ArchiveFileTooLargeError.allowed}. */
  allowed?: number;
}

/**
 * Error thrown when an entry is too large to read into memory: the `maxEntrySize` limit. An
 * {@link ArchiveLimitExceededError} with `limit: "maxEntrySize"`, plus the entry's `path`.
 *
 * A ZIP64 size or offset JavaScript cannot hold exactly is reported the same way, with
 * `allowed` set to `Number.MAX_SAFE_INTEGER` — a bound of the platform, not of an option.
 */
export class ArchiveFileTooLargeError extends ArchiveLimitExceededError {
  override name = "ArchiveFileTooLargeError";

  /** The limit that was hit: always `"maxEntrySize"`. */
  declare readonly limit: "maxEntrySize";

  /**
   * The bound the entry exceeds, in bytes: the configured `maxEntrySize`, or
   * `Number.MAX_SAFE_INTEGER` for a ZIP64 size or offset JavaScript cannot hold; `NaN` when the
   * code that threw did not say.
   */
  declare readonly allowed: number;

  /** The entry that is too large. */
  readonly path: string;

  constructor(path: string, reason: string, options?: ArchiveFileTooLargeErrorOptions) {
    const { allowed = Number.NaN, ...errorOptions } = options ?? {};
    super("maxEntrySize", allowed, options && errorOptions);
    this.message = `File "${path}" is too large to extract into memory (${reason})`;
    this.path = path;
  }
}

/**
 * The file system would not let `ArchiveFile.fromFile` / `writeFile` (or their `Sync` forms)
 * read or write the archive's file — it does not exist, permission is denied, the path is a
 * directory. The system's own error is the `cause` (its `code`, `ENOENT` or `EACCES`, is
 * `error.cause.code`), as `ExcelFileError` and `DocxFileError` keep it. A file that was read
 * but is not a valid archive is a {@link ZipParseError} or a plain {@link ArchiveError}, not
 * this. The `Fs.*` helpers are file-system functions themselves and pass the platform's
 * error on unwrapped.
 */
export class ArchiveFileError extends ArchiveError {
  override name = "ArchiveFileError";

  constructor(
    public readonly path: string,
    public readonly operation: "read" | "write",
    details?: string,
    options?: BaseErrorOptions & {
      /** What the file is to the archive: the archive itself (the default), a `"source"` read into it, or an `"extracted"` entry. */
      readonly fileKind?: "archive" | "source" | "extracted";
    }
  ) {
    const kind = options?.fileKind ?? "archive";
    super(
      details
        ? `Failed to ${operation} ${kind} file "${path}": ${details}`
        : `Failed to ${operation} ${kind} file "${path}"`,
      options?.cause === undefined ? undefined : { cause: options.cause }
    );
  }
}

// -----------------------------------------------------------------------------
// Unsupported Feature Errors
// -----------------------------------------------------------------------------

/**
 * Error thrown when an unsupported compression method is encountered.
 */
export class ZipUnsupportedCompressionError extends ArchiveError {
  override name = "ZipUnsupportedCompressionError";

  constructor(method: number, options?: BaseErrorOptions) {
    super(`Unsupported compression method: ${method}`, options);
  }
}

// -----------------------------------------------------------------------------
// Worker Errors
// -----------------------------------------------------------------------------

/**
 * A worker streaming session (`Compression.Workers.Pool#openStream`) whose worker did
 * not start — it timed out, failed or died first. Nothing was sent to it, so the caller can
 * run the same stream elsewhere. The library's own streams fall back on their own and never
 * report it; a caller of `openStream` can tell it apart with `instanceof`.
 */
export class ArchiveWorkerStreamStartError extends ArchiveError {
  override name = "ArchiveWorkerStreamStartError";
}

/**
 * A `Compression.Workers` member called where the Web Worker pool does not exist — the Node
 * entry, where zlib already runs on libuv's thread pool. `Compression.hasWorkerSupport()`
 * answers `false` there; check it first, or call `Compression.deflate` / `inflate`, which pick
 * the platform's codec on their own.
 */
export class ArchiveWorkerUnsupportedError extends ArchiveError {
  override name = "ArchiveWorkerUnsupportedError";
  /** The public call that was refused, e.g. `Compression.Workers.deflate`. */
  readonly operation: string;

  constructor(operation: string, options?: BaseErrorOptions) {
    super(
      `${operation}: the Web Worker pool is not available in Node.js, which runs zlib on its ` +
        "own thread pool; check Compression.hasWorkerSupport() first, or use Compression.deflate / inflate",
      options
    );
    this.operation = operation;
  }
}
