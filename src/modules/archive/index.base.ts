/**
 * Archive (ZIP) module - shared exports.
 *
 * This module contains exports that are identical across Node.js and browser.
 * Platform-specific entrypoints (index.ts / index.browser.ts) should re-export
 * from this file and then layer their platform-specific bindings.
 *
 * Design principles (aligned with the other domain modules):
 * - The platform-agnostic public value API lives on the `Archive` namespace
 *   (`export * as Archive`), tree-shaken per-member.
 * - Error classes and pure types stay flat at the top level.
 */

import type {
  ZipEntrySizeMismatchReason as ZipEntrySizeMismatchReasonImpl,
  ArchiveFileTooLargeErrorOptions as ArchiveFileTooLargeErrorOptionsImpl
} from "@archive/core/errors";
import {
  ZipEocdNotFoundError as ZipEocdNotFoundErrorImpl,
  ZipCrc32MismatchError as ZipCrc32MismatchErrorImpl,
  ZipDecryptionError as ZipDecryptionErrorImpl,
  ZipPasswordRequiredError as ZipPasswordRequiredErrorImpl,
  ZipEntrySizeMismatchError as ZipEntrySizeMismatchErrorImpl,
  ZipUnsupportedCompressionError as ZipUnsupportedCompressionErrorImpl,
  ArchiveRangeNotSupportedError as ArchiveRangeNotSupportedErrorImpl,
  ArchiveHttpRangeError as ArchiveHttpRangeErrorImpl,
  ArchiveFileTooLargeError as ArchiveFileTooLargeErrorImpl,
  ArchiveWorkerStreamStartError as ArchiveWorkerStreamStartErrorImpl
} from "@archive/core/errors";
import { encodePng as encodePngImpl } from "@archive/png";
import type {
  EOCDInfo as EOCDInfoImpl,
  ZIP64EOCDInfo as ZIP64EOCDInfoImpl
} from "@archive/zip-spec/zip-parser-core";

// =============================================================================
// Core Types
// =============================================================================

// Unified archive I/O
export type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
export type { ArchiveSink } from "@archive/io/archive-sink";

// Random Access / HTTP Range reading
export type {
  RandomAccessReader,
  HttpRangeReaderOptions,
  HttpRangeReaderStats
} from "@archive/io/random-access";

export type {
  RemoteZipReaderOptions,
  RemoteZipOpenOptions,
  RemoteZipStats,
  RemoteZipExtractOptions,
  RawEntryReadOptions,
  RawEntryStreamOptions,
  ExtractOptions
} from "@archive/unzip/remote-zip-reader";
export type { ArchiveOperationBase } from "@archive/core/progress";

// High-level ZIP API types
export type { ArchiveFormat } from "@archive/core/types";
export type { ZipEntryRecord, ZipEntryType } from "@archive/zip-spec/zip-entry-info";
export type { ZipCryptoState } from "@archive/crypto/zip-crypto";
export type { Zip64Mode } from "@archive/zip-spec/zip-records";
export type { ZipPathOptions } from "@archive/zip-spec/zip-path";
export type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
export type { ZipStringEncoding } from "@archive/core/text";
export type { SerializedZipEditOp, SerializedZipEditPlan } from "@archive/zip/zip-edit-plan";
export type {
  ZipOptions,
  ZipEntryOptions,
  ZipEditOptions,
  ZipEditUrlOptions,
  ZipEditWarning,
  ZipEditOp,
  ZipOperation,
  ZipProgress,
  ZipStreamOptions
} from "@archive/zip";
export type {
  UnzipOptions,
  UnzipOperation,
  UnzipProgress,
  UnzipStreamOptions
} from "@archive/unzip";
export type {
  UnzipOptionsTar,
  UnzipOptionsZip,
  ArchivePipeToOptions
} from "@archive/unzip/zip-reader";
export type { ZipOptionsTar, ZipOptionsZip } from "@archive/zip/zip-archive";

// TAR archive types
export type {
  TarType,
  TarEntryInfo,
  TarEntry,
  TarHeaderOptions,
  TarParseOptions,
  TarArchiveOptions,
  TarArchiveEntryOptions,
  TarArchiveProgress,
  TarArchiveStreamOptions,
  TarArchiveOperation,
  TarReaderOptions,
  TarReaderProgress,
  TarReaderStreamOptions,
  TarReaderOperation
} from "@archive/tar/index.browser";
export type { TarEntryInput } from "@archive/tar/tar-archive";
export type { UnpackedTarEntry } from "@archive/tar/tar-parser";
export type { UnpackedZipEntry, UnpackZipOptions } from "@archive/zip-one-shot";

// Encryption types
export type {
  AesKeyStrength,
  AesExtraFieldInfo,
  ZipEncryptionMethod,
  ZipEncryptionInfo,
  ZipPasswordOptions,
  ZipEncryptionOptions
} from "@archive/crypto";

export type { EncodePngOptions } from "@archive/png";
export type { TarGzOptions, ParseTarGzOptions } from "@archive/tar/tar-gzip";

/** @deprecated Use {@link Archive.encodePng}. Removed in 1.0. */
export const encodePng = encodePngImpl;
/**
 * @deprecated No public signature takes or returns it: the end-of-central-directory record is
 * read inside `RemoteZipReader` and never handed out. Removed in 1.0, with no replacement.
 */
export type EOCDInfo = EOCDInfoImpl;
/**
 * @deprecated No public signature takes or returns it: the ZIP64 end-of-central-directory
 * record is read inside `RemoteZipReader` and never handed out. Removed in 1.0, with no
 * replacement.
 */
export type ZIP64EOCDInfo = ZIP64EOCDInfoImpl;

// =============================================================================
// Errors
// =============================================================================

export {
  // Abort error class and its guard (also true for a native `DOMException` named "AbortError")
  AbortError,
  isAbortError,
  // Error classes
  ArchiveError,
  ArchiveArgumentError,
  isArchiveError,
  ZipParseError,
  ZipInvalidSignatureError,
  // Deprecated pre-prefix name; removed in 1.0.
  InvalidZipSignatureError,
  ZipEocdNotFoundError,
  ZipCrc32MismatchError,
  ZipDecryptionError,
  ZipPasswordRequiredError,
  ArchiveRangeNotSupportedError,
  ArchiveHttpRangeError,
  ArchiveFileTooLargeError,
  ArchiveFileError,
  ZipUnsupportedCompressionError,
  ZipEntrySizeMismatchError,
  ArchiveLimitError,
  ArchiveWorkerStreamStartError,
  ArchiveWorkerUnsupportedError,
  // Error types
  type ZipEntrySizeMismatchReason,
  type ArchiveLimit,
  type ArchiveFileTooLargeErrorOptions
} from "@archive/core/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";

// Deprecated error names — a ZIP-only error carries `Zip`, one any archive path raises `Archive`.

/** @deprecated Use {@link ZipEocdNotFoundError}. Removed in 1.0. */
export const EocdNotFoundError = ZipEocdNotFoundErrorImpl;
/** @deprecated Use {@link ZipEocdNotFoundError}. Removed in 1.0. */
export type EocdNotFoundError = ZipEocdNotFoundErrorImpl;
/** @deprecated Use {@link ZipCrc32MismatchError}. Removed in 1.0. */
export const Crc32MismatchError = ZipCrc32MismatchErrorImpl;
/** @deprecated Use {@link ZipCrc32MismatchError}. Removed in 1.0. */
export type Crc32MismatchError = ZipCrc32MismatchErrorImpl;
/** @deprecated Use {@link ZipDecryptionError}. Removed in 1.0. */
export const DecryptionError = ZipDecryptionErrorImpl;
/** @deprecated Use {@link ZipDecryptionError}. Removed in 1.0. */
export type DecryptionError = ZipDecryptionErrorImpl;
/** @deprecated Use {@link ZipPasswordRequiredError}. Removed in 1.0. */
export const PasswordRequiredError = ZipPasswordRequiredErrorImpl;
/** @deprecated Use {@link ZipPasswordRequiredError}. Removed in 1.0. */
export type PasswordRequiredError = ZipPasswordRequiredErrorImpl;
/** @deprecated Use {@link ZipEntrySizeMismatchError}. Removed in 1.0. */
export const EntrySizeMismatchError = ZipEntrySizeMismatchErrorImpl;
/** @deprecated Use {@link ZipEntrySizeMismatchError}. Removed in 1.0. */
export type EntrySizeMismatchError = ZipEntrySizeMismatchErrorImpl;
/** @deprecated Use {@link ZipEntrySizeMismatchReason}. Removed in 1.0. */
export type EntrySizeMismatchReason = ZipEntrySizeMismatchReasonImpl;
/** @deprecated Use {@link ZipUnsupportedCompressionError}. Removed in 1.0. */
export const UnsupportedCompressionError = ZipUnsupportedCompressionErrorImpl;
/** @deprecated Use {@link ZipUnsupportedCompressionError}. Removed in 1.0. */
export type UnsupportedCompressionError = ZipUnsupportedCompressionErrorImpl;
/** @deprecated Use {@link ArchiveRangeNotSupportedError}. Removed in 1.0. */
export const RangeNotSupportedError = ArchiveRangeNotSupportedErrorImpl;
/** @deprecated Use {@link ArchiveRangeNotSupportedError}. Removed in 1.0. */
export type RangeNotSupportedError = ArchiveRangeNotSupportedErrorImpl;
/** @deprecated Use {@link ArchiveHttpRangeError}. Removed in 1.0. */
export const HttpRangeError = ArchiveHttpRangeErrorImpl;
/** @deprecated Use {@link ArchiveHttpRangeError}. Removed in 1.0. */
export type HttpRangeError = ArchiveHttpRangeErrorImpl;
/** @deprecated Use {@link ArchiveFileTooLargeError}. Removed in 1.0. */
export const FileTooLargeError = ArchiveFileTooLargeErrorImpl;
/** @deprecated Use {@link ArchiveFileTooLargeError}. Removed in 1.0. */
export type FileTooLargeError = ArchiveFileTooLargeErrorImpl;
/** @deprecated Use {@link ArchiveFileTooLargeErrorOptions}. Removed in 1.0. */
export type FileTooLargeErrorOptions = ArchiveFileTooLargeErrorOptionsImpl;
/** @deprecated Use {@link ArchiveWorkerStreamStartError}. Removed in 1.0. */
export const WorkerStreamStartError = ArchiveWorkerStreamStartErrorImpl;
/** @deprecated Use {@link ArchiveWorkerStreamStartError}. Removed in 1.0. */
export type WorkerStreamStartError = ArchiveWorkerStreamStartErrorImpl;
