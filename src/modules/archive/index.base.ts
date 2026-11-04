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

import { encodePng as encodePngImpl } from "@archive/png";

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
export type { EOCDInfo, ZIP64EOCDInfo } from "@archive/zip-spec/zip-parser-core";
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
export type { UnpackedZipEntry, UnpackZipOptions } from "@archive/zip-one-shot";
export type { ZipParseOptions } from "@archive/unzip/zip-parser";

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

// =============================================================================
// Errors
// =============================================================================

export {
  // Abort error class
  AbortError,
  // Error classes
  ArchiveError,
  ArchiveArgumentError,
  isArchiveError,
  ZipParseError,
  InvalidZipSignatureError,
  EocdNotFoundError,
  Crc32MismatchError,
  DecryptionError,
  PasswordRequiredError,
  RangeNotSupportedError,
  HttpRangeError,
  FileTooLargeError,
  UnsupportedCompressionError,
  EntrySizeMismatchError,
  ArchiveLimitError,
  WorkerStreamStartError,
  // Error types
  type EntrySizeMismatchReason,
  type ArchiveLimit,
  type FileTooLargeErrorOptions
} from "@archive/core/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
