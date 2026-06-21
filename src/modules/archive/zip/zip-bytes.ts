/**
 * ZIP file format encoder (single-buffer output)
 *
 * Implements ZIP file structure according to PKWARE's APPNOTE.TXT specification
 * https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
 *
 * This module focuses on producing a complete ZIP as a single Uint8Array.
 * For true streaming (push chunks while reading sources), use `zip()` / `ZipArchive.stream()`.
 */

import type { CompressOptions } from "@archive/compression/compress";
import { compress, compressSync } from "@archive/compression/compress";
import { crc32 } from "@archive/compression/crc32";
import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import {
  DEFAULT_ZIP_LEVEL,
  DEFAULT_ZIP_TIMESTAMPS,
  REPRODUCIBLE_ZIP_MOD_TIME
} from "@archive/core/defaults";
import { ArchiveError } from "@archive/core/errors";
import type { ZipStringEncoding } from "@archive/core/text";
import { encodeZipString } from "@archive/core/text";
import type { ZipEncryptionMethod } from "@archive/crypto";
import {
  zipCryptoEncrypt,
  aesEncrypt,
  buildAesExtraField,
  randomBytes,
  isAesEncryption,
  getAesKeyStrength
} from "@archive/crypto";
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import { normalizeZipPath } from "@archive/zip-spec/zip-path";
import {
  FLAG_ENCRYPTED,
  COMPRESSION_AES,
  UINT32_MAX,
  ZIP_LOCAL_FILE_HEADER_FIXED_SIZE,
  buildZip64ExtraField,
  concatExtraFields,
  VERSION_ZIP64,
  writeLocalFileHeaderInto
} from "@archive/zip-spec/zip-records";
import type { Zip64Mode } from "@archive/zip-spec/zip-records";
import { isProbablyIncompressible } from "@archive/zip/compressibility";
import {
  measureCentralDirectoryAndEocd,
  writeCentralDirectoryAndEocdInto
} from "@archive/zip/writer-core";
import { resolveZipExternalAttributesAndVersionMadeBy } from "@archive/zip/zip-entry-attributes";
import {
  buildZipEntryMetadata,
  resolveZipCompressionMethod
} from "@archive/zip/zip-entry-metadata";

interface ProcessedEntry {
  name: Uint8Array;
  /** Case-insensitive sort key for deterministic ordering. */
  sortKey: string;
  uncompressedSize: number;
  compressedData: Uint8Array;
  crc: number;
  compressionMethod: number;
  modTime: number;
  modDate: number;
  extraField: Uint8Array;
  comment: Uint8Array;
  offset: number;
  flags: number;
  externalAttributes: number;
  versionMadeBy?: number;
}

/**
 * ZIP file entry
 */
export interface ZipEntry {
  /** File name (can include directory path, use forward slashes) */
  name: string;
  /** File data (will be compressed unless level=0) */
  data: Uint8Array;
  /** Optional per-entry compression level override */
  level?: number;
  /** File modification time (optional, defaults to current time) */
  modTime?: Date;

  /** Optional access time (used only when timestamps mode supports it). */
  atime?: Date;

  /** Optional metadata change time (used only when timestamps mode supports it). */
  ctime?: Date;

  /** Optional creation time (used by NTFS timestamps mode). */
  birthTime?: Date;
  /** File comment (optional) */
  comment?: string;
  /** Optional string encoding for this entry name/comment. */
  encoding?: ZipStringEncoding;
  /** Per-entry encryption method override */
  encryptionMethod?: ZipEncryptionMethod;
  /** Per-entry password override */
  password?: string | Uint8Array;

  /**
   * Unix mode/permissions for this entry.
   * Accepts either a full `stat.mode` (includes file type bits), or just permission bits.
   */
  mode?: number;

  /** Optional MS-DOS attributes (low 8 bits). */
  msDosAttributes?: number;

  /** Advanced override for the central directory `version made by` field. */
  versionMadeBy?: number;
  /**
   * External file attributes (optional).
   * For Unix symlinks, use: ((mode << 16) | 0x20)
   * where mode is typically 0o120777 for symlinks.
   */
  externalAttributes?: number;
}

// Re-export ZipRawEntry from shared module
export type { ZipRawEntry } from "@archive/zip/raw-entry";
import type { ZipRawEntry } from "@archive/zip/raw-entry";
import { isZipRawEntry } from "@archive/zip/raw-entry";

export type ZipBuildEntry = ZipEntry | ZipRawEntry;

interface ZipBuildSettings {
  level: number;
  timestamps: ZipTimestampMode;
  defaultModTime: Date;
  encryptionMethod: ZipEncryptionMethod;
  password?: string | Uint8Array;
  encoding?: ZipStringEncoding;
}

type ZipPathOptionValue = false | ZipPathOptions;

/**
 * Sort entries alphabetically by name (case-insensitive).
 *
 * Uses a decorate-sort-undecorate step to guarantee stability even when
 * multiple entries share the same `sortKey`.
 */
function sortProcessedEntriesByName(entries: ProcessedEntry[]): void {
  if (entries.length <= 1) {
    return;
  }

  const decorated = entries.map((entry, index) => ({ entry, index }));
  decorated.sort((a, b) => {
    if (a.entry.sortKey < b.entry.sortKey) {
      return -1;
    }
    if (a.entry.sortKey > b.entry.sortKey) {
      return 1;
    }
    return a.index - b.index;
  });
  for (let i = 0; i < decorated.length; i++) {
    entries[i] = decorated[i]!.entry;
  }
}

/**
 * Validate encryption options and throw if invalid.
 */
function validateEncryptionOptions(
  encryptionMethod: ZipEncryptionMethod,
  password: string | Uint8Array | undefined,
  isSync: boolean
): void {
  if (encryptionMethod !== "none" && !password) {
    throw new ArchiveError("Password is required when encryption is enabled");
  }
  if (isSync && isAesEncryption(encryptionMethod)) {
    throw new ArchiveError(
      "AES encryption requires async API. Use createZip() instead of createZipSync()."
    );
  }
}
/**
 * Parse common ZIP options into build settings.
 */
function parseZipBuildOptions(options: ZipOptions): {
  settings: ZipBuildSettings;
  zipComment: Uint8Array;
  zip64Mode: Zip64Mode;
  smartStore: boolean;
  thresholdBytes: number | undefined;
  path: ZipPathOptionValue;
  noSort: boolean;
} {
  const reproducible = options.reproducible ?? false;
  const level = options.level ?? DEFAULT_ZIP_LEVEL;
  const timestamps: ZipTimestampMode =
    options.timestamps ?? (reproducible ? "dos" : DEFAULT_ZIP_TIMESTAMPS);
  const defaultModTime = options.modTime ?? (reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date());

  return {
    settings: {
      level,
      timestamps,
      defaultModTime,
      encryptionMethod: options.encryptionMethod ?? "none",
      password: options.password,
      encoding: options.encoding
    },
    zipComment: encodeZipString(options.comment, options.encoding),
    zip64Mode: options.zip64 ?? "auto",
    smartStore: options.smartStore ?? true,
    thresholdBytes: options.thresholdBytes,
    path: options.path ?? false,
    noSort: options.noSort ?? false
  };
}

function shouldDeflate(level: number, data: Uint8Array): boolean {
  return level > 0 && data.length > 0;
}

async function compressEntryMaybe(
  entry: ZipEntry,
  level: number,
  compressOptions: CompressOptions,
  smartStore: boolean
): Promise<{ compressedData: Uint8Array; deflate: boolean }> {
  if (!shouldDeflate(level, entry.data)) {
    return { compressedData: entry.data, deflate: false };
  }

  if (!smartStore) {
    const compressed = await compress(entry.data, compressOptions);
    return { compressedData: compressed, deflate: true };
  }

  // Heuristic: skip deflate for high-entropy inputs.
  if (isProbablyIncompressible(entry.data)) {
    return { compressedData: entry.data, deflate: false };
  }

  const compressed = await compress(entry.data, compressOptions);
  if (compressed.length >= entry.data.length) {
    return { compressedData: entry.data, deflate: false };
  }

  return { compressedData: compressed, deflate: true };
}

function compressEntryMaybeSync(
  entry: ZipEntry,
  level: number,
  compressOptions: CompressOptions,
  smartStore: boolean
): { compressedData: Uint8Array; deflate: boolean } {
  if (!shouldDeflate(level, entry.data)) {
    return { compressedData: entry.data, deflate: false };
  }

  if (!smartStore) {
    const compressed = compressSync(entry.data, compressOptions);
    return { compressedData: compressed, deflate: true };
  }

  if (isProbablyIncompressible(entry.data)) {
    return { compressedData: entry.data, deflate: false };
  }

  const compressed = compressSync(entry.data, compressOptions);
  if (compressed.length >= entry.data.length) {
    return { compressedData: entry.data, deflate: false };
  }

  return { compressedData: compressed, deflate: true };
}

function buildProcessedEntry(
  entry: ZipEntry,
  settings: ZipBuildSettings,
  path: ZipPathOptionValue,
  compressedData: Uint8Array,
  deflate: boolean,
  encryptionResult?: {
    data: Uint8Array;
    extraField?: Uint8Array;
    compressionMethod: number;
    crcOverride?: number;
  }
): ProcessedEntry {
  const resolvedName = path ? normalizeZipPath(entry.name, path) : entry.name;
  const modDate = entry.modTime ?? settings.defaultModTime;
  const metadata = buildZipEntryMetadata({
    name: resolvedName,
    comment: entry.comment,
    modTime: modDate,
    atime: entry.atime,
    ctime: entry.ctime,
    birthTime: entry.birthTime,
    timestamps: settings.timestamps,
    useDataDescriptor: false,
    deflate,
    codec: entry.encoding ?? settings.encoding
  });

  // Determine final data and compression method based on encryption
  let finalData: Uint8Array;
  let finalCompressionMethod: number;
  let finalExtraField: Uint8Array = metadata.extraField;
  let flags = metadata.flags;
  // CRC-32 stored in the headers. WinZip AE-2 mandates a zero CRC field, so the
  // encryption layer may override the real CRC (crcOverride === 0).
  let finalCrc = crc32(entry.data);

  if (encryptionResult) {
    finalData = encryptionResult.data;
    finalCompressionMethod = encryptionResult.compressionMethod;
    flags |= FLAG_ENCRYPTED;
    if (encryptionResult.extraField) {
      finalExtraField = concatExtraFields(metadata.extraField, encryptionResult.extraField);
    }
    if (encryptionResult.crcOverride !== undefined) {
      finalCrc = encryptionResult.crcOverride;
    }
  } else {
    finalData = compressedData;
    finalCompressionMethod = resolveZipCompressionMethod(deflate);
  }

  const attrs = resolveZipExternalAttributesAndVersionMadeBy({
    name: resolvedName,
    mode: entry.mode,
    msDosAttributes: entry.msDosAttributes,
    externalAttributes: entry.externalAttributes,
    versionMadeBy: entry.versionMadeBy
  });

  return {
    name: metadata.nameBytes,
    sortKey: resolvedName.toLowerCase(),
    uncompressedSize: entry.data.length,
    compressedData: finalData,
    crc: finalCrc,
    compressionMethod: finalCompressionMethod,
    modTime: metadata.dosTime,
    modDate: metadata.dosDate,
    extraField: finalExtraField,
    comment: metadata.commentBytes,
    offset: 0,
    flags,
    externalAttributes: attrs.externalAttributes,
    versionMadeBy: attrs.versionMadeBy
  };
}

function buildProcessedRawEntry(
  entry: ZipRawEntry,
  settings: ZipBuildSettings,
  path: ZipPathOptionValue
): ProcessedEntry {
  const resolvedName = path ? normalizeZipPath(entry.name, path) : entry.name;
  const modDate = entry.modTime ?? settings.defaultModTime;
  const metadata = buildZipEntryMetadata({
    name: resolvedName,
    comment: entry.comment,
    modTime: modDate,
    timestamps: settings.timestamps,
    useDataDescriptor: false,
    deflate: false,
    codec: entry.encoding ?? settings.encoding
  });

  const flags = ((entry.flags ?? 0) | metadata.flags) >>> 0;

  const attrs = resolveZipExternalAttributesAndVersionMadeBy({
    name: resolvedName,
    externalAttributes: entry.externalAttributes,
    versionMadeBy: entry.versionMadeBy
  });

  return {
    name: metadata.nameBytes,
    sortKey: resolvedName.toLowerCase(),
    uncompressedSize: entry.uncompressedSize,
    compressedData: entry.compressedData,
    crc: entry.crc32 >>> 0,
    compressionMethod: entry.compressionMethod,
    modTime: metadata.dosTime,
    modDate: metadata.dosDate,
    extraField: entry.extraField ?? metadata.extraField,
    comment: metadata.commentBytes,
    offset: 0,
    flags,
    externalAttributes: attrs.externalAttributes,
    versionMadeBy: attrs.versionMadeBy
  };
}

/**
 * ZIP encoder options
 */
export interface ZipOptions extends CompressOptions {
  /** ZIP file comment (optional) */
  comment?: string;

  /**
   * Default modification time for entries that don't specify `modTime`.
   *
   * If you need stable output across runs, either pass this explicitly or use `reproducible: true`.
   */
  modTime?: Date;

  /**
   * If true, bias defaults toward reproducible output:
   * - default `modTime` becomes 1980-01-01 00:00:00 (local time)
   * - default `timestamps` becomes "dos" (no UTC extra field)
   */
  reproducible?: boolean;

  /**
   * If true, entries are written in their original input order.
   * If false (default), entries are sorted alphabetically by name.
   */
  noSort?: boolean;

  /**
   * Max number of entries to compress concurrently in `createZip()`.
   * This helps avoid zlib threadpool saturation / memory spikes with many files.
   *
   * Defaults to 4.
   */
  concurrency?: number;

  /** Optional string encoding for entry names/comments and archive comment. */
  encoding?: ZipStringEncoding;

  /**
   * If true (default), automatically STORE incompressible data.
   * If false, always follow `level` (DEFLATE when level > 0).
   */
  smartStore?: boolean;

  /**
   * Timestamp writing strategy.
   * - "dos": only write DOS date/time fields (smallest output)
   * - "dos+utc": also write UTC mtime in 0x5455 extra field
   */
  timestamps?: ZipTimestampMode;

  /**
   * ZIP64 mode:
   * - "auto" (default): write ZIP64 only when required by limits (e.g. >65535 entries).
   * - true: force ZIP64 structures even for small archives (less legacy compatibility).
   * - false: forbid ZIP64; throws if ZIP64 is required.
   */
  zip64?: Zip64Mode;

  /**
   * Encryption method for all entries:
   * - "none" (default): no encryption
   * - "zipcrypto": Traditional PKWARE encryption (weak, for compatibility)
   * - "aes-128", "aes-192", "aes-256": WinZip AES encryption (recommended)
   */
  encryptionMethod?: ZipEncryptionMethod;

  /**
   * Password for encryption. Required when encryptionMethod is not "none".
   */
  password?: string | Uint8Array;

  /**
   * Optional entry name normalization.
   * - `false` (default): do not modify entry names.
   * - `ZipPathOptions`: normalize each entry name before writing.
   */
  path?: ZipPathOptionValue;
}

/**
 * Encrypt compressed data using the specified method.
 */
async function encryptData(
  compressedData: Uint8Array,
  originalCrc: number,
  encryptionMethod: ZipEncryptionMethod,
  password: string | Uint8Array,
  originalCompressionMethod: number
): Promise<{
  data: Uint8Array;
  extraField?: Uint8Array;
  compressionMethod: number;
  crcOverride?: number;
}> {
  if (encryptionMethod === "zipcrypto") {
    // ZipCrypto encryption
    const encrypted = zipCryptoEncrypt(compressedData, password, originalCrc, randomBytes);
    return {
      data: encrypted,
      compressionMethod: originalCompressionMethod
    };
  }

  if (isAesEncryption(encryptionMethod)) {
    // AES encryption (WinZip AE-2). AE-2 does NOT store the real CRC-32 — the
    // field must be 0 (integrity is provided by the HMAC). Writing the real
    // CRC violates the spec and causes strict readers (WinZip, 7-Zip,
    // pyzipper, …) to reject the entry with a CRC error.
    const keyStrength = getAesKeyStrength(encryptionMethod)!;
    const encrypted = await aesEncrypt(compressedData, password, keyStrength);
    const aesExtraField = buildAesExtraField(2, keyStrength, originalCompressionMethod);
    return {
      data: encrypted,
      extraField: aesExtraField,
      compressionMethod: COMPRESSION_AES,
      crcOverride: 0
    };
  }

  // No encryption
  return {
    data: compressedData,
    compressionMethod: originalCompressionMethod
  };
}

/**
 * Encrypt compressed data synchronously (ZipCrypto only).
 */
function encryptDataSync(
  compressedData: Uint8Array,
  originalCrc: number,
  encryptionMethod: ZipEncryptionMethod,
  password: string | Uint8Array,
  originalCompressionMethod: number
): { data: Uint8Array; extraField?: Uint8Array; compressionMethod: number } {
  if (encryptionMethod === "zipcrypto") {
    const encrypted = zipCryptoEncrypt(compressedData, password, originalCrc, randomBytes);
    return {
      data: encrypted,
      compressionMethod: originalCompressionMethod
    };
  }

  if (isAesEncryption(encryptionMethod)) {
    throw new ArchiveError(
      "AES encryption requires async API. Use createZip() instead of createZipSync()."
    );
  }

  return {
    data: compressedData,
    compressionMethod: originalCompressionMethod
  };
}

function finalizeZip(
  processedEntries: ProcessedEntry[],
  zipComment: Uint8Array,
  zip64Mode: Zip64Mode = "auto"
): Uint8Array {
  const forceZip64 = zip64Mode === true;

  // Precompute offsets and effective extra fields (local vs central can differ for ZIP64).
  const localExtraFields: Uint8Array[] = new Array(processedEntries.length);
  const zip64EntryNeeded: boolean[] = new Array(processedEntries.length);

  let localSectionSize = 0;
  for (let i = 0; i < processedEntries.length; i++) {
    const entry = processedEntries[i]!;
    entry.offset = localSectionSize;

    const compressedSize = entry.compressedData.length;
    const needsZip64Entry =
      forceZip64 ||
      entry.offset > UINT32_MAX ||
      compressedSize > UINT32_MAX ||
      entry.uncompressedSize > UINT32_MAX;
    zip64EntryNeeded[i] = needsZip64Entry;

    const zip64LocalExtra = needsZip64Entry
      ? buildZip64ExtraField({
          uncompressedSize: entry.uncompressedSize,
          compressedSize
        })
      : EMPTY_UINT8ARRAY;

    localExtraFields[i] = needsZip64Entry
      ? concatExtraFields(entry.extraField, zip64LocalExtra)
      : entry.extraField;

    const localHeaderSize =
      ZIP_LOCAL_FILE_HEADER_FIXED_SIZE + entry.name.length + localExtraFields[i]!.length;
    localSectionSize += localHeaderSize + compressedSize;
  }

  const centralDirOffset = localSectionSize;
  const cdSizing = measureCentralDirectoryAndEocd(processedEntries, {
    zipComment,
    zip64Mode,
    centralDirOffset
  });

  const totalSize = localSectionSize + cdSizing.totalSize;
  const out = new Uint8Array(totalSize);
  const view = new DataView(out.buffer, out.byteOffset, out.byteLength);

  let offset = 0;

  // Local file headers and data
  for (let i = 0; i < processedEntries.length; i++) {
    const entry = processedEntries[i]!;
    const compressedSize = entry.compressedData.length;
    const needsZip64Entry = zip64EntryNeeded[i]!;

    offset += writeLocalFileHeaderInto(out, view, offset, {
      fileName: entry.name,
      extraField: localExtraFields[i]!,
      flags: entry.flags,
      compressionMethod: entry.compressionMethod,
      dosTime: entry.modTime,
      dosDate: entry.modDate,
      crc32: entry.crc,
      compressedSize: needsZip64Entry ? UINT32_MAX : compressedSize,
      uncompressedSize: needsZip64Entry ? UINT32_MAX : entry.uncompressedSize,
      versionNeeded: needsZip64Entry ? VERSION_ZIP64 : undefined
    });

    out.set(entry.compressedData, offset);
    offset += compressedSize;
  }

  writeCentralDirectoryAndEocdInto(processedEntries, {
    zipComment,
    zip64Mode,
    centralDirOffset,
    out,
    offset
  });

  return out;
}

/**
 * Create a ZIP file from entries (async)
 */
export async function createZip(
  entries: ZipBuildEntry[],
  options: ZipOptions = {}
): Promise<Uint8Array> {
  const { settings, zipComment, zip64Mode, smartStore, thresholdBytes, path, noSort } =
    parseZipBuildOptions(options);
  validateEncryptionOptions(settings.encryptionMethod, settings.password, false);

  const concurrency = options.concurrency ?? 4;
  const limit = Math.max(1, Math.floor(concurrency));
  const processedEntries = new Array<ProcessedEntry>(entries.length);

  if (entries.length > 0) {
    let nextIndex = 0;
    const workerCount = Math.min(limit, entries.length);

    const processEntryAt = async (idx: number): Promise<void> => {
      const entry = entries[idx]!;

      if (isZipRawEntry(entry)) {
        processedEntries[idx] = buildProcessedRawEntry(entry, settings, path);
        return;
      }

      const entryLevel = entry.level ?? settings.level;
      const compressOptions: CompressOptions = {
        level: entryLevel,
        thresholdBytes
      };
      const { compressedData, deflate } = await compressEntryMaybe(
        entry,
        entryLevel,
        compressOptions,
        smartStore
      );

      // Handle encryption
      const entryEncMethod = entry.encryptionMethod ?? settings.encryptionMethod;
      const entryPassword = entry.password ?? settings.password;
      let encryptionResult:
        | { data: Uint8Array; extraField?: Uint8Array; compressionMethod: number }
        | undefined;

      if (entryEncMethod !== "none" && entryPassword) {
        const originalCrc = crc32(entry.data);
        const originalCompressionMethod = resolveZipCompressionMethod(deflate);
        encryptionResult = await encryptData(
          compressedData,
          originalCrc,
          entryEncMethod,
          entryPassword,
          originalCompressionMethod
        );
      }

      processedEntries[idx] = buildProcessedEntry(
        entry,
        settings,
        path,
        compressedData,
        deflate,
        encryptionResult
      );
    };

    const runWorker = async (): Promise<void> => {
      while (true) {
        const idx = nextIndex++;
        if (idx >= entries.length) {
          return;
        }
        await processEntryAt(idx);
      }
    };

    const workers = Array.from({ length: workerCount }, () => runWorker());

    await Promise.all(workers);
  }

  if (!noSort) {
    sortProcessedEntriesByName(processedEntries);
  }

  return finalizeZip(processedEntries, zipComment, zip64Mode);
}

/**
 * Create a ZIP file from entries (sync)
 *
 * This is supported in both Node.js and browser builds.
 * Note: AES encryption is not supported in sync mode.
 */
export function createZipSync(entries: ZipBuildEntry[], options: ZipOptions = {}): Uint8Array {
  const { settings, zipComment, zip64Mode, smartStore, thresholdBytes, path, noSort } =
    parseZipBuildOptions(options);
  validateEncryptionOptions(settings.encryptionMethod, settings.password, true);

  const processedEntries: ProcessedEntry[] = new Array(entries.length);

  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index]!;
    if (isZipRawEntry(entry)) {
      processedEntries[index] = buildProcessedRawEntry(entry, settings, path);
      continue;
    }

    const entryLevel = entry.level ?? settings.level;
    const compressOptions: CompressOptions = {
      level: entryLevel,
      thresholdBytes
    };
    const { compressedData, deflate } = compressEntryMaybeSync(
      entry,
      entryLevel,
      compressOptions,
      smartStore
    );

    // Handle encryption
    const entryEncMethod = entry.encryptionMethod ?? settings.encryptionMethod;
    const entryPassword = entry.password ?? settings.password;
    let encryptionResult:
      | { data: Uint8Array; extraField?: Uint8Array; compressionMethod: number }
      | undefined;

    if (entryEncMethod !== "none" && entryPassword) {
      const originalCrc = crc32(entry.data);
      const originalCompressionMethod = resolveZipCompressionMethod(deflate);
      encryptionResult = encryptDataSync(
        compressedData,
        originalCrc,
        entryEncMethod,
        entryPassword,
        originalCompressionMethod
      );
    }

    processedEntries[index] = buildProcessedEntry(
      entry,
      settings,
      path,
      compressedData,
      deflate,
      encryptionResult
    );
  }

  if (!noSort) {
    sortProcessedEntriesByName(processedEntries);
  }
  return finalizeZip(processedEntries, zipComment, zip64Mode);
}
