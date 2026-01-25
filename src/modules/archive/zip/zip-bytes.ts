/**
 * ZIP file format encoder (single-buffer output)
 *
 * Implements ZIP file structure according to PKWARE's APPNOTE.TXT specification
 * https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
 *
 * This module focuses on producing a complete ZIP as a single Uint8Array.
 * For true streaming (push chunks while reading sources), use `zip()` / `ZipArchive.stream()`.
 */

import { compress, compressSync, type CompressOptions } from "@archive/compression/compress";
import { crc32 } from "@archive/compression/crc32";
import { DEFAULT_ZIP_LEVEL, DEFAULT_ZIP_TIMESTAMPS } from "@archive/defaults";
import { isProbablyIncompressible } from "@archive/utils/compressibility";
import { encodeUtf8 } from "@archive/utils/text";
import { type ZipTimestampMode } from "@archive/utils/timestamps";
import {
  buildZipEntryMetadata,
  resolveZipCompressionMethod
} from "@archive/zip/zip-entry-metadata";
import {
  FLAG_UTF8,
  ZIP_CENTRAL_DIR_HEADER_FIXED_SIZE,
  ZIP_END_OF_CENTRAL_DIR_FIXED_SIZE,
  ZIP_LOCAL_FILE_HEADER_FIXED_SIZE,
  writeCentralDirectoryHeaderInto,
  writeEndOfCentralDirectoryInto,
  writeLocalFileHeaderInto
} from "@archive/zip-spec/zip-records";

const REPRODUCIBLE_ZIP_MOD_TIME = new Date(1980, 0, 1, 0, 0, 0);

interface ProcessedEntry {
  name: Uint8Array;
  uncompressedSize: number;
  compressedData: Uint8Array;
  crc: number;
  compressionMethod: number;
  modTime: number;
  modDate: number;
  extraField: Uint8Array;
  comment: Uint8Array;
  offset: number;
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
  /** File comment (optional) */
  comment?: string;
}

interface ZipBuildSettings {
  level: number;
  timestamps: ZipTimestampMode;
  defaultModTime: Date;
}

function encodeZipComment(comment?: string): Uint8Array {
  // Keep empty comment as empty bytes (no encoding surprises).
  return comment ? encodeUtf8(comment) : new Uint8Array(0);
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

function computeLocalRecordSize(entry: ProcessedEntry): number {
  return (
    ZIP_LOCAL_FILE_HEADER_FIXED_SIZE +
    entry.name.length +
    entry.extraField.length +
    entry.compressedData.length
  );
}

function computeCentralDirHeaderSize(entry: ProcessedEntry): number {
  return (
    ZIP_CENTRAL_DIR_HEADER_FIXED_SIZE +
    entry.name.length +
    entry.extraField.length +
    entry.comment.length
  );
}

function buildProcessedEntry(
  entry: ZipEntry,
  offset: number,
  settings: ZipBuildSettings,
  compressedData: Uint8Array,
  deflate: boolean
): ProcessedEntry {
  const modDate = entry.modTime ?? settings.defaultModTime;
  const metadata = buildZipEntryMetadata({
    name: entry.name,
    comment: entry.comment,
    modTime: modDate,
    timestamps: settings.timestamps,
    useDataDescriptor: false,
    deflate
  });

  return {
    name: metadata.nameBytes,
    uncompressedSize: entry.data.length,
    compressedData,
    crc: crc32(entry.data),
    compressionMethod: resolveZipCompressionMethod(deflate),
    modTime: metadata.dosTime,
    modDate: metadata.dosDate,
    extraField: metadata.extraField,
    comment: metadata.commentBytes,
    offset
  };
}

/**
 * Sort entries alphabetically by name (case-insensitive).
 * Uses Schwartzian transform to avoid repeated decoding during sort.
 */
function sortEntriesByName(entries: ProcessedEntry[]): void {
  if (entries.length <= 1) {
    return;
  }

  const decoder = new TextDecoder();
  const decorated = entries.map(entry => ({
    entry,
    sortKey: decoder.decode(entry.name).toLowerCase()
  }));
  decorated.sort((a, b) => a.sortKey.localeCompare(b.sortKey));

  for (let i = 0; i < decorated.length; i++) {
    entries[i] = decorated[i].entry;
  }
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
}

function finalizeZip(processedEntries: ProcessedEntry[], zipComment: Uint8Array): Uint8Array {
  // Assemble ZIP into a single buffer to reduce allocations and copying.
  let localSectionSize = 0;
  let centralDirSize = 0;
  for (const entry of processedEntries) {
    localSectionSize += computeLocalRecordSize(entry);
    centralDirSize += computeCentralDirHeaderSize(entry);
  }

  // The central directory should start immediately after local section.
  const centralDirOffset = localSectionSize;

  const totalSize =
    localSectionSize + centralDirSize + ZIP_END_OF_CENTRAL_DIR_FIXED_SIZE + zipComment.length;
  const out = new Uint8Array(totalSize);
  const view = new DataView(out.buffer, out.byteOffset, out.byteLength);

  let offset = 0;

  // Local file headers and data
  for (const entry of processedEntries) {
    offset += writeLocalFileHeaderInto(out, view, offset, {
      fileName: entry.name,
      extraField: entry.extraField,
      flags: FLAG_UTF8,
      compressionMethod: entry.compressionMethod,
      dosTime: entry.modTime,
      dosDate: entry.modDate,
      crc32: entry.crc,
      compressedSize: entry.compressedData.length,
      uncompressedSize: entry.uncompressedSize
    });

    out.set(entry.compressedData, offset);
    offset += entry.compressedData.length;
  }

  // Central directory headers
  for (const entry of processedEntries) {
    offset += writeCentralDirectoryHeaderInto(out, view, offset, {
      fileName: entry.name,
      extraField: entry.extraField,
      comment: entry.comment,
      flags: FLAG_UTF8,
      compressionMethod: entry.compressionMethod,
      dosTime: entry.modTime,
      dosDate: entry.modDate,
      crc32: entry.crc,
      compressedSize: entry.compressedData.length,
      uncompressedSize: entry.uncompressedSize,
      localHeaderOffset: entry.offset
    });
  }

  // End of central directory
  writeEndOfCentralDirectoryInto(out, view, offset, {
    entryCount: processedEntries.length,
    centralDirSize,
    centralDirOffset,
    comment: zipComment
  });

  return out;
}

/**
 * Create a ZIP file from entries (async)
 */
export async function createZip(
  entries: ZipEntry[],
  options: ZipOptions = {}
): Promise<Uint8Array> {
  const reproducible = options.reproducible ?? false;
  const noSort = options.noSort ?? false;
  const level = options.level ?? DEFAULT_ZIP_LEVEL;
  const smartStore = options.smartStore ?? true;
  const concurrency = options.concurrency ?? 4;
  const timestamps: ZipTimestampMode =
    options.timestamps ?? (reproducible ? "dos" : DEFAULT_ZIP_TIMESTAMPS);
  const zipComment = encodeZipComment(options.comment);
  const defaultModTime = options.modTime ?? (reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date());

  const settings: ZipBuildSettings = {
    level,
    timestamps,
    defaultModTime
  };

  const thresholdBytes = options.thresholdBytes;

  const limit = Math.max(1, Math.floor(concurrency));
  const processedEntries = new Array<ProcessedEntry>(entries.length);

  if (entries.length > 0) {
    let nextIndex = 0;
    const workerCount = Math.min(limit, entries.length);

    const workers = Array.from({ length: workerCount }, async () => {
      while (true) {
        const idx = nextIndex++;
        if (idx >= entries.length) {
          return;
        }
        const entry = entries[idx]!;
        const entryLevel = entry.level ?? level;
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
        processedEntries[idx] = buildProcessedEntry(entry, 0, settings, compressedData, deflate);
      }
    });

    await Promise.all(workers);
  }

  if (!noSort) {
    sortEntriesByName(processedEntries);
  }

  // Compute offsets after sorting.
  let currentOffset = 0;
  for (let i = 0; i < processedEntries.length; i++) {
    const processedEntry = processedEntries[i]!;
    processedEntry.offset = currentOffset;
    currentOffset += computeLocalRecordSize(processedEntry);
  }
  return finalizeZip(processedEntries, zipComment);
}

/**
 * Create a ZIP file from entries (sync)
 *
 * This is supported in both Node.js and browser builds.
 */
export function createZipSync(entries: ZipEntry[], options: ZipOptions = {}): Uint8Array {
  const reproducible = options.reproducible ?? false;
  const noSort = options.noSort ?? false;
  const level = options.level ?? DEFAULT_ZIP_LEVEL;
  const smartStore = options.smartStore ?? true;
  const timestamps: ZipTimestampMode =
    options.timestamps ?? (reproducible ? "dos" : DEFAULT_ZIP_TIMESTAMPS);
  const zipComment = encodeZipComment(options.comment);
  const defaultModTime = options.modTime ?? (reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date());

  const settings: ZipBuildSettings = {
    level,
    timestamps,
    defaultModTime
  };

  const thresholdBytes = options.thresholdBytes;

  const processedEntries: ProcessedEntry[] = [];

  for (const entry of entries) {
    const entryLevel = entry.level ?? level;
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

    processedEntries.push(buildProcessedEntry(entry, 0, settings, compressedData, deflate));
  }

  if (!noSort) {
    sortEntriesByName(processedEntries);
  }

  // Compute offsets after sorting.
  let currentOffset = 0;
  for (const processedEntry of processedEntries) {
    processedEntry.offset = currentOffset;
    currentOffset += computeLocalRecordSize(processedEntry);
  }
  return finalizeZip(processedEntries, zipComment);
}
