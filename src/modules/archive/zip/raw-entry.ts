import type { ZipStringEncoding } from "@archive/core/text";

/**
 * Shared types for raw ZIP entries.
 *
 * A "raw entry" is a pre-compressed (and optionally encrypted) ZIP entry payload.
 * This is used for passthrough operations where we want to preserve the original
 * compressed bytes without re-compressing.
 *
 * @module
 */

/**
 * A pre-compressed (and optionally already encrypted) ZIP entry.
 *
 * This is primarily used by:
 * - High-level editors/patchers to preserve existing entries without re-compressing
 * - Streaming passthrough operations
 *
 * @example
 * ```ts
 * // Preserve an entry from one ZIP to another
 * const rawData = parser.getRawCompressedData("file.txt");
 * const entry = parser.getEntry("file.txt")!;
 *
 * const raw: ZipRawEntry = {
 *   name: "file.txt",
 *   compressedData: rawData,
 *   crc32: entry.crc32,
 *   uncompressedSize: entry.uncompressedSize,
 *   compressionMethod: entry.compressionMethod,
 *   modTime: entry.lastModified,
 *   extraField: entry.extraField
 * };
 * ```
 */
export interface ZipRawEntry {
  /** File name (can include directory path, use forward slashes) */
  name: string;

  /** Raw entry payload as stored in the ZIP (compressed and/or encrypted). */
  compressedData: Uint8Array;

  /** CRC-32 of the uncompressed data */
  crc32: number;

  /** Uncompressed size */
  uncompressedSize: number;

  /** Compression method (e.g. STORE=0, DEFLATE=8, AES=99) */
  compressionMethod: number;

  /** Optional file modification time */
  modTime?: Date;

  /** Optional file comment */
  comment?: string;

  /** Optional string encoding for this entry name/comment. */
  encoding?: ZipStringEncoding;

  /** Optional extra field (raw bytes) */
  extraField?: Uint8Array;

  /** Optional general purpose flags override */
  flags?: number;

  /** Optional external file attributes */
  externalAttributes?: number;

  /** Optional version made by */
  versionMadeBy?: number;
}

/**
 * Check if a value is a ZipRawEntry.
 */
export function isZipRawEntry(entry: unknown): entry is ZipRawEntry {
  return (
    typeof entry === "object" &&
    entry !== null &&
    "compressedData" in entry &&
    (entry as { compressedData: unknown }).compressedData instanceof Uint8Array
  );
}
