/**
 * Pure Uint8Array-based ZIP parser
 * Works in both Node.js and browser environments
 * No dependency on Node.js stream module
 */

import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { ArchiveLimitError, FileTooLargeError } from "@archive/core/errors";
import type { ZipStringEncoding } from "@archive/core/text";
import {
  processEntryData,
  processEntryDataSync,
  readEntryCompressedData
} from "@archive/unzip/zip-extract-core";
import type { ZipEntryRecord } from "@archive/zip-spec/zip-entry-info";
import { parseZipArchiveFromBuffer } from "@archive/zip-spec/zip-parser-core";
import { assertLimitOption } from "@utils/limits";

const MAX_SAFE_INTEGER_BIGINT = BigInt(Number.MAX_SAFE_INTEGER);

function assertEntryExtractableInMemory(entry: ZipEntryRecord): void {
  // This parser extracts into memory. If ZIP64 values exceed JS safe integers,
  // callers need a random-access + streaming extraction path (not implemented here).
  if (
    entry.uncompressedSize64 !== undefined &&
    entry.uncompressedSize64 > MAX_SAFE_INTEGER_BIGINT
  ) {
    throw new FileTooLargeError(entry.path, "ZIP64 size > 2^53-1");
  }
  if (entry.compressedSize64 !== undefined && entry.compressedSize64 > MAX_SAFE_INTEGER_BIGINT) {
    throw new FileTooLargeError(entry.path, "ZIP64 size > 2^53-1");
  }
  if (
    entry.localHeaderOffset64 !== undefined &&
    entry.localHeaderOffset64 > MAX_SAFE_INTEGER_BIGINT
  ) {
    throw new FileTooLargeError(entry.path, "ZIP64 offset > 2^53-1");
  }
}

export type { ZipEntryRecord } from "@archive/zip-spec/zip-entry-info";

/**
 * ZIP parsing options
 */
export interface ZipParseOptions {
  /** Whether to decode file names as UTF-8 (default: true) */
  decodeStrings?: boolean;

  /** Optional string encoding for legacy (non-UTF8) names/comments. */
  encoding?: ZipStringEncoding;

  /** Password for encrypted entries */
  password?: string | Uint8Array;

  /**
   * Largest uncompressed size, in bytes, of a single entry extracted into memory.
   * An entry declaring more is rejected before inflation; an entry that inflates
   * past its declared size is stopped at that size (`EntrySizeMismatchError`).
   * @default 536870912 (512 MiB)
   */
  maxEntrySize?: number;

  /**
   * Largest combined uncompressed size, in bytes, that `extractAll` /
   * `extractAllSync` / `forEach` may produce (`ArchiveLimitError`). Each entry is
   * already bounded by `maxEntrySize`; this bounds many entries together.
   *
   * Unbounded by default: what a process can afford to hold depends on the
   * host, so no fixed default would be a real guarantee. Set it when reading
   * untrusted archives.
   */
  maxTotalUncompressedSize?: number;

  /**
   * Most entries the archive may contain (`ArchiveLimitError`, limit
   * `"maxEntries"`), checked before entry records are allocated. Unbounded by
   * default; a 46-byte-per-record sanity check only rules out impossible counts.
   */
  maxEntries?: number;
}

/**
 * Result of parsing a ZIP archive.
 */
interface ZipArchiveParseResult {
  entries: ZipEntryRecord[];
  comment: string;
}

/**
 * Parse ZIP archive including entries and archive comment.
 */
function parseZipArchive(data: Uint8Array, options: ZipParseOptions = {}): ZipArchiveParseResult {
  return parseZipArchiveFromBuffer(data, {
    decodeStrings: options.decodeStrings,
    encoding: options.encoding,
    maxEntries: options.maxEntries
  });
}

/**
 * Extraction options with optional password support.
 */
export interface ExtractOptions {
  /** Password for encrypted entries */
  password?: string | Uint8Array;
  /** See {@link ZipParseOptions.maxEntrySize}. */
  maxEntrySize?: number;
}

/**
 * Extract file data for a specific entry (async)
 */
async function extractEntryData(
  data: Uint8Array,
  entry: ZipEntryRecord,
  options: ExtractOptions = {}
): Promise<Uint8Array> {
  if (entry.type === "directory") {
    return EMPTY_UINT8ARRAY;
  }

  assertEntryExtractableInMemory(entry);

  const compressedData = readEntryCompressedData(data, entry);
  return processEntryData(
    entry,
    compressedData,
    options.password,
    false,
    true,
    options.maxEntrySize
  );
}

/**
 * Extract file data synchronously (only supports ZipCrypto, not AES)
 */
function extractEntryDataSync(
  data: Uint8Array,
  entry: ZipEntryRecord,
  options: ExtractOptions = {}
): Uint8Array {
  if (entry.type === "directory") {
    return EMPTY_UINT8ARRAY;
  }

  assertEntryExtractableInMemory(entry);

  const compressedData = readEntryCompressedData(data, entry);
  return processEntryDataSync(entry, compressedData, options.password, true, options.maxEntrySize);
}

/**
 * High-level ZIP parser class
 */
export class ZipParser {
  private data: Uint8Array;
  private entries: ZipEntryRecord[];
  private entryMap: Map<string, ZipEntryRecord>;
  private password?: string | Uint8Array;
  private archiveComment: string;
  private maxEntrySize: number | undefined;
  private maxTotalUncompressedSize: number;

  constructor(data: Uint8Array | ArrayBuffer, options: ZipParseOptions = {}) {
    assertLimitOption("maxEntrySize", options.maxEntrySize);
    assertLimitOption("maxTotalUncompressedSize", options.maxTotalUncompressedSize);
    this.data = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    const result = parseZipArchive(this.data, options);
    this.entries = result.entries;
    this.archiveComment = result.comment;
    this.entryMap = new Map(this.entries.map(e => [e.path, e]));
    this.password = options.password;
    this.maxEntrySize = options.maxEntrySize;
    this.maxTotalUncompressedSize = options.maxTotalUncompressedSize ?? Infinity;
  }

  /**
   * Reject a bulk extraction whose declared sizes exceed the total limit.
   * Each entry is inflated under a bound of its declared size, so the declared
   * sum is an upper bound on what extraction can actually produce.
   */
  private assertTotalWithinLimit(): void {
    let total = 0;
    for (const entry of this.entries) {
      if (entry.type !== "directory") {
        total += entry.uncompressedSize;
      }
    }
    if (total > this.maxTotalUncompressedSize) {
      throw new ArchiveLimitError("maxTotalUncompressedSize", this.maxTotalUncompressedSize);
    }
  }

  /**
   * Set the password for encrypted entries.
   */
  setPassword(password: string | Uint8Array | undefined): void {
    this.password = password;
  }

  /**
   * Get all entries in the ZIP file
   */
  getEntries(): ZipEntryRecord[] {
    return this.entries;
  }

  /**
   * Get entry by path
   */
  getEntry(path: string): ZipEntryRecord | undefined {
    return this.entryMap.get(path);
  }

  /**
   * Get a zero-copy view of the raw (compressed) entry payload.
   *
   * Notes:
   * - This returns the bytes as stored in the ZIP (compressed and possibly encrypted).
   * - The returned Uint8Array is a view into the original ZIP buffer.
   * - This does NOT include the local file header, extra field, or data descriptor.
   */
  getRawCompressedData(path: string): Uint8Array | null {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    // This helper returns a subarray view into the original `data`.
    return readEntryCompressedData(this.data, entry);
  }

  /**
   * Get raw (compressed) payload together with its parsed entry info.
   */
  getRawEntry(path: string): { info: ZipEntryRecord; compressedData: Uint8Array } | null {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return { info: entry, compressedData: readEntryCompressedData(this.data, entry) };
  }

  /**
   * Check if entry exists
   */
  hasEntry(path: string): boolean {
    return this.entryMap.has(path);
  }

  /**
   * Get the number of child entries in a directory.
   *
   * Returns the count of entries whose paths start with the directory prefix,
   * excluding the directory entry itself. For non-directory entries, returns 0.
   *
   * @param path - Directory path (with or without trailing slash)
   * @returns Number of child entries
   */
  childCount(path: string): number {
    const direct = this.entryMap.get(path);
    if (direct && direct.type !== "directory") {
      return 0;
    }

    const slashPath = path.endsWith("/") ? path : path + "/";
    const dirEntry = direct?.type === "directory" ? direct : this.entryMap.get(slashPath);

    // If there is no explicit directory entry, still support implicit directories
    // as long as there are entries under the prefix.
    const prefix = (dirEntry?.path ?? slashPath).endsWith("/")
      ? (dirEntry?.path ?? slashPath)
      : (dirEntry?.path ?? slashPath) + "/";

    let count = 0;
    for (const e of this.entries) {
      if (e.path.startsWith(prefix) && e.path !== prefix) {
        count++;
      }
    }
    return count;
  }

  /**
   * Get the archive comment.
   */
  getZipComment(): string {
    return this.archiveComment;
  }

  /**
   * Check if the archive contains encrypted entries
   */
  hasEncryptedEntries(): boolean {
    return this.entries.some(e => e.isEncrypted);
  }

  /**
   * Get all encrypted entries
   */
  getEncryptedEntries(): ZipEntryRecord[] {
    return this.entries.filter(e => e.isEncrypted);
  }

  /**
   * List all file paths
   */
  listFiles(): string[] {
    return this.entries.map(e => e.path);
  }

  /**
   * Extract a single file (async)
   * @param path - File path within the archive
   * @param password - Optional password for this entry (overrides constructor password)
   */
  async extract(path: string, password?: string | Uint8Array): Promise<Uint8Array | null> {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryData(this.data, entry, {
      password: password ?? this.password,
      maxEntrySize: this.maxEntrySize
    });
  }

  /**
   * Extract a single file (sync)
   *
   * Note: AES-encrypted files cannot be extracted synchronously.
   * Use the async extract() method for AES-encrypted files.
   *
   * @param path - File path within the archive
   * @param password - Optional password for this entry (overrides constructor password)
   */
  extractSync(path: string, password?: string | Uint8Array): Uint8Array | null {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryDataSync(this.data, entry, {
      password: password ?? this.password,
      maxEntrySize: this.maxEntrySize
    });
  }

  /**
   * Extract all files (async)
   * @param password - Optional password for encrypted entries (overrides constructor password)
   */
  async extractAll(password?: string | Uint8Array): Promise<Map<string, Uint8Array>> {
    const result = new Map<string, Uint8Array>();
    const pw = password ?? this.password;
    this.assertTotalWithinLimit();
    for (const entry of this.entries) {
      const data = await extractEntryData(this.data, entry, {
        password: pw,
        maxEntrySize: this.maxEntrySize
      });
      result.set(entry.path, data);
    }
    return result;
  }

  /**
   * Extract all files (sync)
   * Returns object with file paths as keys and Uint8Array content as values
   *
   * Note: AES-encrypted files cannot be extracted synchronously.
   * Use the async extractAll() method if the archive contains AES-encrypted files.
   *
   * @param password - Optional password for encrypted entries (overrides constructor password)
   */
  extractAllSync(password?: string | Uint8Array): Record<string, Uint8Array> {
    const result: Record<string, Uint8Array> = {};
    const pw = password ?? this.password;
    this.assertTotalWithinLimit();
    for (const entry of this.entries) {
      result[entry.path] = extractEntryDataSync(this.data, entry, {
        password: pw,
        maxEntrySize: this.maxEntrySize
      });
    }
    return result;
  }

  /**
   * Iterate over entries with async callback
   * @param callback - Callback for each entry
   * @param password - Optional password for encrypted entries (overrides constructor password)
   */
  async forEach(
    callback: (
      entry: ZipEntryRecord,
      getData: () => Promise<Uint8Array>
    ) => Promise<boolean | void>,
    password?: string | Uint8Array
  ): Promise<void> {
    const pw = password ?? this.password;
    let produced = 0;
    for (const entry of this.entries) {
      let dataPromise: Promise<Uint8Array> | null = null;
      const getData = () => {
        if (!dataPromise) {
          dataPromise = extractEntryData(this.data, entry, {
            password: pw,
            maxEntrySize: this.maxEntrySize
          }).then(bytes => {
            produced += bytes.length;
            if (produced > this.maxTotalUncompressedSize) {
              throw new ArchiveLimitError(
                "maxTotalUncompressedSize",
                this.maxTotalUncompressedSize
              );
            }
            return bytes;
          });
        }
        return dataPromise;
      };

      const shouldContinue = await callback(entry, getData);
      if (shouldContinue === false) {
        break;
      }
    }
  }
}
