/**
 * Pure Uint8Array-based ZIP parser
 * Works in both Node.js and browser environments
 * No dependency on Node.js stream module
 */

import { decompress, decompressSync } from "@archive/compression/compress";
import { BinaryReader } from "@archive/utils/binary";
import { resolveZipLastModifiedDateFromUnixSeconds } from "@archive/utils/timestamps";
import { parseZipExtraFields } from "@archive/utils/zip-extra-fields";
import type { ZipEntryInfo } from "@archive/zip-spec/zip-entry-info";
import {
  CENTRAL_DIR_HEADER_SIG,
  COMPRESSION_DEFLATE,
  COMPRESSION_STORE,
  LOCAL_FILE_HEADER_SIG,
  UINT16_MAX,
  UINT32_MAX,
  ZIP64_END_OF_CENTRAL_DIR_SIG
} from "@archive/zip-spec/zip-records";

export type { ZipEntryInfo };

/**
 * ZIP parsing options
 */
export interface ZipParseOptions {
  /** Whether to decode file names as UTF-8 (default: true) */
  decodeStrings?: boolean;
}

/**
 * Find the End of Central Directory record
 * Searches backwards from the end of the file
 */
function findEndOfCentralDir(data: Uint8Array): number {
  // EOCD is at least 22 bytes, search backwards
  // Comment can be up to 65535 bytes
  const minOffset = Math.max(0, data.length - 65557);

  // Signature bytes (little-endian): 0x06054b50 -> 50 4b 05 06
  const b0 = 0x50;
  const b1 = 0x4b;
  const b2 = 0x05;
  const b3 = 0x06;

  for (let i = data.length - 22; i >= minOffset; i--) {
    if (data[i] === b0 && data[i + 1] === b1 && data[i + 2] === b2 && data[i + 3] === b3) {
      return i;
    }
  }

  return -1;
}

/**
 * Find ZIP64 End of Central Directory Locator
 */
function findZip64EOCDLocator(data: Uint8Array, eocdOffset: number): number {
  // ZIP64 EOCD Locator is 20 bytes and appears right before EOCD
  const locatorOffset = eocdOffset - 20;
  if (locatorOffset < 0) {
    return -1;
  }

  // Signature bytes (little-endian): 0x07064b50 -> 50 4b 06 07
  if (
    data[locatorOffset] === 0x50 &&
    data[locatorOffset + 1] === 0x4b &&
    data[locatorOffset + 2] === 0x06 &&
    data[locatorOffset + 3] === 0x07
  ) {
    return locatorOffset;
  }

  return -1;
}

/**
 * Parse ZIP file entries from Central Directory
 */
function parseZipEntries(data: Uint8Array, options: ZipParseOptions = {}): ZipEntryInfo[] {
  const { decodeStrings = true } = options;
  let entries: ZipEntryInfo[] = [];

  // Find End of Central Directory
  const eocdOffset = findEndOfCentralDir(data);
  if (eocdOffset === -1) {
    throw new Error("Invalid ZIP file: End of Central Directory not found");
  }

  const reader = new BinaryReader(data, eocdOffset);

  // Read EOCD
  // Offset  Size  Description
  // 0       4     EOCD signature (0x06054b50)
  // 4       2     Number of this disk
  // 6       2     Disk where central directory starts
  // 8       2     Number of central directory records on this disk
  // 10      2     Total number of central directory records
  // 12      4     Size of central directory (bytes)
  // 16      4     Offset of start of central directory
  // 20      2     Comment length
  reader.skip(4); // signature
  reader.skip(2); // disk number
  reader.skip(2); // disk where central dir starts
  reader.skip(2); // entries on this disk
  let totalEntries = reader.readUint16(); // total entries
  reader.skip(4); // central directory size (unused)
  let centralDirOffset = reader.readUint32();

  // Check for ZIP64
  const zip64LocatorOffset = findZip64EOCDLocator(data, eocdOffset);
  if (zip64LocatorOffset !== -1) {
    const locatorReader = new BinaryReader(data, zip64LocatorOffset);
    locatorReader.skip(4); // signature
    locatorReader.skip(4); // disk number with ZIP64 EOCD
    const zip64EOCDOffset = Number(locatorReader.readBigUint64());

    // Read ZIP64 EOCD
    const zip64Reader = new BinaryReader(data, zip64EOCDOffset);
    const zip64Sig = zip64Reader.readUint32();
    if (zip64Sig === ZIP64_END_OF_CENTRAL_DIR_SIG) {
      zip64Reader.skip(8); // size of ZIP64 EOCD
      zip64Reader.skip(2); // version made by
      zip64Reader.skip(2); // version needed
      zip64Reader.skip(4); // disk number
      zip64Reader.skip(4); // disk with central dir
      const zip64TotalEntries = Number(zip64Reader.readBigUint64());
      zip64Reader.skip(8); // central directory size (unused)
      const zip64CentralDirOffset = Number(zip64Reader.readBigUint64());

      // Use ZIP64 values if standard values are maxed out
      if (totalEntries === UINT16_MAX) {
        totalEntries = zip64TotalEntries;
      }
      if (centralDirOffset === UINT32_MAX) {
        centralDirOffset = zip64CentralDirOffset;
      }
    }
  }

  // Preallocate to avoid repeated array growth on large archives.
  if (totalEntries > 0) {
    entries = new Array(totalEntries);
  }

  // Read Central Directory entries
  const centralReader = new BinaryReader(data, centralDirOffset);

  for (let i = 0; i < totalEntries; i++) {
    const sig = centralReader.readUint32();
    if (sig !== CENTRAL_DIR_HEADER_SIG) {
      throw new Error(`Invalid Central Directory header signature at entry ${i}`);
    }

    // Central Directory File Header format:
    // Offset  Size  Description
    // 0       4     Central directory file header signature (0x02014b50)
    // 4       2     Version made by
    // 6       2     Version needed to extract
    // 8       2     General purpose bit flag
    // 10      2     Compression method
    // 12      2     File last modification time
    // 14      2     File last modification date
    // 16      4     CRC-32
    // 20      4     Compressed size
    // 24      4     Uncompressed size
    // 28      2     File name length
    // 30      2     Extra field length
    // 32      2     File comment length
    // 34      2     Disk number where file starts
    // 36      2     Internal file attributes
    // 38      4     External file attributes
    // 42      4     Relative offset of local file header
    // 46      n     File name
    // 46+n    m     Extra field
    // 46+n+m  k     File comment

    centralReader.skip(2); // version made by
    centralReader.skip(2); // version needed
    const flags = centralReader.readUint16();
    const compressionMethod = centralReader.readUint16();
    const lastModTime = centralReader.readUint16();
    const lastModDate = centralReader.readUint16();
    const crc32 = centralReader.readUint32();
    let compressedSize = centralReader.readUint32();
    let uncompressedSize = centralReader.readUint32();
    const fileNameLength = centralReader.readUint16();
    const extraFieldLength = centralReader.readUint16();
    const commentLength = centralReader.readUint16();
    centralReader.skip(2); // disk number start
    centralReader.skip(2); // internal attributes
    const externalAttributes = centralReader.readUint32();
    let localHeaderOffset = centralReader.readUint32();

    // Check for UTF-8 flag (bit 11)
    const isUtf8 = (flags & 0x800) !== 0;
    const useUtf8 = decodeStrings && isUtf8;

    const fileName = fileNameLength > 0 ? centralReader.readString(fileNameLength, useUtf8) : "";

    let extraFields = {} as ReturnType<typeof parseZipExtraFields>;
    if (extraFieldLength > 0) {
      const extraField = centralReader.readBytes(extraFieldLength);
      const vars = {
        compressedSize,
        uncompressedSize,
        offsetToLocalFileHeader: localHeaderOffset
      };
      extraFields = parseZipExtraFields(extraField, vars);

      compressedSize = vars.compressedSize;
      uncompressedSize = vars.uncompressedSize;
      localHeaderOffset = vars.offsetToLocalFileHeader ?? localHeaderOffset;
    }

    const comment = commentLength > 0 ? centralReader.readString(commentLength, useUtf8) : "";

    const isDirectory = fileName.endsWith("/") || (externalAttributes & 0x10) !== 0;
    const isEncrypted = (flags & 0x01) !== 0;

    const unixSecondsMtime = extraFields.mtimeUnixSeconds;
    const lastModified = resolveZipLastModifiedDateFromUnixSeconds(
      lastModDate,
      lastModTime,
      unixSecondsMtime
    );

    entries[i] = {
      path: fileName,
      isDirectory,
      compressedSize,
      uncompressedSize,
      compressionMethod,
      crc32,
      lastModified,
      localHeaderOffset,
      comment,
      externalAttributes,
      isEncrypted
    };
  }

  return entries;
}

/**
 * Extract file data for a specific entry
 */
async function extractEntryData(data: Uint8Array, entry: ZipEntryInfo): Promise<Uint8Array> {
  if (entry.isDirectory) {
    return new Uint8Array(0);
  }

  if (entry.isEncrypted) {
    throw new Error(`File "${entry.path}" is encrypted and cannot be extracted`);
  }

  const compressedData = readEntryCompressedData(data, entry);

  if (entry.compressionMethod === COMPRESSION_STORE) {
    return compressedData;
  }
  if (entry.compressionMethod === COMPRESSION_DEFLATE) {
    return decompress(compressedData);
  }
  throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
}

/**
 * Extract file data synchronously
 */
function extractEntryDataSync(data: Uint8Array, entry: ZipEntryInfo): Uint8Array {
  if (entry.isDirectory) {
    return new Uint8Array(0);
  }

  if (entry.isEncrypted) {
    throw new Error(`File "${entry.path}" is encrypted and cannot be extracted`);
  }

  const compressedData = readEntryCompressedData(data, entry);

  if (entry.compressionMethod === COMPRESSION_STORE) {
    return compressedData;
  }
  if (entry.compressionMethod === COMPRESSION_DEFLATE) {
    return decompressSync(compressedData);
  }
  throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
}

function readEntryCompressedData(data: Uint8Array, entry: ZipEntryInfo): Uint8Array {
  const reader = new BinaryReader(data, entry.localHeaderOffset);

  const sig = reader.readUint32();
  if (sig !== LOCAL_FILE_HEADER_SIG) {
    throw new Error(`Invalid local file header signature for "${entry.path}"`);
  }

  reader.skip(2); // version needed
  reader.skip(2); // flags
  reader.skip(2); // compression method
  reader.skip(2); // last mod time
  reader.skip(2); // last mod date
  reader.skip(4); // crc32
  reader.skip(4); // compressed size
  reader.skip(4); // uncompressed size
  const fileNameLength = reader.readUint16();
  const extraFieldLength = reader.readUint16();

  reader.skip(fileNameLength);
  reader.skip(extraFieldLength);

  return reader.readBytes(entry.compressedSize);
}

/**
 * High-level ZIP parser class
 */
export class ZipParser {
  private data: Uint8Array;
  private entries: ZipEntryInfo[];
  private entryMap: Map<string, ZipEntryInfo>;

  constructor(data: Uint8Array | ArrayBuffer, options: ZipParseOptions = {}) {
    this.data = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    this.entries = parseZipEntries(this.data, options);
    this.entryMap = new Map(this.entries.map(e => [e.path, e]));
  }

  /**
   * Get all entries in the ZIP file
   */
  getEntries(): ZipEntryInfo[] {
    return this.entries;
  }

  /**
   * Get entry by path
   */
  getEntry(path: string): ZipEntryInfo | undefined {
    return this.entryMap.get(path);
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
  getChildCount(path: string): number {
    const entry = this.entryMap.get(path);
    if (!entry?.isDirectory) {
      return 0;
    }
    // Ensure prefix ends with "/" for correct matching
    const prefix = entry.path.endsWith("/") ? entry.path : entry.path + "/";
    let count = 0;
    for (const e of this.entries) {
      if (e.path.startsWith(prefix) && e !== entry) {
        count++;
      }
    }
    return count;
  }

  /**
   * List all file paths
   */
  listFiles(): string[] {
    return this.entries.map(e => e.path);
  }

  /**
   * Extract a single file (async)
   */
  async extract(path: string): Promise<Uint8Array | null> {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryData(this.data, entry);
  }

  /**
   * Extract a single file (sync)
   *
   * This is supported in both Node.js and browser builds.
   */
  extractSync(path: string): Uint8Array | null {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryDataSync(this.data, entry);
  }

  /**
   * Extract all files (async)
   */
  async extractAll(): Promise<Map<string, Uint8Array>> {
    const result = new Map<string, Uint8Array>();
    for (const entry of this.entries) {
      const data = await extractEntryData(this.data, entry);
      result.set(entry.path, data);
    }
    return result;
  }

  /**
   * Extract all files (sync)
   * Returns object with file paths as keys and Uint8Array content as values
   */
  extractAllSync(): Record<string, Uint8Array> {
    const result: Record<string, Uint8Array> = {};
    for (const entry of this.entries) {
      result[entry.path] = extractEntryDataSync(this.data, entry);
    }
    return result;
  }

  /**
   * Iterate over entries with async callback
   */
  async forEach(
    callback: (entry: ZipEntryInfo, getData: () => Promise<Uint8Array>) => Promise<boolean | void>
  ): Promise<void> {
    for (const entry of this.entries) {
      let dataPromise: Promise<Uint8Array> | null = null;
      const getData = () => {
        if (!dataPromise) {
          dataPromise = extractEntryData(this.data, entry);
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
