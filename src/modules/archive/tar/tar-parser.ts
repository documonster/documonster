/**
 * TAR Parser
 *
 * Parses TAR archives from various input sources.
 * Supports POSIX ustar, GNU tar (long filenames), and PAX extensions.
 */

import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { ArchiveError, FileTooLargeError, createAbortError } from "@archive/core/errors";
import { TAR_BLOCK_SIZE, TAR_TYPE } from "@archive/tar/tar-constants";
import type { TarEntryInfo } from "@archive/tar/tar-entry-info";
import { decodeHeader, isZeroBlock, calculatePadding } from "@archive/tar/tar-header";
import { textDecoder } from "@utils/binary";

// Helper to strip trailing null characters without using control char regex
const NULL_CHAR = String.fromCharCode(0);
function stripTrailingNulls(str: string): string {
  let end = str.length;
  while (end > 0 && str[end - 1] === NULL_CHAR) {
    end--;
  }
  return str.slice(0, end);
}

/**
 * Apply PAX extended attributes to entry info
 */
function applyPaxAttributes(info: TarEntryInfo, pax: Record<string, string>): void {
  if (pax.path) {
    info.path = pax.path;
  }
  if (pax.linkpath) {
    info.linkname = pax.linkpath;
  }
  if (pax.size) {
    info.size = parseInt(pax.size, 10);
  }
  if (pax.mtime) {
    info.mtime = new Date(parseFloat(pax.mtime) * 1000);
  }
  if (pax.uid) {
    info.uid = parseInt(pax.uid, 10);
  }
  if (pax.gid) {
    info.gid = parseInt(pax.gid, 10);
  }
  if (pax.uname) {
    info.uname = pax.uname;
  }
  if (pax.gname) {
    info.gname = pax.gname;
  }
  info.pax = pax;
}

export interface TarParseOptions {
  /** Maximum file size to extract into memory (default: 100MB) */
  maxFileSize?: number;

  /** Abort signal for cancellation */
  signal?: AbortSignal;
}

export interface TarEntry {
  /** Entry metadata */
  info: TarEntryInfo;

  /** Extract entry data as Uint8Array */
  data(): Promise<Uint8Array>;

  /** Extract entry data as string (UTF-8) */
  text(): Promise<string>;

  /** Skip this entry without reading data */
  skip(): Promise<void>;
}

/**
 * Parse TAR archive from Uint8Array
 */
export function parseTar(data: Uint8Array, options: TarParseOptions = {}): TarEntry[] {
  const { maxFileSize = 100 * 1024 * 1024 } = options;
  const entries: TarEntry[] = [];

  let offset = 0;
  let pendingLongName: string | undefined;
  let pendingLongLink: string | undefined;
  let pendingPax: Record<string, string> | undefined;

  while (offset + TAR_BLOCK_SIZE <= data.length) {
    const headerBlock = data.subarray(offset, offset + TAR_BLOCK_SIZE);

    // Check for end of archive (two zero blocks)
    if (isZeroBlock(headerBlock)) {
      break;
    }

    const info = decodeHeader(headerBlock);
    if (!info) {
      break;
    }

    offset += TAR_BLOCK_SIZE;

    // Calculate data size and padding
    const dataSize = info.size;
    const padding = calculatePadding(dataSize);
    const totalSize = dataSize + padding;

    // Handle special entry types
    if (info.type === TAR_TYPE.GNU_LONG_NAME) {
      // GNU long filename - read and store for next entry
      pendingLongName = stripTrailingNulls(
        textDecoder.decode(data.subarray(offset, offset + dataSize))
      );
      offset += totalSize;
      continue;
    }

    if (info.type === TAR_TYPE.GNU_LONG_LINK) {
      // GNU long link name - read and store for next entry
      pendingLongLink = stripTrailingNulls(
        textDecoder.decode(data.subarray(offset, offset + dataSize))
      );
      offset += totalSize;
      continue;
    }

    if (info.type === TAR_TYPE.PAX_EXTENDED || info.type === TAR_TYPE.PAX_GLOBAL) {
      // PAX extended header - parse and store for next entry
      const paxData = textDecoder.decode(data.subarray(offset, offset + dataSize));
      pendingPax = parsePaxHeader(paxData);
      offset += totalSize;
      if (info.type === TAR_TYPE.PAX_GLOBAL) {
        // Global headers apply to all subsequent entries (not implemented yet)
        pendingPax = undefined;
      }
      continue;
    }

    // Apply pending long name/link
    if (pendingLongName) {
      info.path = pendingLongName;
      pendingLongName = undefined;
    }
    if (pendingLongLink) {
      info.linkname = pendingLongLink;
      pendingLongLink = undefined;
    }
    if (pendingPax) {
      applyPaxAttributes(info, pendingPax);
      pendingPax = undefined;
    }

    // Check file size limit
    if (dataSize > maxFileSize) {
      throw new FileTooLargeError(
        info.path,
        `exceeds maximum file size (${dataSize} > ${maxFileSize})`
      );
    }

    // Capture data range for this entry
    const dataStart = offset;
    const dataEnd = offset + dataSize;

    // Create entry object
    const entry: TarEntry = {
      info,
      async data(): Promise<Uint8Array> {
        return data.subarray(dataStart, dataEnd);
      },
      async text(): Promise<string> {
        return textDecoder.decode(data.subarray(dataStart, dataEnd));
      },
      async skip(): Promise<void> {
        // No-op for in-memory parsing
      }
    };

    entries.push(entry);
    offset += totalSize;
  }

  return entries;
}

/**
 * Parse TAR archive from async iterable (streaming)
 */
export async function* parseTarStream(
  source: AsyncIterable<Uint8Array>,
  options: TarParseOptions = {}
): AsyncGenerator<TarEntry & { data(): Promise<Uint8Array> }> {
  const { maxFileSize = 100 * 1024 * 1024, signal } = options;

  // Buffer for accumulating data
  const buffer: Uint8Array[] = [];
  let bufferHead = 0;
  let bufferSize = 0;
  let bufferOffset = 0;

  // State for special headers
  let pendingLongName: string | undefined;
  let pendingLongLink: string | undefined;
  let pendingPax: Record<string, string> | undefined;

  // Helper to read exactly n bytes from buffer
  async function* readFromSource(): AsyncGenerator<Uint8Array> {
    for await (const chunk of source) {
      if (signal?.aborted) {
        throw createAbortError();
      }
      yield chunk;
    }
  }

  const sourceIterator = readFromSource();

  async function ensureBuffer(needed: number): Promise<boolean> {
    while (bufferSize - bufferOffset < needed) {
      const { value, done } = await sourceIterator.next();
      if (done) {
        return false;
      }
      buffer.push(value);
      bufferSize += value.length;
    }
    return true;
  }

  function consumeCurrentChunk(chunkLength: number): void {
    if (bufferOffset >= chunkLength) {
      bufferHead += 1;
      bufferSize -= chunkLength;
      bufferOffset = 0;
      if (bufferHead > 32 && bufferHead * 2 >= buffer.length) {
        buffer.splice(0, bufferHead);
        bufferHead = 0;
      }
    }
  }

  function readBytes(count: number): Uint8Array {
    if (count === 0) {
      return EMPTY_UINT8ARRAY;
    }

    const first = buffer[bufferHead]!;
    const firstAvailable = first.length - bufferOffset;
    if (firstAvailable >= count) {
      const out = first.subarray(bufferOffset, bufferOffset + count);
      bufferOffset += count;
      consumeCurrentChunk(first.length);
      return out;
    }

    const result = new Uint8Array(count);
    let written = 0;

    while (written < count) {
      const chunk = buffer[bufferHead]!;
      const available = chunk.length - bufferOffset;
      const toRead = Math.min(available, count - written);

      result.set(chunk.subarray(bufferOffset, bufferOffset + toRead), written);
      written += toRead;
      bufferOffset += toRead;
      consumeCurrentChunk(chunk.length);
    }

    return result;
  }

  function skipBytes(count: number): void {
    let remaining = count;
    while (remaining > 0) {
      const chunk = buffer[bufferHead]!;
      const available = chunk.length - bufferOffset;
      const toSkip = Math.min(available, remaining);
      remaining -= toSkip;
      bufferOffset += toSkip;
      consumeCurrentChunk(chunk.length);
    }
  }

  // Main parsing loop
  while (true) {
    // Read header block
    if (!(await ensureBuffer(TAR_BLOCK_SIZE))) {
      break;
    }

    const headerBlock = readBytes(TAR_BLOCK_SIZE);

    // Check for end of archive
    if (isZeroBlock(headerBlock)) {
      break;
    }

    const info = decodeHeader(headerBlock);
    if (!info) {
      break;
    }

    // Calculate sizes
    const dataSize = info.size;
    const padding = calculatePadding(dataSize);

    // Handle special entry types
    if (info.type === TAR_TYPE.GNU_LONG_NAME) {
      if (!(await ensureBuffer(dataSize + padding))) {
        throw new ArchiveError("Unexpected end of TAR archive in long name");
      }
      const nameData = readBytes(dataSize);
      if (padding > 0) {
        skipBytes(padding);
      }
      pendingLongName = stripTrailingNulls(textDecoder.decode(nameData));
      continue;
    }

    if (info.type === TAR_TYPE.GNU_LONG_LINK) {
      if (!(await ensureBuffer(dataSize + padding))) {
        throw new ArchiveError("Unexpected end of TAR archive in long link");
      }
      const linkData = readBytes(dataSize);
      if (padding > 0) {
        skipBytes(padding);
      }
      pendingLongLink = stripTrailingNulls(textDecoder.decode(linkData));
      continue;
    }

    if (info.type === TAR_TYPE.PAX_EXTENDED || info.type === TAR_TYPE.PAX_GLOBAL) {
      if (!(await ensureBuffer(dataSize + padding))) {
        throw new ArchiveError("Unexpected end of TAR archive in PAX header");
      }
      const paxData = textDecoder.decode(readBytes(dataSize));
      if (padding > 0) {
        skipBytes(padding);
      }
      pendingPax = parsePaxHeader(paxData);
      if (info.type === TAR_TYPE.PAX_GLOBAL) {
        pendingPax = undefined;
      }
      continue;
    }

    // Apply pending metadata
    if (pendingLongName) {
      info.path = pendingLongName;
      pendingLongName = undefined;
    }
    if (pendingLongLink) {
      info.linkname = pendingLongLink;
      pendingLongLink = undefined;
    }
    if (pendingPax) {
      applyPaxAttributes(info, pendingPax);
      pendingPax = undefined;
    }

    // Check file size limit
    if (dataSize > maxFileSize) {
      throw new FileTooLargeError(info.path, `exceeds maximum file size`);
    }

    // Read entry data
    if (!(await ensureBuffer(dataSize + padding))) {
      throw new ArchiveError(`Unexpected end of TAR archive reading "${info.path}"`);
    }

    const entryData = readBytes(dataSize);
    if (padding > 0) {
      skipBytes(padding);
    }

    yield {
      info,
      async data(): Promise<Uint8Array> {
        return entryData;
      },
      async text(): Promise<string> {
        return textDecoder.decode(entryData);
      },
      async skip(): Promise<void> {
        // Data already read
      }
    };
  }
}

/**
 * Parse PAX extended header
 */
function parsePaxHeader(data: string): Record<string, string> {
  const result: Record<string, string> = {};
  let pos = 0;

  while (pos < data.length) {
    // Find the space after length
    const spaceIdx = data.indexOf(" ", pos);
    if (spaceIdx === -1) {
      break;
    }

    // Parse length
    const length = parseInt(data.substring(pos, spaceIdx), 10);
    if (isNaN(length) || length <= 0) {
      break;
    }

    // Extract the record (length includes the length field itself and newline)
    const record = data.substring(spaceIdx + 1, pos + length - 1);

    // Find the = separator
    const eqIdx = record.indexOf("=");
    if (eqIdx !== -1) {
      const key = record.substring(0, eqIdx);
      const value = record.substring(eqIdx + 1);
      result[key] = value;
    }

    pos += length;
  }

  return result;
}

/**
 * Parse TAR archive and return all entries with their data
 */
export async function untar(
  source: Uint8Array | AsyncIterable<Uint8Array>,
  options: TarParseOptions = {}
): Promise<Map<string, { info: TarEntryInfo; data: Uint8Array }>> {
  const result = new Map<string, { info: TarEntryInfo; data: Uint8Array }>();

  if (source instanceof Uint8Array) {
    const entries = parseTar(source, options);
    for (const entry of entries) {
      result.set(entry.info.path, {
        info: entry.info,
        data: await entry.data()
      });
    }
  } else {
    for await (const entry of parseTarStream(source, options)) {
      result.set(entry.info.path, {
        info: entry.info,
        data: await entry.data()
      });
    }
  }

  return result;
}
