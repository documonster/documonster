/**
 * TAR Parser
 *
 * Parses TAR archives from various input sources.
 * Supports POSIX ustar, GNU tar (long filenames), and PAX extensions.
 */

import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { checkReadOptions, TAR_READ_LIMITS } from "@archive/core/entry-checks";
import {
  ArchiveArgumentError,
  ArchiveError,
  ArchiveFileTooLargeError,
  createAbortError
} from "@archive/core/errors";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  archiveBytesToUint8Array,
  isArchiveBytes,
  isArchiveSource,
  isInMemoryArchiveSource,
  iterateArchiveSource,
  toUint8Array
} from "@archive/io/archive-source";
import { TAR_BLOCK_SIZE, TAR_TYPE } from "@archive/tar/tar-constants";
import type { TarEntryInfo } from "@archive/tar/tar-entry-info";
import { decodeHeader, isZeroBlock, calculatePadding } from "@archive/tar/tar-header";
import { isReadableStream } from "@stream/core/type-guards";
import { textDecoder, asPlainBytes } from "@utils/binary";
import { describeValue as describeSource } from "@utils/describe-value";

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
  /**
   * Largest entry, in bytes, that may be read into memory (`ArchiveFileTooLargeError`). The same
   * option as ZIP's `maxEntrySize`.
   * @default 104857600 (100 MiB)
   */
  maxEntrySize?: number;

  /** @deprecated Use {@link maxEntrySize}. Removed in 1.0. */
  maxFileSize?: number;

  /** Abort signal for cancellation */
  signal?: AbortSignal;
}

/** The largest entry a parse reads into memory: `maxEntrySize`, else the deprecated `maxFileSize`. */
export function entrySizeLimit(options: {
  readonly maxEntrySize?: number;
  readonly maxFileSize?: number;
}): number {
  return options.maxEntrySize ?? options.maxFileSize ?? DEFAULT_MAX_TAR_ENTRY_SIZE;
}

const DEFAULT_MAX_TAR_ENTRY_SIZE = 100 * 1024 * 1024;

/**
 * One entry of {@link parseTar} / {@link parseTarStream} (and their tar.gz forms). Its reading
 * methods carry the names `TarReaderEntry` and `UnzipEntry` use.
 */
export interface TarEntry {
  /** Entry metadata */
  info: TarEntryInfo;

  /** The entry's data. */
  bytes(): Promise<Uint8Array>;

  /** The entry's data as a string (UTF-8). */
  text(): Promise<string>;

  /**
   * Let go of this entry without reading it. A parse reads each entry's data as it reaches
   * it, so there is nothing left to skip; this exists so code written against
   * `TarReaderEntry` / `UnzipEntry` works unchanged.
   */
  discard(): void;

  /** @deprecated Use {@link TarEntry.bytes}. Removed in 1.0. */
  data(): Promise<Uint8Array>;

  /** @deprecated Use {@link TarEntry.discard}. Removed in 1.0. */
  skip(): Promise<void>;
}

/** One entry of {@link untar} / {@link untarSync} (and their tar.gz forms): its header and its data. */
export interface UnpackedTarEntry {
  info: TarEntryInfo;
  data: Uint8Array<ArrayBuffer>;
}

/** A parsed entry over data already read. */
function createTarEntry(info: TarEntryInfo, entryData: Uint8Array): TarEntry {
  return {
    info,
    async bytes(): Promise<Uint8Array> {
      return entryData;
    },
    async text(): Promise<string> {
      return textDecoder.decode(entryData);
    },
    discard(): void {
      // Data already read
    },
    async data(): Promise<Uint8Array> {
      return entryData;
    },
    async skip(): Promise<void> {
      // Data already read
    }
  };
}

/**
 * Parse TAR archive from bytes in memory (any `ArrayBufferView` or an `ArrayBuffer`).
 * Synchronous; the entries' readers return promises so they match `parseTarStream`'s. An
 * archive cut short is an `ArchiveError`.
 */
export function parseTar(data: ArchiveBytes, options: TarParseOptions = {}): TarEntry[] {
  if (!isArchiveBytes(data)) {
    throw new ArchiveArgumentError(
      "data",
      `Archive.parseTar: data must be bytes (an ArrayBufferView or ArrayBuffer), got ${describeSource(data)}; ` +
        "use Archive.parseTarStream, Archive.unpackTar or Archive.openTar for other sources"
    );
  }
  checkReadOptions("Archive.parseTar", options, TAR_READ_LIMITS);
  return parseTarBytes(archiveBytesToUint8Array(data), options);
}

function parseTarBytes(data: Uint8Array, options: TarParseOptions): TarEntry[] {
  return parseTarRecords(data, options).map(({ info, start, end }) =>
    createTarEntry(info, data.subarray(start, end))
  );
}

/** One entry located by {@link parseTarRecords}: its header and the byte range of its data. */
interface TarRecord {
  info: TarEntryInfo;
  start: number;
  end: number;
}

/** A gzip stream starts 0x1f 0x8b; no TAR header does (that would be a control character name). */
export function refuseGzipTar(firstBlock: Uint8Array): void {
  if (firstBlock[0] === 0x1f && firstBlock[1] === 0x8b) {
    throw new ArchiveError(
      "TAR data is gzip-compressed; read it with unpackTarGz, parseTarGz or openTarGz"
    );
  }
}

/** Bytes left over after the last whole block: an error unless they are padding (all zero). */
function refusePartialBlock(rest: Uint8Array): void {
  if (rest.some(byte => byte !== 0)) {
    throw new ArchiveError("Unexpected end of TAR archive inside a header");
  }
}

/** Walk an in-memory TAR and locate every data entry, resolving GNU/PAX headers. */
function parseTarRecords(data: Uint8Array, options: TarParseOptions): TarRecord[] {
  const maxFileSize = entrySizeLimit(options);
  const entries: TarRecord[] = [];
  refuseGzipTar(data);
  let ended = false;

  let offset = 0;
  let pendingLongName: string | undefined;
  let pendingLongLink: string | undefined;
  let pendingPax: Record<string, string> | undefined;

  while (offset + TAR_BLOCK_SIZE <= data.length) {
    const headerBlock = data.subarray(offset, offset + TAR_BLOCK_SIZE);

    // Check for end of archive (two zero blocks)
    if (isZeroBlock(headerBlock)) {
      ended = true;
      break;
    }

    const info = decodeHeader(headerBlock);
    if (!info) {
      ended = true;
      break;
    }

    offset += TAR_BLOCK_SIZE;

    // Calculate data size and padding
    const dataSize = info.size;
    const padding = calculatePadding(dataSize);
    const totalSize = dataSize + padding;

    // A record whose data runs past the end is a truncated archive, as the streaming parser
    // reports it — not an entry to return shortened.
    if (offset + totalSize > data.length) {
      throw new ArchiveError(`Unexpected end of TAR archive reading "${info.path}"`);
    }

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
      throw new ArchiveFileTooLargeError(
        info.path,
        `exceeds maximum file size (${dataSize} > ${maxFileSize})`,
        { allowed: maxFileSize }
      );
    }

    entries.push({ info, start: offset, end: offset + dataSize });
    offset += totalSize;
  }

  if (!ended) {
    refusePartialBlock(data.subarray(offset));
  }
  return entries;
}

/**
 * Parse TAR archive from any archive source, streaming: one entry is held at a time. The
 * source and options are checked when this is called, before it is iterated.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function parseTarStream(
  source: ArchiveSource | ArchiveBytes,
  options: TarParseOptions = {}
): AsyncGenerator<TarEntry> {
  if (!isArchiveSource(source)) {
    throw new ArchiveArgumentError(
      "source",
      `Archive.parseTarStream: source must be an archive source (bytes, a string, a Blob, a ReadableStream or an async iterable), got ${describeSource(source)}`
    );
  }
  checkReadOptions("Archive.parseTarStream", options, TAR_READ_LIMITS);
  return readTarStream(source, options);
}

/**
 * {@link parseTarStream} without its argument checks, for readers that checked their own.
 * @internal
 */
export async function* readTarStream(
  source: ArchiveSource | ArchiveBytes,
  options: TarParseOptions
): AsyncGenerator<TarEntry> {
  const { signal } = options;
  const maxFileSize = entrySizeLimit(options);

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
    for await (const chunk of iterateArchiveSource(source, { signal })) {
      if (signal?.aborted) {
        throw createAbortError();
      }
      yield chunk;
    }
  }

  const sourceIterator = readFromSource();
  let sourceEnded = false;

  async function ensureBuffer(needed: number): Promise<boolean> {
    while (bufferSize - bufferOffset < needed) {
      const { value, done } = await sourceIterator.next();
      if (done) {
        sourceEnded = true;
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

  try {
    // Main parsing loop
    let first = true;
    while (true) {
      // Read header block
      if (!(await ensureBuffer(TAR_BLOCK_SIZE))) {
        if (bufferSize - bufferOffset > 0) {
          refusePartialBlock(readBytes(bufferSize - bufferOffset));
        }
        break;
      }

      const headerBlock = readBytes(TAR_BLOCK_SIZE);
      if (first) {
        refuseGzipTar(headerBlock);
        first = false;
      }

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
        throw new ArchiveFileTooLargeError(
          info.path,
          `exceeds maximum file size (${dataSize} > ${maxFileSize})`,
          { allowed: maxFileSize }
        );
      }

      // Read entry data
      if (!(await ensureBuffer(dataSize + padding))) {
        throw new ArchiveError(`Unexpected end of TAR archive reading "${info.path}"`);
      }

      const entryData = readBytes(dataSize);
      if (padding > 0) {
        skipBytes(padding);
      }

      yield createTarEntry(info, entryData);
    }
  } finally {
    // The parse owns its source, so leaving it before the source's end — the consumer stopped
    // early, the archive ended before its trailing bytes, or it failed — lets go of the source
    // the way a `for await` over it does. Returning the iterator destroys a Node stream and
    // returns an async iterator, but only releases a Web stream's reader (`toAsyncIterable`
    // hands a stream back unlocked), so a Web stream is then cancelled here, as ZIP and
    // tar.gz cancel theirs. It used to be iterated directly, which cancelled it.
    await sourceIterator.return(undefined);
    if (!sourceEnded && isReadableStream(source)) {
      await source.cancel().catch(() => {
        // Already errored or cancelled: nothing left to let go of
      });
    }
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
 * Parse TAR archive and return all entries with their data. Takes any `ArchiveSource`, as
 * `unpackTarGz` does. For bytes in memory each `data` is a view into them, not a copy (as
 * with {@link untarSync}); copy it (`data.slice()`) before mutating either. A streamed
 * source yields fresh buffers. An archive cut short is an `ArchiveError`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export async function untar(
  source: ArchiveSource | ArchiveBytes,
  options: TarParseOptions = {}
): Promise<Map<string, UnpackedTarEntry>> {
  if (!isArchiveSource(source)) {
    throw new ArchiveArgumentError(
      "source",
      `Archive.unpackTar: unsupported archive source (${describeSource(source)})`
    );
  }
  checkReadOptions("Archive.unpackTar", options, TAR_READ_LIMITS);
  if (isInMemoryArchiveSource(source)) {
    return untarSync(await toUint8Array(source), options);
  }
  const result = new Map<string, UnpackedTarEntry>();
  for await (const entry of readTarStream(source, options)) {
    result.set(entry.info.path, {
      info: entry.info,
      data: asPlainBytes(await entry.bytes())
    });
  }
  return result;
}

/**
 * Parse an in-memory TAR archive synchronously and return every entry with its data —
 * the synchronous form of {@link untar} for a `Uint8Array` source. Each `data` is a view
 * into `source`, not a copy (unless `source` is in a `SharedArrayBuffer`, which is copied — see `asPlainBytes`).
 */
export function untarSync(
  source: ArchiveBytes,
  options: Omit<TarParseOptions, "signal"> = {}
): Map<string, UnpackedTarEntry> {
  if (!isArchiveBytes(source)) {
    throw new ArchiveArgumentError(
      "source",
      `Archive.unpackTarSync: source must be bytes (an ArrayBufferView or ArrayBuffer), got ${describeSource(source)}; ` +
        "use Archive.unpackTar for other sources"
    );
  }
  checkReadOptions("Archive.unpackTarSync", options, TAR_READ_LIMITS);
  const bytes = archiveBytesToUint8Array(source);
  const result = new Map<string, UnpackedTarEntry>();
  for (const { info, start, end } of parseTarRecords(bytes, options)) {
    result.set(info.path, { info, data: asPlainBytes(bytes.subarray(start, end)) });
  }
  return result;
}
