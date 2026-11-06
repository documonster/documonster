import { checkEntryPath } from "@archive/core/entry-checks";
import {
  ArchiveArgumentError,
  ArchiveError,
  createAbortError,
  ArchiveFileTooLargeError,
  oneShotSourceReadError,
  createLinkedAbortController,
  throwIfAborted,
  toError,
  suppressUnhandledRejection
} from "@archive/core/errors";
import { ProgressEmitter } from "@archive/core/progress";
import type { ZipStringEncoding } from "@archive/core/text";
import { decodeEntryText } from "@archive/core/text";
import type { ArchiveFormat } from "@archive/core/types";
import type { ArchivePipeToOptions, ArchiveSink } from "@archive/io/archive-sink";
import { checkSinkArgument, pipeIterableToSink } from "@archive/io/archive-sink";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  asyncIterableToReadableStream,
  isArchiveSource,
  isInMemoryArchiveSource,
  resolveArchiveSourceToBuffer,
  iterateArchiveSource,
  toUint8Array
} from "@archive/io/archive-source";
import type { TarReaderProgress } from "@archive/tar/tar-archive";
import { hasDataDescriptorFlag } from "@archive/unzip/parser-core";
import type { UnzipOperation, UnzipProgress, UnzipStreamOptions } from "@archive/unzip/progress";
import type { ParseOptions, ZipEntry as ParseZipEntry } from "@archive/unzip/stream";
import { createParse } from "@archive/unzip/stream";
import { DEFAULT_UNZIP_STREAM_HIGH_WATER_MARK } from "@archive/unzip/stream.base";
import {
  processEntryData,
  processEntryDataStream,
  readEntryCompressedData
} from "@archive/unzip/zip-extract-core";
import type { ZipEntryRecord, ZipParseOptions } from "@archive/unzip/zip-parser";
import { ZipParser } from "@archive/unzip/zip-parser";
import type { ZipEntryEncryptionMethod, ZipEntryType } from "@archive/zip-spec/zip-entry-info";
import { isSymlink } from "@archive/zip-spec/zip-entry-info";
import { COMPRESSION_AES } from "@archive/zip-spec/zip-records";
import { eventedReadableToAsyncIterableNoDestroy } from "@stream/core/evented-readable-to-async-iterable";
import { isReadableStream, isWritableStream } from "@stream/core/type-guards";
import { concatUint8Arrays, getTextDecoder } from "@utils/binary";

function attachAbortToParseEntry(entry: ParseZipEntry, signal: AbortSignal): void {
  let cleanedUp = false;

  const cleanup = () => {
    if (cleanedUp) {
      return;
    }
    cleanedUp = true;
    signal.removeEventListener("abort", onAbort);
  };

  const onAbort = () => {
    cleanup();
    try {
      entry.destroy(createAbortError(signal.reason));
    } catch {
      entry.autodrain();
    }
  };

  if (signal.aborted) {
    onAbort();
    return;
  }

  signal.addEventListener("abort", onAbort, { once: true });
  entry.once("end", cleanup);
  entry.once("close", cleanup);
  entry.once("error", cleanup);
}

/**
 * The most of an unread entry a streaming reader keeps when the caller moves past it.
 *
 * A streamed archive can only move to its next entry once the current one has been read, so an entry the caller
 * does not read is read for them: kept, so it can still be read later, while it is small, and skipped once it is
 * not. The limit is what the parser always buffered for an entry, so an entry that could be read after moving past
 * it before still can — where a larger one used to stall the whole iteration.
 */
const KEEP_UNREAD_ENTRY_LIMIT = DEFAULT_UNZIP_STREAM_HIGH_WATER_MARK;

/**
 * Read `entry` to its end in memory, up to `limit` bytes. Resolves with the bytes, or with `undefined` once there
 * are more than that — having called `onOverflow`, which must stop the entry or drain it elsewhere.
 */
function keepEntry(
  entry: ParseZipEntry,
  limit: number,
  onOverflow: () => void
): Promise<Uint8Array | undefined> {
  return new Promise((resolve, reject) => {
    const parts: Uint8Array[] = [];
    let total = 0;
    const cleanup = (): void => {
      entry.removeListener("data", onData);
      entry.removeListener("end", onEnd);
      entry.removeListener("error", onError);
      entry.removeListener("close", onClose);
    };
    const onData = (chunk: Uint8Array): void => {
      total += chunk.length;
      if (total > limit) {
        cleanup();
        parts.length = 0;
        onOverflow();
        resolve(undefined);
        return;
      }
      parts.push(chunk);
    };
    const onEnd = (): void => {
      cleanup();
      resolve(parts.length === 1 ? asPlainBytes(parts[0]!) : concatUint8Arrays(parts));
    };
    const onError = (error: Error): void => {
      cleanup();
      reject(error);
    };
    const onClose = (): void => {
      cleanup();
      reject(new ArchiveError(`Entry "${entry.path}" was closed before it was read to the end`));
    };
    entry.on("data", onData);
    entry.once("end", onEnd);
    entry.once("error", onError);
    entry.once("close", onClose);
  });
}

/**
 * In Node.js an entry's data may be a `Buffer`, which causes deep-equality mismatches
 * against `Uint8Array`; hand back a plain view of the same bytes.
 */
function asPlainBytes(data: Uint8Array): Uint8Array {
  return typeof Buffer !== "undefined" && data instanceof Buffer
    ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
    : data;
}

/** Resolves once `entry` has been read to its end, or is gone. */
function entryDone(entry: ParseZipEntry): Promise<void> {
  if (entry.destroyed || entry.readableEnded) {
    return Promise.resolve();
  }
  return new Promise(resolve => {
    const done = (): void => {
      entry.removeListener("end", done);
      entry.removeListener("close", done);
      entry.removeListener("error", done);
      resolve();
    };
    entry.once("end", done);
    entry.once("close", done);
    entry.once("error", done);
  });
}

/**
 * Let go of the caller's source once a streaming read is over early. A Node stream is destroyed, as leaving a
 * `for await` over it would; a Web stream is cancelled through the operation's signal, which the feed reads it under.
 */
function releaseSource(source: ArchiveSource | ArchiveBytes): void {
  // Not a source at all: the read has already failed with an archive error, which this must not mask.
  if (source == null || isReadableStream(source)) {
    return;
  }
  const destroy = (source as { destroy?: unknown }).destroy;
  if (typeof destroy === "function") {
    try {
      (destroy as () => void).call(source);
    } catch {
      // Best effort: the read is over either way
    }
  }
}

/**
 * Build a ZipEntryRecord from a streaming ParseZipEntry's local file header data.
 * This enables processEntryData (decrypt + decompress) for streaming-mode entries.
 */
function buildEntryInfoFromParseEntry(entry: ParseZipEntry): ZipEntryRecord {
  const vars = entry.vars;
  const flags = vars.flags ?? 0;
  const compressionMethod = vars.compressionMethod ?? 0;
  const isEncrypted = (flags & 0x01) !== 0 || compressionMethod === COMPRESSION_AES;
  const aesInfo = entry.extraFields?.aesInfo;

  let encryptionMethod: ZipEntryEncryptionMethod = "none";
  if (isEncrypted) {
    if (compressionMethod === COMPRESSION_AES && aesInfo) {
      encryptionMethod = "aes";
    } else {
      encryptionMethod = "zipcrypto";
    }
  }

  return {
    path: entry.path,
    type: entry.type === "Directory" ? "directory" : "file",
    compressedSize: vars.compressedSize ?? 0,
    uncompressedSize: vars.uncompressedSize ?? 0,
    compressionMethod,
    crc32: vars.crc32 ?? 0,
    lastModified: vars.lastModifiedDateTime ?? new Date(0),
    localHeaderOffset: 0,
    comment: "",
    externalAttributes: 0,
    mode: 0,
    isEncrypted,
    encryptionMethod,
    aesVersion: aesInfo?.version,
    aesKeyStrength: aesInfo?.keyStrength,
    originalCompressionMethod: aesInfo?.compressionMethod,
    dosTime: vars.lastModifiedTime ?? undefined
  };
}

export type { ArchivePipeToOptions } from "@archive/io/archive-sink";

export interface UnzipOptions {
  /**
   * Archive format: "zip" (default) or "tar".
   * Read only by the deprecated `Archive.unzip`; `Archive.openZip` and `Archive.openTar` name the format instead.
   */
  format?: ArchiveFormat;

  decodeStrings?: boolean;
  /** Optional string encoding for legacy (non-UTF8) names/comments. */
  encoding?: ZipStringEncoding;
  parse?: ParseOptions;

  /** Password for encrypted entries (ZIP only). */
  password?: string | Uint8Array;

  /**
   * Whether to verify each entry against its recorded CRC-32 as it is read, throwing
   * `ZipCrc32MismatchError` on a mismatch — from `bytes()`, or from `stream()` once the
   * entry's last chunk has been produced. Only entries that are read are verified, over any
   * source: a mismatch fails that read, not the iteration. Set `false` to read damaged archives as far as
   * they go. AES entries carry an HMAC instead; ZipCrypto entries are always verified.
   * @default true
   */
  checkCrc32?: boolean;

  /**
   * Largest uncompressed size, in bytes, of one entry that `bytes()` / `text()` read into
   * memory (`ArchiveFileTooLargeError`), as `unpackZip`'s `maxEntrySize`. Over bytes in memory an entry
   * declaring more is refused before inflation and one inflating past its size is stopped; over
   * a stream an entry is refused when its header declares more, or once it has read more.
   * @default 536870912 (512 MiB)
   */
  maxEntrySize?: number;

  /**
   * Most entries the archive may contain (`ArchiveLimitExceededError`), checked against the central
   * directory before its records are read — so for a streamed source only by `list` / `get` /
   * `bytes`, which need it. Unbounded by default.
   */
  maxEntries?: number;

  /** Default abort signal used by streaming operations. */
  signal?: AbortSignal;

  /** Default progress callback used by streaming operations. */
  onProgress?: (p: UnzipProgress) => void;

  /** Default throttle for progress callbacks. */
  progressIntervalMs?: number;
}

export type { UnzipOperation, UnzipProgress, UnzipStreamOptions } from "@archive/unzip/progress";

export class UnzipEntry {
  readonly path: string;
  /** Entry type: file, directory, or symlink */
  readonly type: ZipEntryType;
  /** Whether the entry is a directory (`type === "directory"`; matches `TarReaderEntry.isDirectory`). */
  get isDirectory(): boolean {
    return this.type === "directory";
  }
  /**
   * For symlinks, returns the target path after calling bytes().
   * Before extraction, this is undefined.
   */
  linkTarget?: string;
  /**
   * Unix file mode/permissions (0 if unavailable).
   */
  readonly mode: number;
  /** Whether this entry is encrypted (ZipCrypto or AES). */
  readonly isEncrypted: boolean;

  private readonly _data?: Uint8Array;
  private readonly _info?: ZipEntryRecord;
  private readonly _password?: string | Uint8Array;
  private readonly _parseEntry?: ParseZipEntry;
  private readonly _onBytesOut?: (path: string, type: ZipEntryType, bytes: number) => void;
  private readonly _signal?: AbortSignal;
  private readonly _checkCrc32: boolean;
  private readonly _maxEntrySize?: number;
  private _infoCopy?: Readonly<ZipEntryRecord>;

  // Streaming entries only: whether the caller has started reading, and what became of it if they moved on first.
  private _state: "fresh" | "reading" | "kept" | "skipped" | "discarded" = "fresh";
  private _kept?: Promise<Uint8Array | undefined>;
  private _failure?: Error;
  /** Whether the caller moved past this entry by leaving the iteration, rather than by asking for the next one. */
  private _leftOn = false;

  /** @internal How a reader makes an entry; the constructor is not public API. */
  static _create(
    args: ConstructorParameters<typeof UnzipEntry>[0],
    hooks?: ConstructorParameters<typeof UnzipEntry>[1]
  ): UnzipEntry {
    return new UnzipEntry(args, hooks);
  }

  /**
   * @deprecated Entries come from a reader — `ZipReader.entries()` / `get()`, or iterating
   * `Archive.openZip(…)` — and the shapes this takes are the reader's own. Removed in 1.0.
   */
  constructor(
    args:
      | { kind: "buffer"; data: Uint8Array; info: ZipEntryRecord; password?: string | Uint8Array }
      | { kind: "stream"; entry: ParseZipEntry; password?: string | Uint8Array },
    hooks: {
      onBytesOut?: (path: string, type: ZipEntryType, bytes: number) => void;
      signal?: AbortSignal;
      checkCrc32?: boolean;
      maxEntrySize?: number;
    } = {}
  ) {
    if (args.kind === "buffer") {
      this._data = args.data;
      this._info = args.info;
      this._password = args.password;
      this.path = args.info.path;
      this.type = args.info.type;
      this.mode = args.info.mode;
      this.isEncrypted = args.info.isEncrypted;
    } else {
      this._parseEntry = args.entry;
      this._password = args.password;
      this.path = args.entry.path;
      // Streaming parser cannot detect symlinks (requires Central Directory)
      this.type = args.entry.type === "Directory" ? "directory" : "file";
      this.mode = 0;
      const flags = args.entry.vars.flags ?? 0;
      this.isEncrypted = (flags & 0x01) !== 0 || args.entry.vars.compressionMethod === 99;

      // For encrypted entries the streaming parser passes through raw ciphertext
      // (inflate is skipped). Build a ZipEntryRecord so processEntryData can
      // handle decryption + decompression in bytes()/stream().
      if (this.isEncrypted) {
        this._info = buildEntryInfoFromParseEntry(args.entry);
      }
    }

    this._onBytesOut = hooks.onBytesOut;
    this._signal = hooks.signal;
    this._checkCrc32 = hooks.checkCrc32 ?? true;
    this._maxEntrySize = hooks.maxEntrySize;

    if (this._parseEntry) {
      // This entry hears every failure of its data and hands it to whoever reads it — including one found after the
      // caller moved on — so a CRC-32 mismatch fails the entry, not the iteration, as it does over bytes in memory.
      this._parseEntry.reportsOwnFailures = true;
      this._parseEntry.on("error", (error: Error) => {
        this._failure ??= error;
      });
      // Interrupt it on abort so consumers don't hang waiting for more chunks.
      if (this._signal) {
        attachAbortToParseEntry(this._parseEntry, this._signal);
      }
    }
  }

  /**
   * The entry's record: from the central directory when the reader has the whole archive
   * (`ZipReader.get`, or `entries()` over bytes in memory), from the local header while it
   * streams. A streamed entry's local header has no comment or external attributes, and an
   * entry written with a data descriptor (general-purpose flag bit 3) has its sizes and
   * CRC-32 after its data, so those read 0 here. `lastModified` comes from the UT field when
   * there is one and from the DOS fields (local time) otherwise. A frozen copy.
   */
  get info(): Readonly<ZipEntryRecord> {
    this._infoCopy ??= Object.freeze({
      ...(this._info ?? buildEntryInfoFromParseEntry(this._parseEntry!))
    });
    return this._infoCopy;
  }

  /**
   * The caller has moved past this streamed entry — asked for the next one, or left the iteration. One they never
   * started reading is kept in memory while it is small, so it can still be read, and skipped once it is not; reading
   * it after that throws instead of stalling. Resolves once the entry is through: read to its end, kept, skipped or closed.
   *
   * `leaving`: the iteration is over, so a skipped entry is closed rather than skipped through to its end.
   *
   * @internal
   */
  _moveOn(leaving: boolean): Promise<void> {
    const entry = this._parseEntry;
    if (!entry) {
      return Promise.resolve();
    }
    const stop = (): void => {
      if (leaving) {
        entry.destroy();
      } else {
        entry.autodrain();
      }
    };
    if (this._state === "fresh") {
      this._leftOn = leaving;
      if (this._failure) {
        this._state = "kept";
        this._kept = Promise.reject(this._failure);
        suppressUnhandledRejection(this._kept);
      } else if (entry.size !== undefined && entry.size > KEEP_UNREAD_ENTRY_LIMIT) {
        this._state = "skipped";
        stop();
      } else {
        this._state = "kept";
        this._kept = keepEntry(entry, KEEP_UNREAD_ENTRY_LIMIT, () => {
          this._state = "skipped";
          stop();
        });
        suppressUnhandledRejection(this._kept);
      }
    } else if (leaving && entry.__autodraining && !entry.destroyed) {
      // Left part-read: what remains is not wanted, and skipping through it would read the rest of the archive's
      // copy of it for nothing
      entry.destroy();
    }
    return entryDone(entry);
  }

  /** The bytes of a streamed entry the caller moved past before reading it. */
  private async _keptBytes(): Promise<Uint8Array> {
    const kept = this._state === "kept" ? await this._kept : undefined;
    if (kept !== undefined) {
      return kept;
    }
    if (this._state === "discarded") {
      throw new ArchiveError(`Entry "${this.path}" was discarded`);
    }
    // Two different mistakes end here, and the fix for each is different: asking for the next entry before reading
    // this one, or leaving the loop on this entry and reading it afterwards.
    const tooLarge = `it is too large to keep (over ${KEEP_UNREAD_ENTRY_LIMIT} bytes)`;
    const inMemory =
      "or read the archive from bytes in memory, where every entry can be read at any time.";
    throw new ArchiveError(
      this._leftOn
        ? `Entry "${this.path}" was skipped: the iteration ended before it was read, and ${tooLarge}. ` +
            `Read a streamed entry before leaving the loop, ${inMemory}`
        : `Entry "${this.path}" was skipped: the reader moved past it unread to the next entry, and ${tooLarge}. ` +
            `Read a streamed entry before asking for the next one, ${inMemory}`
    );
  }

  /** Start reading a streamed entry; `false` when it was moved past, and is read with {@link _keptBytes}. */
  private _startReading(): boolean {
    if (this._state !== "fresh" && this._state !== "reading") {
      return false;
    }
    this._state = "reading";
    if (this._failure) {
      throw this._failure;
    }
    return true;
  }

  /**
   * Process extracted bytes: populate linkTarget for symlinks and notify progress.
   */
  private _processExtractedBytes(bytes: Uint8Array): Uint8Array {
    // For symlinks, the data content is the target path
    if (isSymlink(this.type) && bytes.length > 0) {
      this.linkTarget = getTextDecoder().decode(bytes);
    }
    if (this._onBytesOut && bytes.length) {
      this._onBytesOut(this.path, this.type, bytes.length);
    }
    return bytes;
  }

  async bytes(): Promise<Uint8Array> {
    if (this._data && this._info) {
      // Use shared extraction core for buffer mode
      const compressedData = readEntryCompressedData(this._data, this._info);
      const bytes = await processEntryData(
        this._info,
        compressedData,
        this._password,
        this._checkCrc32,
        true,
        this._maxEntrySize
      );
      return this._processExtractedBytes(bytes);
    }
    if (this._parseEntry) {
      if (!this._startReading()) {
        return this._processExtractedBytes(await this._decryptStreamed(await this._keptBytes()));
      }
      const limit = this._maxEntrySize;
      const declared = this._parseEntry.vars.uncompressedSize;
      if (
        limit !== undefined &&
        !hasDataDescriptorFlag(this._parseEntry.vars.flags) &&
        typeof declared === "number" &&
        declared > limit
      ) {
        this._parseEntry.autodrain();
        throw new ArchiveFileTooLargeError(
          this.path,
          `declares an uncompressed size of ${declared} bytes, more than maxEntrySize (${limit})`,
          { allowed: limit }
        );
      }
      const data = await this._parseEntry.buffer();
      let bytes = asPlainBytes(data);
      if (limit !== undefined && bytes.length > limit) {
        throw new ArchiveFileTooLargeError(
          this.path,
          `inflates to ${bytes.length} bytes, more than maxEntrySize (${limit})`,
          { allowed: limit }
        );
      }

      bytes = await this._decryptStreamed(bytes);

      return this._processExtractedBytes(bytes);
    }
    return new Uint8Array(0);
  }

  /**
   * Encrypted entries in streaming mode carry raw ciphertext (inflate is skipped by readFileRecord). Decrypt +
   * decompress via processEntryData, the same path used by buffer mode.
   */
  private async _decryptStreamed(bytes: Uint8Array): Promise<Uint8Array> {
    if (this._info && this.isEncrypted) {
      return processEntryData(this._info, bytes, this._password);
    }
    return bytes;
  }

  async *stream(): AsyncIterable<Uint8Array> {
    if (this._data && this._info) {
      // Symlinks are tiny and their target is resolved by bytes() via
      // _processExtractedBytes; keep the buffered path for them.
      if (this._info.type === "symlink") {
        const data = await this.bytes();
        if (data.length) {
          yield data;
        }
        return;
      }

      // True streaming for buffer-backed entries: inflate incrementally rather
      // than materializing the whole uncompressed entry up-front. This keeps
      // peak memory at O(inflate chunk) for large entries (e.g. a multi-MB
      // `word/document.xml` / xlsx `sheetN.xml`), which is what makes
      // SAX-driven streaming readers genuinely O(largest element) on an
      // in-memory package instead of O(full uncompressed part).
      //
      // The compressed bytes are fed in fixed-size slices rather than as a
      // single chunk. Feeding the whole entry in one `write()` lets the
      // decompressor (notably the browser's async `DecompressionStream`,
      // whose read loop runs ahead of the consumer) produce *all* output
      // before the consumer pulls the first chunk — collapsing the stream
      // back to O(full uncompressed part). Slicing the input interleaves
      // production with consumption so output is delivered incrementally on
      // every platform.
      const compressedData = readEntryCompressedData(this._data, this._info);
      const FEED_CHUNK = 65536;
      const feed = (async function* () {
        for (let off = 0; off < compressedData.length; off += FEED_CHUNK) {
          yield compressedData.subarray(off, off + FEED_CHUNK);
        }
      })();
      const outStream = processEntryDataStream(this._info, feed, {
        password: this._password,
        checkCrc32: this._checkCrc32,
        signal: this._signal
      });

      // Robustness: the browser's native streaming `DecompressionStream` can
      // intermittently reject input that is in fact valid deflate (observed in
      // Chromium under heavy concurrent stream creation). The non-streaming
      // `processEntryData` path already recovers from this by falling back to
      // the pure-JS inflater. Mirror that here: if the streaming inflate fails
      // *before producing any output* we have not yielded anything yet, so we
      // can safely re-decode the whole (already in-memory) entry via the
      // buffered path and yield the result as a single chunk. Once output has
      // started flowing we keep streaming; a mid-stream failure on valid data
      // has not been observed and would indicate genuine corruption.
      //
      // The first pull is handled separately from the rest so the two error
      // regimes are structurally distinct: a throw from the first `next()`
      // means "before any output" (recover); any throw once we are streaming
      // propagates unchanged. `iterator.return()` in `finally` reproduces the
      // automatic cleanup `for await` performs on early break/return/throw,
      // so the underlying decompress stream is always torn down.
      const iterator = outStream[Symbol.asyncIterator]();
      try {
        let first: IteratorResult<Uint8Array>;
        try {
          first = await iterator.next();
        } catch {
          // Fall back to the buffered decode (which itself falls back to
          // pure-JS inflate on native failure). Throws if genuinely corrupt.
          const recovered = await this.bytes();
          if (recovered.length) {
            if (this._onBytesOut) {
              this._onBytesOut(this.path, this.type, recovered.length);
            }
            yield recovered;
          }
          return;
        }
        // Streaming succeeded past the first pull: yield the first chunk and
        // the remainder. Any failure from here on is genuine mid-stream
        // corruption and propagates to the caller.
        for (let step = first; !step.done; step = await iterator.next()) {
          const chunk = step.value;
          if (this._onBytesOut && chunk.length) {
            this._onBytesOut(this.path, this.type, chunk.length);
          }
          yield chunk;
        }
      } finally {
        await iterator.return?.();
      }
      return;
    }

    if (this._parseEntry) {
      if (!this._startReading()) {
        const data = await this.bytes();
        if (data.length) {
          yield data;
        }
        return;
      }
      // ParseZipEntry is always a Node PassThrough, so it satisfies the evented
      // readable contract; adapt it to an AsyncIterable without destroying it.
      const iterable: AsyncIterable<Uint8Array> =
        eventedReadableToAsyncIterableNoDestroy<Uint8Array>(this._parseEntry);

      // Encrypted entries carry raw ciphertext in streaming mode (inflate is
      // skipped by readFileRecord). Pipe through processEntryDataStream which
      // handles decrypt + decompress in a true streaming fashion for ZipCrypto.
      // AES still buffers internally (HMAC needs full ciphertext) but the API
      // surface remains consistent.
      if (this.isEncrypted && this._info) {
        const outStream = processEntryDataStream(this._info, iterable, {
          password: this._password,
          signal: this._signal
        });

        let completed = false;
        try {
          for await (const chunk of outStream) {
            if (this._onBytesOut && chunk.length) {
              this._onBytesOut(this.path, this.type, chunk.length);
            }
            yield chunk;
          }
          completed = true;
        } finally {
          if (!completed) {
            try {
              this._parseEntry.autodrain();
            } catch {
              // Best effort cleanup only.
            }
          }
        }
        return;
      }

      let completed = false;
      try {
        for await (const chunk of iterable) {
          if (this._onBytesOut && chunk.length) {
            this._onBytesOut(this.path, this.type, chunk.length);
          }
          yield chunk;
        }
        completed = true;
      } finally {
        if (!completed) {
          try {
            this._parseEntry.autodrain();
          } catch {
            // Best effort cleanup only.
          }
        }
      }
    }
  }

  async pipeTo(sink: WritableStream<Uint8Array>, options?: ArchivePipeToOptions): Promise<void>;
  async pipeTo(sink: ArchiveSink): Promise<void>;
  async pipeTo(sink: ArchiveSink, options?: ArchivePipeToOptions): Promise<void> {
    checkSinkArgument("Archive.UnzipEntry#pipeTo", sink);
    // Prefer native Web Streams piping semantics when a WHATWG WritableStream is provided.
    // This supports standard options like `signal` / `preventClose` / `preventAbort`.
    if (isWritableStream(sink) && typeof this.readableStream().pipeTo === "function") {
      await this.readableStream().pipeTo(sink, options);
      return;
    }

    // Fallback to the library sink piping (supports Node-style Writable too).
    await pipeIterableToSink(this.stream(), sink);
  }

  readableStream(): ReadableStream<Uint8Array> {
    const parseEntry = this._parseEntry;

    return asyncIterableToReadableStream(this.stream(), reason => {
      if (parseEntry) {
        try {
          parseEntry.destroy?.(createAbortError(reason));
        } catch {
          try {
            parseEntry.autodrain?.();
          } catch {
            // ignore
          }
        }
      }
    });
  }

  async text(encoding?: string): Promise<string> {
    const bytes = await this.bytes();
    return decodeEntryText(bytes, encoding);
  }

  /**
   * Skip this entry without reading it. Over a streamed source the entry cannot be read afterwards — reading it then
   * throws an `ArchiveError` — and what was kept of it is released.
   */
  discard(): void {
    if (!this._parseEntry) {
      return;
    }
    const wasKept = this._state === "kept" || this._state === "skipped";
    if (this._state === "fresh" || this._state === "kept") {
      this._state = "discarded";
      this._kept = undefined;
    }
    if (!wasKept) {
      this._parseEntry.autodrain();
    }
  }
}

/**
 * Reads a ZIP archive (`Archive.openZip` returns one).
 *
 * Over a one-shot source — a `ReadableStream`, an async iterable or a Node stream — the order of
 * calls matters, because the source can be read only once:
 *
 * - `list()`, `get()` or `bytes()` first reads the whole source into memory. Every later call,
 *   `entries()` included, works from that buffer, as often as asked.
 * - Iterating `entries()`, `entriesStream()` or `operation().iterable` first streams the source
 *   and spends it: any later call throws an `ArchiveError` ("already consumed"). Only starting
 *   the iteration reads; asking for the iterable does not.
 *
 * Bytes, a string or a `Blob` are read again by every call and are unaffected.
 */
export class ZipReader {
  private readonly _source: ArchiveSource | ArchiveBytes;
  private readonly _options: UnzipOptions;
  private _bufferParser: ZipParser | null = null;
  private _bufferData: Uint8Array | null = null;
  // A one-shot source (a stream) can be read once: by a streaming pass, or buffered for list/get
  private _sourceRead = false;

  constructor(source: ArchiveSource | ArchiveBytes, options: UnzipOptions = {}) {
    this._source = source;
    this._options = options;
  }

  private get _encoding() {
    return this._options.encoding ?? this._options.parse?.encoding;
  }

  /**
   * The entries, in archive order. Over a one-shot source that `list` / `get` / `bytes` did not
   * buffer first, iterating streams the source and spends it (see {@link ZipReader}).
   */
  entries(options: UnzipStreamOptions = {}): AsyncIterable<UnzipEntry> {
    return this.operation(options).iterable;
  }

  /**
   * {@link ZipReader.entries} as a Web `ReadableStream`; reading it spends a one-shot source the
   * same way.
   */
  entriesStream(options: UnzipStreamOptions = {}): ReadableStream<UnzipEntry> {
    const op = this.operation(options);

    return asyncIterableToReadableStream(op.iterable, reason => {
      try {
        op.abort(reason ?? "cancelled");
      } catch {
        // ignore
      }
    });
  }

  /**
   * {@link ZipReader.entries} with abort and progress control; iterating its `iterable` spends a
   * one-shot source the same way.
   */
  operation(options: UnzipStreamOptions = {}): UnzipOperation {
    const { controller, cleanup: cleanupAbortLink } = createLinkedAbortController(
      options.signal ?? this._options.signal
    );
    const signal = controller.signal;

    const onProgress = options.onProgress ?? this._options.onProgress;
    const progress = new ProgressEmitter<UnzipProgress>(
      {
        type: "unzip",
        phase: "running",
        bytesIn: 0,
        bytesOut: 0,
        entriesEmitted: 0
      },
      onProgress,
      { intervalMs: options.progressIntervalMs ?? this._options.progressIntervalMs }
    );

    const onBytesOut = (path: string, entryType: ZipEntryType, bytes: number): void => {
      progress.mutate(s => {
        s.bytesOut += bytes;
        const prev = s.currentEntry;
        s.currentEntry =
          prev && prev.path === path
            ? { ...prev, bytesOut: prev.bytesOut + bytes }
            : { path, entryType, bytesOut: bytes };
      });
    };

    const checkCrc32 = this._options.checkCrc32;
    // Set once a streaming pass has taken over cancelling the operation: it may have to wait for the caller to finish
    // reading the entry they left the loop with, and the abort would destroy that entry.
    let releasesItself = false;
    const iterable = async function* (this: ZipReader): AsyncIterable<UnzipEntry> {
      try {
        throwIfAborted(signal);

        const inMemory = isInMemoryArchiveSource(this._source);
        if (!inMemory && this._bufferData === null) {
          if (this._sourceRead) {
            throw oneShotSourceReadError("Archive.ZipReader#entries");
          }
          this._sourceRead = true;
        }

        // Buffer mode — an in-memory source, or a stream an earlier list/get buffered
        if (inMemory || this._bufferData !== null) {
          const bytes = inMemory ? await toUint8Array(this._source) : this._bufferData!;
          throwIfAborted(signal);
          progress.update({ bytesIn: bytes.length });
          const parser = new ZipParser(bytes, {
            decodeStrings: this._options.decodeStrings,
            encoding: this._encoding,
            checkCrc32: this._options.checkCrc32,
            maxEntries: this._options.maxEntries
          } satisfies ZipParseOptions);
          const password = this._options.password;

          for (const info of parser.getEntries()) {
            throwIfAborted(signal);
            progress.mutate(s => {
              s.entriesEmitted += 1;
              s.currentEntry = { path: info.path, entryType: info.type, bytesOut: 0 };
            });
            yield UnzipEntry._create(
              { kind: "buffer", data: bytes, info, password },
              { onBytesOut, signal, checkCrc32, maxEntrySize: this._options.maxEntrySize }
            );
          }

          if (progress.snapshot.phase === "running") {
            progress.update({ phase: "done" });
          }
          return;
        }

        // Streaming mode
        // The parser checks each entry's CRC-32 itself (see `ParseOptions.checkCrc32`)
        const parse = createParse({
          checkCrc32,
          ...(this._options.parse ?? {}),
          forceStream: true
        });

        const parseDonePromise = parse.promise();
        // Ensure abort-driven rejections from the parser itself never surface as unhandled.
        suppressUnhandledRejection(parseDonePromise);

        const onAbort = () => {
          const err = createAbortError(signal.reason);
          progress.update({ phase: "aborted" });
          try {
            parse.destroy(err);
          } catch {
            // ignore
          }
        };
        signal.addEventListener("abort", onAbort, { once: true });

        const feedPromise = (async () => {
          try {
            for await (const chunk of iterateArchiveSource(this._source, {
              signal,
              onChunk: c =>
                progress.mutate(s => {
                  s.bytesIn += c.length;
                })
            })) {
              throwIfAborted(signal);
              await new Promise<void>((resolve, reject) => {
                parse.write(chunk, (err?: Error | null) => {
                  if (err) {
                    reject(err);
                  } else {
                    resolve();
                  }
                });
              });
            }

            throwIfAborted(signal);
            parse.end();
            await parseDonePromise;
          } catch (e) {
            const err = toError(e);
            parse.destroy(err);
            throw err;
          }
        })();

        // Avoid unhandled rejection warnings when the operation is aborted.
        suppressUnhandledRejection(feedPromise);

        const parseIter: AsyncIterator<ParseZipEntry> = (parse as AsyncIterable<ParseZipEntry>)[
          Symbol.asyncIterator
        ]();

        const source = this._source;
        let released = false;
        // Tear the read down: the parser (and with it the entry it is in the middle of, and its inflater), the
        // caller's source, and the operation's signal, which cancels a Web stream the feed is reading.
        const release = (): void => {
          if (released) {
            return;
          }
          released = true;
          try {
            parse.destroy(createAbortError("cancelled"));
          } catch {
            // ignore
          }
          releaseSource(source);
          if (!signal.aborted) {
            controller.abort("cancelled");
          }
          cleanupAbortLink();
        };
        releasesItself = true;

        let current: UnzipEntry | undefined;
        let completed = false;
        try {
          while (true) {
            const { value, done } = await parseIter.next();
            if (done) {
              break;
            }
            const entry = value;
            throwIfAborted(signal);
            progress.mutate(s => {
              s.entriesEmitted += 1;
              s.currentEntry = {
                path: entry.path,
                entryType: entry.type === "Directory" ? "directory" : "file",
                bytesOut: 0
              };
            });
            current = UnzipEntry._create(
              { kind: "stream", entry, password: this._options.password },
              { onBytesOut, signal, checkCrc32, maxEntrySize: this._options.maxEntrySize }
            );
            yield current;
            // Asked for the next entry: the archive cannot move on until this one has been read
            void current._moveOn(false);
            current = undefined;
          }

          await feedPromise;
          completed = true;
          if (progress.snapshot.phase === "running") {
            progress.update({ phase: "done" });
          }
        } finally {
          signal.removeEventListener("abort", onAbort);

          // Closes the parser's entry queue; it settles at once.
          await parseIter.return?.().catch(() => {});

          if (completed) {
            cleanupAbortLink();
          } else if (current && !signal.aborted) {
            // The caller left the loop (`break`, `return`, a `throw`) holding `current`. Nothing here waits — the loop
            // settles now — but the parser and the source are kept until that entry is through: being read, or
            // kept for a later read while small. A failure or an abort tears everything down at once instead.
            void current._moveOn(true).then(release, release);
          } else {
            release();
          }
        }
      } catch (e) {
        const err = toError(e);
        if (err.name === "AbortError") {
          progress.update({ phase: "aborted" });
        } else {
          progress.update({ phase: "error" });
        }
        throw err;
      } finally {
        if (progress.snapshot.phase === "running" && !signal.aborted) {
          if (!releasesItself) {
            try {
              controller.abort("cancelled");
            } catch {
              // ignore
            }
          }
          progress.update({ phase: "aborted" });
        }
        if (!releasesItself) {
          cleanupAbortLink();
        }
        progress.emitNow();
      }
    }.call(this);

    return {
      iterable,
      signal,
      abort(reason?: unknown) {
        controller.abort(reason);
      },
      pointer() {
        return progress.snapshot.bytesIn;
      },
      progress() {
        return progress.snapshotCopy();
      }
    };
  }

  private async _ensureBufferParser(
    where: "list" | "get" | "bytes"
  ): Promise<{ parser: ZipParser; data: Uint8Array }> {
    if (this._bufferParser && this._bufferData) {
      return { parser: this._bufferParser, data: this._bufferData };
    }

    if (!isArchiveSource(this._source)) {
      throw new ArchiveArgumentError("source", "Unsupported archive source");
    }
    const inMemory = isInMemoryArchiveSource(this._source);
    if (!inMemory) {
      if (this._sourceRead) {
        throw oneShotSourceReadError(`Archive.ZipReader#${where}`);
      }
      this._sourceRead = true;
    }
    const bytes = inMemory
      ? await toUint8Array(this._source)
      : await resolveArchiveSourceToBuffer(this._source, { signal: this._options.signal });
    this._bufferData = bytes;
    this._bufferParser = new ZipParser(bytes, {
      decodeStrings: this._options.decodeStrings,
      encoding: this._encoding,
      checkCrc32: this._options.checkCrc32,
      maxEntries: this._options.maxEntries
    } satisfies ZipParseOptions);
    return { parser: this._bufferParser, data: bytes };
  }

  /**
   * Every entry's path, in the order of the central directory (matches `TarReader.list`).
   * Needs random access, so a one-shot source is read into memory on the first `list` / `get` /
   * `bytes`, and later calls (`entries()` too) use that buffer; after an iteration has spent the
   * source it throws an `ArchiveError` (see {@link ZipReader}).
   */
  async list(): Promise<string[]> {
    const { parser } = await this._ensureBufferParser("list");
    return parser.getEntries().map(entry => entry.path);
  }

  /**
   * Get one entry by path; `undefined` when the archive has no such entry. Of two entries with one
   * name the later is the one returned, as `unpackZip` and every extractor keep it.
   * Buffers a one-shot source as {@link ZipReader.list} does.
   */
  async get(path: string): Promise<UnzipEntry | undefined> {
    checkEntryPath("Archive.ZipReader#get", path);
    return this._get(path, "get");
  }

  /**
   * Read one entry's decompressed bytes; `undefined` when the archive has no such entry. Buffers a
   * one-shot source as {@link ZipReader.list} does.
   */
  async bytes(path: string): Promise<Uint8Array | undefined> {
    checkEntryPath("Archive.ZipReader#bytes", path);
    // Through `get`, so a subclass that overrides it is honoured
    const entry = await this.get(path);
    if (!entry) {
      return undefined;
    }
    return entry.bytes();
  }

  private async _get(path: string, where: "get" | "bytes"): Promise<UnzipEntry | undefined> {
    const { parser, data } = await this._ensureBufferParser(where);
    const info = parser.getEntry(path);
    if (!info) {
      return undefined;
    }
    return UnzipEntry._create(
      { kind: "buffer", data, info, password: this._options.password },
      { checkCrc32: this._options.checkCrc32, maxEntrySize: this._options.maxEntrySize }
    );
  }

  async close(): Promise<void> {
    // No persistent resources in buffer mode.
  }
}

/** Unzip options with format: "tar" */
export interface UnzipOptionsTar extends Omit<UnzipOptions, "onProgress"> {
  format: "tar";

  /**
   * Default progress callback used by streaming operations.
   *
   * TAR archives emit {@link TarReaderProgress} rather than {@link UnzipProgress}.
   */
  onProgress?: (p: TarReaderProgress) => void;
}

/** Unzip options with format: "zip" (or default) */
export interface UnzipOptionsZip extends UnzipOptions {
  format?: "zip";
}
