/**
 * TAR Archive - Unified API compatible with ZIP
 *
 * Provides TarArchive and TarReader classes with the same interface
 * as ZipArchive and ZipReader, allowing seamless format switching.
 */

import {
  OPTIONS,
  checkArchiveArgument,
  dateSchema,
  limitSchema,
  streamOptionShape,
  uintSchema
} from "@archive/core/arguments";
import {
  checkArchiveSource,
  checkBytes,
  checkEntryName,
  checkEntryPath,
  checkString
} from "@archive/core/entry-checks";
import {
  ArchiveArgumentError,
  ArchiveError,
  createAbortError,
  createLinkedAbortController,
  oneShotSourceReadError,
  throwIfAborted
} from "@archive/core/errors";
import type {
  ArchiveProgressPhase,
  ArchiveStreamOptions,
  ArchiveOperationBase
} from "@archive/core/progress";
import { decodeEntryText } from "@archive/core/text";
import type { ArchivePipeToOptions, ArchiveSink } from "@archive/io/archive-sink";
import { checkSinkArgument, collect, pipeIterableToSink } from "@archive/io/archive-sink";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  asyncIterableToReadableStream,
  checkSourceArgument,
  isInMemoryArchiveSource,
  isSyncArchiveSource,
  iterateArchiveSource,
  toUint8Array,
  toUint8ArraySync
} from "@archive/io/archive-source";
import type { TarType } from "@archive/tar/tar-constants";
import { TAR_TYPE, DEFAULT_TAR_MODE, DEFAULT_TAR_DIR_MODE } from "@archive/tar/tar-constants";
import type { TarEntryInfo } from "@archive/tar/tar-entry-info";
import { isDataEntry, isDirectory } from "@archive/tar/tar-entry-info";
import type { TarHeaderOptions } from "@archive/tar/tar-header";
import { encodeHeader, createPadding, createEndOfArchive } from "@archive/tar/tar-header";
import type { TarParseOptions } from "@archive/tar/tar-parser";
import { entrySizeLimit, parseTar, readTarStream } from "@archive/tar/tar-parser";
import { isWritableStream } from "@stream/core/type-guards";
import { concatUint8Arrays, textEncoder, asPlainBytes, plainChunks } from "@utils/binary";
import { defineAsyncDispose } from "@utils/dispose";
import type { AsyncDisposeMember } from "@utils/dispose";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// ============================================================================
// Types
// ============================================================================

export interface TarArchiveOptions {
  /** Default modification time for entries */
  modTime?: Date;

  /** Default abort signal used by streaming operations */
  signal?: AbortSignal;

  /** Default progress callback */
  onProgress?: (p: TarArchiveProgress) => void;

  /** Default throttle for progress callbacks */
  progressIntervalMs?: number;
}

/** The largest value a ustar 8-byte octal field holds: the limit for `mode`. */
const TAR_OCTAL_7_MAX = 0o7777777;

/**
 * The options of a TAR builder. A TAR mtime is any valid `Date`: one outside the octal field's
 * 1970–2242 is written in GNU base-256, which GNU tar, bsdtar and libarchive read.
 */
export const tarArchiveOptionShape = {
  modTime: dateSchema.optional(),
  ...streamOptionShape<TarArchiveProgress>()
};
const tarArchiveOptionsSchema = s.object(tarArchiveOptionShape, OPTIONS);
const _tarArchiveOptionsMatch: SchemaMatches<typeof tarArchiveOptionsSchema, TarArchiveOptions> =
  true;
void _tarArchiveOptionsMatch;

/** The options of a TAR builder (`createTar`, `packTar`). */
export function checkTarArchiveOptions(
  where: string,
  options: TarArchiveOptions | undefined
): void {
  checkArchiveArgument(tarArchiveOptionsSchema.optional(), options, where, "options");
}

/** The entry types a caller may give: everything but the GNU and PAX metadata records. */
const TAR_ENTRY_TYPES: readonly string[] = [
  TAR_TYPE.FILE,
  TAR_TYPE.FILE_OLD,
  TAR_TYPE.HARD_LINK,
  TAR_TYPE.SYMLINK,
  TAR_TYPE.CHAR_DEVICE,
  TAR_TYPE.BLOCK_DEVICE,
  TAR_TYPE.DIRECTORY,
  TAR_TYPE.FIFO,
  TAR_TYPE.CONTIGUOUS
];

/** A ustar `uname` / `gname` field: 32 bytes with a terminating NUL. */
const TAR_NAME_FIELD_MAX_BYTES = 31;

function checkTarNameField(where: string, field: string, value: string | undefined): void {
  if (value === undefined) {
    return;
  }
  if (textEncoder.encode(value).length > TAR_NAME_FIELD_MAX_BYTES) {
    throw new ArchiveArgumentError(
      field,
      `${where}: ${field} ${JSON.stringify(value)} is longer than the ${TAR_NAME_FIELD_MAX_BYTES} bytes a TAR header holds`
    );
  }
}

const tarEntryOptionsSchema = s.object(
  {
    mode: uintSchema(TAR_OCTAL_7_MAX).optional(),
    // Past 2097151 an id is written in GNU base-256, which holds every safe integer.
    uid: uintSchema(Number.MAX_SAFE_INTEGER).optional(),
    gid: uintSchema(Number.MAX_SAFE_INTEGER).optional(),
    uname: s.string().optional(),
    gname: s.string().optional(),
    modTime: dateSchema.optional(),
    mtime: dateSchema.optional(),
    type: s.oneOf(TAR_ENTRY_TYPES as readonly TarType[]).optional(),
    linkname: s.string().optional()
  },
  OPTIONS
);
const _tarEntryOptionsMatch: SchemaMatches<typeof tarEntryOptionsSchema, TarArchiveEntryOptions> =
  true;
void _tarEntryOptionsMatch;

function checkTarEntry(
  where: string,
  name: string,
  options: TarArchiveEntryOptions | undefined
): void {
  checkEntryName(where, name);
  checkArchiveArgument(tarEntryOptionsSchema.optional(), options, where, "options");
  checkTarNameField(where, "options.uname", options?.uname);
  checkTarNameField(where, "options.gname", options?.gname);
}

export interface TarArchiveEntryOptions {
  /** File mode/permissions (default: 0644 for files, 0755 for directories) */
  mode?: number;

  /** User ID (default: 0) */
  uid?: number;

  /** Group ID (default: 0) */
  gid?: number;

  /** User name */
  uname?: string;

  /** Group name */
  gname?: string;

  /** Modification time (default: archive default or now) */
  modTime?: Date;

  /** Alias for modTime - matches TAR field name */
  mtime?: Date;

  /** Entry type (auto-detected from path if not specified) */
  type?: TarType;

  /** Link target (for symlinks) */
  linkname?: string;
}

/**
 * Progress phase for TAR operations.
 */
export type TarProgressPhase = ArchiveProgressPhase;

/**
 * Progress information for TAR creation.
 */
export interface TarArchiveProgress {
  type: "tar";
  phase: TarProgressPhase;
  entriesTotal: number;
  entriesDone: number;
  bytesIn: number;
  bytesOut: number;
  currentEntry?: { name: string; index: number; bytesIn: number };
}

/**
 * Streaming options for TAR creation.
 */
export type TarArchiveStreamOptions = ArchiveStreamOptions<TarArchiveProgress>;

/**
 * Operation handle for streaming TAR creation.
 */
export type TarArchiveOperation = ArchiveOperationBase<TarArchiveProgress> & {
  /** Async iterable of TAR output chunks */
  iterable: AsyncIterable<Uint8Array<ArrayBuffer>>;
};

export interface TarReaderOptions {
  /**
   * Largest entry, in bytes, that may be read into memory (`ArchiveFileTooLargeError`). The same
   * option as ZIP's `maxEntrySize`.
   * @default 104857600 (100 MiB)
   */
  maxEntrySize?: number;

  /** @deprecated Use {@link maxEntrySize}. Removed in 1.0. */
  maxFileSize?: number;

  /** Default abort signal */
  signal?: AbortSignal;

  /** Default progress callback */
  onProgress?: (p: TarReaderProgress) => void;

  /** Default throttle for progress callbacks */
  progressIntervalMs?: number;
}

export interface TarReaderProgress {
  type: "untar";
  phase: "running" | "done" | "error" | "aborted";
  entriesTotal: number;
  entriesDone: number;
  bytesIn: number;
  bytesOut: number;
  currentEntry?: { path: string; isDirectory: boolean };
}

export interface TarReaderStreamOptions {
  signal?: AbortSignal;
  onProgress?: (p: TarReaderProgress) => void;
  progressIntervalMs?: number;
}

export interface TarReaderOperation {
  iterable: AsyncIterable<TarReaderEntry>;
  signal: AbortSignal;
  abort(reason?: unknown): void;
  pointer(): number;
  progress(): TarReaderProgress;
}

// ============================================================================
// TarArchive - For creating TAR archives
// ============================================================================

type TarInput = {
  name: string;
  source: ArchiveSource;
  options?: TarArchiveEntryOptions;
};

/**
 * TarArchive - Create TAR archives with ZIP-compatible API
 *
 * @example
 * ```ts
 * const archive = Archive.createTar();
 * archive.add("file.txt", "Hello, World!");
 * archive.add("dir/", null); // Directory
 * const bytes = await archive.toBuffer();
 * ```
 */
export class TarArchive {
  private readonly _options: TarArchiveOptions;
  private readonly _entries: TarInput[] = [];
  private _sealed = false;

  constructor(options: TarArchiveOptions = {}) {
    checkTarArchiveOptions("Archive.createTar", options);
    this._options = {
      modTime: options.modTime ?? new Date(),
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    };
  }

  /**
   * Add an entry to the archive
   */
  add(name: string, source: ArchiveSource | null, options?: TarArchiveEntryOptions): this {
    return this._addAt("Archive.TarArchive#add", name, source, options);
  }

  /** @internal {@link add}, naming `where` — the public call — in what it refuses. */
  _addAt(
    where: string,
    name: string,
    source: ArchiveSource | null,
    options?: TarArchiveEntryOptions
  ): this {
    if (this._sealed) {
      throw new ArchiveError(`${where}: cannot add entries after output has started`);
    }
    if (!name) {
      throw new ArchiveArgumentError("name", `${where}: name must be a non-empty string`);
    }
    checkTarEntry(where, name, options);
    if (source !== null) {
      checkArchiveSource(where, "source", source);
    }

    // Normalize directory names
    const isDir = source === null || name.endsWith("/");
    const normalizedName = isDir && !name.endsWith("/") ? name + "/" : name;

    this._entries.push({
      name: normalizedName,
      source: source ?? new Uint8Array(0),
      options: {
        ...options,
        type: options?.type ?? (isDir ? TAR_TYPE.DIRECTORY : TAR_TYPE.FILE),
        mode: options?.mode ?? (isDir ? DEFAULT_TAR_DIR_MODE : DEFAULT_TAR_MODE)
      }
    });

    return this;
  }

  /**
   * Add a directory entry
   */
  addDirectory(name: string, options?: Omit<TarArchiveEntryOptions, "type">): this {
    const dirName = typeof name === "string" && !name.endsWith("/") ? name + "/" : name;
    return this._addAt("Archive.TarArchive#addDirectory", dirName, null, {
      ...options,
      type: TAR_TYPE.DIRECTORY
    });
  }

  /**
   * Add a symbolic link
   */
  addSymlink(
    name: string,
    target: string,
    options?: Omit<TarArchiveEntryOptions, "type" | "linkname">
  ): this {
    checkString("Archive.TarArchive#addSymlink", "target", target);
    return this._addAt("Archive.TarArchive#addSymlink", name, new Uint8Array(0), {
      ...options,
      type: TAR_TYPE.SYMLINK,
      linkname: target,
      mode: options?.mode ?? 0o777
    });
  }

  /**
   * Generate archive as async iterable (matches ZipArchive.stream)
   */
  stream(options: TarArchiveStreamOptions = {}): AsyncIterable<Uint8Array<ArrayBuffer>> {
    return this.operation(options).iterable;
  }

  /**
   * Get operation object with abort/progress control (matches ZipArchive.operation)
   */
  operation(options: TarArchiveStreamOptions = {}): TarArchiveOperation {
    this._sealed = true;

    const signalOpt = options.signal ?? this._options.signal;
    const { controller, cleanup: cleanupAbortLink } = createLinkedAbortController(signalOpt);
    const signal = controller.signal;

    const progress: TarArchiveProgress = {
      type: "tar",
      phase: "running",
      entriesTotal: this._entries.length,
      entriesDone: 0,
      bytesIn: 0,
      bytesOut: 0
    };

    const onProgress = options.onProgress ?? this._options.onProgress;
    const entries = this._entries;
    const archiveOptions = this._options;

    async function* generate(): AsyncIterable<Uint8Array> {
      try {
        for (let i = 0; i < entries.length; i++) {
          if (signal.aborted) {
            progress.phase = "aborted";
            throw createAbortError();
          }

          const { name, source, options: entryOpts } = entries[i];

          // Get source data
          let data: Uint8Array;
          if (isInMemoryArchiveSource(source)) {
            data = await toUint8Array(source);
          } else {
            data = await collect(iterateArchiveSource(source, { signal }));
          }

          progress.bytesIn += data.length;
          progress.currentEntry = { name, index: i, bytesIn: data.length };

          // For directories and links, size is 0
          const size = isDataEntry(entryOpts?.type) ? data.length : 0;

          // Build header
          const headerOptions: TarHeaderOptions = {
            path: name,
            size,
            mode: entryOpts?.mode,
            uid: entryOpts?.uid,
            gid: entryOpts?.gid,
            mtime: entryOpts?.modTime ?? entryOpts?.mtime ?? archiveOptions.modTime,
            type: entryOpts?.type,
            linkname: entryOpts?.linkname,
            uname: entryOpts?.uname,
            gname: entryOpts?.gname
          };

          const { header, longName } = encodeHeader(headerOptions);

          // Yield long name header if needed
          if (longName) {
            yield longName;
            progress.bytesOut += longName.length;
          }

          // Yield main header
          yield header;
          progress.bytesOut += header.length;

          // Yield data and padding
          if (size > 0) {
            yield data;
            progress.bytesOut += data.length;

            const padding = createPadding(size);
            if (padding.length > 0) {
              yield padding;
              progress.bytesOut += padding.length;
            }
          }

          progress.entriesDone++;
          if (onProgress) {
            onProgress({ ...progress });
          }
        }

        // End of archive marker
        const endMarker = createEndOfArchive();
        yield endMarker;
        progress.bytesOut += endMarker.length;

        progress.phase = "done";
        if (onProgress) {
          onProgress({ ...progress });
        }
      } catch (e) {
        if (progress.phase === "running") {
          progress.phase = "error";
        }
        throw e;
      } finally {
        cleanupAbortLink();
      }
    }

    return {
      iterable: plainChunks(generate()),
      signal,
      abort(reason?: unknown) {
        controller.abort(reason);
      },
      pointer() {
        return progress.bytesOut;
      },
      progress() {
        return { ...progress };
      }
    };
  }

  /**
   * Generate archive as Uint8Array (matches ZipArchive.toBuffer)
   */
  async toBuffer(options: TarArchiveStreamOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
    return asPlainBytes(await collect(this.stream(options)));
  }

  /**
   * @deprecated Use {@link toBuffer}. Removed in 1.0.
   */
  async bytes(options: TarArchiveStreamOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
    return this.toBuffer(options);
  }

  /**
   * @deprecated Use {@link toBufferSync}. Removed in 1.0.
   */
  bytesSync(): Uint8Array<ArrayBuffer> {
    return asPlainBytes(this._buildSync("bytesSync"));
  }

  /**
   * Generate archive synchronously (matches ZipArchive.toBufferSync)
   */
  toBufferSync(): Uint8Array<ArrayBuffer> {
    return asPlainBytes(this._buildSync("toBufferSync"));
  }

  /** Build synchronously; `method` names the public entry point in error messages. */
  protected _buildSync(method: string): Uint8Array {
    this._sealed = true;

    const chunks: Uint8Array[] = [];

    for (const { name, source, options } of this._entries) {
      // Get source data synchronously
      let data: Uint8Array;
      if (isSyncArchiveSource(source)) {
        data = toUint8ArraySync(source);
      } else {
        throw new ArchiveArgumentError(
          "source",
          `Archive.TarArchive#${method}: a source must be bytes or a string to be written synchronously (a Blob or a stream needs toBuffer)`
        );
      }

      const size = isDataEntry(options?.type) ? data.length : 0;

      const headerOptions: TarHeaderOptions = {
        path: name,
        size,
        mode: options?.mode,
        uid: options?.uid,
        gid: options?.gid,
        mtime: options?.modTime ?? options?.mtime ?? this._options.modTime,
        type: options?.type,
        linkname: options?.linkname,
        uname: options?.uname,
        gname: options?.gname
      };

      const { header, longName } = encodeHeader(headerOptions);

      if (longName) {
        chunks.push(longName);
      }

      chunks.push(header);

      if (size > 0) {
        chunks.push(data);

        const padding = createPadding(size);
        if (padding.length > 0) {
          chunks.push(padding);
        }
      }
    }

    // End of archive
    chunks.push(createEndOfArchive());

    return concatUint8Arrays(chunks);
  }

  /**
   * Pipe archive to sink (matches ZipArchive.pipeTo)
   */
  async pipeTo(sink: ArchiveSink, options: TarArchiveStreamOptions = {}): Promise<void> {
    checkSinkArgument("Archive.TarArchive#pipeTo", sink);
    await pipeIterableToSink(this.stream(options), sink);
  }
}

// ============================================================================
// TarReaderEntry - Entry from TAR archive (matches UnzipEntry interface)
// ============================================================================

export interface TarReaderEntry {
  readonly path: string;
  readonly isDirectory: boolean;
  readonly info: TarEntryInfo;
  /** Get entry data as Uint8Array (matches UnzipEntry.bytes) */
  bytes(): Promise<Uint8Array<ArrayBuffer>>;
  /** Get entry data as string (matches UnzipEntry.text) */
  text(encoding?: string): Promise<string>;
  /** Stream entry data (matches UnzipEntry.stream) */
  stream(): AsyncIterable<Uint8Array<ArrayBuffer>>;
  /** Entry data as a Web `ReadableStream` (matches UnzipEntry.readableStream) */
  readableStream(): ReadableStream<Uint8Array>;
  /**
   * Pipe entry to a Web `WritableStream`, with its `pipeTo` options (matches
   * UnzipEntry.pipeTo)
   */
  pipeTo(sink: WritableStream<Uint8Array>, options?: ArchivePipeToOptions): Promise<void>;
  /** Pipe entry to sink (matches UnzipEntry.pipeTo) */
  pipeTo(sink: ArchiveSink): Promise<void>;
  /** Discard entry (matches UnzipEntry.discard) */
  discard(): void;
}

export function createTarReaderEntry(info: TarEntryInfo, data: Uint8Array): TarReaderEntry {
  if (typeof info !== "object" || info === null || typeof info.path !== "string") {
    throw new ArchiveArgumentError(
      "info",
      "Archive.createTarReaderEntry: info must be a TarEntryInfo with a string path"
    );
  }
  checkBytes("Archive.createTarReaderEntry", "data", data);
  const entry: TarReaderEntry = {
    info,
    path: info.path,
    isDirectory: isDirectory(info) || info.path.endsWith("/"),
    async bytes(): Promise<Uint8Array<ArrayBuffer>> {
      return asPlainBytes(data);
    },
    async text(encoding?: string): Promise<string> {
      return decodeEntryText(data, encoding);
    },
    async *stream(): AsyncIterable<Uint8Array<ArrayBuffer>> {
      if (data.length > 0) {
        yield asPlainBytes(data);
      }
    },
    readableStream(): ReadableStream<Uint8Array> {
      return asyncIterableToReadableStream(entry.stream());
    },
    async pipeTo(sink: ArchiveSink, options?: ArchivePipeToOptions): Promise<void> {
      checkSinkArgument("Archive.TarReaderEntry#pipeTo", sink);
      if (isWritableStream(sink)) {
        await entry.readableStream().pipeTo(sink, options);
        return;
      }
      await pipeIterableToSink(entry.stream(), sink);
    },
    discard(): void {
      // No-op for TAR (data already loaded)
    }
  };
  return entry;
}

// ============================================================================
// TarReader - For reading TAR archives (matches ZipReader interface)
// ============================================================================

/**
 * Reads a TAR archive (`Archive.openTar` returns one).
 *
 * Over a one-shot source — a `ReadableStream`, an async iterable or a Node stream — the order of
 * calls matters, because the source can be read only once:
 *
 * - `list()`, `get()` or `bytes()` first reads the whole archive and keeps its entries. Every
 *   later call, `entries()` included, works from what was kept, as often as asked, until
 *   `close()` drops it.
 * - Iterating `entries()`, `entriesStream()` or `operation().iterable` first streams the source,
 *   keeping nothing, and spends it: any later call throws an `ArchiveError` ("already
 *   consumed"). Only starting the iteration reads; asking for the iterable does not.
 *
 * Bytes, a string or a `Blob` are read again by every call and are unaffected, and
 * `Archive.openTarGz` keeps the decompressed archive, so its calls work in any order.
 */
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging
export class TarReader {
  private readonly _source: ArchiveSource | ArchiveBytes;
  private readonly _options: TarReaderOptions;
  private _parsedEntries: TarReaderEntry[] | null = null;
  // Bumped by close(): a pass that started before it must not refill the cache it cleared
  private _generation = 0;
  // A one-shot source (a stream) can be read by one pass; `get` / `list` keep what that pass read
  private _sourceRead = false;

  constructor(source: ArchiveSource | ArchiveBytes, options: TarReaderOptions = {}) {
    this._source = source;
    this._options = options;
  }

  /**
   * Iterate over entries (matches `ZipReader.entries`). Over a one-shot source that `list` / `get`
   * / `bytes` did not read first, iterating streams the source and spends it (see
   * {@link TarReader}).
   */
  entries(options: TarReaderStreamOptions = {}): AsyncIterable<TarReaderEntry> {
    return this._operation(options, "entries").iterable;
  }

  /**
   * The entries as a Web `ReadableStream` (matches `ZipReader.entriesStream`); reading it spends
   * a one-shot source as {@link TarReader.entries} does.
   */
  entriesStream(options: TarReaderStreamOptions = {}): ReadableStream<TarReaderEntry> {
    const op = this._operation(options, "entriesStream");
    return asyncIterableToReadableStream(op.iterable, reason => {
      op.abort(reason ?? "cancelled");
    });
  }

  /**
   * Get operation object with abort/progress control (matches `ZipReader.operation`); iterating
   * its `iterable` spends a one-shot source as {@link TarReader.entries} does.
   */
  operation(options: TarReaderStreamOptions = {}): TarReaderOperation {
    return this._operation(options, "operation");
  }

  private _operation(options: TarReaderStreamOptions, where: string): TarReaderOperation {
    const signalOpt = options.signal ?? this._options.signal;
    const { controller, cleanup: cleanupAbortLink } = createLinkedAbortController(signalOpt);
    const signal = controller.signal;

    const progress: TarReaderProgress = {
      type: "untar",
      phase: "running",
      entriesTotal: 0,
      entriesDone: 0,
      bytesIn: 0,
      bytesOut: 0
    };

    const onProgress = options.onProgress ?? this._options.onProgress;
    const resolveSource = () => this._resolveSource(signal);
    const claimOneShotSource = (): void => {
      if (this._sourceRead) {
        throw oneShotSourceReadError(`Archive.TarReader#${where}`);
      }
      this._sourceRead = true;
    };
    // The entries a `get` / `list` over a one-shot source kept: the source itself is spent
    const kept = isInMemoryArchiveSource(this._source) ? null : this._parsedEntries;
    const parseOptions: TarParseOptions = {
      maxEntrySize: entrySizeLimit(this._options),
      signal
    };

    // Every entry of one complete pass, cached for get() / list()
    const collected: TarReaderEntry[] = [];
    const generation = this._generation;
    const setParsedEntries = (entries: TarReaderEntry[]) => {
      if (this._generation === generation) {
        this._parsedEntries = entries;
      }
    };

    async function* generate(): AsyncIterable<TarReaderEntry> {
      try {
        if (kept !== null) {
          for (const entry of kept) {
            throwIfAborted(signal);
            progress.entriesTotal++;
            progress.bytesOut += entry.info.size;
            progress.currentEntry = { path: entry.path, isDirectory: entry.isDirectory };
            yield entry;
            progress.entriesDone++;
            if (onProgress) {
              onProgress({ ...progress });
            }
          }
          progress.phase = "done";
          if (onProgress) {
            onProgress({ ...progress });
          }
          return;
        }
        const source = await resolveSource();
        if (!isInMemoryArchiveSource(source)) {
          claimOneShotSource();
        }
        if (isInMemoryArchiveSource(source)) {
          // Buffer mode - parse all at once
          const data = await toUint8Array(source);
          progress.bytesIn = data.length;

          const entries = parseTar(data, parseOptions);
          progress.entriesTotal = entries.length;

          for (const entry of entries) {
            if (signal.aborted) {
              progress.phase = "aborted";
              throw createAbortError();
            }

            const entryData = await entry.bytes();
            progress.bytesOut += entryData.length;
            progress.currentEntry = {
              path: entry.info.path,
              isDirectory: isDirectory(entry.info)
            };

            const readerEntry = createTarReaderEntry(entry.info, entryData);

            collected.push(readerEntry);

            yield readerEntry;

            progress.entriesDone++;
            if (onProgress) {
              onProgress({ ...progress });
            }
          }
          setParsedEntries(collected);
        } else {
          // Streaming mode — entries are not retained, so memory stays bounded
          for await (const entry of readTarStream(
            source as AsyncIterable<Uint8Array>,
            parseOptions
          )) {
            if (signal.aborted) {
              progress.phase = "aborted";
              throw createAbortError();
            }

            const entryData = await entry.bytes();
            progress.bytesOut += entryData.length;
            progress.entriesTotal++;
            progress.currentEntry = {
              path: entry.info.path,
              isDirectory: isDirectory(entry.info)
            };

            const readerEntry = createTarReaderEntry(entry.info, entryData);
            yield readerEntry;

            progress.entriesDone++;
            if (onProgress) {
              onProgress({ ...progress });
            }
          }
        }

        progress.phase = "done";
        if (onProgress) {
          onProgress({ ...progress });
        }
      } catch (e) {
        if (progress.phase === "running") {
          progress.phase = "error";
        }
        throw e;
      } finally {
        cleanupAbortLink();
      }
    }

    return {
      iterable: generate(),
      signal,
      abort(reason?: unknown) {
        controller.abort(reason);
      },
      pointer() {
        return progress.bytesIn;
      },
      progress() {
        return { ...progress };
      }
    };
  }

  /**
   * Get a specific entry by path (matches `ZipReader.get`); `undefined` when the archive has no such
   * entry. Of two members with one name the later is returned, as `tar -x` and `unpackTar` keep
   * it. Parses (and caches) the whole archive on first call — which is what lets a one-shot
   * source be read again by later calls; after an iteration has spent it, throws an
   * `ArchiveError` (see {@link TarReader}).
   */
  async get(path: string): Promise<TarReaderEntry | undefined> {
    checkEntryPath("Archive.TarReader#get", path);
    return this._get(path, "get");
  }

  /**
   * List all entry paths, in archive order (matches `ZipReader.list`). Parses and caches the
   * archive as {@link TarReader.get} does.
   */
  async list(): Promise<string[]> {
    return (await this._ensureParsed("list")).map(e => e.path);
  }

  /**
   * Read one entry's bytes (matches `ZipReader.bytes`); `undefined` when the archive has no such
   * entry. Parses and caches the archive as {@link TarReader.get} does.
   */
  async bytes(path: string): Promise<Uint8Array<ArrayBuffer> | undefined> {
    checkEntryPath("Archive.TarReader#bytes", path);
    // Through `get`, so a subclass that overrides it is honoured
    const entry = await this.get(path);
    return entry ? entry.bytes() : undefined;
  }

  private async _get(path: string, where: string): Promise<TarReaderEntry | undefined> {
    return (await this._ensureParsed(where)).findLast(e => e.path === path);
  }

  /**
   * The source each pass parses. A subclass that must transform a one-shot source (such as
   * gunzipping it) overrides this to do so once and return a replayable result.
   */
  protected _resolveSource(
    _signal: AbortSignal
  ): ArchiveSource | ArchiveBytes | Promise<ArchiveSource | ArchiveBytes> {
    return this._source;
  }

  private async _ensureParsed(where: string): Promise<TarReaderEntry[]> {
    const cached = this._parsedEntries;
    if (cached) {
      return cached;
    }
    // A buffered source caches itself on a complete pass; a streaming one does not
    // (iterating it must not retain every entry), so collect here.
    const generation = this._generation;
    const all: TarReaderEntry[] = [];
    for await (const entry of this._operation({}, where).iterable) {
      all.push(entry);
    }
    if (this._generation !== generation) {
      // close() ran during the pass: answer this call, but keep the cache it released empty
      return all;
    }
    return (this._parsedEntries ??= all);
  }

  /**
   * Close reader and release resources (matches ZipReader.close)
   */
  async close(): Promise<void> {
    // No persistent resources to release for TAR; drop the cache, including what an
    // in-flight pass would otherwise store when it finishes
    this._generation++;
    this._parsedEntries = null;
  }
}

// ============================================================================
// Factory functions
// ============================================================================

/**
 * Create a new TAR archive builder
 */
export function createTarArchive(options?: TarArchiveOptions): TarArchive {
  return new TarArchive(options);
}

/**
 * Create a new TAR reader
 */
export function createTarReader(source: ArchiveSource, options?: TarReaderOptions): TarReader {
  return new TarReader(source, options);
}

/**
 * Open a TAR archive for reading, checking the source and options when it is opened.
 * Published as `Archive.openTar`.
 *
 * Over a one-shot source (a `ReadableStream`, an async iterable or a Node stream) the first call
 * decides how it is read. `list()`, `get()` or `bytes()` reads it once and keeps the entries
 * for every later call, `entries()` included, until `close()`. Iterating `entries()`,
 * `entriesStream()` or `operation().iterable` streams it instead, keeping nothing, and spends it:
 * any later call throws an `ArchiveError`. Bytes, a string or a `Blob` are unaffected, and
 * `Archive.openTarGz` replays its source in any order. See `TarReader`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function openTarReader(
  source: ArchiveSource | ArchiveBytes,
  options?: TarReaderOptions
): TarReader {
  checkTarReaderArguments("Archive.openTar", source, options);
  return new TarReader(source, options);
}

const tarReaderOptionsSchema = s.object(
  {
    maxEntrySize: limitSchema.optional(),
    maxFileSize: limitSchema.optional(),
    ...streamOptionShape<TarReaderProgress>()
  },
  OPTIONS
);
const _tarReaderOptionsMatch: SchemaMatches<typeof tarReaderOptionsSchema, TarReaderOptions> = true;
void _tarReaderOptionsMatch;

/** What `openTar` / `openTarGz` are given, checked when the reader is opened. */
export function checkTarReaderArguments(
  where: string,
  source: unknown,
  options: TarReaderOptions | undefined
): void {
  checkSourceArgument(where, source);
  checkArchiveArgument(tarReaderOptionsSchema.optional(), options, where, "options");
}

// ============================================================================
// Convenience functions for quick TAR creation
// ============================================================================

/**
 * One entry for the one-shot builders (`packZip`, `packTar`, `packTarGz`, …): an object
 * `{ name, source }` or a `[name, source]` pair.
 */
export type TarEntryInput = { name: string; source: ArchiveSource } | [string, ArchiveSource];

/** Helper to add entries to archive from Map or Iterable */
/**
 * Walk the `entries` a `pack*` function was given — a map of name to source (a `Map` or any
 * `ReadonlyMap`, read by iteration), or an
 * iterable of `[name, source]` pairs or `{ name, source }` objects — handing each to `add`.
 * Anything else is an `ArchiveArgumentError` naming `entries` or the entry at fault.
 */
export function forEachEntryInput(
  where: string,
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  add: (name: string, source: ArchiveSource) => void
): void {
  if (entries instanceof Map) {
    for (const [name, source] of entries) {
      add(name, source);
    }
    return;
  }
  if (
    entries === null ||
    typeof entries !== "object" ||
    typeof (entries as Iterable<unknown>)[Symbol.iterator] !== "function"
  ) {
    throw new ArchiveArgumentError(
      "entries",
      `${where}: entries must be a Map or an iterable of [name, source] pairs or { name, source } objects, got ${
        entries === null ? "null" : typeof entries
      }`
    );
  }
  let index = 0;
  for (const entry of entries as Iterable<unknown>) {
    if (Array.isArray(entry)) {
      add(entry[0], entry[1]);
    } else if (entry !== null && typeof entry === "object" && "name" in entry) {
      const { name, source } = entry as { name: string; source: ArchiveSource };
      add(name, source);
    } else {
      throw new ArchiveArgumentError(
        `entries[${index}]`,
        `${where}: entries[${index}] must be a [name, source] pair or a { name, source } object`
      );
    }
    index++;
  }
}

export function addEntries(
  archive: TarArchive,
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  where = "Archive.packTar"
): void {
  forEachEntryInput(where, entries, (name, source) => archive._addAt(where, name, source));
}

/**
 * Create a TAR archive from entries (async)
 */
export async function tar(
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  options?: TarArchiveOptions
): Promise<Uint8Array<ArrayBuffer>> {
  checkTarArchiveOptions("Archive.packTar", options);
  const archive = new TarArchive(options);
  addEntries(archive, entries);
  return archive.toBuffer();
}

/**
 * Create a TAR archive from entries (sync)
 */
export function tarSync(
  entries: ReadonlyMap<string, string | Uint8Array | ArrayBuffer> | Iterable<TarEntryInput>,
  options?: TarArchiveOptions
): Uint8Array<ArrayBuffer> {
  checkTarArchiveOptions("Archive.packTarSync", options);
  const archive = new TarArchive(options);
  addEntries(archive, entries, "Archive.packTarSync");
  return archive.toBufferSync();
}

// The type of what `defineAsyncDispose` installs below: `await using` disposes of a `TarReader`.
// oxlint-disable-next-line typescript/no-empty-interface, typescript/no-empty-object-type, typescript/no-unsafe-declaration-merging
export interface TarReader extends AsyncDisposeMember {}

defineAsyncDispose(TarReader.prototype, function () {
  return this.close();
});
