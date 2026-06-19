/**
 * TAR Archive - Unified API compatible with ZIP
 *
 * Provides TarArchive and TarReader classes with the same interface
 * as ZipArchive and ZipReader, allowing seamless format switching.
 */

import { ArchiveError, createAbortError, createLinkedAbortController } from "@archive/core/errors";
import type {
  ArchiveProgressPhase,
  ArchiveStreamOptions,
  ArchiveOperationBase
} from "@archive/core/progress";
import type { ArchiveSink } from "@archive/io/archive-sink";
import { collect, pipeIterableToSink } from "@archive/io/archive-sink";
import type { ArchiveSource } from "@archive/io/archive-source";
import { toAsyncIterable, toUint8Array, isInMemoryArchiveSource } from "@archive/io/archive-source";
import { concatUint8Arrays, textEncoder, getTextDecoder } from "@utils/binary";

import type { TarType } from "./tar-constants";
import { TAR_TYPE, DEFAULT_TAR_MODE, DEFAULT_TAR_DIR_MODE } from "./tar-constants";
import type { TarEntryInfo } from "./tar-entry-info";
import { isDataEntry, isDirectory } from "./tar-entry-info";
import type { TarHeaderOptions } from "./tar-header";
import { encodeHeader, createPadding, createEndOfArchive } from "./tar-header";
import type { TarParseOptions } from "./tar-parser";
import { parseTar, parseTarStream } from "./tar-parser";

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
  iterable: AsyncIterable<Uint8Array>;
};

export interface TarReaderOptions {
  /** Maximum file size to extract into memory (default: 100MB) */
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
 * const archive = new TarArchive();
 * archive.add("file.txt", "Hello, World!");
 * archive.add("dir/", null); // Directory
 * const bytes = await archive.bytes();
 * ```
 */
export class TarArchive {
  private readonly _options: TarArchiveOptions;
  private readonly _entries: TarInput[] = [];
  private _sealed = false;

  constructor(options: TarArchiveOptions = {}) {
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
    if (this._sealed) {
      throw new ArchiveError("Cannot add entries after output has started");
    }
    if (!name) {
      throw new ArchiveError("Entry name is required");
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
    return this.add(name.endsWith("/") ? name : name + "/", null, {
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
    return this.add(name, new Uint8Array(0), {
      ...options,
      type: TAR_TYPE.SYMLINK,
      linkname: target,
      mode: options?.mode ?? 0o777
    });
  }

  /**
   * Generate archive as async iterable (matches ZipArchive.stream)
   */
  stream(options: TarArchiveStreamOptions = {}): AsyncIterable<Uint8Array> {
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
            data = await collect(toAsyncIterable(source, { signal }));
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
      iterable: generate(),
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
   * Generate archive as Uint8Array (matches ZipArchive.bytes)
   */
  async bytes(options: TarArchiveStreamOptions = {}): Promise<Uint8Array> {
    return collect(this.stream(options));
  }

  /**
   * Generate archive synchronously (matches ZipArchive.bytesSync)
   */
  bytesSync(): Uint8Array {
    this._sealed = true;

    const chunks: Uint8Array[] = [];

    for (const { name, source, options } of this._entries) {
      // Get source data synchronously
      let data: Uint8Array;
      if (source instanceof Uint8Array) {
        data = source;
      } else if (typeof source === "string") {
        data = textEncoder.encode(source);
      } else if (source instanceof ArrayBuffer) {
        data = new Uint8Array(source);
      } else {
        throw new ArchiveError("bytesSync() only supports Uint8Array/ArrayBuffer/string sources");
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
    await pipeIterableToSink(this.stream(options), sink);
  }
}

// ============================================================================
// TarReaderEntry - Entry from TAR archive (matches UnzipEntry interface)
// ============================================================================

export class TarReaderEntry {
  readonly path: string;
  readonly isDirectory: boolean;
  readonly info: TarEntryInfo;

  private readonly _data: Uint8Array;

  constructor(info: TarEntryInfo, data: Uint8Array) {
    this.info = info;
    this.path = info.path;
    this.isDirectory = isDirectory(info) || info.path.endsWith("/");
    this._data = data;
  }

  /**
   * Get entry data as Uint8Array (matches UnzipEntry.bytes)
   */
  async bytes(): Promise<Uint8Array> {
    return this._data;
  }

  /**
   * Get entry data as string (matches UnzipEntry.text)
   */
  async text(encoding?: string): Promise<string> {
    return getTextDecoder(encoding).decode(this._data);
  }

  /**
   * Stream entry data (matches UnzipEntry.stream)
   */
  async *stream(): AsyncIterable<Uint8Array> {
    if (this._data.length > 0) {
      yield this._data;
    }
  }

  /**
   * Pipe entry to sink (matches UnzipEntry.pipeTo)
   */
  async pipeTo(sink: ArchiveSink): Promise<void> {
    await pipeIterableToSink(this.stream(), sink);
  }

  /**
   * Discard entry (matches UnzipEntry.discard)
   */
  discard(): void {
    // No-op for TAR (data already loaded)
  }
}

// ============================================================================
// TarReader - For reading TAR archives (matches ZipReader interface)
// ============================================================================

export class TarReader {
  private readonly _source: ArchiveSource;
  private readonly _options: TarReaderOptions;
  private _parsedEntries: TarReaderEntry[] | null = null;

  constructor(source: ArchiveSource, options: TarReaderOptions = {}) {
    this._source = source;
    this._options = options;
  }

  /**
   * Iterate over entries (matches ZipReader.entries)
   */
  entries(options: TarReaderStreamOptions = {}): AsyncIterable<TarReaderEntry> {
    return this.operation(options).iterable;
  }

  /**
   * Get operation object with abort/progress control (matches ZipReader.operation)
   */
  operation(options: TarReaderStreamOptions = {}): TarReaderOperation {
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
    const source = this._source;
    const parseOptions: TarParseOptions = {
      maxFileSize: this._options.maxFileSize,
      signal
    };

    // Capture reference for caching parsed entries
    const parsedEntriesRef = { entries: this._parsedEntries };
    const setParsedEntries = (entries: TarReaderEntry[]) => {
      this._parsedEntries = entries;
      parsedEntriesRef.entries = entries;
    };

    async function* generate(): AsyncIterable<TarReaderEntry> {
      try {
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

            const entryData = await entry.data();
            progress.bytesOut += entryData.length;
            progress.currentEntry = {
              path: entry.info.path,
              isDirectory: isDirectory(entry.info)
            };

            const readerEntry = new TarReaderEntry(entry.info, entryData);

            // Cache for get() method
            if (!parsedEntriesRef.entries) {
              setParsedEntries([]);
            }
            parsedEntriesRef.entries!.push(readerEntry);

            yield readerEntry;

            progress.entriesDone++;
            if (onProgress) {
              onProgress({ ...progress });
            }
          }
        } else {
          // Streaming mode
          for await (const entry of parseTarStream(
            source as AsyncIterable<Uint8Array>,
            parseOptions
          )) {
            if (signal.aborted) {
              progress.phase = "aborted";
              throw createAbortError();
            }

            const entryData = await entry.data();
            progress.bytesOut += entryData.length;
            progress.entriesTotal++;
            progress.currentEntry = {
              path: entry.info.path,
              isDirectory: isDirectory(entry.info)
            };

            const readerEntry = new TarReaderEntry(entry.info, entryData);
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
   * Get a specific entry by path (matches ZipReader.get)
   */
  async get(path: string): Promise<TarReaderEntry | null> {
    // If we haven't parsed yet, do so now
    if (!this._parsedEntries) {
      // Consume entries() which will cache them in _parsedEntries
      for await (const _ of this.entries()) {
        // entries() already caches to _parsedEntries
      }
    }

    return this._parsedEntries!.find(e => e.path === path) ?? null;
  }

  /**
   * List all entry paths (matches common pattern)
   */
  async list(): Promise<string[]> {
    if (!this._parsedEntries) {
      // Consume entries() which will cache them in _parsedEntries
      for await (const _ of this.entries()) {
        // entries() already caches to _parsedEntries
      }
    }
    return this._parsedEntries!.map(e => e.path);
  }

  /**
   * Get bytes of a specific entry (matches ZipReader.bytes)
   */
  async bytes(path: string): Promise<Uint8Array | null> {
    const entry = await this.get(path);
    if (!entry) {
      return null;
    }
    return entry.bytes();
  }

  /**
   * Close reader and release resources (matches ZipReader.close)
   */
  async close(): Promise<void> {
    // No persistent resources to release for TAR
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

// ============================================================================
// Convenience functions for quick TAR creation
// ============================================================================

export type TarEntryInput = { name: string; source: ArchiveSource } | [string, ArchiveSource];

/** Helper to add entries to archive from Map or Iterable */
export function addEntries(
  archive: TarArchive,
  entries: Map<string, any> | Iterable<TarEntryInput>
): void {
  if (entries instanceof Map) {
    for (const [name, source] of entries) {
      archive.add(name, source);
    }
  } else {
    for (const entry of entries) {
      if (Array.isArray(entry)) {
        archive.add(entry[0], entry[1]);
      } else {
        archive.add(entry.name, entry.source);
      }
    }
  }
}

/**
 * Create a TAR archive from entries (async)
 */
export async function tar(
  entries: Map<string, ArchiveSource> | Iterable<TarEntryInput>,
  options?: TarArchiveOptions
): Promise<Uint8Array> {
  const archive = new TarArchive(options);
  addEntries(archive, entries);
  return archive.bytes();
}

/**
 * Create a TAR archive from entries (sync)
 */
export function tarSync(
  entries: Map<string, string | Uint8Array | ArrayBuffer> | Iterable<TarEntryInput>,
  options?: TarArchiveOptions
): Uint8Array {
  const archive = new TarArchive(options);
  addEntries(archive, entries);
  return archive.bytesSync();
}
