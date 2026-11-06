/**
 * TAR + Gzip (tar.gz / tgz) Support
 *
 * Provides utilities for creating and extracting gzip-compressed TAR archives.
 * Uses the unified compress module for compression/decompression.
 */

import { gunzip, gunzipSync } from "@archive/compression/compress";
import { createGunzipStream } from "@archive/compression/streaming-compress";
import { OPTIONS, checkArchiveArgument, levelSchema } from "@archive/core/arguments";
import { DEFAULT_COMPRESS_LEVEL } from "@archive/core/defaults";
import { ArchiveArgumentError, ArchiveError } from "@archive/core/errors";
import { collect } from "@archive/io/archive-sink";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  archiveBytesToUint8Array,
  isArchiveBytes,
  isInMemoryArchiveSource,
  iterateArchiveSource,
  toUint8Array
} from "@archive/io/archive-source";
import { gzipTarChunks, gzipTarStream, gzipTarSync } from "@archive/tar/gzip-tar-stream";
import type { TarArchiveOptions, TarEntryInput, TarReaderOptions } from "@archive/tar/tar-archive";
import {
  TarArchive,
  TarReader,
  addEntries,
  tarArchiveOptionShape,
  checkTarReaderArguments
} from "@archive/tar/tar-archive";
import type { TarEntry, TarParseOptions, UnpackedTarEntry } from "@archive/tar/tar-parser";
import { parseTar, readTarStream, untar, untarSync } from "@archive/tar/tar-parser";
import { createAbortError, isAbortError, throwIfAborted } from "@utils/errors";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

export interface TarGzOptions extends TarArchiveOptions {
  /** Compression level (0-9, default: 6) */
  level?: number;
}

const tarGzOptionsSchema = s.object(
  { ...tarArchiveOptionShape, level: levelSchema.optional() },
  OPTIONS
);
const _tarGzOptionsMatch: SchemaMatches<typeof tarGzOptionsSchema, TarGzOptions> = true;
void _tarGzOptionsMatch;

function checkTarGzOptions(where: string, options: TarGzOptions | undefined): void {
  checkArchiveArgument(tarGzOptionsSchema.optional(), options, where, "options");
}

// Use TarParseOptions directly since we don't need gzip-specific parse options
export type ParseTarGzOptions = TarParseOptions;

/**
 * TarGz Archive Builder
 *
 * Creates gzip-compressed TAR archives (.tar.gz / .tgz)
 *
 * @example
 * ```ts
 * const archive = Archive.createTarGz({ level: 6 });
 * archive.add("file.txt", "Hello, World!");
 * const bytes = await archive.toBuffer();
 * ```
 */
export class TarGzArchive extends TarArchive {
  private readonly _gzLevel: number;

  constructor(options: TarGzOptions = {}) {
    checkTarGzOptions("Archive.createTarGz", options);
    super(options);
    this._gzLevel = options.level ?? DEFAULT_COMPRESS_LEVEL;
  }

  /**
   * Generate compressed archive as async iterable (true streaming)
   */
  override async *stream(): AsyncIterable<Uint8Array> {
    yield* gzipTarStream(super.stream(), this._gzLevel);
  }

  /**
   * Compress the synchronously built TAR, for both `toBufferSync` and `bytesSync`. In Node
   * the bytes are those `stream()` and `toBuffer()` produce.
   */
  protected override _buildSync(method: string): Uint8Array {
    return gzipTarSync(super._buildSync(method), this._gzLevel);
  }
}

/** Create a {@link TarGzArchive} builder. Published as `Archive.createTarGz`. */
export function createTarGzArchive(options?: TarGzOptions): TarGzArchive {
  return new TarGzArchive(options);
}

/**
 * Create a gzip-compressed TAR archive. Accepts the same entries as `Archive.packTar`.
 * Published as `Archive.packTarGz`.
 */
export async function targz(
  entries: ReadonlyMap<string, ArchiveSource> | Iterable<TarEntryInput>,
  options: TarGzOptions = {}
): Promise<Uint8Array> {
  checkTarGzOptions("Archive.packTarGz", options);
  const archive = new TarGzArchive(options);
  addEntries(archive, entries, "Archive.packTarGz");
  return archive.toBuffer();
}

/**
 * Helper to get compressed data from any archive source
 */
async function getCompressedData(
  source: ArchiveSource | ArchiveBytes,
  signal?: AbortSignal
): Promise<Uint8Array> {
  if (isInMemoryArchiveSource(source)) {
    return toUint8Array(source);
  }
  return collect(iterateArchiveSource(source, { signal }));
}

/**
 * Parse a gzip-compressed TAR archive
 *
 * @param source - Compressed archive data. A string is the archive's content, encoded as UTF-8 —
 *   never a file path or URL; read a file's bytes first (in Node, or open it with
 *   `ArchiveFile.fromFile`).
 * @param options - Parse options
 * @returns Array of TAR entries
 */
export async function parseTarGz(
  source: ArchiveSource | ArchiveBytes,
  options: ParseTarGzOptions = {}
): Promise<TarEntry[]> {
  checkTarReaderArguments("Archive.parseTarGz", source, options);
  const compressed = await getCompressedData(source, options.signal);
  return parseTar(await gunzipTar(compressed), options);
}

/**
 * Decompress a gzip source incrementally — the compressed input is never held whole.
 */
function gunzipChunks(
  source: ArchiveSource | ArchiveBytes,
  signal?: AbortSignal
): AsyncIterable<Uint8Array> {
  return {
    [Symbol.asyncIterator]: () =>
      gzipTarChunks(iterateArchiveSource(source, { signal }), createGunzipStream())[
        Symbol.asyncIterator
      ]()
  };
}

/**
 * Report a decoder failure as an {@link ArchiveError}. Errors already in the archive's
 * vocabulary — TAR errors, limits — and aborts pass through unchanged.
 */
function toGzipError(error: unknown): unknown {
  if (error instanceof ArchiveError || isAbortError(error) || !(error instanceof Error)) {
    return error;
  }
  return new ArchiveError(`Invalid gzip data: ${error.message}`, { cause: error });
}

/** Decompress a whole `.tar.gz`, reporting a damaged gzip stream as an {@link ArchiveError}. */
async function gunzipTar(compressed: Uint8Array): Promise<Uint8Array> {
  try {
    return await gunzip(compressed);
  } catch (error) {
    throw toGzipError(error);
  }
}

/** @internal */
export function gunzipTarSync(compressed: Uint8Array): Uint8Array {
  try {
    return gunzipSync(compressed);
  } catch (error) {
    throw toGzipError(error);
  }
}

/** `iterator` as an iterable that leaves closing it to its owner. */
function borrow<T>(iterator: AsyncIterator<T>): AsyncIterable<T> {
  return { [Symbol.asyncIterator]: () => ({ next: () => iterator.next() }) };
}

/**
 * Parse a gzip-compressed TAR archive as an async iterable, decompressing the whole archive
 * before the first entry is yielded, so a damaged archive fails before iteration starts —
 * the same result as {@link parseTarGz}. Kept for the deprecated flat `parseTarGzStream`;
 * `Archive.parseTarGzStream` is {@link parseTarGzIncremental}.
 */
export async function* parseTarGzStream(
  source: ArchiveSource | ArchiveBytes,
  options: ParseTarGzOptions = {}
): AsyncGenerator<TarEntry> {
  yield* await parseTarGz(source, options);
}

/**
 * Parse a gzip-compressed TAR archive incrementally: decompression and parsing proceed
 * together, so memory stays bounded by the largest entry rather than the archive — as
 * `parseTarStream` reads a plain TAR. A gzip or TAR error surfaces when iteration reaches
 * it (after earlier entries have been yielded); iterate {@link parseTarGz} instead to fail
 * before the first entry. An archive cut off inside an entry throws. Once the TAR ends, the
 * rest of the gzip stream is still read, so a damaged or missing gzip trailer is reported
 * too; a damaged gzip stream is reported as an `ArchiveError`. Published as
 * `Archive.parseTarGzStream`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function parseTarGzIncremental(
  source: ArchiveSource | ArchiveBytes,
  options: ParseTarGzOptions = {}
): AsyncGenerator<TarEntry> {
  checkTarReaderArguments("Archive.parseTarGzStream", source, options);
  return readTarGzStream(source, options);
}

async function* readTarGzStream(
  source: ArchiveSource | ArchiveBytes,
  options: ParseTarGzOptions
): AsyncGenerator<TarEntry> {
  const inflated = gunzipChunks(source, options.signal)[Symbol.asyncIterator]();
  let drained = false;
  try {
    yield* readTarStream(borrow(inflated), options);
    while (!(await inflated.next()).done) {
      // Padding after the end-of-archive blocks; only the gzip stream's end matters
    }
    drained = true;
  } catch (error) {
    throw toGzipError(error);
  } finally {
    if (!drained) {
      await inflated.return?.();
    }
  }
}

/**
 * Extract gzip-compressed TAR archive to Map. The whole archive is decompressed in memory;
 * for large files iterate {@link parseTarGzIncremental} (`Archive.parseTarGzStream`) instead.
 *
 * @param source - Compressed archive data. A string is the archive's content, encoded as UTF-8 —
 *   never a file path or URL; read a file's bytes first (in Node, or open it with
 *   `ArchiveFile.fromFile`).
 * @param options - Parse options
 * @returns Map of path → { info, data }; each `data` is a view into the decompressed buffer
 */
export async function untargz(
  source: ArchiveSource | ArchiveBytes,
  options: ParseTarGzOptions = {}
): Promise<Map<string, UnpackedTarEntry>> {
  checkTarReaderArguments("Archive.unpackTarGz", source, options);
  const compressed = await getCompressedData(source, options.signal);
  return untar(await gunzipTar(compressed), options);
}

/**
 * Create a gzip-compressed TAR archive synchronously from in-memory entries. Published as
 * `Archive.packTarGzSync`.
 */
export function targzSync(
  entries: ReadonlyMap<string, string | Uint8Array | ArrayBuffer> | Iterable<TarEntryInput>,
  options: TarGzOptions = {}
): Uint8Array {
  checkTarGzOptions("Archive.packTarGzSync", options);
  const archive = new TarGzArchive(options);
  addEntries(archive, entries, "Archive.packTarGzSync");
  return archive.toBufferSync();
}

/**
 * Extract an in-memory gzip-compressed TAR archive synchronously. Each `data` is a view
 * into the decompressed buffer. Published as `Archive.unpackTarGzSync`.
 */
export function untargzSync(
  source: ArchiveBytes,
  options: Omit<ParseTarGzOptions, "signal"> = {}
): Map<string, UnpackedTarEntry> {
  if (!isArchiveBytes(source)) {
    throw new ArchiveArgumentError(
      "source",
      `Archive.unpackTarGzSync: source must be bytes (an ArrayBufferView or ArrayBuffer), got ${
        source === null ? "null" : typeof source
      }; use Archive.unpackTarGz for other sources`
    );
  }
  checkTarReaderArguments("Archive.unpackTarGzSync", source, options);
  return untarSync(gunzipTarSync(archiveBytesToUint8Array(source)), options);
}

/** The one decompression every pass of a {@link TarGzReader} shares. */
interface SharedLoad {
  readonly result: Promise<Uint8Array>;
  /** Cancels the read; only used once no caller is waiting any more. */
  readonly controller: AbortController;
  waiting: number;
  /**
   * Set synchronously when the read settles, before any caller's callback runs, so an
   * abort landing between the two never cancels (and discards) a finished read.
   */
  settled: boolean;
}

/**
 * A {@link TarReader} over a gzip source, decompressed once on first use and kept until
 * `close()`.
 *
 * The decompression belongs to the reader, not to the call that started it: each caller
 * races its own signal against the shared result, so one caller aborting never fails
 * another. The read of a one-shot source is cancelled only once every waiting caller has
 * aborted before it completed (also after `close()`), and later calls then fail with an
 * `ArchiveError` (not an `AbortError`); a read that has already completed is never discarded,
 * so aborting then only rejects that caller. An in-memory source is never cancelled (its read cannot be). A failure is
 * cached for a one-shot source — the stream is consumed or errored and cannot be read
 * again — so every later call rethrows the original error; an in-memory source is read
 * afresh on the next call.
 */
class TarGzReader extends TarReader {
  private readonly _compressed: ArchiveSource | ArchiveBytes;
  private readonly _replayable: boolean;
  private _load: SharedLoad | undefined;
  private _closed = false;
  private _cancelled = false;

  constructor(source: ArchiveSource | ArchiveBytes, options: TarReaderOptions) {
    super(source, options);
    this._compressed = source;
    this._replayable = isInMemoryArchiveSource(source);
  }

  protected override _resolveSource(signal: AbortSignal): Promise<Uint8Array> {
    if (this._closed && !this._replayable) {
      return Promise.reject(
        new ArchiveError("TAR.GZ reader is closed and its one-shot source cannot be read again")
      );
    }
    if (signal.aborted) {
      return Promise.reject(createAbortError(signal.reason));
    }
    if (this._cancelled) {
      return Promise.reject(
        new ArchiveError(
          "TAR.GZ source was cancelled by an earlier caller and its one-shot stream cannot be read again"
        )
      );
    }
    const load = (this._load ??= this._startLoad());
    load.waiting++;
    return new Promise<Uint8Array>((resolve, reject) => {
      const onAbort = (): void => {
        settle();
        reject(createAbortError(signal.reason));
        if (load.waiting === 0 && !load.settled && !this._replayable) {
          this._cancelLoad(load, signal.reason);
        }
      };
      let settled = false;
      const settle = (): void => {
        if (!settled) {
          settled = true;
          load.waiting--;
          signal.removeEventListener("abort", onAbort);
        }
      };
      signal.addEventListener("abort", onAbort, { once: true });
      load.result.then(
        tar => {
          settle();
          resolve(tar);
        },
        (error: unknown) => {
          settle();
          reject(error);
        }
      );
    });
  }

  /** Cancel a one-shot read nobody waits for any more; later callers must not join it. */
  private _cancelLoad(load: SharedLoad, reason?: unknown): void {
    if (this._load === load) {
      this._load = undefined;
      this._cancelled = true;
    }
    load.controller.abort(reason);
  }

  private _startLoad(): SharedLoad {
    const controller = new AbortController();
    // An in-memory read ignores the signal, so it is never cancelled
    // A cancelled read may end early rather than throw; report it as the abort it was
    const read = getCompressedData(this._compressed, controller.signal).then(
      compressed => {
        throwIfAborted(controller.signal);
        return gunzipTar(compressed);
      },
      (error: unknown) => {
        throwIfAborted(controller.signal);
        throw error;
      }
    );
    const result = read.then(
      tar => {
        load.settled = true;
        return tar;
      },
      (error: unknown) => {
        load.settled = true;
        throw error;
      }
    );
    const load: SharedLoad = { result, controller, waiting: 0, settled: false };
    result.catch(() => {
      // A replayable source is read afresh next time; a one-shot one keeps its error
      if (this._replayable && this._load === load) {
        this._load = undefined;
      }
    });
    return load;
  }

  /** Release the decompressed archive. A one-shot source cannot be read afterwards. */
  override async close(): Promise<void> {
    this._closed = true;
    const load = this._load;
    this._load = undefined;
    // Callers still waiting keep the read; the last of them to abort cancels it
    if (load && load.waiting === 0 && !load.settled && !this._replayable) {
      load.controller.abort();
    }
    await super.close();
  }
}

/**
 * Open a gzip-compressed TAR archive for reading — the `.tar.gz` counterpart of
 * `Archive.openTar`. The source is read and decompressed once, on first use, and the
 * **whole decompressed TAR is held in memory** until `close()` (as are the parsed entries
 * once `get` or `list` has run), so `entries()`, `get` and `list` can be
 * called in any order and any number of times even over a one-shot source such as a
 * `ReadableStream` or a `fetch` body. Memory is therefore proportional to the
 * uncompressed archive; to process a large `.tar.gz` in memory bounded by its largest
 * entry, iterate `Archive.parseTarGzStream` instead.
 *
 * Concurrent calls share one decompression, and a caller's `signal` cancels only that
 * caller — with one exception for a one-shot source: if every caller waiting on its
 * read aborts before decompression completes, the stream is cancelled, and later calls
 * throw an `ArchiveError` — not an `AbortError` — because it cannot be read again. Once
 * decompression has completed, an abort only rejects that caller. An in-memory source is never cancelled. A failed read of a one-shot source is final — later calls rethrow the same
 * error, since the stream cannot be read twice — while an in-memory source is retried.
 * After `close()` an in-memory source is decompressed again on next use; a one-shot
 * source throws an `ArchiveError`.
 * Published as `Archive.openTarGz`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function openTarGzReader(
  source: ArchiveSource | ArchiveBytes,
  options: TarReaderOptions = {}
): TarReader {
  checkTarReaderArguments("Archive.openTarGz", source, options);
  return new TarGzReader(source, options);
}
