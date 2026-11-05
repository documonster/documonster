import { hasDeflateRawCompressionStream } from "@archive/compression/compress.base";
import { crc32Finalize, crc32Update } from "@archive/compression/crc32";
import { ByteQueue } from "@archive/core/byte-queue";
import {
  DEFAULT_ZIP_LEVEL,
  DEFAULT_ZIP_TIMESTAMPS,
  REPRODUCIBLE_ZIP_MOD_TIME
} from "@archive/core/defaults";
import {
  checkArchiveSource,
  checkEntryName,
  checkString,
  checkZipEntryOptions,
  checkZipOptions
} from "@archive/core/entry-checks";
import { ArchiveArgumentError, throwIfAborted, ArchiveError } from "@archive/core/errors";
import type { ZipStringEncoding } from "@archive/core/text";
import { encodeZipString } from "@archive/core/text";
import type { ArchiveFormat } from "@archive/core/types";
import type { ArchiveSink } from "@archive/io/archive-sink";
import { checkSinkArgument, collect, pipeIterableToSink } from "@archive/io/archive-sink";
import type { ArchiveSource } from "@archive/io/archive-source";
import {
  iterateArchiveSource,
  toUint8Array,
  toUint8ArraySync,
  isSyncArchiveSource,
  isInMemoryArchiveSource
} from "@archive/io/archive-source";
import type { TarArchiveProgress } from "@archive/tar/tar-archive";
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import { normalizeZipPath } from "@archive/zip-spec/zip-path";
import type { Zip64Mode } from "@archive/zip-spec/zip-records";
import {
  buildDataDescriptor,
  FLAG_DATA_DESCRIPTOR,
  UINT16_MAX,
  UINT32_MAX,
  writeLocalFileHeaderInto
} from "@archive/zip-spec/zip-records";
import type { ZipOperation, ZipProgress, ZipStreamOptions } from "@archive/zip/progress";
import { ZipDeflateFile } from "@archive/zip/stream";
import type { ZipCentralDirectoryEntryInput } from "@archive/zip/writer-core";
import {
  measureCentralDirectoryAndEocd,
  writeCentralDirectoryAndEocdInto
} from "@archive/zip/writer-core";
import { createZip, createZipSync } from "@archive/zip/zip-bytes";
import type { ZipEntry } from "@archive/zip/zip-bytes";
import { resolveZipExternalAttributesAndVersionMadeBy } from "@archive/zip/zip-entry-attributes";
import { buildZipEntryMetadata } from "@archive/zip/zip-entry-metadata";
import { buildZipDeflateFileOptions } from "@archive/zip/zip-entry-options";
import { createZipOperation } from "@archive/zip/zip-output-pipeline";
import { stringToUint8Array as encodeUtf8 } from "@utils/binary";
import { isNode } from "@utils/env";

/** Archive options */
export interface ZipOptions {
  /**
   * Archive format: "zip" (default) or "tar".
   * Read only by the deprecated `Archive.zip`; `Archive.createZip` and `Archive.createTar` name the format instead.
   */
  format?: ArchiveFormat;

  level?: number;
  timestamps?: ZipTimestampMode;
  comment?: string;

  /** Optional entry name normalization. `false` keeps names as-is. */
  path?: false | ZipPathOptions;

  /** Optional string encoding for entry names/comments and archive comment. */
  encoding?: ZipStringEncoding;

  /** Default abort signal used by streaming operations. */
  signal?: AbortSignal;

  /** Default progress callback used by streaming operations. */
  onProgress?: (p: ZipProgress) => void;

  /** Default throttle for progress callbacks. */
  progressIntervalMs?: number;

  /**
   * ZIP64 mode:
   * - "auto" (default): write ZIP64 only when required by limits.
   * - true: force ZIP64 structures even for small archives.
   * - false: forbid ZIP64; throws if ZIP64 is required.
   */
  zip64?: Zip64Mode;

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
   * If true (default), automatically STORE incompressible data.
   * If false, always follow `level` (DEFLATE when level > 0).
   */
  smartStore?: boolean;

  /**
   * If true, entries are written in their original input order.
   * If false (default), entries are sorted alphabetically by name.
   *
   * Note: streaming output preserves the input order.
   */
  noSort?: boolean;
}

export interface ZipEntryOptions {
  level?: number;
  modTime?: Date;
  atime?: Date;
  ctime?: Date;
  birthTime?: Date;
  comment?: string;

  /** Optional Unix mode/permissions for this entry. */
  mode?: number;

  /** Optional MS-DOS attributes (low 8 bits). */
  msDosAttributes?: number;

  /** Advanced override for external attributes. */
  externalAttributes?: number;

  /** Advanced override for versionMadeBy. */
  versionMadeBy?: number;

  /** Per-entry ZIP64 override. Defaults to the archive-level zip64 mode. */
  zip64?: Zip64Mode;

  /** Optional string encoding for this entry name/comment. */
  encoding?: ZipStringEncoding;
}

export type { ZipOperation, ZipProgress, ZipStreamOptions } from "@archive/zip/progress";

type ZipInput = {
  name: string;
  source: ArchiveSource;
  options?: ZipEntryOptions;
};

/**
 * Decide whether {@link ZipArchive.bytes} may take the browser Blob fast path,
 * which streams each source instead of materialising every Blob through
 * `Blob.arrayBuffer()`.
 *
 * A pure function because `isNode` is one of its terms, so the decision cannot
 * be reached from a Node test otherwise.
 *
 * `hasDeflateRawCompressionStream` is a thunk so a STORE-only archive never
 * consults it: the probe constructs a `CompressionStream`, and an archive with
 * `needsDeflate: false` neither needs the codec nor may be refused on it.
 * Chromium 80–102 has the class but not the `deflate-raw` format.
 */
export function canUseBrowserBlobFastPath(facts: {
  isNode: boolean;
  hasBlobSource: boolean;
  allSourcesInMemory: boolean;
  smartStore: boolean;
  zip64Requested: boolean;
  needsDeflate: boolean;
  hasDeflateRawCompressionStream: () => boolean;
}): boolean {
  if (facts.isNode || !facts.hasBlobSource || !facts.allSourcesInMemory) {
    return false;
  }
  if (facts.smartStore || facts.zip64Requested) {
    return false;
  }
  return !facts.needsDeflate || facts.hasDeflateRawCompressionStream();
}

export class ZipArchive {
  private readonly _options: Required<Pick<ZipOptions, "level" | "timestamps">> & {
    comment?: string;
    modTime: Date;
    smartStore: boolean;
    zip64: Zip64Mode;
    path: false | ZipPathOptions;
    encoding?: ZipStringEncoding;
    noSort: boolean;
  };
  private readonly _streamDefaults: {
    signal?: AbortSignal;
    onProgress?: (p: ZipProgress) => void;
    progressIntervalMs?: number;
  };
  private readonly _entries: ZipInput[] = [];
  // The names entries are stored under, to refuse a second entry of one name
  private readonly _names = new Set<string>();
  private _sealed = false;

  constructor(options: ZipOptions = {}) {
    checkZipOptions("Archive.createZip", options);
    const reproducible = options.reproducible ?? false;
    this._options = {
      level: options.level ?? DEFAULT_ZIP_LEVEL,
      timestamps: options.timestamps ?? (reproducible ? "dos" : DEFAULT_ZIP_TIMESTAMPS),
      comment: options.comment,
      modTime: options.modTime ?? (reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date()),
      smartStore: options.smartStore ?? true,
      zip64: options.zip64 ?? "auto",
      path: options.path ?? false,
      encoding: options.encoding,
      noSort: options.noSort ?? false
    };
    this._streamDefaults = {
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    };
  }

  /**
   * The in-memory builder's form of an entry, carrying every {@link ZipEntryOptions} field the
   * streaming writer reads (see `buildZipDeflateFileOptions`), so `toBuffer()`,
   * `toBufferSync()` and `stream()` write the same headers.
   */
  private _toBuildEntry(entry: ZipInput, data: Uint8Array): ZipEntry {
    const options = entry.options;
    return {
      name: entry.name,
      data,
      level: options?.level,
      modTime: options?.modTime,
      atime: options?.atime,
      ctime: options?.ctime,
      birthTime: options?.birthTime,
      comment: options?.comment,
      encoding: options?.encoding ?? this._options.encoding,
      mode: options?.mode,
      msDosAttributes: options?.msDosAttributes,
      externalAttributes: options?.externalAttributes,
      versionMadeBy: options?.versionMadeBy,
      zip64: options?.zip64
    };
  }

  private _getCreateZipOptions() {
    return {
      level: this._options.level,
      timestamps: this._options.timestamps,
      modTime: this._options.modTime,
      comment: this._options.comment,
      smartStore: this._options.smartStore,
      zip64: this._options.zip64,
      encoding: this._options.encoding,
      // The in-memory fast path writes through here; without it, `path` applied to
      // `stream()` and was ignored by `toBuffer()` / `toBufferSync()`.
      path: this._options.path,
      noSort: this._options.noSort
    };
  }

  /**
   * Add an entry. A ZIP holds one entry per name, so a name already added is refused
   * (`ArchiveArgumentError`, field `name`): readers disagree about which of two same-named
   * entries they return, and an editor keeps only one.
   */
  add(name: string, source: ArchiveSource, options?: ZipEntryOptions): this {
    return this._addAt("Archive.ZipArchive#add", name, source, options);
  }

  /** @internal {@link add}, naming `where` — the public call — in what it refuses. */
  _addAt(where: string, name: string, source: ArchiveSource, options?: ZipEntryOptions): this {
    if (this._sealed) {
      throw new ArchiveError(`${where}: cannot add entries after output has started`);
    }
    if (!name) {
      throw new ArchiveArgumentError("name", `${where}: name must be a non-empty string`);
    }
    if (this._options.path === false) {
      checkEntryName(where, name);
    }
    checkArchiveSource(where, "source", source);
    checkZipEntryOptions(where, options);
    const stored = this._options.path ? normalizeZipPath(name, this._options.path) : name;
    if (this._names.has(stored)) {
      throw new ArchiveArgumentError(
        "name",
        `${where}: the archive already has an entry named ${JSON.stringify(stored)}; a ZIP holds one entry per name`
      );
    }
    this._names.add(stored);
    this._entries.push({ name, source, options });
    return this;
  }

  /**
   * Add a directory entry.
   *
   * Unified API consistent with TarArchive.
   */
  addDirectory(name: string, options?: Omit<ZipEntryOptions, "level">): this {
    const dirName = typeof name === "string" && !name.endsWith("/") ? name + "/" : name;
    return this._addAt("Archive.ZipArchive#addDirectory", dirName, new Uint8Array(0), {
      ...options,
      level: 0
    });
  }

  /**
   * Add a symbolic link entry.
   *
   * Unified API consistent with TarArchive.
   */
  addSymlink(name: string, target: string, options?: Omit<ZipEntryOptions, "level">): this {
    checkString("Archive.ZipArchive#addSymlink", "target", target);
    return this._addAt("Archive.ZipArchive#addSymlink", name, encodeUtf8(target), {
      ...options,
      level: 0,
      mode: options?.mode ?? 0o120777
    });
  }

  stream(options: ZipStreamOptions = {}): AsyncIterable<Uint8Array> {
    return this.operation(options).iterable;
  }

  operation(options: ZipStreamOptions = {}): ZipOperation {
    this._sealed = true;

    const signalOpt = options.signal ?? this._streamDefaults.signal;
    const onProgress = options.onProgress ?? this._streamDefaults.onProgress;
    const progressIntervalMs =
      options.progressIntervalMs ?? this._streamDefaults.progressIntervalMs;

    return createZipOperation(
      this._entries.length,
      {
        comment: this._options.comment,
        zip64: this._options.zip64,
        encoding: this._options.encoding
      },
      { signal: signalOpt, onProgress, progressIntervalMs },
      async ({ zip, signal, progress }) => {
        for (let i = 0; i < this._entries.length; i++) {
          throwIfAborted(signal);

          const entry = this._entries[i]!;

          let entryBytesIn = 0;
          progress.update({ currentEntry: { name: entry.name, index: i, bytesIn: 0 } });

          const file = new ZipDeflateFile(
            entry.name,
            buildZipDeflateFileOptions(entry.options, {
              level: this._options.level,
              modTime: this._options.modTime,
              timestamps: this._options.timestamps,
              smartStore: this._options.smartStore,
              zip64: this._options.zip64,
              path: this._options.path,
              encoding: this._options.encoding
            })
          );

          zip.add(file);

          const onChunk = (chunk: Uint8Array) => {
            entryBytesIn += chunk.length;
            progress.mutate(s => {
              s.bytesIn += chunk.length;
              s.currentEntry = { name: entry.name, index: i, bytesIn: entryBytesIn };
            });
          };

          if (isSyncArchiveSource(entry.source)) {
            const bytes = toUint8ArraySync(entry.source);
            throwIfAborted(signal);
            onChunk(bytes);
            await file.push(bytes, true);
            await file.outputCheckpoint();
          } else {
            for await (const chunk of iterateArchiveSource(entry.source, { signal, onChunk })) {
              throwIfAborted(signal);
              await file.push(chunk, false);
              await file.outputCheckpoint();
            }
            throwIfAborted(signal);
            await file.push(new Uint8Array(0), true);
            await file.outputCheckpoint();
          }

          await file.complete();
          progress.set("entriesDone", progress.snapshot.entriesDone + 1);
        }

        throwIfAborted(signal);
        zip.end();
      }
    );
  }

  /**
   * Browser-only fast path: stream sources through CompressionStream while
   * computing CRC incrementally. Avoids Blob.arrayBuffer() materialization.
   */
  private async _browserStreamingBytes(): Promise<Uint8Array> {
    const out = new ByteQueue();
    const cdEntries: ZipCentralDirectoryEntryInput[] = [];

    if (this._entries.length > UINT16_MAX) {
      throw new Error("Too many entries for non-ZIP64 fast path");
    }

    for (const input of this._entries) {
      const level = input.options?.level ?? this._options.level;
      const deflate = level > 0;

      const name = this._options.path
        ? normalizeZipPath(input.name, this._options.path)
        : input.name;
      const attrs = resolveZipExternalAttributesAndVersionMadeBy({
        name,
        mode: input.options?.mode,
        msDosAttributes: input.options?.msDosAttributes,
        externalAttributes: input.options?.externalAttributes,
        versionMadeBy: input.options?.versionMadeBy
      });
      const meta = buildZipEntryMetadata({
        name,
        comment: input.options?.comment,
        modTime: input.options?.modTime ?? this._options.modTime,
        atime: input.options?.atime,
        ctime: input.options?.ctime,
        birthTime: input.options?.birthTime,
        timestamps: this._options.timestamps,
        useDataDescriptor: true,
        deflate,
        codec: input.options?.encoding ?? this._options.encoding
      });

      const localHeaderOffset = out.length;
      if (localHeaderOffset > UINT32_MAX) {
        throw new Error("ZIP64 required for offsets");
      }

      const localHeader = new Uint8Array(30 + meta.nameBytes.length + meta.extraField.length);
      const view = new DataView(localHeader.buffer, localHeader.byteOffset, localHeader.byteLength);
      writeLocalFileHeaderInto(localHeader, view, 0, {
        fileName: meta.nameBytes,
        extraField: meta.extraField,
        flags: meta.flags,
        compressionMethod: meta.compressionMethod,
        dosTime: meta.dosTime,
        dosDate: meta.dosDate,
        crc32: 0,
        compressedSize: 0,
        uncompressedSize: 0
      });
      out.append(localHeader);

      let crcState = 0xffffffff;
      let uncompressedSize = 0;
      let compressedSize = 0;

      if (!deflate) {
        for await (const chunk of iterateArchiveSource(input.source)) {
          if (chunk.length === 0) {
            continue;
          }
          crcState = crc32Update(crcState, chunk);
          uncompressedSize += chunk.length;
          compressedSize += chunk.length;
          out.append(chunk);
        }
      } else {
        const cs = new CompressionStream("deflate-raw");
        const writer = cs.writable.getWriter();
        const reader = cs.readable.getReader();

        const readPromise = (async () => {
          while (true) {
            const { done, value } = await reader.read();
            if (done) {
              return;
            }
            compressedSize += value.length;
            out.append(value);
          }
        })();

        try {
          for await (const chunk of iterateArchiveSource(input.source)) {
            if (chunk.length === 0) {
              continue;
            }
            crcState = crc32Update(crcState, chunk);
            uncompressedSize += chunk.length;
            // `Uint8Array<ArrayBufferLike>` is a valid `BufferSource` at runtime;
            // the cast only reconciles the lib's `ArrayBufferLike` backing-store
            // generic with `BufferSource`'s `ArrayBuffer`.
            await writer.write(chunk as BufferSource);
          }
          await writer.close();
          await readPromise;
        } finally {
          try {
            writer.releaseLock();
          } catch {
            // ignore
          }
          try {
            reader.releaseLock();
          } catch {
            // ignore
          }
        }
      }

      const crc32 = crc32Finalize(crcState);

      if (uncompressedSize > UINT32_MAX || compressedSize > UINT32_MAX || out.length > UINT32_MAX) {
        throw new Error("ZIP64 required for sizes");
      }

      out.append(buildDataDescriptor(crc32, compressedSize, uncompressedSize));

      cdEntries.push({
        fileName: meta.nameBytes,
        extraField: meta.extraField,
        comment: meta.commentBytes,
        flags: (meta.flags | FLAG_DATA_DESCRIPTOR) >>> 0,
        crc32,
        compressedSize,
        uncompressedSize,
        compressionMethod: meta.compressionMethod,
        dosTime: meta.dosTime,
        dosDate: meta.dosDate,
        localHeaderOffset,
        zip64: false,
        externalAttributes: attrs.externalAttributes,
        versionMadeBy: attrs.versionMadeBy
      });
    }

    const centralDirOffset = out.length;
    if (centralDirOffset > UINT32_MAX) {
      throw new Error("ZIP64 required for central directory offset");
    }

    const zipComment = encodeZipString(this._options.comment, this._options.encoding);
    const cdSizing = measureCentralDirectoryAndEocd(cdEntries, {
      zipComment,
      zip64Mode: this._options.zip64,
      centralDirOffset
    });

    const cdChunk = new Uint8Array(cdSizing.totalSize);
    writeCentralDirectoryAndEocdInto(cdEntries, {
      zipComment,
      zip64Mode: this._options.zip64,
      centralDirOffset,
      out: cdChunk,
      offset: 0
    });
    out.append(cdChunk);

    return out.read(out.length);
  }

  /**
   * Build the archive and return the complete ZIP bytes.
   */
  async toBuffer(options: ZipStreamOptions = {}): Promise<Uint8Array> {
    this._sealed = true;

    const signalOpt = options.signal ?? this._streamDefaults.signal;
    const onProgress = options.onProgress ?? this._streamDefaults.onProgress;

    // If progress/abort is requested, prefer the streaming pipeline.
    if (onProgress || signalOpt) {
      return collect(this.stream(options));
    }

    const allSourcesInMemory = this._entries.every(e => isInMemoryArchiveSource(e.source));

    const hasBlobSource =
      typeof Blob !== "undefined" && this._entries.some(e => e.source instanceof Blob);

    // Browser fast path: stream Blob sources instead of materialising them.
    // `needsDeflate` mirrors the per-entry `level > 0` test inside
    // `_browserStreamingBytes`, which is the only thing there that reaches
    // for `CompressionStream("deflate-raw")`.
    const needsDeflate = this._entries.some(e => (e.options?.level ?? this._options.level) > 0);

    if (
      canUseBrowserBlobFastPath({
        isNode: isNode(),
        hasBlobSource,
        allSourcesInMemory,
        smartStore: this._options.smartStore,
        // This path writes no ZIP64 records, so any entry asking for them goes elsewhere.
        zip64Requested:
          this._options.zip64 === true || this._entries.some(e => e.options?.zip64 === true),
        needsDeflate,
        hasDeflateRawCompressionStream
      })
    ) {
      try {
        return await this._browserStreamingBytes();
      } catch {
        // Fall back to the buffered builder
      }
    }

    // Fast-path: when all sources are already in memory, use single-buffer ZIP builder.
    if (allSourcesInMemory) {
      // Prefer the sync builder when possible (Node.js hot path): it avoids
      // async/Promise overhead and uses zlib sync fast paths.
      if (!hasBlobSource) {
        const entries = this._entries.map(entry =>
          this._toBuildEntry(entry, toUint8ArraySync(toSyncSource(entry.source)))
        );

        return createZipSync(entries, this._getCreateZipOptions());
      }

      const entries = await Promise.all(
        this._entries.map(async entry =>
          this._toBuildEntry(entry, await toUint8Array(toInMemorySource(entry.source)))
        )
      );

      return createZip(entries, this._getCreateZipOptions());
    }

    return collect(this.stream());
  }

  /**
   * @deprecated Use {@link toBuffer}. Removed in 1.0.
   */
  async bytes(options: ZipStreamOptions = {}): Promise<Uint8Array> {
    return this.toBuffer(options);
  }

  /**
   * @deprecated Use {@link toBufferSync}. Removed in 1.0.
   */
  bytesSync(): Uint8Array {
    return this._buildSync("bytesSync");
  }

  /**
   * Build the archive synchronously (in-memory sources only).
   */
  toBufferSync(): Uint8Array {
    return this._buildSync("toBufferSync");
  }

  /** Build synchronously; `method` names the public entry point in error messages. */
  protected _buildSync(method: string): Uint8Array {
    this._sealed = true;

    const entries = this._entries.map(entry => {
      if (!isSyncArchiveSource(entry.source)) {
        throw new ArchiveArgumentError(
          "source",
          `Archive.ZipArchive#${method}: a source must be bytes or a string to be written synchronously (a Blob or a stream needs toBuffer)`
        );
      }
      return this._toBuildEntry(entry, toUint8ArraySync(entry.source));
    });

    return createZipSync(entries, this._getCreateZipOptions());
  }

  async pipeTo(sink: ArchiveSink, options: ZipStreamOptions = {}): Promise<void> {
    checkSinkArgument("Archive.ZipArchive#pipeTo", sink);
    await pipeIterableToSink(this.stream(options), sink);
  }
}

/**
 * Narrow an {@link ArchiveSource} to a synchronously-resolvable source.
 *
 * Only called on code paths that have already verified all sources are sync
 * (no Blob/stream), so a failure here indicates an internal invariant break.
 */
function toSyncSource(source: ArchiveSource): Uint8Array | ArrayBuffer | string {
  if (isSyncArchiveSource(source)) {
    return source;
  }
  throw new ArchiveArgumentError(
    "source",
    "Expected a synchronous archive source (Uint8Array/ArrayBuffer/string)"
  );
}

/**
 * Narrow an {@link ArchiveSource} to an in-memory source (sync types or Blob).
 */
function toInMemorySource(source: ArchiveSource): Uint8Array | ArrayBuffer | string | Blob {
  if (isInMemoryArchiveSource(source)) {
    return source;
  }
  throw new ArchiveArgumentError(
    "source",
    "Expected an in-memory archive source (Uint8Array/ArrayBuffer/string/Blob)"
  );
}

/** ZIP options with format: "tar" */
export interface ZipOptionsTar extends Omit<ZipOptions, "onProgress"> {
  format: "tar";

  /**
   * Default progress callback used by streaming operations.
   *
   * TAR archives emit {@link TarArchiveProgress} rather than {@link ZipProgress}.
   */
  onProgress?: (p: TarArchiveProgress) => void;
}

/** ZIP options with format: "zip" (or default) */
export interface ZipOptionsZip extends ZipOptions {
  format?: "zip";
}
