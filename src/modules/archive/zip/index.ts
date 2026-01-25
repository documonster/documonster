import type { ZipTimestampMode } from "@archive/utils/timestamps";
import { DEFAULT_ZIP_LEVEL, DEFAULT_ZIP_TIMESTAMPS } from "@archive/defaults";
import { ZipDeflateFile, StreamingZip } from "@archive/zip/stream";
import { createZip, createZipSync } from "@archive/zip/zip-bytes";
import { collect, pipeIterableToSink, type ArchiveSink } from "@archive/io/archive-sink";
import {
  toAsyncIterable,
  toUint8Array,
  toUint8ArraySync,
  type ArchiveSource
} from "@archive/io/archive-source";
import { createAsyncQueue } from "@archive/utils/async-queue";

const REPRODUCIBLE_ZIP_MOD_TIME = new Date(1980, 0, 1, 0, 0, 0);

export interface ZipOptions {
  level?: number;
  timestamps?: ZipTimestampMode;
  comment?: string;

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
   */
  noSort?: boolean;
}

export interface ZipEntryOptions {
  level?: number;
  modTime?: Date;
  comment?: string;
}

type ZipInput = {
  name: string;
  source: ArchiveSource;
  options?: ZipEntryOptions;
};

export class ZipArchive {
  private readonly _options: Required<Pick<ZipOptions, "level" | "timestamps">> & {
    comment?: string;
    modTime: Date;
    smartStore: boolean;
    noSort: boolean;
  };
  private readonly _entries: ZipInput[] = [];
  private _sealed = false;

  constructor(options: ZipOptions = {}) {
    const reproducible = options.reproducible ?? false;
    this._options = {
      level: options.level ?? DEFAULT_ZIP_LEVEL,
      timestamps: options.timestamps ?? (reproducible ? "dos" : DEFAULT_ZIP_TIMESTAMPS),
      comment: options.comment,
      modTime: options.modTime ?? (reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date()),
      smartStore: options.smartStore ?? true,
      noSort: options.noSort ?? false
    };
  }

  add(name: string, source: ArchiveSource, options?: ZipEntryOptions): this {
    if (this._sealed) {
      throw new Error("Cannot add entries after output has started");
    }
    if (!name) {
      throw new Error("Entry name is required");
    }
    this._entries.push({ name, source, options });
    return this;
  }

  stream(): AsyncIterable<Uint8Array> {
    this._sealed = true;

    const queue = createAsyncQueue<Uint8Array>();

    const zip = new StreamingZip(
      (err, data, final) => {
        if (err) {
          queue.fail(err);
          return;
        }
        if (data.length) {
          queue.push(data);
        }
        if (final) {
          queue.close();
        }
      },
      { comment: this._options.comment }
    );

    (async () => {
      try {
        for (const entry of this._entries) {
          const level = entry.options?.level ?? this._options.level;
          const file = new ZipDeflateFile(entry.name, {
            level,
            modTime: entry.options?.modTime ?? this._options.modTime,
            timestamps: this._options.timestamps,
            comment: entry.options?.comment,
            smartStore: this._options.smartStore
          });

          zip.add(file);

          // Feed data
          if (
            entry.source instanceof Uint8Array ||
            entry.source instanceof ArrayBuffer ||
            typeof entry.source === "string" ||
            (typeof Blob !== "undefined" && entry.source instanceof Blob)
          ) {
            const bytes = await toUint8Array(entry.source as any);
            await file.push(bytes, true);
          } else {
            for await (const chunk of toAsyncIterable(entry.source)) {
              await file.push(chunk, false);
            }
            await file.push(new Uint8Array(0), true);
          }

          await file.complete();
        }

        zip.end();
      } catch (e) {
        queue.fail(e instanceof Error ? e : new Error(String(e)));
      }
    })();

    return queue.iterable;
  }

  async bytes(): Promise<Uint8Array> {
    this._sealed = true;

    const allSourcesInMemory = this._entries.every(
      e =>
        e.source instanceof Uint8Array ||
        e.source instanceof ArrayBuffer ||
        typeof e.source === "string" ||
        (typeof Blob !== "undefined" && e.source instanceof Blob)
    );
    const hasBlobSource = this._entries.some(
      e => typeof Blob !== "undefined" && e.source instanceof Blob
    );

    // Fast-path: when all sources are already in memory and there are no
    // per-entry compression overrides, use the single-buffer ZIP builder.
    // This avoids the overhead of chunking + collecting from the streaming writer.
    if (allSourcesInMemory) {
      // Prefer the sync builder when possible (Node.js hot path): it avoids
      // async/Promise overhead and uses zlib sync fast paths.
      if (!hasBlobSource) {
        const entries = this._entries.map(e => ({
          name: e.name,
          data: toUint8ArraySync(e.source as any),
          level: e.options?.level,
          modTime: e.options?.modTime,
          comment: e.options?.comment
        }));

        return createZipSync(entries, {
          level: this._options.level,
          timestamps: this._options.timestamps,
          modTime: this._options.modTime,
          comment: this._options.comment,
          smartStore: this._options.smartStore,
          noSort: this._options.noSort
        });
      }

      const entries = await Promise.all(
        this._entries.map(async e => ({
          name: e.name,
          data: await toUint8Array(e.source as any),
          level: e.options?.level,
          modTime: e.options?.modTime,
          comment: e.options?.comment
        }))
      );

      return createZip(entries, {
        level: this._options.level,
        timestamps: this._options.timestamps,
        modTime: this._options.modTime,
        comment: this._options.comment,
        smartStore: this._options.smartStore,
        noSort: this._options.noSort
      });
    }

    return collect(this.stream());
  }

  bytesSync(): Uint8Array {
    this._sealed = true;

    const entries = this._entries.map(e => {
      if (
        !(e.source instanceof Uint8Array) &&
        !(e.source instanceof ArrayBuffer) &&
        typeof e.source !== "string"
      ) {
        throw new Error("bytesSync() only supports Uint8Array/ArrayBuffer/string sources");
      }
      return {
        name: e.name,
        data: toUint8ArraySync(e.source as any),
        modTime: e.options?.modTime,
        comment: e.options?.comment
      };
    });

    return createZipSync(entries, {
      level: this._options.level,
      timestamps: this._options.timestamps,
      modTime: this._options.modTime,
      comment: this._options.comment,
      smartStore: this._options.smartStore,
      noSort: this._options.noSort
    });
  }

  async pipeTo(sink: ArchiveSink): Promise<void> {
    await pipeIterableToSink(this.stream(), sink);
  }
}

export function zip(options?: ZipOptions): ZipArchive {
  return new ZipArchive(options);
}
