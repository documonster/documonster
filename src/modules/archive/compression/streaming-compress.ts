/**
 * Node.js True Streaming Compression
 *
 * Uses zlib.createDeflateRaw() with explicit flush() calls for real chunk-by-chunk streaming.
 * Each write() immediately produces compressed output without waiting for end().
 */

import {
  createDeflateRaw,
  createInflateRaw,
  createGzip,
  createGunzip,
  createDeflate,
  createInflate,
  deflateRawSync,
  constants,
  type Gunzip,
  type Inflate
} from "zlib";
import { Transform, type TransformCallback } from "@stream";

import { DEFAULT_COMPRESS_LEVEL } from "@archive/shared/defaults";

export type {
  DeflateStream,
  InflateStream,
  StreamCompressOptions,
  StreamingCodec
} from "@archive/compression/streaming-compress.base";
import type {
  DeflateStream,
  InflateStream,
  StreamCompressOptions,
  SyncDeflaterLike
} from "@archive/compression/streaming-compress.base";

export type { SyncDeflaterLike };

// Reusable type for zlib streams with flush() method
type ZlibFlushable = {
  write: (chunk: Buffer, cb: (err?: Error | null) => void) => void;
  flush: (mode: number, cb: () => void) => void;
  end: (cb: () => void) => void;
  on(event: "data", listener: (chunk: Buffer) => void): void;
  on(event: "error", listener: (err: Error) => void): void;
};

/**
 * Generic wrapper around zlib streams that flushes after every write.
 * This ensures true streaming behavior - data is emitted immediately, not buffered.
 */
class TrueStreamingZlib<T extends ZlibFlushable> extends Transform {
  constructor(private readonly zstream: T) {
    super();
    zstream.on("data", chunk => this.push(chunk));
    zstream.on("error", err => this.destroy(err));
  }

  _transform(chunk: Buffer, _encoding: BufferEncoding, callback: TransformCallback): void {
    this.zstream.write(chunk, writeErr => {
      if (writeErr) {
        callback(writeErr);
        return;
      }
      this.zstream.flush(constants.Z_SYNC_FLUSH, () => callback());
    });
  }

  _flush(callback: TransformCallback): void {
    this.zstream.flush(constants.Z_FINISH, () => {
      this.zstream.end(() => callback());
    });
  }
}

/**
 * Create a true streaming DEFLATE compressor
 * Returns a Transform stream that emits compressed data immediately after each write
 */
export function createDeflateStream(options: StreamCompressOptions = {}): DeflateStream {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  return new TrueStreamingZlib(createDeflateRaw({ level }));
}

/**
 * Create a true streaming INFLATE decompressor
 *
 * @param options - Decompression options (useWorker is ignored in Node.js)
 */
export function createInflateStream(options: StreamCompressOptions = {}): InflateStream {
  // Note: options.useWorker is ignored in Node.js (zlib uses native thread pool)
  void options;
  return createInflateRaw();
}

/**
 * Check if true streaming deflate-raw is available
 * In Node.js, zlib is always available, so this always returns true
 */
export function hasDeflateRaw(): boolean {
  return true;
}

// =============================================================================
// GZIP Streaming
// =============================================================================

export type GzipStream = Transform;
export type GunzipStream = Gunzip;

/**
 * Create a streaming GZIP compressor
 */
export function createGzipStream(options: StreamCompressOptions = {}): GzipStream {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  return new TrueStreamingZlib(createGzip({ level }));
}

/**
 * Create a streaming GZIP decompressor
 */
export function createGunzipStream(_options: StreamCompressOptions = {}): GunzipStream {
  return createGunzip();
}

// =============================================================================
// ZLIB Streaming (RFC 1950)
// =============================================================================

export type ZlibStream = Transform;
export type UnzlibStream = Inflate;

/**
 * Create a streaming Zlib compressor
 */
export function createZlibStream(options: StreamCompressOptions = {}): ZlibStream {
  const level = options.level ?? DEFAULT_COMPRESS_LEVEL;
  return new TrueStreamingZlib(createDeflate({ level }));
}

/**
 * Create a streaming Zlib decompressor
 */
export function createUnzlibStream(_options: StreamCompressOptions = {}): UnzlibStream {
  return createInflate();
}

// =============================================================================
// Synchronous stateful deflater (Node.js — native zlib)
// =============================================================================

/**
 * Minimum batch size before flushing to the native zlib compressor.
 *
 * Small chunks (e.g. one spreadsheet row ≈ 200-400 bytes) compress very
 * poorly when each is given its own zlib context because the LZ77 dictionary
 * starts empty every time. Batching into ≥ 64 KB mega-chunks gives zlib
 * enough history to find good matches, bringing compression ratios within
 * ~1% of single-shot compression.
 *
 * 64 KB is chosen as a sweet spot: large enough for good compression,
 * small enough to keep memory bounded and latency low.
 */
const SYNC_DEFLATE_BATCH_SIZE = 65536;

/**
 * Node.js synchronous deflater that batches small writes for better
 * compression.
 *
 * Previous implementation compressed each `write()` call independently
 * with `deflateRawSync()`, creating a fresh zlib context every time.
 * For streaming workloads that push many small chunks (e.g. WorkbookWriter
 * writing one row at a time), this destroyed the LZ77 dictionary between
 * chunks and caused compression ratios to drop from ~82% to ~58%.
 *
 * The new implementation accumulates incoming data into an internal buffer
 * and only calls `deflateRawSync()` when the buffer reaches 64 KB (or on
 * `finish()`). Each batch is still compressed independently, but 64 KB
 * is enough for zlib to build a good dictionary — the compression ratio
 * is within ~1% of a single-shot compression of the entire input.
 *
 * The trade-off is slightly higher latency (compressed output is not
 * returned byte-for-byte immediately), but this is acceptable because
 * the ZIP writer buffers output anyway and the streaming contract only
 * requires data to flow *eventually*, not after every single write.
 */
export class SyncDeflater implements SyncDeflaterLike {
  private _level: number;
  private _pending: Uint8Array[] = [];
  private _pendingSize = 0;

  constructor(level = DEFAULT_COMPRESS_LEVEL) {
    this._level = level;
  }

  write(data: Uint8Array): Uint8Array {
    if (data.length === 0) {
      return new Uint8Array(0);
    }

    this._pending.push(data);
    this._pendingSize += data.length;

    if (this._pendingSize >= SYNC_DEFLATE_BATCH_SIZE) {
      return this._flushBatch(false);
    }

    return new Uint8Array(0);
  }

  finish(): Uint8Array {
    return this._flushBatch(true);
  }

  private _flushBatch(final: boolean): Uint8Array {
    let input: Buffer;

    if (this._pending.length === 0) {
      input = Buffer.alloc(0);
    } else if (this._pending.length === 1) {
      input = Buffer.from(this._pending[0]);
    } else {
      input = Buffer.concat(this._pending);
    }

    this._pending.length = 0;
    this._pendingSize = 0;

    if (input.length === 0 && !final) {
      return new Uint8Array(0);
    }

    const result = deflateRawSync(input, {
      level: this._level,
      finishFlush: final ? constants.Z_FINISH : constants.Z_SYNC_FLUSH
    });

    // deflateRawSync returns a Buffer sharing a 16 KB slab ArrayBuffer.
    // Copy to a tight Uint8Array so the slab can be reclaimed.
    return new Uint8Array(result);
  }
}
