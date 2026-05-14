/**
 * Browser True Streaming Compression
 *
 * Uses native CompressionStream("deflate-raw") for real chunk-by-chunk streaming.
 * Falls back to buffered compression if not supported.
 *
 * Worker Pool: Optional off-main-thread streaming compression/decompression
 * to prevent UI blocking.
 *
 * API compatible with Node.js version - supports .on("data"), .on("end"), .write(callback), .end()
 */

import {
  hasDeflateRawWebStreams,
  hasDeflateRawCompressionStream,
  hasGzipCompressionStream,
  hasGzipDecompressionStream,
  hasDeflateCompressionStream,
  hasDeflateDecompressionStream
} from "@archive/compression/compress.base";
import { gzipSync, gunzipSync, zlibSync, unzlibSync } from "@archive/compression/compress.browser";
import {
  deflateRawCompressed,
  inflateRaw,
  SyncDeflater as PureJsSyncDeflater
} from "@archive/compression/deflate-fallback";
import type { WorkerPool, WorkerTaskType } from "@archive/compression/worker-pool/index.browser";
import {
  hasWorkerSupport,
  getDefaultWorkerPool
} from "@archive/compression/worker-pool/index.browser";
import { DEFAULT_COMPRESS_LEVEL } from "@archive/shared/defaults";
import { concatUint8Arrays } from "@utils/binary";
import { EventEmitter } from "@utils/event-emitter";

export type {
  DeflateStream,
  InflateStream,
  StreamingCodec,
  StreamCompressOptions,
  SyncDeflaterLike
} from "@archive/compression/streaming-compress.base";
import {
  toError,
  type DeflateStream,
  type InflateStream,
  type StreamCallback,
  type StreamCompressOptions
} from "@archive/compression/streaming-compress.base";
import { EMPTY_UINT8ARRAY } from "@archive/shared/bytes";

export { hasWorkerSupport };

/** Shared error message constant */
const WRITE_AFTER_END_ERROR = "write after end";

/** Helper to handle errors with optional callback */
function handleError(emitter: EventEmitter, err: unknown, callback?: StreamCallback): void {
  const error = toError(err);
  if (callback) {
    callback(error);
  } else {
    emitter.emit("error", error);
  }
}

/**
 * Check if deflate-raw streaming compression is supported by this library.
 */
export function hasDeflateRaw(): boolean {
  return true;
}

// =============================================================================
// Base Codec - shared write/end/destroy lifecycle
// =============================================================================

/**
 * Async codec interface - abstracted write/end/close operations.
 */
interface AsyncCodecBackend {
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(err?: Error): void;
}

/**
 * Base streaming codec with unified lifecycle management.
 * Backend can emit events via the returned codec reference.
 *
 * Backpressure: each `write()` chains a deflate task onto `writeChain`.
 * Without throttling, a fast producer could enqueue thousands of tasks
 * (each retaining its input chunk) before the backend catches up. To
 * surface backpressure to the caller, we count pending writes; once the
 * count reaches `HIGH_WATER_MARK`, `write()` returns `false`. When the
 * count drops back below the mark, we emit `'drain'` so callers using
 * the standard Node Writable contract (`pipe()` / `pipeline()`) can
 * resume. This keeps memory bounded on slow backends or slow consumers.
 */
class AsyncStreamCodec extends EventEmitter {
  // Allow up to this many in-flight chunks before signalling backpressure.
  // 16 chunks × typical 64 KiB ≈ 1 MiB of in-flight retention, enough to
  // saturate the deflate backend without unbounded growth.
  private static readonly HIGH_WATER_MARK = 16;

  private ended = false;
  private destroyed = false;
  private writeChain: Promise<void> = Promise.resolve();
  private _backend: AsyncCodecBackend | null = null;
  private _pendingWrites = 0;
  private _needsDrain = false;

  setBackend(backend: AsyncCodecBackend): void {
    this._backend = backend;
  }

  write(chunk: Uint8Array, callback?: StreamCallback): boolean {
    if (this.ended) {
      handleError(this, new Error(WRITE_AFTER_END_ERROR), callback);
      return false;
    }

    if (chunk.byteLength === 0) {
      if (callback) {
        queueMicrotask(callback);
      }
      return true;
    }

    const backend = this._backend;
    if (!backend) {
      throw new Error("Backend not initialized");
    }

    this._pendingWrites++;
    const promise = this.writeChain.then(() => backend.write(chunk));
    this.writeChain = promise;

    promise
      .then(() => {
        if (!this.destroyed) {
          callback?.();
        }
      })
      .catch(err => {
        if (!this.destroyed) {
          handleError(this, err, callback);
        }
      })
      .finally(() => {
        this._pendingWrites--;
        if (
          this._needsDrain &&
          this._pendingWrites < AsyncStreamCodec.HIGH_WATER_MARK &&
          !this.destroyed
        ) {
          this._needsDrain = false;
          this.emit("drain");
        }
      });

    if (this._pendingWrites >= AsyncStreamCodec.HIGH_WATER_MARK) {
      this._needsDrain = true;
      return false;
    }
    return true;
  }

  end(callback?: StreamCallback): void {
    if (this.ended) {
      callback?.();
      return;
    }
    this.ended = true;

    const backend = this._backend;
    if (!backend) {
      throw new Error("Backend not initialized");
    }

    void this.writeChain
      .then(() => backend.close())
      .then(() => callback?.())
      .catch(err => handleError(this, err, callback));
  }

  destroy(err?: Error): void {
    this.ended = true;
    this.destroyed = true;
    this._backend?.abort(err);
    if (err) {
      this.emit("error", err);
    }
  }
}

// =============================================================================
// WebStream Codec - uses native CompressionStream/DecompressionStream
// =============================================================================

type WebStreamFormat = "deflate-raw" | "deflate" | "gzip";

function createNativeWebStreamCodec(format: WebStreamFormat, isCompress: boolean): DeflateStream {
  const stream = isCompress
    ? new CompressionStream(format as CompressionFormat)
    : new DecompressionStream(format as CompressionFormat);
  const writer = stream.writable.getWriter() as WritableStreamDefaultWriter<Uint8Array>;
  const reader = stream.readable.getReader();

  const codec = new AsyncStreamCodec();
  let readLoopError: Error | null = null;

  const readLoop = (async (): Promise<void> => {
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) {
          break;
        }
        if (value) {
          codec.emit("data", value);
        }
      }
      codec.emit("end");
    } catch (err) {
      const error = toError(err);
      readLoopError = error;
      codec.emit("error", error);
    }
  })();

  codec.setBackend({
    write: chunk => writer.write(chunk),
    close: async () => {
      await writer.close();
      await readLoop;
      if (readLoopError) {
        throw toError(readLoopError);
      }
    },
    abort: err => {
      reader.cancel(err).catch(() => {});
      writer.abort(err).catch(() => {});
    }
  });

  return codec;
}

// =============================================================================
// Worker Codec - uses WorkerPool.openStream for true streaming in worker
// =============================================================================

function createWorkerStreamCodec(
  type: WorkerTaskType,
  pool: WorkerPool | undefined,
  level: number | undefined,
  allowTransfer: boolean | undefined
): DeflateStream {
  const effectivePool = pool ?? getDefaultWorkerPool();

  let endResolve: (() => void) | null = null;
  let endReject: ((err: Error) => void) | null = null;
  const endPromise = new Promise<void>((resolve, reject) => {
    endResolve = resolve;
    endReject = reject;
  });

  const codec = new AsyncStreamCodec();
  const workerStream = effectivePool.openStream(type, {
    level,
    allowTransfer,
    onData: chunk => codec.emit("data", chunk),
    onEnd: () => {
      codec.emit("end");
      endResolve?.();
    },
    onError: err => {
      codec.emit("error", err);
      endReject?.(err);
    }
  });

  codec.setBackend({
    write: chunk => workerStream.write(chunk),
    close: async () => {
      await workerStream.end();
      await endPromise;
    },
    abort: err => {
      endResolve?.();
      workerStream.abort(err?.message);
    }
  });

  return codec;
}

// =============================================================================
// Buffered Codec - fallback when no native streaming available
// =============================================================================

class BufferedCodec extends EventEmitter {
  private readonly chunks: Uint8Array[] = [];
  private ended = false;

  constructor(private readonly process: (data: Uint8Array) => Uint8Array) {
    super();
  }

  write(chunk: Uint8Array, callback?: StreamCallback): boolean {
    if (this.ended) {
      handleError(this, new Error(WRITE_AFTER_END_ERROR), callback);
      return false;
    }

    if (chunk.byteLength === 0) {
      if (callback) {
        queueMicrotask(callback);
      }
      return true;
    }

    this.chunks.push(chunk);
    if (callback) {
      queueMicrotask(callback);
    }
    return true;
  }

  end(callback?: StreamCallback): void {
    if (this.ended) {
      callback?.();
      return;
    }
    this.ended = true;

    const chunkCount = this.chunks.length;
    const data =
      chunkCount === 0
        ? EMPTY_UINT8ARRAY
        : chunkCount === 1
          ? this.chunks[0]
          : concatUint8Arrays(this.chunks);
    this.chunks.length = 0;

    try {
      const result = this.process(data);
      this.emit("data", result);
      this.emit("end");
      callback?.();
    } catch (err) {
      handleError(this, err, callback);
    }
  }

  destroy(err?: Error): void {
    this.ended = true;
    this.chunks.length = 0;
    if (err) {
      this.emit("error", err);
    }
  }
}

// =============================================================================
// Factory - select best codec based on environment and options
// =============================================================================

function createStreamCodec(
  type: "deflate" | "inflate",
  options: StreamCompressOptions
): DeflateStream {
  const level = type === "deflate" ? (options.level ?? DEFAULT_COMPRESS_LEVEL) : undefined;

  if (options.useWorker && hasWorkerSupport()) {
    return createWorkerStreamCodec(
      type,
      options.workerPool as WorkerPool | undefined,
      level,
      options.allowTransfer
    );
  }

  // Use native CompressionStream/DecompressionStream when the required
  // direction is available. Compression only needs CompressionStream;
  // decompression only needs DecompressionStream.
  if (type === "deflate" ? hasDeflateRawCompressionStream() : hasDeflateRawWebStreams()) {
    return createNativeWebStreamCodec("deflate-raw", type === "deflate");
  }

  return new BufferedCodec(
    type === "deflate" ? data => deflateRawCompressed(data, level) : inflateRaw
  );
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Create a streaming DEFLATE compressor
 */
export function createDeflateStream(options: StreamCompressOptions = {}): DeflateStream {
  return createStreamCodec("deflate", options);
}

/**
 * Create a streaming INFLATE decompressor
 */
export function createInflateStream(options: StreamCompressOptions = {}): InflateStream {
  return createStreamCodec("inflate", options);
}

// =============================================================================
// GZIP / ZLIB Streaming - unified factory pattern
// =============================================================================

export type GzipStream = DeflateStream;
export type GunzipStream = InflateStream;
export type ZlibStream = DeflateStream;
export type UnzlibStream = InflateStream;

interface WrappedCodecConfig {
  format: WebStreamFormat;
  hasNative: () => boolean;
  compressFallback: (data: Uint8Array, level: number) => Uint8Array;
  decompressFallback: (data: Uint8Array) => Uint8Array;
}

const GZIP_CONFIG: WrappedCodecConfig = {
  format: "gzip",
  hasNative: () => hasGzipCompressionStream() && hasGzipDecompressionStream(),
  compressFallback: (data, level) => gzipSync(data, { level }),
  decompressFallback: gunzipSync
};

const ZLIB_CONFIG: WrappedCodecConfig = {
  format: "deflate",
  hasNative: () => hasDeflateCompressionStream() && hasDeflateDecompressionStream(),
  compressFallback: (data, level) => zlibSync(data, { level }),
  decompressFallback: unzlibSync
};

function createWrappedStream(
  config: WrappedCodecConfig,
  isCompress: boolean,
  options: StreamCompressOptions
): DeflateStream {
  if (config.hasNative()) {
    return createNativeWebStreamCodec(config.format, isCompress);
  }
  const level = isCompress ? (options.level ?? DEFAULT_COMPRESS_LEVEL) : DEFAULT_COMPRESS_LEVEL;
  return new BufferedCodec(
    isCompress ? data => config.compressFallback(data, level) : config.decompressFallback
  );
}

/** Create a streaming GZIP compressor */
export function createGzipStream(options: StreamCompressOptions = {}): GzipStream {
  return createWrappedStream(GZIP_CONFIG, true, options);
}

/** Create a streaming GZIP decompressor */
export function createGunzipStream(options: StreamCompressOptions = {}): GunzipStream {
  return createWrappedStream(GZIP_CONFIG, false, options);
}

/** Create a streaming Zlib compressor */
export function createZlibStream(options: StreamCompressOptions = {}): ZlibStream {
  return createWrappedStream(ZLIB_CONFIG, true, options);
}

/** Create a streaming Zlib decompressor */
export function createUnzlibStream(options: StreamCompressOptions = {}): UnzlibStream {
  return createWrappedStream(ZLIB_CONFIG, false, options);
}

// =============================================================================
// Synchronous stateful deflater (Browser — pure JS)
// =============================================================================

/**
 * Browser synchronous deflater — re-exports the pure-JS `SyncDeflater`
 * from deflate-fallback.ts which maintains a LZ77 sliding window and
 * bit-stream state across `write()` calls.
 */
export { PureJsSyncDeflater as SyncDeflater };

/**
 * Returns true when the browser supports native `CompressionStream("deflate-raw")`,
 * signalling that `push()` should prefer the async path over `SyncDeflater`.
 *
 * Only checks for compression support — decompression is not needed for writing.
 */
export function hasNativeAsyncDeflate(): boolean {
  return hasDeflateRawCompressionStream();
}
