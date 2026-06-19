/**
 * Shared types for true streaming compression.
 *
 * Kept in a dedicated base module so Node.js and browser implementations
 * don't depend on each other.
 */

// Re-export error helper for browser use
export { toError } from "@archive/core/errors";

export interface StreamCompressOptions {
  level?: number;

  /**
   * Use Web Workers for streaming compression/decompression (browser only).
   * This offloads the work to a background thread, keeping the main thread responsive.
   *
   * Note: Worker-based streaming buffers chunks and processes them at end()
   * for better throughput, trading memory for main thread responsiveness.
   *
   * Note: This option is ignored in Node.js.
   *
   * Defaults to false for backward compatibility.
   */
  useWorker?: boolean;

  /**
   * Custom worker pool instance (browser only).
   * If not provided, uses a shared pool.
   *
   * Note: This option is ignored in Node.js.
   */
  workerPool?: unknown;

  /**
   * Allow transferring the input buffer to the worker (browser only).
   * When true, chunks written to the stream may have their underlying buffer transferred.
   *
   * Note: This option is ignored in Node.js.
   */
  allowTransfer?: boolean;
}

export type StreamCallback = (err?: Error | null) => void;

/**
 * Minimal cross-platform streaming codec surface.
 *
 * Both Node.js (zlib / stream.Transform) and browser implementations
 * support this subset.
 */
export interface StreamingCodec {
  on(event: "data", listener: (chunk: Uint8Array) => void): this;
  on(event: "end", listener: () => void): this;
  on(event: "error", listener: (err: Error) => void): this;
  on(event: string, listener: (...args: any[]) => void): this;

  once(event: "data", listener: (chunk: Uint8Array) => void): this;
  once(event: "end", listener: () => void): this;
  once(event: "error", listener: (err: Error) => void): this;
  once(event: string, listener: (...args: any[]) => void): this;

  off(event: "data", listener: (chunk: Uint8Array) => void): this;
  off(event: "end", listener: () => void): this;
  off(event: "error", listener: (err: Error) => void): this;
  off(event: string, listener: (...args: any[]) => void): this;

  write(chunk: Uint8Array, callback?: StreamCallback): boolean;
  end(callback?: StreamCallback): unknown;
  destroy(err?: Error): unknown;
}

export type DeflateStream = StreamingCodec;
export type InflateStream = StreamingCodec;

/**
 * Stateful synchronous DEFLATE compressor interface.
 *
 * Unlike one-shot `compressSync`, this maintains compression state across
 * multiple `write()` calls (LZ77 window, bit position) to produce a single
 * valid DEFLATE stream. Used by the streaming ZIP writer (`pushSync`) to
 * achieve constant-memory streaming.
 *
 * Platform implementations:
 * - **Node.js**: wraps `zlib.deflateRawSync` with `Z_SYNC_FLUSH` per chunk.
 * - **Browser**: pure-JS LZ77 + fixed Huffman with a sliding window.
 */
export interface SyncDeflaterLike {
  /** Compress a chunk. Returns compressed bytes (may be empty if buffered). */
  write(data: Uint8Array): Uint8Array;
  /** Finalize the stream. Returns remaining bytes including the final block. */
  finish(): Uint8Array;
}
