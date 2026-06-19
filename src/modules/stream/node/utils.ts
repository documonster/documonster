/**
 * Node.js Stream - Utilities
 *
 * Stream utility functions, type guards, consumers, and state inspection.
 */

import { Readable, Transform, Duplex } from "stream";

import { createAddAbortSignal } from "@stream/core/add-abort-signal";
import { toStreamBytes } from "@stream/core/binary-chunk";
import { createConsumers } from "@stream/core/consumers";
import {
  createIsTransform,
  createIsDuplex,
  createIsStream,
  isAsyncIterable,
  isReadableStream
} from "@stream/core/type-guards";
import { getDefaultHighWaterMark } from "@stream/core/utils";
import { UnsupportedStreamTypeError } from "@stream/errors";
import { pipeline, finished } from "@stream/node/pipeline";
import { Writable } from "@stream/node/writable";
import type {
  DuplexStreamOptions,
  IDuplex,
  ITransform,
  ReadableLike,
  WritableLike
} from "@stream/types";
import { createTextDecoder, concatUint8Arrays } from "@utils/binary";

// =============================================================================
// Utility Functions
// =============================================================================

/** Convert a stream to a promise that resolves when finished */
export const streamToPromise = finished;

/** Copy from a readable stream to a writable stream */
export const copyStream = pipeline;

/**
 * Collect all data from a readable stream into a Uint8Array
 * (Node.js equivalent of browser streamToBuffer)
 */
export async function streamToBuffer(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>
): Promise<Uint8Array> {
  let iterable: AsyncIterable<Uint8Array>;
  if (isReadableStream(stream)) {
    iterable = (Readable as any).fromWeb(stream as any) as AsyncIterable<Uint8Array>;
  } else if (isAsyncIterable(stream)) {
    iterable = stream;
  } else {
    throw new UnsupportedStreamTypeError("streamToBuffer", typeof stream);
  }
  const chunks: Uint8Array[] = [];
  let totalLength = 0;
  for await (const chunk of iterable as any) {
    const bytes = toStreamBytes(chunk);
    if (!bytes) {
      throw new UnsupportedStreamTypeError("streamToBuffer", typeof chunk);
    }
    chunks.push(bytes);
    totalLength += bytes.byteLength;
  }

  return concatUint8Arrays(chunks, totalLength);
}

/**
 * Collect all data from a readable stream into a Uint8Array
 */
export async function streamToUint8Array(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>
): Promise<Uint8Array> {
  return streamToBuffer(stream);
}

/**
 * Collect all data from a readable stream into a string
 */
export async function streamToString(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
  encoding?: string
): Promise<string> {
  let iterable: AsyncIterable<Uint8Array>;
  if (isReadableStream(stream)) {
    iterable = (Readable as any).fromWeb(stream as any) as AsyncIterable<Uint8Array>;
  } else if (isAsyncIterable(stream)) {
    iterable = stream;
  } else {
    throw new UnsupportedStreamTypeError("streamToString", typeof stream);
  }

  const decoder = createTextDecoder(encoding);
  let text = "";

  for await (const chunk of iterable as any) {
    const bytes = toStreamBytes(chunk);
    if (!bytes) {
      throw new UnsupportedStreamTypeError("streamToString", typeof chunk);
    }
    text += decoder.decode(bytes, { stream: true });
  }

  text += decoder.decode();
  return text;
}

/**
 * Drain a stream (consume all data without processing)
 */
export async function drainStream(
  stream: AsyncIterable<unknown> | ReadableStream<unknown>
): Promise<void> {
  let iterable: AsyncIterable<unknown>;
  if (isReadableStream(stream)) {
    iterable = (Readable as any).fromWeb(stream as any) as AsyncIterable<unknown>;
  } else if (isAsyncIterable(stream)) {
    iterable = stream;
  } else {
    throw new UnsupportedStreamTypeError("drainStream", typeof stream);
  }

  for await (const _chunk of iterable) {
    // Consume data
  }
}

// =============================================================================
// Additional Utility Functions
// =============================================================================

/** Add abort signal handling to any stream */
export const addAbortSignal = createAddAbortSignal({
  add(emitter, event, listener) {
    (emitter as any).once?.(event, listener);
  },
  remove(emitter, event, listener) {
    (emitter as any).off?.(event, listener);
  }
});

// =============================================================================
// Type Guards
// =============================================================================

/**
 * Check if an object is a readable stream that is still in a readable state.
 * Returns false for destroyed or ended/finished streams (matches Node.js behavior).
 */
export function isReadable(obj: unknown): obj is ReadableLike {
  if (obj == null) {
    return false;
  }
  const s = obj as any;
  if (s.destroyed) {
    return false;
  }
  if (!(obj instanceof Readable || obj instanceof Transform || obj instanceof Duplex)) {
    if (!(typeof s.read === "function" && typeof s.pipe === "function")) {
      return false;
    }
  }
  if (s.readableEnded === true) {
    return false;
  }
  return true;
}

/**
 * Check if an object is a writable stream that is still in a writable state.
 * Returns false for destroyed or ended/finished streams (matches Node.js behavior).
 */
export function isWritable(obj: unknown): obj is WritableLike {
  if (obj == null) {
    return false;
  }
  const s = obj as any;
  if (s.destroyed) {
    return false;
  }
  if (!(obj instanceof Writable || obj instanceof Transform || obj instanceof Duplex)) {
    if (!(typeof s.write === "function" && typeof s.end === "function")) {
      return false;
    }
  }
  if (s.writableEnded === true) {
    return false;
  }
  return true;
}

/** Check if an object is a transform stream */
export const isTransform: (obj: unknown) => obj is ITransform<any, any> =
  createIsTransform(Transform);

/** Check if an object is a duplex stream */
export const isDuplex: (obj: unknown) => obj is IDuplex<any, any> = createIsDuplex(
  Duplex,
  Transform
);

/** Check if an object is any kind of stream */
export const isStream: (obj: unknown) => obj is ReadableLike | WritableLike = createIsStream(
  Readable,
  Writable
);

// =============================================================================
// Stream State Inspection Functions
// =============================================================================

/**
 * Check if a readable stream has been disturbed (read from or cancelled).
 * Matches Node.js native: readableDidRead || readableAborted.
 */
export function isDisturbed(stream: unknown): boolean {
  if ((stream as any)?.locked !== undefined) {
    return !!(stream as ReadableStream).locked;
  }
  const s = stream as any;
  return !!(s?.readableDidRead || s?._didRead || s?.readableAborted);
}

// =============================================================================
// Duplex Pair
// =============================================================================

/**
 * Create a pair of connected Duplex streams
 * Data written to one stream can be read from the other
 */
export function duplexPair<T = any>(options?: DuplexStreamOptions): [IDuplex<T, T>, IDuplex<T, T>] {
  const objectMode =
    options?.readableObjectMode ?? options?.writableObjectMode ?? options?.objectMode ?? false;
  const highWaterMark =
    options?.readableHighWaterMark ??
    options?.writableHighWaterMark ??
    options?.highWaterMark ??
    getDefaultHighWaterMark(objectMode);

  // Holder object allows both streams to reference each other via closure
  // while satisfying the `const` constraint (each variable is assigned once).
  const pair: { s1?: Duplex; s2?: Duplex } = {};

  const duplex1 = new Duplex({
    allowHalfOpen: options?.allowHalfOpen,
    readableHighWaterMark: options?.readableHighWaterMark ?? highWaterMark,
    writableHighWaterMark: options?.writableHighWaterMark ?? highWaterMark,
    readableObjectMode: options?.readableObjectMode ?? objectMode,
    writableObjectMode: options?.writableObjectMode ?? objectMode,
    read(): void {
      // Data will be pushed from duplex2's write()
    },
    write(chunk: T, _encoding: BufferEncoding, callback: (error?: Error | null) => void): void {
      // Push to peer; if peer signals backpressure, defer callback until drain.
      if (!pair.s2!.push(chunk)) {
        pair.s2!.once("drain", () => callback());
      } else {
        callback();
      }
    },
    final(callback: (error?: Error | null) => void): void {
      pair.s2!.push(null);
      callback();
    }
  });

  const duplex2 = new Duplex({
    allowHalfOpen: options?.allowHalfOpen,
    readableHighWaterMark: options?.readableHighWaterMark ?? highWaterMark,
    writableHighWaterMark: options?.writableHighWaterMark ?? highWaterMark,
    readableObjectMode: options?.readableObjectMode ?? objectMode,
    writableObjectMode: options?.writableObjectMode ?? objectMode,
    read(): void {
      // Data will be pushed from duplex1's write()
    },
    write(chunk: T, _encoding: BufferEncoding, callback: (error?: Error | null) => void): void {
      // Push to peer; if peer signals backpressure, defer callback until drain.
      if (!duplex1.push(chunk)) {
        duplex1.once("drain", () => callback());
      } else {
        callback();
      }
    },
    final(callback: (error?: Error | null) => void): void {
      duplex1.push(null);
      callback();
    }
  });

  pair.s1 = duplex1;
  pair.s2 = duplex2;

  return [duplex1, duplex2];
}

// =============================================================================
// Stream Consumers (like stream/consumers in Node.js)
// =============================================================================

export const consumers = createConsumers({ streamToUint8Array, streamToString });

// =============================================================================
// Promises API (like stream/promises in Node.js)
// =============================================================================

export const promises = { pipeline, finished };
