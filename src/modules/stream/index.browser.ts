/**
 * Stream Module (browser entry)
 *
 * The same names as `index.ts`, backed by the browser kernel in `./browser`,
 * a port of Node's `lib/internal/streams`.
 */

import {
  Collector as CollectorImpl,
  createCollector as createCollectorImpl
} from "@stream/browser/collector";
import {
  PullStream as PullStreamImpl,
  BufferedStream as BufferedStreamImpl,
  createStringChunk as createStringChunkImpl,
  createByteChunk as createByteChunkImpl,
  createReadable as createReadableImpl,
  createReadableFromAsyncIterable as createReadableFromAsyncIterableImpl,
  createReadableFromArray as createReadableFromArrayImpl,
  createWritable as createWritableImpl,
  createTransform as createTransformImpl,
  createPassThrough as createPassThroughImpl,
  createPullStream as createPullStreamImpl,
  createBufferedStream as createBufferedStreamImpl,
  createDuplex as createDuplexImpl,
  createReadableFromGenerator as createReadableFromGeneratorImpl,
  createReadableFromPromise as createReadableFromPromiseImpl,
  createEmptyReadable as createEmptyReadableImpl,
  createNullWritable as createNullWritableImpl,
  filter as filterImpl,
  fromBytes as fromBytesImpl,
  fromJSON as fromJSONImpl,
  fromString as fromStringImpl,
  transform as transformImpl
} from "@stream/browser/factories";
import { finishedAll as finishedAllImpl } from "@stream/browser/pipeline";
import {
  streamToPromise as streamToPromiseImpl,
  streamToUint8Array as streamToUint8ArrayImpl,
  streamToBuffer as streamToBufferImpl,
  streamToString as streamToStringImpl,
  drainStream as drainStreamImpl,
  copyStream as copyStreamImpl,
  isTransform as isTransformImpl,
  isDuplex as isDuplexImpl,
  isStream as isStreamImpl
} from "@stream/browser/utils";
import {
  readableFromWeb as readableFromWebImpl,
  readableToWeb as readableToWebImpl,
  toWritable as toWritableImpl,
  writableFromWeb as writableFromWebImpl,
  writableToWeb as writableToWebImpl
} from "@stream/browser/web";

// Shared type + platform-independent exports
export * from "@stream/index.base";
// The guard for the `AbortError` above: true for any error named `"AbortError"`, so a native
// `DOMException` from `fetch` or `signal.throwIfAborted()` answers too.
export { isAbortError } from "@utils/errors";

// Everything node:stream does not name
export * as Streams from "@stream/surface/streams.browser";

// Core stream classes
export { Readable, Duplex, Transform, PassThrough } from "@stream/kernel/classes";
export type { ReadableOptions, DuplexOptions, TransformOptions } from "@stream/kernel/classes";
export { Writable } from "@stream/kernel/writable";
// `WritableSide` is the interface the browser `Writable` implements, so it is part of that
// class's published shape and has to be nameable from this entry.
export type { WritableOptions, WritableSide } from "@stream/kernel/writable";
export { getDefaultHighWaterMark, setDefaultHighWaterMark } from "@stream/kernel/state";

// Pipeline & Finished
export { pipeline, finished } from "@stream/browser/pipeline";

// Compose
export { compose } from "@stream/browser/compose";
export type { ComposedDuplex } from "@stream/browser/compose";

// node:stream utilities
export {
  addAbortSignal,
  isDisturbed,
  isReadable,
  isWritable,
  consumers,
  promises
} from "@stream/browser/utils";

export { duplexPair } from "@stream/browser/duplex-pair";

// =============================================================================
// Deprecated flat forwards — use the `Streams` namespace
// =============================================================================

/**
 * @deprecated Use {@link Streams.PullStream} (`new Streams.PullStream(options)`).
 * Removed in 1.0.
 */
export const PullStream = PullStreamImpl;
/**
 * @deprecated Use {@link Streams.PullStream} (`new Streams.PullStream(options)`).
 * Removed in 1.0.
 */
export type PullStream = PullStreamImpl;
/**
 * @deprecated Use {@link Streams.BufferedStream} (`new Streams.BufferedStream(options)`).
 * Removed in 1.0.
 */
export const BufferedStream = BufferedStreamImpl;
/**
 * @deprecated Use {@link Streams.BufferedStream} (`new Streams.BufferedStream(options)`).
 * Removed in 1.0.
 */
export type BufferedStream = BufferedStreamImpl;
/** @deprecated Use {@link Streams.createStringChunk}. Removed in 1.0. */
export const createStringChunk = createStringChunkImpl;
/** @deprecated Use {@link Streams.createByteChunk}. Removed in 1.0. */
export const createByteChunk = createByteChunkImpl;
/** @deprecated Use {@link Streams.createReadable}. Removed in 1.0. */
export const createReadable = createReadableImpl;
/** @deprecated Use {@link Streams.createReadableFromAsyncIterable}. Removed in 1.0. */
export const createReadableFromAsyncIterable = createReadableFromAsyncIterableImpl;
/** @deprecated Use {@link Streams.createReadableFromArray}. Removed in 1.0. */
export const createReadableFromArray = createReadableFromArrayImpl;
/** @deprecated Use {@link Streams.createWritable}. Removed in 1.0. */
export const createWritable = createWritableImpl;
/** @deprecated Use {@link Streams.createTransform}. Removed in 1.0. */
export const createTransform = createTransformImpl;
/** @deprecated Use {@link Streams.createPassThrough}. Removed in 1.0. */
export const createPassThrough = createPassThroughImpl;
/** @deprecated Use `new Streams.PullStream(options)`. Removed in 1.0. */
export const createPullStream = createPullStreamImpl;
/** @deprecated Use `new Streams.BufferedStream(options)`. Removed in 1.0. */
export const createBufferedStream = createBufferedStreamImpl;
/** @deprecated Use {@link Streams.createDuplex}. Removed in 1.0. */
export const createDuplex = createDuplexImpl;
/** @deprecated Use {@link Streams.createReadableFromGenerator}. Removed in 1.0. */
export const createReadableFromGenerator = createReadableFromGeneratorImpl;
/** @deprecated Use {@link Streams.createReadableFromPromise}. Removed in 1.0. */
export const createReadableFromPromise = createReadableFromPromiseImpl;
/** @deprecated Use {@link Streams.createEmptyReadable}. Removed in 1.0. */
export const createEmptyReadable = createEmptyReadableImpl;
/** @deprecated Use {@link Streams.createNullWritable}. Removed in 1.0. */
export const createNullWritable = createNullWritableImpl;
/** @deprecated Use {@link Streams.filter}. Removed in 1.0. */
export const filter = filterImpl;
/** @deprecated Use {@link Streams.fromBytes}. Removed in 1.0. */
export const fromBytes = fromBytesImpl;
/** @deprecated Use {@link Streams.fromJson}. Removed in 1.0. */
export const fromJSON = fromJSONImpl;
/** @deprecated Use {@link Streams.fromString}. Removed in 1.0. */
export const fromString = fromStringImpl;
/** @deprecated Use {@link Streams.createTransform}, which this duplicated. Removed in 1.0. */
export const transform = transformImpl;
/**
 * @deprecated Use {@link Streams.Collector} (`new Streams.Collector(options)`).
 * Removed in 1.0.
 */
export const Collector = CollectorImpl;
/**
 * @deprecated Use {@link Streams.Collector} (`new Streams.Collector(options)`).
 * Removed in 1.0.
 */
export type Collector<T = Uint8Array> = CollectorImpl<T>;
/** @deprecated Use `new Streams.Collector(options)`. Removed in 1.0. */
export const createCollector = createCollectorImpl;
/** @deprecated Use {@link Streams.waitForFinished}. Removed in 1.0. */
export const finishedAll = finishedAllImpl;
/** @deprecated Use {@link finished}, which this duplicated. Removed in 1.0. */
export const streamToPromise = streamToPromiseImpl;
/** @deprecated Use {@link Streams.bytes}. Removed in 1.0. */
export const streamToUint8Array = streamToUint8ArrayImpl;
/**
 * @deprecated Use {@link Streams.bytes}, which this duplicated: both resolve to a `Uint8Array`,
 * not a `Buffer`. Removed in 1.0.
 */
export const streamToBuffer = streamToBufferImpl;
/** @deprecated Use {@link Streams.text}; pass the encoding as `{ encoding }`. Removed in 1.0. */
export const streamToString = streamToStringImpl;
/** @deprecated Use {@link Streams.discard}. Removed in 1.0. */
export const drainStream = drainStreamImpl;
/** @deprecated Use {@link pipeline}, which this duplicated. Removed in 1.0. */
export const copyStream = copyStreamImpl;
/** @deprecated Use {@link Streams.isTransform}. Removed in 1.0. */
export const isTransform = isTransformImpl;
/** @deprecated Use {@link Streams.isDuplex}. Removed in 1.0. */
export const isDuplex = isDuplexImpl;
/** @deprecated Use {@link Streams.isStream}. Removed in 1.0. */
export const isStream = isStreamImpl;
/** @deprecated Use {@link Streams.fromWebReadable}. Removed in 1.0. */
export const readableFromWeb = readableFromWebImpl;
/** @deprecated Use {@link Streams.toWebReadable}. Removed in 1.0. */
export const readableToWeb = readableToWebImpl;
/** @deprecated Use {@link Streams.toWritable}. Removed in 1.0. */
export const toWritable = toWritableImpl;
/** @deprecated Use {@link Streams.fromWebWritable}. Removed in 1.0. */
export const writableFromWeb = writableFromWebImpl;
/** @deprecated Use {@link Streams.toWebWritable}. Removed in 1.0. */
export const writableToWeb = writableToWebImpl;
