/**
 * Stream Module (shared exports)
 *
 * Type exports and platform-independent re-exports shared by both
 * `index.ts` (Node) and `index.browser.ts`.
 */

import {
  ChunkedBuilder as ChunkedBuilderImpl,
  TransactionalChunkedBuilder as TransactionalChunkedBuilderImpl
} from "@stream/chunked-builder";
import { onceEvent as onceEventImpl } from "@stream/core/event-utils";
import { eventedReadableToAsyncIterableNoDestroy as eventedReadableToAsyncIterableNoDestroyImpl } from "@stream/core/evented-readable-to-async-iterable";
import { SinkAdapter as SinkAdapterImpl } from "@stream/core/sink-adapter";
import {
  bytes as bytesImpl,
  collect as collectImpl,
  isReadableStreamLike as isReadableStreamLikeImpl,
  json as jsonImpl,
  readableStreamToAsyncIterable as readableStreamToAsyncIterableImpl,
  text as textImpl
} from "@stream/core/stream-helpers";
import {
  isAsyncIterable as isAsyncIterableImpl,
  isReadableStream as isReadableStreamImpl,
  isTransformStream as isTransformStreamImpl,
  isWritableStream as isWritableStreamImpl
} from "@stream/core/type-guards";
import { promisify as promisifyImpl } from "@stream/core/utils";

export type {
  StreamOptions,
  ReadableStreamOptions,
  WritableStreamOptions,
  TransformStreamOptions,
  DuplexStreamOptions,
  PullStreamOptions,
  PullUntilOptions,
  BufferedStreamOptions,
  TransformCallback,
  FlushCallback,
  WriteCallback,
  DestroyCallback,
  IEventEmitter,
  ReadableLike,
  WritableLike,
  PipelineStreamLike,
  PipelineStageLike,
  PipelineGeneratorStage,
  IReadable,
  IWritable,
  ITransform,
  IDuplex,
  IPushableReadable,
  IPushableDuplex,
  IPullStream,
  IBufferedStream,
  IPassThrough,
  ICollector,
  StreamDataChunk,
  StreamEventListener,
  DataChunk,
  EventListener,
  PipelineSource,
  PipelineTransform,
  PipelineDestination
} from "@stream/types";

// Node-compatible (node:stream has these names)
export { isDestroyed, isErrored } from "@stream/core/utils";
export type {
  StreamFactories,
  StreamByteChunk,
  StreamObjectModeOptions,
  StreamByteModeOptions,
  StreamDestroyHook
} from "@stream/core/factories";
export type { StreamConsumers } from "@stream/core/consumers";
export type {
  ReadableFromWeb,
  WebReadableOptions,
  WebWritableOptions,
  WritableFromWeb
} from "@stream/core/web-conversions";
export type { PipelineOptions, FinishedOptions, PipelineCallback } from "@stream/core/options";

export type { ChunkedBuilderOptions, BuilderSnapshot } from "@stream/chunked-builder";

export type {
  AsyncIterableStream,
  ReadableStreamLike,
  StreamTextOptions
} from "@stream/core/stream-helpers";

export type { StreamErrorOptions } from "@stream/errors";
export {
  StreamError,
  StreamStateError,
  StreamTypeError,
  UnsupportedStreamTypeError,
  isStreamError
} from "@stream/errors";
// What every stream in this entry rejects with when it is aborted (`addAbortSignal`, a
// `signal` option) — the same class `documonster/archive` exports, so one `instanceof` check
// works for both. `isStreamError` does not answer for it: cancellation is not a stream fault.
export { AbortError } from "@utils/errors";

export type { EventedReadableLike } from "@stream/core/evented-readable-to-async-iterable";
export type { AnySink, NodeWritableLike, DuckSinkLike } from "@stream/core/sink-adapter";
export type { EventEmitterLike } from "@stream/types";

// =============================================================================
// Deprecated flat forwards — everything node:stream does not name now lives on
// the platform entry's `Streams` namespace.
// =============================================================================

/** @deprecated Use {@link Streams.ChunkedBuilder}. Removed in 1.0. */
export const ChunkedBuilder = ChunkedBuilderImpl;
export type ChunkedBuilder = ChunkedBuilderImpl;
/** @deprecated Use {@link Streams.TransactionalChunkedBuilder}. Removed in 1.0. */
export const TransactionalChunkedBuilder = TransactionalChunkedBuilderImpl;
export type TransactionalChunkedBuilder = TransactionalChunkedBuilderImpl;
/** @deprecated Use {@link Streams.bytes}. Removed in 1.0. */
export const bytes = bytesImpl;
/** @deprecated Use {@link Streams.collect}. Removed in 1.0. */
export const collect = collectImpl;
/** @deprecated Use {@link Streams.isReadableStream}, which this duplicated. Removed in 1.0. */
export const isReadableStreamLike = isReadableStreamLikeImpl;
/** @deprecated Use {@link Streams.json}. Removed in 1.0. */
export const json = jsonImpl;
/** @deprecated Use {@link Streams.readableStreamToAsyncIterable}. Removed in 1.0. */
export const readableStreamToAsyncIterable = readableStreamToAsyncIterableImpl;
/** @deprecated Use {@link Streams.text}. Removed in 1.0. */
export const text = textImpl;
/**
 * @deprecated Use {@link Streams.LowLevel.fromCallback}: `promisify` read as `node:util`'s,
 * which wraps a function rather than running one. Removed in 1.0.
 */
export const promisify = promisifyImpl;
/** @deprecated Use {@link Streams.isReadableStream}. Removed in 1.0. */
export const isReadableStream = isReadableStreamImpl;
/** @deprecated Use {@link Streams.isWritableStream}. Removed in 1.0. */
export const isWritableStream = isWritableStreamImpl;
/** @deprecated Use {@link Streams.isAsyncIterable}. Removed in 1.0. */
export const isAsyncIterable = isAsyncIterableImpl;
/** @deprecated Use {@link Streams.isTransformStream}. Removed in 1.0. */
export const isTransformStream = isTransformStreamImpl;
/** @deprecated Use {@link Streams.LowLevel.onceEvent}. Removed in 1.0. */
export const onceEvent = onceEventImpl;
/**
 * @deprecated Use {@link Streams.LowLevel.eventsToAsyncIterableNoDestroy}. Removed in 1.0.
 */
export const eventedReadableToAsyncIterableNoDestroy = eventedReadableToAsyncIterableNoDestroyImpl;
/** @deprecated Use {@link Streams.LowLevel.SinkAdapter}. Removed in 1.0. */
export const SinkAdapter = SinkAdapterImpl;
export type SinkAdapter = SinkAdapterImpl;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
