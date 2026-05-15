/**
 * Stream Module (shared exports)
 *
 * Type exports and platform-independent re-exports shared by both
 * `index.ts` (Node) and `index.browser.ts`.
 */

export type {
  StreamOptions,
  ReadableStreamOptions,
  WritableStreamOptions,
  TransformStreamOptions,
  DuplexStreamOptions,
  PullStreamOptions,
  BufferedStreamOptions,
  TransformCallback,
  FlushCallback,
  WriteCallback,
  DestroyCallback,
  IEventEmitter,
  IReadable,
  IWritable,
  ITransform,
  IDuplex,
  IPullStream,
  IBufferedStream,
  IPassThrough,
  ICollector,
  DataChunk,
  EventListener,
  PipelineSource,
  PipelineTransform,
  PipelineDestination
} from "@stream/types";

// Common re-exports (shared between Node.js and browser)
export {
  isDestroyed,
  isErrored,
  getDefaultHighWaterMark,
  setDefaultHighWaterMark,
  promisify
} from "@stream/common/utils";
export type { PipelineOptions, FinishedOptions } from "@stream/common/options";

export { ChunkedBuilder, TransactionalChunkedBuilder } from "@stream/chunked-builder";
export type { ChunkedBuilderOptions, BuilderSnapshot } from "@stream/chunked-builder";

export {
  collect,
  text,
  json,
  bytes,
  fromString,
  fromJSON,
  fromBytes,
  transform,
  filter,
  isReadableStreamLike,
  readableStreamToAsyncIterable
} from "@stream/utils";

export { StreamStateError, StreamTypeError, UnsupportedStreamTypeError } from "@stream/errors";

// Internal utilities exposed for cross-module use (e.g. archive, word/excel writers)
export {
  isReadableStream,
  isWritableStream,
  isAsyncIterable,
  isTransformStream
} from "@stream/internal/type-guards";
export { onceEvent } from "@stream/internal/event-utils";
export {
  eventedReadableToAsyncIterableNoDestroy,
  type EventedReadableLike
} from "@stream/internal/evented-readable-to-async-iterable";
export { SinkAdapter } from "@stream/internal/sink-adapter";
export type { AnySink, NodeWritableLike, DuckSinkLike } from "@stream/internal/sink-adapter";
export type { EventEmitterLike } from "@stream/types";
