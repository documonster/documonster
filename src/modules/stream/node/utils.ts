/**
 * Node.js Stream - utilities
 */

import {
  Duplex,
  Readable,
  Transform,
  Writable,
  duplexPair as nodeDuplexPair,
  isReadable as nodeIsReadable,
  isWritable as nodeIsWritable
} from "stream";

import { createAddAbortSignal } from "@stream/core/add-abort-signal";
import { createConsumers } from "@stream/core/consumers";
import { createIsDuplex, createIsStream, createIsTransform } from "@stream/core/type-guards";
import { finished, pipeline } from "@stream/node/pipeline";
import type {
  DuplexStreamOptions,
  IDuplex,
  ITransform,
  ReadableLike,
  WritableLike
} from "@stream/types";

export {
  drainStream,
  isDisturbed,
  streamToBuffer,
  streamToString,
  streamToUint8Array
} from "@stream/core/stream-helpers";

/**
 * `true` for a stream that can still be read from — `node:stream`'s answer, Web
 * `ReadableStream`s included. A type guard, so a value that is not a stream is `false`
 * where Node's returns `null`.
 */
export function isReadable(stream: unknown): stream is ReadableLike {
  return nodeIsReadable(stream as Readable) === true;
}

/**
 * `true` for a stream that can still be written to — `node:stream`'s answer, Web
 * `WritableStream`s included. A type guard, so a value that is not a stream is `false`
 * where Node's returns `null`.
 */
export function isWritable(stream: unknown): stream is WritableLike {
  return nodeIsWritable(stream as Writable) === true;
}

/** Resolve when a stream has finished. */
export const streamToPromise = finished;

/** Copy from a readable stream to a writable stream. */
export const copyStream = pipeline;

export const isTransform: (obj: unknown) => obj is ITransform<unknown, unknown> =
  /* @__PURE__ */ createIsTransform(Transform);

export const isDuplex: (obj: unknown) => obj is IDuplex<unknown, unknown> = createIsDuplex(Duplex);

export const isStream: (obj: unknown) => obj is ReadableLike | WritableLike = createIsStream(
  Readable,
  Writable
);

export const addAbortSignal = createAddAbortSignal({
  add(emitter, event, listener) {
    emitter.once?.(event, listener);
  },
  remove(emitter, event, listener) {
    (emitter.off ?? emitter.removeListener)?.call(emitter, event, listener);
  }
});

/** Two connected Duplex streams: each one reads what the other writes. */
export function duplexPair<T = unknown>(
  options?: DuplexStreamOptions
): [IDuplex<T, T>, IDuplex<T, T>] {
  return nodeDuplexPair(options) as unknown as [IDuplex<T, T>, IDuplex<T, T>];
}

export const consumers = createConsumers({
  toBuffer: bytes => Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
});

export const promises = { pipeline, finished };
