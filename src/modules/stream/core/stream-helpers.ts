/**
 * Stream helpers shared by the Node and browser entry points.
 *
 * Everything here works on the public shape of a stream — async iteration,
 * a Web `ReadableStream`'s reader, the `readable*` / `writable*` getters — so
 * it is the same code on both platforms. A Web stream is read through its own
 * reader; it is not wrapped in a Node-style stream first.
 */

import { toStreamBytes } from "@stream/core/binary-chunk";
import { isAsyncIterable, isReadableStream } from "@stream/core/type-guards";
import { StreamUnsupportedTypeError } from "@stream/errors";
import type { ReadableLike, WritableLike } from "@stream/types";
import { concatUint8Arrays, createTextDecoder } from "@utils/binary";

/** A stream or anything else that can be iterated asynchronously. */
export type AsyncIterableStream = { [Symbol.asyncIterator](): AsyncIterator<Uint8Array> };

/**
 * What `readableStreamToAsyncIterable` reads: a Web `ReadableStream` — whose
 * element type is inferred from it — or anything with the same kind of reader,
 * such as a stream from another realm.
 */
export type ReadableStreamLike<T = unknown> =
  | ReadableStream<T>
  | {
      getReader(): {
        read(): PromiseLike<{ done: true; value?: unknown } | { done: false; value: T }>;
        cancel?(reason?: unknown): PromiseLike<void>;
        releaseLock?(): void;
      };
    };

/** `true` for an object with a Web `ReadableStream`'s `getReader()`. */
export const isReadableStreamLike = isReadableStream as (
  value: unknown
) => value is ReadableStreamLike;

/**
 * Read a Web `ReadableStream` with `for await`, whether or not the platform's
 * `ReadableStream` is async-iterable itself. Every value is yielded as read —
 * `null` and `undefined` included, as `ReadableStream.prototype.values()` does —
 * so the caller decides what a value it cannot use means. Leaving the loop
 * early cancels the stream; the lock is released either way.
 */
export async function* readableStreamToAsyncIterable<T>(
  stream: ReadableStreamLike<T>
): AsyncGenerator<T, void, unknown> {
  if (!isReadableStream(stream)) {
    throw new StreamUnsupportedTypeError(
      "Streams.readableStreamToAsyncIterable",
      stream === null ? "null" : typeof stream,
      { field: "stream" }
    );
  }
  const reader = stream.getReader();
  let done = false;
  try {
    while (true) {
      const result = await reader.read();
      if (result.done) {
        done = true;
        return;
      }
      yield result.value;
    }
  } finally {
    if (!done) {
      await Promise.resolve(reader.cancel?.()).catch(() => {});
    }
    reader.releaseLock?.();
  }
}

/** Every chunk of a stream, in an array. */
export async function collect<T>(stream: {
  [Symbol.asyncIterator](): AsyncIterator<T>;
}): Promise<T[]> {
  if (!isAsyncIterable(stream)) {
    throw new StreamUnsupportedTypeError(
      "Streams.collect",
      stream === null ? "null" : typeof stream,
      {
        field: "stream"
      }
    );
  }
  const result: T[] = [];
  for await (const chunk of stream) {
    result.push(chunk);
  }
  return result;
}

/** A whole stream decoded as UTF-8 text. */
export function text(stream: AsyncIterableStream): Promise<string> {
  return streamToString(stream as AsyncIterable<Uint8Array>);
}

/** A whole stream parsed as JSON. */
export async function json<T = unknown>(stream: AsyncIterableStream): Promise<T> {
  return JSON.parse(await text(stream)) as T;
}

/** A whole stream's bytes. */
export function bytes(stream: AsyncIterableStream): Promise<Uint8Array> {
  return streamToUint8Array(stream as AsyncIterable<Uint8Array>);
}

/** Options for `Streams.text` and `Streams.json`. */
export interface StreamTextOptions {
  /** Text encoding, as `TextDecoder` names it. Default `"utf-8"`. */
  encoding?: string;
}

/**
 * A whole stream decoded as text, as `node:stream/consumers.text` decodes it: a leading BOM
 * is stripped (as `TextDecoder`, `Response.text()` and `Blob.text()` strip it), string chunks
 * are taken as they are, and the stream may be an async iterable (a Node-style readable is
 * one), a Web `ReadableStream` or a plain iterable. Published as `Streams.text`, and behind
 * `consumers.text`.
 */
export function readText(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array> | Iterable<Uint8Array | string>,
  options?: StreamTextOptions
): Promise<string> {
  return readTextAt("Streams.text", stream, options);
}

/** {@link readText}, naming `where` — the public call — when it refuses the stream. */
export async function readTextAt(
  where: string,
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array> | Iterable<Uint8Array | string>,
  options?: StreamTextOptions
): Promise<string> {
  const decoder = createTextDecoder(options?.encoding, { ignoreBOM: false });
  let text = "";
  for await (const chunk of toAsyncIterable(stream, where)) {
    if (typeof chunk === "string") {
      text += chunk;
      continue;
    }
    const bytes = toStreamBytes(chunk);
    if (!bytes) {
      throw new StreamUnsupportedTypeError(where, typeof chunk);
    }
    text += decoder.decode(bytes, { stream: true });
  }
  return text + decoder.decode();
}

/**
 * A whole stream decoded as {@link readText} does and parsed as JSON. Published as
 * `Streams.json`, and as `consumers.json`.
 */
export async function readJson<T = unknown>(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array> | Iterable<Uint8Array | string>,
  options?: StreamTextOptions
): Promise<T> {
  return JSON.parse(await readTextAt("Streams.json", stream, options)) as T;
}

/**
 * A stream as an async iterable: a Web `ReadableStream`, an async iterable, or — as
 * `node:stream/consumers` accepts — a plain iterable such as an array or a string.
 */
function toAsyncIterable<T>(
  stream: AsyncIterable<T> | ReadableStream<T> | Iterable<T>,
  name: string
): AsyncIterable<T> {
  if (isReadableStream(stream)) {
    return readableStreamToAsyncIterable<T>(stream);
  }
  if (isAsyncIterable(stream)) {
    return stream as AsyncIterable<T>;
  }
  if (
    typeof stream === "string" ||
    (stream !== null &&
      typeof stream === "object" &&
      typeof (stream as Iterable<T>)[Symbol.iterator] === "function")
  ) {
    return (async function* () {
      yield* stream as Iterable<T>;
    })();
  }
  throw new StreamUnsupportedTypeError(name, stream === null ? "null" : typeof stream, {
    field: "stream"
  });
}

/**
 * Collect every byte of a stream (string chunks as UTF-8). Published as `Streams.bytes`, and
 * as `consumers.bytes`.
 */
export async function streamToUint8Array(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array> | Iterable<Uint8Array | string>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let totalLength = 0;
  for await (const chunk of toAsyncIterable(stream, "Streams.bytes")) {
    const bytes = toStreamBytes(chunk);
    if (!bytes) {
      throw new StreamUnsupportedTypeError("Streams.bytes", typeof chunk);
    }
    chunks.push(bytes);
    totalLength += bytes.byteLength;
  }
  return concatUint8Arrays(chunks, totalLength);
}

/**
 * Alias of {@link streamToUint8Array}, kept for the deprecated `streamToBuffer` export. Despite the name it resolves to a
 * `Uint8Array`, not a Node `Buffer` — exactly what `Streams.bytes` returns.
 */
export function streamToBuffer(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>
): Promise<Uint8Array> {
  return streamToUint8Array(stream);
}

/** Decode a whole stream as text. */
export async function streamToString(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
  encoding?: string
): Promise<string> {
  const decoder = createTextDecoder(encoding);
  let text = "";
  for await (const chunk of toAsyncIterable(stream, "streamToString")) {
    const bytes = toStreamBytes(chunk);
    if (!bytes) {
      throw new StreamUnsupportedTypeError("streamToString", typeof chunk);
    }
    text += decoder.decode(bytes, { stream: true });
  }
  return text + decoder.decode();
}

/**
 * Consume a stream to its end, discarding the data. Published as `Streams.discard` —
 * not `drain`, which names the writable-side backpressure event.
 */
export async function drainStream(
  stream: AsyncIterable<unknown> | ReadableStream<unknown>
): Promise<void> {
  for await (const _chunk of toAsyncIterable(stream, "Streams.discard")) {
    // discard
  }
}

/** `true` for a Node-style readable stream that can still be read from. */
export function isReadable(stream: unknown): stream is ReadableLike {
  const s = stream as {
    read?: unknown;
    pipe?: unknown;
    destroyed?: boolean;
    readable?: boolean;
    readableEnded?: boolean;
  } | null;
  return (
    !!s &&
    typeof s.read === "function" &&
    typeof s.pipe === "function" &&
    !s.destroyed &&
    s.readable !== false &&
    s.readableEnded !== true
  );
}

/** `true` for a Node-style writable stream that can still be written to. */
export function isWritable(stream: unknown): stream is WritableLike {
  const s = stream as {
    write?: unknown;
    end?: unknown;
    destroyed?: boolean;
    writable?: boolean;
    writableEnded?: boolean;
  } | null;
  return (
    !!s &&
    typeof s.write === "function" &&
    typeof s.end === "function" &&
    !s.destroyed &&
    s.writable !== false &&
    s.writableEnded !== true
  );
}

const kIsDisturbed = Symbol.for("nodejs.stream.disturbed");

/**
 * `true` once a readable stream has been read from or aborted.
 *
 * Node's own Web streams answer through an internal symbol, which is consulted
 * first. Elsewhere a Web `ReadableStream` has no public "disturbed" flag, but
 * the Fetch spec makes it observable: `new Response(stream)` must throw for a
 * stream that is disturbed or locked, and does nothing to one that is neither —
 * it takes no lock and reads nothing. So an unlocked stream gets an exact
 * answer. A locked one cannot be told apart (locked-and-read and
 * locked-but-unread both throw) and reports `false`, as Node does for a stream
 * that is locked but not yet read.
 */
export function isDisturbed(stream: unknown): boolean {
  const s = stream as {
    [kIsDisturbed]?: boolean;
    readableDidRead?: boolean;
    readableAborted?: boolean;
  } | null;
  if (!s) {
    return false;
  }
  const known = s[kIsDisturbed] ?? (s.readableDidRead || s.readableAborted);
  if (known !== undefined) {
    return !!known;
  }
  if (isReadableStream(s) && !s.locked && typeof Response === "function") {
    try {
      new Response(s);
      return false;
    } catch {
      return true;
    }
  }
  return false;
}
