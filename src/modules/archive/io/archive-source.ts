import { OPTIONS, checkArchiveArgument } from "@archive/core/arguments";
import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { ArchiveArgumentError, createAbortError, toError } from "@archive/core/errors";
import { isAsyncIterable, isReadableStream } from "@stream/core/type-guards";
import { stringToUint8Array as encodeUtf8, concatUint8Arrays } from "@utils/binary";
import { s } from "@utils/schema";

/**
 * Anything an archive can be read from, or an entry's content taken from. A string is content,
 * encoded as UTF-8 — never a file path or URL.
 */
export type ArchiveSource =
  | Uint8Array
  | ArrayBuffer
  | string
  | Blob
  | AsyncIterable<unknown>
  | ReadableStream<unknown>
  | { [Symbol.asyncIterator](): AsyncIterator<unknown> };

/**
 * Convert an ArchiveSource to a web ReadableStream of Uint8Array chunks.
 *
 * This is a convenience adapter for environments where consumers expect
 * `ReadableStream` instead of `AsyncIterable`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function toReadableStream(
  source: ArchiveSource | ArchiveBytes,
  options: { signal?: AbortSignal } = {}
): ReadableStream<Uint8Array> {
  checkSourceArgument("Archive.toReadableStream", source);
  checkSourceOptions("Archive.toReadableStream", options);
  const iterable = iterateArchiveSource(source, options);
  const iterator = iterable[Symbol.asyncIterator]();

  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      const { value, done } = await iterator.next();
      if (done) {
        controller.close();
        return;
      }
      controller.enqueue(value);
    },
    async cancel() {
      try {
        await iterator.return?.();
      } catch {
        // ignore
      }
    }
  });
}

/**
 * Bytes held in memory, recognised as the platform's own readers recognise them: any
 * `ArrayBufferView` (a `Uint8Array`, a Node `Buffer`, a `DataView`, another typed array) or an
 * `ArrayBuffer`, including one from another realm, which fails `instanceof`.
 */
export type ArchiveBytes = ArrayBufferView | ArrayBuffer;

/** Whether `value` is {@link ArchiveBytes}. */
export function isArchiveBytes(value: unknown): value is ArchiveBytes {
  return ArrayBuffer.isView(value) || isArrayBufferLike(value);
}

function isArrayBufferLike(value: unknown): value is ArrayBuffer {
  if (value instanceof ArrayBuffer) {
    return true;
  }
  const tag = Object.prototype.toString.call(value);
  return tag === "[object ArrayBuffer]" || tag === "[object SharedArrayBuffer]";
}

/** {@link ArchiveBytes} as a `Uint8Array` over the same memory (no copy). */
export function archiveBytesToUint8Array(bytes: ArchiveBytes): Uint8Array {
  if (bytes instanceof Uint8Array) {
    return bytes;
  }
  if (ArrayBuffer.isView(bytes)) {
    return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  return new Uint8Array(bytes);
}

export function isInMemoryArchiveSource(
  source: ArchiveSource | ArchiveBytes
): source is Uint8Array | ArrayBuffer | string | Blob {
  return (
    isArchiveBytes(source) ||
    typeof source === "string" ||
    (typeof Blob !== "undefined" && source instanceof Blob)
  );
}

/** Refuse, where it is given, a value that is not an {@link ArchiveSource}. */
export function checkSourceArgument(where: string, source: unknown): void {
  checkArchiveArgument(
    s.custom(
      isArchiveSource,
      "a Uint8Array, ArrayBuffer, string, Blob, ReadableStream or async iterable"
    ),
    source,
    where,
    "source"
  );
}

const sourceOptionsSchema = s.object(
  {
    signal: s.abortSignal().optional(),
    onChunk: s.function<(chunk: Uint8Array) => void>().optional()
  },
  OPTIONS
);

/**
 * Refuse, where it is given, options that are not `{ signal?, onChunk? }`: read later, inside the
 * iteration, a `null` failed to destructure and a non-signal failed on `addEventListener`.
 */
export function checkSourceOptions(where: string, options: unknown): void {
  checkArchiveArgument(sourceOptionsSchema.optional(), options, where, "options");
}

/** `true` for every value {@link toAsyncIterable} can read: the {@link ArchiveSource} union. */
export function isArchiveSource(value: unknown): value is ArchiveSource {
  return (
    isInMemoryArchiveSource(value as ArchiveSource) ||
    isReadableStream(value) ||
    isAsyncIterable(value)
  );
}

function normalizeChunk(value: unknown): Uint8Array | null {
  if (!value) {
    return null;
  }

  if (value instanceof Uint8Array) {
    return value.length ? value : null;
  }

  if (typeof value === "string") {
    const bytes = encodeUtf8(value);
    return bytes.length ? bytes : null;
  }

  if (value instanceof ArrayBuffer) {
    return value.byteLength ? new Uint8Array(value) : null;
  }

  if (ArrayBuffer.isView(value)) {
    return value.byteLength
      ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
      : null;
  }

  // Strict mode: reject unknown chunk types.
  // Treating arbitrary array-like values as bytes can silently corrupt data
  // (e.g. number[] values get clamped to 0-255 when copied into Uint8Array).
  throw new ArchiveArgumentError(
    "source",
    `Unsupported archive source chunk type: ${Object.prototype.toString.call(value)}`
  );
}

export function toUint8ArraySync(source: Uint8Array | ArrayBuffer | string): Uint8Array {
  if (typeof source === "string") {
    return encodeUtf8(source);
  }
  return archiveBytesToUint8Array(source);
}

export function isSyncArchiveSource(source: unknown): source is Uint8Array | ArrayBuffer | string {
  return isArchiveBytes(source) || typeof source === "string";
}

export async function toUint8Array(
  source: Uint8Array | ArrayBuffer | string | Blob
): Promise<Uint8Array> {
  if (typeof source === "string") {
    return encodeUtf8(source);
  }
  if (isArchiveBytes(source)) {
    return archiveBytesToUint8Array(source);
  }
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/**
 * Resolve any ArchiveSource to a single Uint8Array buffer.
 *
 * This collects all chunks from streaming sources (AsyncIterable, ReadableStream)
 * into a single buffer. For large sources, this may use significant memory.
 *
 * @param source - Any ArchiveSource type
 * @param options - Options including abort signal
 * @returns Complete buffer containing all source data
 */
export async function resolveArchiveSourceToBuffer(
  source: ArchiveSource | ArchiveBytes,
  options: { signal?: AbortSignal } = {}
): Promise<Uint8Array> {
  // Fast path for in-memory sources
  if (isInMemoryArchiveSource(source)) {
    return toUint8Array(source as Uint8Array | ArrayBuffer | string | Blob);
  }

  // Streaming sources - collect all chunks
  const chunks: Uint8Array[] = [];
  let totalLength = 0;

  for await (const chunk of iterateArchiveSource(source, options)) {
    chunks.push(chunk);
    totalLength += chunk.length;
  }

  // Fast paths for common cases to avoid unnecessary allocation/copy.
  if (chunks.length === 0) {
    return EMPTY_UINT8ARRAY;
  }
  if (chunks.length === 1) {
    return chunks[0]!;
  }

  return concatUint8Arrays(chunks, totalLength);
}

export async function collectUint8ArrayStream(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
  options: { signal?: AbortSignal } = {}
): Promise<Uint8Array> {
  // Delegate to the general ArchiveSource collector so abort/cancellation semantics stay consistent.
  return resolveArchiveSourceToBuffer(stream, options);
}

/**
 * The bytes of any archive source, as `Uint8Array` chunks. Leaving early releases a caller's
 * Web stream (it stays theirs) and cancels a stream opened from a `Blob`.
 *
 * @param source - A string is the archive's content, encoded as UTF-8 — never a file path or URL;
 *   read a file's bytes first (in Node, or open it with `ArchiveFile.fromFile`).
 */
export function toAsyncIterable(
  source: ArchiveSource | ArchiveBytes,
  options: { signal?: AbortSignal; onChunk?: (chunk: Uint8Array) => void } = {}
): AsyncIterable<Uint8Array> {
  checkSourceArgument("Archive.toAsyncIterable", source);
  checkSourceOptions("Archive.toAsyncIterable", options);
  return iterateArchiveSource(source, options);
}

/** {@link toAsyncIterable} once its arguments are checked; the iteration starts on the first `next()`. */
export async function* iterateArchiveSource(
  source: ArchiveSource | ArchiveBytes,
  options: { signal?: AbortSignal; onChunk?: (chunk: Uint8Array) => void } = {}
): AsyncIterable<Uint8Array> {
  const { signal, onChunk } = options;

  const checkAborted = (): void => {
    if (signal?.aborted) {
      throw createAbortError(signal?.reason);
    }
  };

  checkAborted();

  if (isArchiveBytes(source) || typeof source === "string") {
    const bytes =
      typeof source === "string" ? encodeUtf8(source) : archiveBytesToUint8Array(source);
    if (onChunk) {
      onChunk(bytes);
    }
    yield bytes;
    return;
  }
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    // Prefer streaming the Blob to avoid a full Blob->ArrayBuffer copy.
    // This reduces memory use and improves performance for both small and large inputs.
    // `Blob.stream` may be missing on very old runtimes despite the type, so probe it.
    const maybeStream: unknown = source.stream;
    if (typeof maybeStream === "function") {
      // This stream is ours, not the caller's: leaving early (or failing) cancels it once the
      // inner read has released it, where a caller's stream is only released. Read to its
      // end, there is nothing to cancel.
      const stream = source.stream();
      let ended = false;
      try {
        yield* iterateArchiveSource(stream, options);
        ended = true;
      } finally {
        if (!ended) {
          await stream.cancel().catch(() => {
            // Already cancelled by an abort, or errored: nothing left to let go of
          });
        }
      }
      return;
    }

    // Fallback for very old environments.
    const bytes = await toUint8Array(source);
    checkAborted();
    if (onChunk) {
      onChunk(bytes);
    }
    yield bytes;
    return;
  }

  if (isReadableStream(source)) {
    const reader = source.getReader();

    let aborted = false;
    const onAbort = () => {
      aborted = true;
      try {
        reader.cancel();
      } catch {
        // ignore
      }
    };

    if (signal) {
      if (signal.aborted) {
        onAbort();
      } else {
        signal.addEventListener("abort", onAbort, { once: true });
      }
    }

    try {
      while (true) {
        if (aborted) {
          throw createAbortError(signal?.reason);
        }
        checkAborted();
        const { done, value } = await reader.read();
        if (done) {
          return;
        }

        const chunk = normalizeChunk(value);
        if (chunk) {
          if (aborted) {
            throw createAbortError(signal?.reason);
          }
          checkAborted();
          if (onChunk) {
            onChunk(chunk);
          }
          yield chunk;
        }
      }
    } finally {
      if (signal) {
        try {
          signal.removeEventListener("abort", onAbort);
        } catch {
          // ignore
        }
      }
      // Released, not cancelled, when the consumer stops early: the stream is the
      // caller's, and handing it back unlocked lets them keep reading it or cancel it
      // themselves. Only an abort cancels it, above. (A reader that owns the whole
      // parse, such as `unzip(stream)` or Excel's streaming workbook reader, cancels
      // the source itself once it is done with it.)
      try {
        reader.releaseLock();
      } catch {
        // Ignore
      }
    }
  }

  if (isAsyncIterable(source)) {
    for await (const value of source) {
      checkAborted();
      const chunk = normalizeChunk(value);
      if (chunk) {
        if (onChunk) {
          onChunk(chunk);
        }
        yield chunk;
      }
    }
    return;
  }

  throw new ArchiveArgumentError("source", "Unsupported archive source");
}

/**
 * Convert an AsyncIterable to a WHATWG ReadableStream.
 */
export function asyncIterableToReadableStream<T>(
  iterable: AsyncIterable<T>,
  onCancel?: (reason: unknown) => void
): ReadableStream<T> {
  const iterator = iterable[Symbol.asyncIterator]();
  let cancelled = false;

  return new ReadableStream<T>({
    async pull(controller) {
      if (cancelled) {
        controller.close();
        return;
      }

      try {
        const { value, done } = await iterator.next();
        if (done) {
          controller.close();
          return;
        }
        controller.enqueue(value);
      } catch (e) {
        controller.error(toError(e));
      }
    },
    async cancel(reason) {
      cancelled = true;
      try {
        await iterator.return?.();
      } catch {
        // ignore
      }
      onCancel?.(reason);
    }
  });
}
