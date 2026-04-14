import { EMPTY_UINT8ARRAY } from "@archive/shared/bytes";
import { createAbortError } from "@archive/shared/errors";
import { isAsyncIterable, isReadableStream } from "@stream/internal/type-guards";
import { stringToUint8Array as encodeUtf8, concatUint8Arrays } from "@utils/binary";

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
 */
export function toReadableStream(
  source: ArchiveSource,
  options: { signal?: AbortSignal } = {}
): ReadableStream<Uint8Array> {
  const iterable = toAsyncIterable(source, options);
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

export function isInMemoryArchiveSource(
  source: ArchiveSource
): source is Uint8Array | ArrayBuffer | string | Blob {
  return (
    source instanceof Uint8Array ||
    source instanceof ArrayBuffer ||
    typeof source === "string" ||
    (typeof Blob !== "undefined" && source instanceof Blob)
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
  throw new TypeError(
    `Unsupported archive source chunk type: ${Object.prototype.toString.call(value)}`
  );
}

export function toUint8ArraySync(source: Uint8Array | ArrayBuffer | string): Uint8Array {
  if (source instanceof Uint8Array) {
    return source;
  }
  if (typeof source === "string") {
    return encodeUtf8(source);
  }
  return new Uint8Array(source);
}

export function isSyncArchiveSource(source: unknown): source is Uint8Array | ArrayBuffer | string {
  return (
    source instanceof Uint8Array || source instanceof ArrayBuffer || typeof source === "string"
  );
}

export async function toUint8Array(
  source: Uint8Array | ArrayBuffer | string | Blob
): Promise<Uint8Array> {
  if (source instanceof Uint8Array) {
    return source;
  }
  if (typeof source === "string") {
    return encodeUtf8(source);
  }
  if (source instanceof ArrayBuffer) {
    return new Uint8Array(source);
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
  source: ArchiveSource,
  options: { signal?: AbortSignal } = {}
): Promise<Uint8Array> {
  // Fast path for in-memory sources
  if (isInMemoryArchiveSource(source)) {
    return toUint8Array(source as Uint8Array | ArrayBuffer | string | Blob);
  }

  // Streaming sources - collect all chunks
  const chunks: Uint8Array[] = [];
  let totalLength = 0;

  for await (const chunk of toAsyncIterable(source, options)) {
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
  return resolveArchiveSourceToBuffer(stream as unknown as ArchiveSource, options);
}

export async function* toAsyncIterable(
  source: ArchiveSource,
  options: { signal?: AbortSignal; onChunk?: (chunk: Uint8Array) => void } = {}
): AsyncIterable<Uint8Array> {
  const { signal, onChunk } = options;

  const checkAborted = (): void => {
    if (signal?.aborted) {
      throw createAbortError((signal as any).reason);
    }
  };

  checkAborted();

  if (source instanceof Uint8Array) {
    if (onChunk) {
      onChunk(source);
    }
    yield source;
    return;
  }
  if (typeof source === "string") {
    const bytes = encodeUtf8(source);
    if (onChunk) {
      onChunk(bytes);
    }
    yield bytes;
    return;
  }
  if (source instanceof ArrayBuffer) {
    const bytes = new Uint8Array(source);
    if (onChunk) {
      onChunk(bytes);
    }
    yield bytes;
    return;
  }
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    // Prefer streaming the Blob to avoid a full Blob->ArrayBuffer copy.
    // This reduces memory use and improves performance for both small and large inputs.
    const maybeStream = (source as any).stream;
    if (typeof maybeStream === "function") {
      const stream = (source as any).stream();
      yield* toAsyncIterable(stream, options);
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
          throw createAbortError((signal as any).reason);
        }
        checkAborted();
        const { done, value } = await reader.read();
        if (done) {
          return;
        }

        const chunk = normalizeChunk(value);
        if (chunk) {
          if (aborted) {
            throw createAbortError((signal as any).reason);
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

  throw new Error("Unsupported archive source");
}
