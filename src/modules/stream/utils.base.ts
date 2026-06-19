/**
 * Stream Utilities (shared base)
 *
 * Platform-independent helpers re-exported by both `utils.ts` (Node) and
 * `utils.browser.ts`.  The two platform files supply the concrete
 * `createReadableFromArray`, `createTransform` and `consumers` bindings.
 */

import { isReadableStream } from "@stream/core/type-guards";
import type { IReadable, ITransform } from "@stream/types";
import { stringToUint8Array as _stringToUint8Array } from "@utils/binary";

// Platform-specific dependencies injected by the caller.
export interface UtilsDeps {
  createReadableFromArray: <T>(data: T[], options?: { objectMode?: boolean }) => IReadable<T>;
  createTransform: <TIn, TOut>(
    fn: (chunk: TIn) => TOut | undefined | Promise<TOut | undefined>,
    options?: { objectMode?: boolean }
  ) => ITransform<TIn, TOut>;
  consumers: {
    text: (stream: AsyncIterable<Uint8Array>) => Promise<string>;
    json: (stream: AsyncIterable<Uint8Array>) => Promise<unknown>;
    buffer: (stream: AsyncIterable<Uint8Array>) => Promise<Uint8Array>;
  };
}

// =============================================================================
// High-Level Stream Consumers
// =============================================================================

export async function collect<T>(stream: {
  [Symbol.asyncIterator](): AsyncIterator<T>;
}): Promise<T[]> {
  const result: T[] = [];
  for await (const chunk of stream) {
    result.push(chunk);
  }
  return result;
}

export function createText(deps: UtilsDeps) {
  return async function text(stream: {
    [Symbol.asyncIterator](): AsyncIterator<Uint8Array>;
  }): Promise<string> {
    return deps.consumers.text(stream as AsyncIterable<Uint8Array>);
  };
}

export function createJson(deps: UtilsDeps) {
  return async function json<T = unknown>(stream: {
    [Symbol.asyncIterator](): AsyncIterator<Uint8Array>;
  }): Promise<T> {
    return deps.consumers.json(stream as AsyncIterable<Uint8Array>) as Promise<T>;
  };
}

export function createBytes(deps: UtilsDeps) {
  return async function bytes(stream: {
    [Symbol.asyncIterator](): AsyncIterator<Uint8Array>;
  }): Promise<Uint8Array> {
    return deps.consumers.buffer(stream as AsyncIterable<Uint8Array>);
  };
}

// =============================================================================
// Stream Factory Helpers
// =============================================================================

export function createFromString(deps: UtilsDeps) {
  return function fromString(str: string): IReadable<Uint8Array> {
    return deps.createReadableFromArray([_stringToUint8Array(str)], {
      objectMode: false
    });
  };
}

export function createFromJSON(deps: UtilsDeps) {
  const fromString = createFromString(deps);
  return function fromJSON(data: unknown): IReadable<Uint8Array> {
    return fromString(JSON.stringify(data));
  };
}

export function createFromBytes(deps: UtilsDeps) {
  return function fromBytes(data: Uint8Array): IReadable<Uint8Array> {
    return deps.createReadableFromArray([data], { objectMode: false });
  };
}

export function createTransformHelper(deps: UtilsDeps) {
  return function transform<TIn = Uint8Array, TOut = TIn>(
    fn: (chunk: TIn) => TOut | Promise<TOut>
  ): ITransform<TIn, TOut> {
    return deps.createTransform<TIn, TOut>(fn);
  };
}

export function createFilter(deps: UtilsDeps) {
  return function filter<T>(predicate: (chunk: T) => boolean | Promise<boolean>): ITransform<T, T> {
    return deps.createTransform<T, T>(
      async chunk => {
        if (await predicate(chunk)) {
          return chunk;
        }
        return undefined;
      },
      { objectMode: true }
    );
  };
}

// =============================================================================
// ReadableStream Conversion
// =============================================================================

/**
 * Type guard for browser ReadableStream-like objects.
 * Re-exported from internal/type-guards for public API compatibility.
 */
export const isReadableStreamLike = isReadableStream as (
  value: unknown
) => value is { getReader: () => any };

/**
 * Convert a browser ReadableStream to an AsyncIterable.
 * This is useful for consuming fetch response bodies in a streaming fashion.
 *
 * @example
 * ```ts
 * const response = await fetch(url);
 * for await (const chunk of readableStreamToAsyncIterable(response.body)) {
 *   // process chunk
 * }
 * ```
 */
export async function* readableStreamToAsyncIterable<T = Uint8Array>(stream: {
  getReader: () => any;
}): AsyncGenerator<T, void, unknown> {
  const reader = stream.getReader();
  try {
    while (true) {
      const result = await reader.read();
      if (result.done) {
        return;
      }
      // Only yield defined, non-null values — avoids silently dropping
      // falsy-but-valid values like 0 or "".
      if (result.value !== undefined && result.value !== null) {
        yield result.value as T;
      }
    }
  } finally {
    // Cancel the reader before releasing the lock so the underlying stream
    // can free resources (e.g. a fetch body).  Without cancel(), the stream
    // stays in a "disturbed but not cancelled" state per the WHATWG spec.
    try {
      await reader.cancel?.();
    } catch {
      // ignore
    }
    try {
      reader.releaseLock?.();
    } catch {
      // ignore
    }
  }
}
