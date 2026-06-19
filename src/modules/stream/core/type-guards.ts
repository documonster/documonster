/**
 * Stream Module - Common Type Guard Factories
 *
 * Creates type guard functions parameterized by platform-specific
 * class constructors (Node.js native vs browser polyfill classes).
 */

import type { IDuplex, ITransform, ReadableLike, WritableLike } from "@stream/types";

// =============================================================================
// Types
// =============================================================================

type Constructor = abstract new (...args: any[]) => any;

// =============================================================================
// Factories
// =============================================================================

/**
 * Create `isTransform` bound to platform-specific Transform class(es).
 */
export function createIsTransform(
  ...classes: Constructor[]
): (obj: unknown) => obj is ITransform<any, any> {
  return function isTransform(obj: unknown): obj is ITransform<any, any> {
    if (obj == null) {
      return false;
    }
    for (const cls of classes) {
      if (obj instanceof cls) {
        return true;
      }
    }
    const o = obj as Record<string, unknown>;
    return (
      typeof o.read === "function" &&
      typeof o.pipe === "function" &&
      typeof o.write === "function" &&
      typeof o.end === "function" &&
      typeof o._transform === "function"
    );
  };
}

/**
 * Create `isDuplex` bound to platform-specific Duplex/Transform class(es).
 */
export function createIsDuplex(
  ...classes: Constructor[]
): (obj: unknown) => obj is IDuplex<any, any> {
  return function isDuplex(obj: unknown): obj is IDuplex<any, any> {
    if (obj == null) {
      return false;
    }
    for (const cls of classes) {
      if (obj instanceof cls) {
        return true;
      }
    }
    const o = obj as Record<string, unknown>;
    return (
      typeof o.read === "function" &&
      typeof o.pipe === "function" &&
      typeof o.write === "function" &&
      typeof o.end === "function"
    );
  };
}

/**
 * Create `isStream` bound to platform-specific Readable/Writable class(es).
 */
export function createIsStream(
  ...classes: Constructor[]
): (obj: unknown) => obj is ReadableLike | WritableLike {
  return function isStream(obj: unknown): obj is ReadableLike | WritableLike {
    if (obj == null) {
      return false;
    }
    for (const cls of classes) {
      if (obj instanceof cls) {
        return true;
      }
    }
    const o = obj as Record<string, unknown>;
    return (
      (typeof o.read === "function" && typeof o.pipe === "function") ||
      (typeof o.write === "function" && typeof o.end === "function")
    );
  };
}

// =============================================================================
// Web Stream Guards
// =============================================================================

/**
 * Lightweight runtime type guards shared across modules.
 *
 * Keep these dependency-free to maximize deduping in bundled builds.
 */

export function isReadableStream(value: unknown): value is ReadableStream<unknown> {
  return !!value && typeof value === "object" && typeof (value as any).getReader === "function";
}

export function isWritableStream(value: unknown): value is WritableStream<unknown> {
  return !!value && typeof value === "object" && typeof (value as any).getWriter === "function";
}

export function isAsyncIterable(value: unknown): value is AsyncIterable<unknown> {
  return (
    !!value &&
    (typeof value === "object" || typeof value === "function") &&
    typeof (value as any)[Symbol.asyncIterator] === "function"
  );
}

export function isTransformStream(value: unknown): value is TransformStream<unknown, unknown> {
  return (
    !!value &&
    typeof value === "object" &&
    !!(value as any).readable &&
    !!(value as any).writable &&
    isReadableStream((value as any).readable) &&
    isWritableStream((value as any).writable)
  );
}
