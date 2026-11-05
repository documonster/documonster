/**
 * Stream Module - Common Utilities
 *
 * Platform-neutral utility functions shared by Node.js and Browser implementations:
 * - Stream state inspection (isDestroyed, isErrored)
 * - Callback-to-promise conversion (promisify)
 */

import { StreamArgumentError } from "@stream/errors";

// =============================================================================
// State Inspection
// =============================================================================

/**
 * Check if a stream has been destroyed
 */
export function isDestroyed(stream: { destroyed?: boolean } | null | undefined): boolean {
  return !!stream?.destroyed;
}

/**
 * Check if a stream has an error
 */
export function isErrored(stream: { errored?: unknown } | null | undefined): boolean {
  return !!stream?.errored;
}

// =============================================================================
// Promisify
// =============================================================================

/**
 * Convert a callback-based operation to a promise
 */
export function promisify<T>(
  fn: (callback: (error?: Error | null, result?: T) => void) => void
): Promise<T> {
  if (typeof fn !== "function") {
    return Promise.reject(
      new StreamArgumentError("a function", fn === null ? "null" : typeof fn, {
        field: "fn",
        where: "Streams.LowLevel.fromCallback"
      })
    );
  }
  return new Promise((resolve, reject) => {
    fn((error, result) => {
      if (error) {
        reject(error);
      } else {
        resolve(result as T);
      }
    });
  });
}
