/**
 * Stream Module - Common Utilities
 *
 * Platform-neutral utility functions shared by Node.js and Browser implementations:
 * - Stream state inspection (isDestroyed, isErrored)
 * - Callback-to-promise conversion (promisify)
 * - High water mark management (getDefaultHighWaterMark, setDefaultHighWaterMark)
 */

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

// =============================================================================
// High Water Mark
// =============================================================================

let _defaultHighWaterMark = 65536; // 64KB default (matches Node.js 25)
let _defaultHighWaterMarkObjectMode = 16; // 16 objects default

/**
 * Get the default high water mark for streams
 */
export function getDefaultHighWaterMark(objectMode: boolean): number {
  return objectMode ? _defaultHighWaterMarkObjectMode : _defaultHighWaterMark;
}

/**
 * Set the default high water mark for streams
 */
export function setDefaultHighWaterMark(objectMode: boolean, value: number): void {
  if (objectMode) {
    _defaultHighWaterMarkObjectMode = value;
  } else {
    _defaultHighWaterMark = value;
  }
}
