/**
 * Browser Stream - PassThrough
 */

import { Transform } from "@stream/browser/transform";

// =============================================================================
// PassThrough Stream
// =============================================================================

/**
 * A passthrough stream that passes data through unchanged.
 * Uses a prototype _transform override (matching Node.js PassThrough behavior).
 */
export class PassThrough<T = Uint8Array> extends Transform<T, T> {
  override _transform(
    chunk: T,
    _encoding: string,
    callback: (error?: Error | null, data?: T) => void
  ): void {
    callback(null, chunk);
  }
}
