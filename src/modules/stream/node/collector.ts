/**
 * Node.js Stream - Collector
 *
 * A writable stream that collects all chunks.
 */

import { toBinaryChunk } from "@stream/common/binary-chunk";
import { StreamTypeError } from "@stream/errors";
import type { WritableStreamOptions, ICollector } from "@stream/types";
import { chunksToString, concatUint8Arrays } from "@utils/binary";

import { Writable } from "./writable";

// =============================================================================
// Collector Stream - Collects all chunks into an array
// =============================================================================

/**
 * A writable stream that collects all chunks
 */
export class Collector<T = Uint8Array> extends Writable<T> {
  public chunks: T[] = [];

  constructor(options?: WritableStreamOptions) {
    super({
      ...options,
      objectMode: options?.objectMode ?? true,
      write: ((chunk: T, _encoding: BufferEncoding, callback: (error?: Error | null) => void) => {
        this.chunks.push(chunk);
        callback();
      }) as any
    });
  }

  /**
   * Get all collected data as a single Uint8Array (for binary mode)
   */
  toUint8Array(): Uint8Array {
    const chunks = this.chunks;
    const len = chunks.length;
    if (len === 0) {
      return new Uint8Array(0);
    }

    const binaryChunks = new Array<Uint8Array>(len);
    let totalLength = 0;
    for (let i = 0; i < len; i++) {
      const normalized = toBinaryChunk(chunks[i]);
      if (!normalized) {
        throw new StreamTypeError("Uint8Array", "non-binary data");
      }
      binaryChunks[i] = normalized;
      totalLength += normalized.length;
    }

    if (len === 1) {
      return binaryChunks[0]!;
    }

    return concatUint8Arrays(binaryChunks, totalLength);
  }

  /**
   * Get all collected data as a string
   */
  override toString(): string {
    return chunksToString(this.chunks, () => this.toUint8Array());
  }

  /**
   * Whether the collector has finished receiving data
   */
  get isFinished(): boolean {
    return this.writableFinished;
  }
}

/**
 * Create a collector stream
 */
export function createCollector<T = Uint8Array>(options?: WritableStreamOptions): ICollector<T> {
  return new Collector<T>(options);
}
