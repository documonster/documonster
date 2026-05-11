/**
 * StreamBuf - Minimal Event-Driven Stream Buffer for DOCX Streaming
 *
 * A lightweight stream buffer that emits data events when written to.
 * This is specifically designed for the streaming DOCX writer's needs:
 * - Write StringBuf or Uint8Array data
 * - Emit "data" events with Uint8Array chunks
 * - Emit "finish" on end()
 * - Support custom event listeners (once/on/emit)
 *
 * This is a word-module-local implementation to avoid depending on @excel/utils.
 */

import { EventEmitter } from "@utils/event-emitter";

import type { StringBuf } from "./string-buf";

interface StreamBufOptions {
  bufSize?: number;
}

/**
 * Minimal streaming buffer for piping XML data to ZIP compression.
 * Extends EventEmitter for event-driven data flow.
 */
class StreamBuf extends EventEmitter {
  private _ended: boolean;

  constructor(_options?: StreamBufOptions) {
    super();
    this._ended = false;
  }

  /** Returns true if the stream is writable (not ended). */
  get writable(): boolean {
    return !this._ended;
  }

  /**
   * Write data to the stream. Emits a "data" event with the bytes.
   */
  write(
    data: Uint8Array | string | StringBuf | ArrayBuffer | ArrayBufferView,
    _encoding?: string | ((...args: unknown[]) => unknown),
    callback?: (...args: unknown[]) => unknown
  ): void {
    let bytes: Uint8Array;

    if (data instanceof Uint8Array) {
      bytes = data;
    } else if (typeof data === "string") {
      bytes = new TextEncoder().encode(data);
    } else if (ArrayBuffer.isView(data)) {
      bytes = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    } else if (data instanceof ArrayBuffer) {
      bytes = new Uint8Array(data);
    } else if (data && typeof (data as { toBuffer?: unknown }).toBuffer === "function") {
      // StringBuf-like object
      bytes = (data as { toBuffer: () => Uint8Array }).toBuffer();
    } else {
      bytes = new Uint8Array(0);
    }

    if (bytes.length > 0) {
      this.emit("data", bytes);
    }

    if (callback) {
      callback();
    }
  }

  /**
   * End the stream. Emits "finish" event.
   */
  end(
    chunk?: Uint8Array | string | StringBuf | ArrayBuffer | ArrayBufferView,
    _encoding?: string,
    callback?: (...args: unknown[]) => unknown
  ): void {
    if (chunk) {
      this.write(chunk);
    }

    this._ended = true;
    this.emit("finish");
    this.emit("close");

    if (callback) {
      callback();
    }
  }
}

export { StreamBuf };
export type { StreamBufOptions };
