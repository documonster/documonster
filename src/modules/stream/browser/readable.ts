/**
 * Browser Stream - Readable
 */

import { getDuplexFrom } from "@stream/browser/_lazy";
import { ChunkBuffer } from "@stream/browser/chunk-buffer";
import { deferTask, inDeferredContext } from "@stream/browser/microtask-context";
import { PipeManager } from "@stream/browser/pipe-manager";
import { stringToEncodedBytes } from "@stream/core/binary-chunk";
import { getDefaultHighWaterMark } from "@stream/core/utils";
import type { IDuplex, ReadableStreamOptions, WritableLike } from "@stream/types";
import { createStreamDecoder, decodeBytesToString } from "@utils/binary";
import type { StreamDecoder } from "@utils/binary";
import { createAbortError, toError } from "@utils/errors";
import { EventEmitter } from "@utils/event-emitter";

/**
 * Shared toString implementation for Uint8Array chunks converted from strings.
 * Uses `this`-binding to avoid per-chunk closure allocation.
 * Supports all Node.js Buffer encodings (hex, base64, base64url, ascii, etc.).
 */
function encodedBytesToString(this: Uint8Array, enc?: string): string {
  return decodeBytesToString(this, enc ?? "utf-8");
}

// =============================================================================
// Readable Stream Wrapper
// =============================================================================

/**
 * Shared state for async iteration (Node.js parity).
 *
 * In Node.js, multiple `[Symbol.asyncIterator]()` calls return independent
 * iterators that *compete* for chunks from the same underlying stream — each
 * chunk is delivered to exactly one waiting iterator (FIFO).  We replicate
 * this by sharing a single `data` listener and routing incoming chunks to
 * a queue of per-iterator resolvers.
 */
interface _AsyncIterState<T> {
  /** Chunks that arrived when no iterator was waiting. */
  dataQueue: any[];
  dataQueueIndex: number;
  queuedSize: number;
  /** Per-iterator resolve/reject callbacks waiting for the next chunk. */
  resolverQueue: Array<{
    resolve: (v: any) => void;
    reject: (e: Error) => void;
  }>;
  done: boolean;
  streamError: Error | null;
  /** Number of active (not-yet-returned) iterators. */
  activeCount: number;
  /** Whether the shared event listeners have been attached. */
  listenersAttached: boolean;
  pausedByIterator: boolean;
  /** Bound handlers for cleanup. */
  dataHandler: (chunk: T) => void;
  doneHandler: () => void;
  errorHandler: (err: Error) => void;
}

/**
 * A wrapper around Web ReadableStream that provides Node.js-like API
 */
export class Readable<T = Uint8Array> extends EventEmitter {
  /**
   * Allow duck-typed instanceof checks.
   * Node.js Duplex passes `instanceof Readable` via prototype chain.
   * Our browser Duplex composes a Readable, so we use Symbol.hasInstance
   * to check for key Readable-like methods/properties.
   */
  static [Symbol.hasInstance](instance: any): boolean {
    if (instance == null || typeof instance !== "object") {
      return false;
    }
    // Fast path: actual Readable prototype
    if (Object.prototype.isPrototypeOf.call(Readable.prototype, instance)) {
      return true;
    }
    // Duck-type: must have key Readable methods and the stream brand
    return (
      instance.__documonster_stream === true &&
      typeof instance.read === "function" &&
      typeof instance.pipe === "function" &&
      typeof instance.on === "function" &&
      "readableFlowing" in instance
    );
  }

  private _stream: ReadableStream<T> | null;
  private _reader: ReadableStreamDefaultReader<T> | null = null;
  private _buf!: ChunkBuffer<T>;
  private _reading: boolean = false;
  private _ended: boolean = false;
  private _endEmitted: boolean = false;
  private _destroyed: boolean = false;
  /** @internal Set by _callRead catch to signal afterDestroy to emit synchronously */
  private _internalDestroy: boolean = false;
  private _errored: Error | null = null;
  private _closed: boolean = false;
  private _flowing: boolean = false;
  private _resumeScheduled: boolean = false;
  private _hasFlowed: boolean = false;
  private _pipes!: PipeManager<T>;
  private _encoding: string | null = null;
  private _decoder: StreamDecoder | null = null;
  private _didRead: boolean = false;
  // Whether this stream uses push() mode (true) or Web Stream mode (false)
  private _pushMode: boolean = false;
  // Whether this stream was created from an external Web Stream (true) or is controllable (false)
  private _webStreamMode: boolean = false;
  private _objectMode: boolean;
  private _highWaterMark: number;
  private _autoDestroy: boolean;
  private _emitClose: boolean;
  // User-provided read function (Node.js compatibility)
  private _read?: (size?: number) => void;
  // True when a user-provided _read exists (via options or subclass override).
  // Used to gate _callRead invocations: the prototype _read throws
  // ERR_METHOD_NOT_IMPLEMENTED (Node.js parity), so we must not call it
  // speculatively during the flow loop for push-driven / Web-Stream-backed
  // streams that never intend to supply a pull-based _read.
  private _hasReadImpl: boolean = false;
  // User-provided construct function (Node.js compatibility)
  private _constructFunc?: (callback: (error?: Error | null) => void) => void;
  private _constructed: boolean = true;

  // Shared async-iterator state (Node.js parity: multiple iterators compete for chunks)
  private _asyncIterState: _AsyncIterState<T> | null = null;

  constructor(
    options?: ReadableStreamOptions & {
      stream?: ReadableStream<T>;
      autoDestroy?: boolean;
      emitClose?: boolean;
      signal?: AbortSignal;
      encoding?: string;
      read?: (this: Readable<T>, size?: number) => void;
      destroy?: (
        this: Readable<T>,
        error: Error | null,
        callback: (error?: Error | null) => void
      ) => void;
      construct?: (this: Readable<T>, callback: (error?: Error | null) => void) => void;
    }
  ) {
    super();
    (this as any).__documonster_stream = true;
    this._objectMode = options?.objectMode ?? false;
    this._highWaterMark = options?.highWaterMark ?? getDefaultHighWaterMark(this._objectMode);
    this._buf = new ChunkBuffer<T>(this._objectMode);
    this._pipes = new PipeManager<T>(this);
    this._autoDestroy = options?.autoDestroy ?? true;
    this._emitClose = options?.emitClose ?? true;

    // Store user-provided read function
    if (options?.read) {
      this._read = options.read.bind(this);
      this._hasReadImpl = true;
      this._pushMode = true; // User will call push()
    } else if (this._hasSubclassRead()) {
      // Subclass defines _read on its prototype — bind and mark as available.
      // This matches Node.js where subclass._read() is the canonical pull API.
      this._hasReadImpl = true;
      this._pushMode = true;
    }

    // Store user-provided destroy function
    if (options?.destroy) {
      this._destroy = options.destroy.bind(this);
    }

    // Store user-provided construct function
    if (options?.construct) {
      this._constructFunc = options.construct.bind(this);
    }

    if (options?.stream) {
      this._stream = options.stream;
      this._webStreamMode = true; // Created from external Web Stream
    } else {
      // Controllable stream - no need to eagerly create a Web ReadableStream.
      // The webStream getter will create one lazily when accessed.
      this._stream = null;
    }

    // M2: encoding constructor option — call setEncoding() if provided
    if (options?.encoding) {
      this.setEncoding(options.encoding);
    }

    // M1: signal constructor option — destroy stream when signal aborts
    if (options?.signal) {
      this._setupAbortSignal(options.signal);
    }

    // L2: _construct hook — if provided, delay _read until constructed
    this._maybeConstruct();
  }

  /**
   * Run _construct if provided (via options or subclass override).
   * Delays _read calls until the callback fires.
   */
  private _maybeConstruct(): void {
    const hasConstructHook = this._constructFunc || this._hasSubclassConstruct();
    if (!hasConstructHook) {
      return;
    }
    this._constructed = false;
    // Call _construct on next microtask (matches Node.js which uses process.nextTick)
    deferTask(() => {
      const fn = this._constructFunc ?? (this as any)._construct.bind(this);
      fn(err => {
        if (err) {
          this.destroy(err);
          return;
        }
        this._constructed = true;
        // If _read was requested while not yet constructed, call it now
        if (this._hasReadImpl && this._flowing && !this._ended && !this._destroyed) {
          this._callRead(this._highWaterMark);
        }
      });
    });
  }

  /**
   * Check if a subclass defines _construct on its own prototype.
   * Node.js does NOT have _construct on any stream prototype — it only exists
   * when provided via constructor options or defined by a subclass.
   */
  private _hasSubclassConstruct(): boolean {
    let proto = Object.getPrototypeOf(this);
    while (proto && proto !== Readable.prototype && proto !== Object.prototype) {
      if (Object.prototype.hasOwnProperty.call(proto, "_construct")) {
        return true;
      }
      proto = Object.getPrototypeOf(proto);
    }
    return false;
  }

  /**
   * Check if a subclass defines _read on its own prototype.
   * The base Readable.prototype._read throws ERR_METHOD_NOT_IMPLEMENTED,
   * so we must stop before it. Only return true if an intermediate prototype
   * (the user's subclass) owns _read.
   */
  private _hasSubclassRead(): boolean {
    let proto = Object.getPrototypeOf(this);
    while (proto && proto !== Readable.prototype && proto !== Object.prototype) {
      if (Object.prototype.hasOwnProperty.call(proto, "_read")) {
        return true;
      }
      proto = Object.getPrototypeOf(proto);
    }
    return false;
  }

  /**
   * Wire up an AbortSignal to destroy this stream on abort.
   */
  private _setupAbortSignal(signal: AbortSignal): void {
    if (signal.aborted) {
      this.destroy(createAbortError((signal as any).reason));
      return;
    }

    const onAbort = (): void => {
      cleanup();
      this.destroy(createAbortError((signal as any).reason));
    };

    const onDone = (): void => {
      cleanup();
    };

    const cleanup = (): void => {
      signal.removeEventListener("abort", onAbort);
      this.off("close", onDone);
      this.off("end", onDone);
      this.off("error", onDone);
    };

    signal.addEventListener("abort", onAbort, { once: true });
    this.on("close", onDone);
    this.on("end", onDone);
    this.on("error", onDone);
  }

  /**
   * Create a Readable from an iterable (static factory method)
   */
  static from<T>(
    iterable: Iterable<T> | AsyncIterable<T> | ReadableStream<T> | Blob,
    options?: ReadableStreamOptions
  ): Readable<T> {
    // Node.js also supports creating from a Web ReadableStream.
    // Detect it explicitly (do not rely on Symbol.asyncIterator presence).
    if (iterable && typeof (iterable as any).getReader === "function") {
      return Readable.fromWeb(iterable as ReadableStream<T>, options);
    }

    // Node.js: Blob → stream via its ReadableStream.
    // Must be checked before the generic Iterable path because Blob is not
    // iterable but has a .stream() method that returns a ReadableStream.
    if (typeof Blob !== "undefined" && iterable instanceof Blob) {
      return Readable.fromWeb((iterable as Blob).stream() as ReadableStream<T>, options);
    }

    // Validate argument type early (Node.js throws ERR_INVALID_ARG_TYPE).
    if (iterable == null || (typeof iterable !== "object" && typeof iterable !== "string")) {
      const err = new TypeError(
        `The "iterable" argument must be an instance of Iterable. Received type ${typeof iterable} (${String(
          iterable
        )})`
      ) as TypeError & { code: string };
      err.code = "ERR_INVALID_ARG_TYPE";
      throw err;
    }
    const hasIter = typeof (iterable as any)[Symbol.iterator] === "function";
    const hasAsyncIter = typeof (iterable as any)[Symbol.asyncIterator] === "function";
    if (!hasIter && !hasAsyncIter && typeof iterable !== "string") {
      const name = (iterable as any)?.constructor?.name ?? "Object";
      const err = new TypeError(
        `The "iterable" argument must be an instance of Iterable. Received an instance of ${name}`
      ) as TypeError & { code: string };
      err.code = "ERR_INVALID_ARG_TYPE";
      throw err;
    }

    const readable = new Readable<T>({
      ...options,
      objectMode: options?.objectMode ?? true
    });

    const iter = iterable as unknown;

    // Node.js treats strings as a single chunk, not as Iterable<char>.
    // Match that behavior by wrapping strings in an array.
    const source =
      typeof iter === "string"
        ? toAsyncIterable([iter] as Iterable<T>)
        : toAsyncIterable(iter as Iterable<T> | AsyncIterable<T>);
    pumpAsyncIterableToReadable(readable, source);

    return readable;
  }

  /**
   * Check if a stream has been disturbed (read from).
   * Matches Node.js native: readableDidRead || readableAborted.
   * readableAborted = (destroyed || errored) && !endEmitted.
   */
  static isDisturbed(stream: Readable<any>): boolean {
    return (
      stream._didRead || ((stream._destroyed || stream._errored !== null) && !stream._endEmitted)
    );
  }

  /**
   * Check if a stream is still readable (not destroyed, not ended).
   * Matches Node.js Readable.isReadable() (available since v21).
   * Returns null for non-stream inputs (matching Node.js behavior).
   */
  static isReadable(stream: unknown): boolean | null {
    if (stream == null || typeof stream !== "object") {
      return null;
    }
    const s = stream as any;
    // Not a stream-like object — return null (matches Node.js)
    if (typeof s.read !== "function") {
      return null;
    }
    // If destroyed or errored, not readable
    if (s.destroyed || s._destroyed) {
      return false;
    }
    // If ended, not readable
    if (s.readableEnded || s._endEmitted) {
      return false;
    }
    // Must have readable property
    return s.readable !== false;
  }

  /**
   * Convert a Web ReadableStream to Node.js Readable
   */
  static fromWeb<T>(webStream: ReadableStream<T>, options?: ReadableStreamOptions): Readable<T> {
    return new Readable<T>({ ...options, stream: webStream });
  }

  /**
   * Convert a Node.js Readable to Web ReadableStream
   */
  static toWeb<T>(nodeStream: Readable<T>): ReadableStream<T> {
    return nodeStream._webStream;
  }

  /**
   * Static wrap method - wraps an old-style stream into a Readable.
   * Matches Node.js Readable.wrap(stream, options).
   */
  static wrap<T>(src: any, options?: ReadableStreamOptions): Readable<T> {
    return new Readable<T>({
      objectMode: src.readableObjectMode ?? src.objectMode ?? true,
      ...options,
      destroy(err, callback) {
        if (typeof src.destroy === "function") {
          src.destroy(err ?? undefined);
        }
        callback(err);
      }
    }).wrap(src);
  }

  /**
   * Push data to the stream (when using controllable stream)
   */
  push(chunk: T | null, encoding?: string): boolean {
    if (this._destroyed) {
      return false;
    }

    // Mark as push mode when push() is called
    this._pushMode = true;

    // Handle string encoding (Node.js compatibility)
    // Node.js always converts strings to Buffer in binary mode, defaulting to utf8.
    // We override toString() so it behaves like Node.js Buffer.toString().
    if (chunk !== null && typeof chunk === "string" && !this._objectMode) {
      const encoded = stringToEncodedBytes(chunk as string, encoding || "utf8");
      (encoded as any).toString = encodedBytesToString;
      chunk = encoded as any;
    }

    // Reject push() after EOF (matches Node.js ERR_STREAM_PUSH_AFTER_EOF)
    if (this._ended && chunk !== null) {
      const err = new Error("stream.push() after EOF") as Error & {
        code: string;
      };
      err.code = "ERR_STREAM_PUSH_AFTER_EOF";
      deferTask(() => this.emit("error", err));
      return false;
    }

    if (chunk === null) {
      // Prevent duplicate end handling
      if (this._ended) {
        return false;
      }
      this._ended = true;

      // Emit 'end' only after buffered data is fully drained.
      // This avoids premature 'end' when producers push null while paused.
      // Defer via deferTask to match Node.js process.nextTick behavior —
      // synchronous code after push(null) should not see 'end' yet.
      if (this._buf.length === 0) {
        deferTask(() => this._emitEndOnce());
      }
      // Note: Don't call destroy() here, let the stream be consumed naturally
      // The reader will return done:true when it finishes reading
      return false;
    }

    // Node.js: In binary mode, pushing a zero-length chunk (empty Buffer/Uint8Array
    // or string that encodes to empty) is a no-op — it is silently ignored.
    // In object mode, all values (including empty buffers) are valid chunks.
    if (!this._objectMode) {
      const len =
        typeof chunk === "string"
          ? (chunk as string).length
          : ((chunk as any)?.byteLength ?? (chunk as any)?.length ?? 0);
      if (len === 0) {
        return true;
      }
    }

    if (this._flowing) {
      // IMPORTANT: If there is still buffered data waiting to be drained (from a
      // resume() microtask that hasn't fired yet), we must NOT emit this chunk
      // directly — doing so would deliver it out-of-order, ahead of the buffered
      // chunks.  Buffer it instead and let the drain microtask deliver everything
      // in the correct sequence.
      if (this._buf.length > 0) {
        this._buf.push(chunk);
        return this._objectMode
          ? this._buf.length < this._highWaterMark
          : this._buf.byteSize < this._highWaterMark;
      }

      // In flowing mode with empty buffer, emit data directly
      this._didRead = true;
      this.emit("data", this._applyEncoding(chunk));
      // Check if stream was paused during emit (backpressure from consumer)
      if (!this._flowing) {
        return false;
      }
      // After emitting data, call _read again if available (Node.js behavior)
      if (this._hasReadImpl && !this._ended) {
        deferTask(() => {
          if (this._flowing && !this._ended && !this._destroyed) {
            this._callRead(this._highWaterMark);
            // Sync end check after _callRead — see resume() for rationale.
            if (this._ended && this._buf.length === 0) {
              this._emitEndOnce();
            }
          }
        });
      }
      // In flowing mode, return true (no backpressure since data is immediately consumed)
      return true;
    } else {
      // In paused mode, buffer for later
      const wasEmpty = this._buf.length === 0;
      this._buf.push(chunk);

      // Emit readable event when buffer goes from empty to having data
      if (wasEmpty) {
        deferTask(() => this.emit("readable"));
      }
      // Return false if buffer exceeds high water mark (backpressure signal)
      // Fast path for object mode - just count items
      if (this._objectMode) {
        return this._buf.length < this._highWaterMark;
      }
      // For binary mode, use tracked buffer size (O(1))
      return this._buf.byteSize < this._highWaterMark;
    }
  }

  private _emitEndOnce(): void {
    if (this._endEmitted || this._destroyed) {
      return;
    }
    // Guard: if data was unshifted after push(null) scheduled us,
    // the buffer is non-empty — defer 'end' until data is consumed.
    if (this._buf.length > 0) {
      return;
    }
    // Flush the stream decoder's remainder (e.g. partial base64 group).
    // This mirrors Node.js StringDecoder.end() which is called before 'end'.
    const flushed = this._flushDecoder();
    if (flushed) {
      this.emit("data", flushed);
    }
    this._endEmitted = true;
    this.emit("end");

    // Match Node.js autoDestroy behavior: automatically destroy after end
    if (this._autoDestroy) {
      this.destroy();
    }
  }

  /**
   * Put a chunk back at the front of the buffer
   */
  unshift(chunk: T, encoding?: string): void {
    if (this._destroyed) {
      return;
    }
    // Node.js: unshift(null) signals EOF, just like push(null).
    // Note: unshift(undefined) does NOT signal EOF in Node.js — only null does.
    if (chunk === null) {
      if (!this._ended) {
        this._ended = true;
        if (this._buf.length === 0) {
          deferTask(() => this._emitEndOnce());
        }
      }
      return;
    }
    // Node.js emits ERR_STREAM_UNSHIFT_AFTER_END_EVENT when unshifting data
    // after the 'end' event has been emitted (readableEnded === true).
    // Note: unshift(data) right after push(null) but before end event is
    // silently ignored by Node.js — we match that by checking _endEmitted.
    if (this._endEmitted) {
      const err = new Error("stream.unshift() after end event") as Error & {
        code: string;
      };
      err.code = "ERR_STREAM_UNSHIFT_AFTER_END_EVENT";
      deferTask(() => this.emit("error", err));
      return;
    }
    // Handle string encoding (Node.js compatibility)
    // Node.js always converts strings to Buffer in binary mode, defaulting to utf8.
    if (typeof chunk === "string" && !this._objectMode) {
      const encoded = stringToEncodedBytes(chunk as string, encoding || "utf8");
      (encoded as any).toString = encodedBytesToString;
      chunk = encoded as any;
    }
    const wasEmpty = this._buf.length === 0;
    this._buf.unshift(chunk);
    if (this._flowing) {
      // In flowing mode, schedule a drain microtask to deliver the unshifted data.
      // This is needed when unshift() is called after push(null) — the original
      // drain loop has already finished, so we must schedule a new one to emit
      // the data before 'end'.
      deferTask(() => {
        while (this._buf.length > 0 && this._flowing) {
          const c = this._buf.shift();
          this._didRead = true;
          this.emit("data", this._applyEncoding(c));
        }
        if (this._ended && this._buf.length === 0) {
          this._emitEndOnce();
        }
      });
    } else if (wasEmpty) {
      // Node.js emits 'readable' when data is unshifted into an empty buffer
      // while in paused mode, so that consumers know data is available.
      deferTask(() => this.emit("readable"));
    }
  }

  /**
   * Safely invoke _read(), catching synchronous exceptions.
   * Node.js catches sync throws in _read and emits an error event, then
   * destroys the stream.  We mirror that behaviour here.
   */
  private _callRead(size: number): void {
    try {
      this._read!(size);
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      // _callRead is always invoked from within a deferTask callback, so
      // we can destroy synchronously instead of scheduling another deferTask.
      // This ensures error/close fire before interleaved Promise.then
      // callbacks, matching Node.js process.nextTick nesting behavior.
      // Set _internalDestroy to signal afterDestroy to emit synchronously.
      this._internalDestroy = true;
      this.destroy(error);
    }
  }

  /**
   * Read data from the stream.
   *
   * Node.js behavior:
   * - read() / read(undefined) — in object mode returns one object;
   *   in binary mode returns all buffered data as one chunk
   * - read(0) — returns null, may trigger internal _read()
   * - read(n) — returns exactly n bytes (binary) or null if not enough;
   *   in object mode returns one object (size ignored)
   */
  read(size?: number): T | null {
    this._didRead = true;

    // Node.js normalization: parseInt then isNaN check.
    // NaN/null/undefined/Infinity/-Infinity/false/"" all become NaN after parseInt,
    // which means "read all available data" (same as read() with no args).
    // Negative integers return null. Zero returns null but may trigger _read.
    let n: number | undefined;
    if (size != null) {
      n = parseInt(size as any, 10);
      if (isNaN(n)) {
        n = undefined; // treat as read() — read all
      } else if (n <= 0) {
        // size === 0: return null but may trigger internal _read
        if (n === 0 && !this._ended && !this._destroyed && this._constructed) {
          const bufSize = this._objectMode ? this._buf.length : this._buf.byteSize;
          if (bufSize < this._highWaterMark || bufSize === 0) {
            this._callRead(this._highWaterMark);
          }
        }
        return null;
      }
    }

    if (this._buf.length === 0) {
      // Buffer is empty — trigger _read() to fill it (matches Node.js behavior
      // where read() on an empty buffer requests more data from the source).
      if (!this._ended && !this._destroyed && this._constructed) {
        this._callRead(this._highWaterMark);
      }
      return null;
    }

    // Object mode: always return a single object, ignore size
    if (this._objectMode) {
      const chunk = this._buf.shift();
      if (this._ended && this._buf.length === 0) {
        deferTask(() => this._emitEndOnce());
      }
      // Node.js triggers _read() to refill buffer after consuming an object
      if (this._hasReadImpl && !this._ended && !this._destroyed && this._constructed) {
        if (this._buf.length < this._highWaterMark) {
          deferTask(() => {
            if (!this._ended && !this._destroyed) {
              this._callRead(this._highWaterMark);
            }
          });
        }
      }
      return chunk;
    }

    // Binary mode
    let result: T;

    if (n == null) {
      // read() with no size: return all buffered data as one chunk.
      // Note: Node.js 26+ changed to return a single chunk, but our browser
      // implementation preserves the pre-26 behavior (consumeAll) for backward
      // compatibility and because it matches what most consumers expect.
      result = this._applyEncoding(this._buf.consumeAll());
    } else {
      // read(n): return exactly n bytes, or null if not enough
      if (this._buf.byteSize < n) {
        // Not enough data buffered
        if (this._ended) {
          // Stream ended — return whatever is available
          result = this._applyEncoding(this._buf.consumeAll());
        } else {
          // Trigger internal read to fill buffer
          // Node.js passes Math.max(n, highWaterMark) to _read so that
          // custom _read implementations know the consumer's requested size.
          if (this._hasReadImpl && this._constructed) {
            this._callRead(Math.max(n, this._highWaterMark));
          }
          return null;
        }
      } else {
        result = this._applyEncoding(this._buf.consumeBytes(n));
      }
    }

    // Trigger _read to refill buffer if below HWM
    if (this._hasReadImpl && !this._ended && !this._destroyed) {
      const bufSize = this._buf.byteSize;
      if (bufSize < this._highWaterMark) {
        deferTask(() => {
          if (!this._ended && !this._destroyed) {
            this._callRead(this._highWaterMark);
          }
        });
      }
    }

    if (this._ended && this._buf.length === 0) {
      deferTask(() => this._emitEndOnce());
    }

    // Node.js re-emits 'readable' when there is still data in the buffer
    // after a read(), so consumers know more data is available.
    if (!this._flowing && this._buf.length > 0) {
      deferTask(() => this.emit("readable"));
    }

    return result;
  }

  /**
   * Flush the stream decoder's internal remainder and clear it.
   * Returns the flushed string, or null if nothing to flush.
   * Called when the stream ends, both from the event path (_emitEndOnce)
   * and the async iterator path (_initAsyncIterState).
   */
  private _flushDecoder(): string | null {
    if (!this._decoder) {
      return null;
    }
    const final = this._decoder.decode(new Uint8Array(0), { stream: false });
    this._decoder = null;
    return final || null;
  }

  /**
   * Set encoding for string output.
   * Creates a private TextDecoder instance for streaming decode to avoid
   * sharing mutable state across concurrent streams.
   */
  setEncoding(encoding: string): this {
    this._encoding = encoding;
    // Always create a fresh decoder for this stream — the { stream: true }
    // option in _applyEncoding makes the decoder stateful (it retains
    // incomplete multi-byte sequences between calls).  Sharing a cached
    // singleton across streams would corrupt multi-byte boundaries.
    this._decoder = createStreamDecoder(encoding);
    return this;
  }

  private _applyEncoding(chunk: T): T {
    if (this._encoding && chunk instanceof Uint8Array) {
      if (!this._decoder) {
        this._decoder = createStreamDecoder(this._encoding);
      }
      // Pass {stream: true} to handle multi-byte characters that may span
      // chunk boundaries (matches Node.js StringDecoder behavior).
      return this._decoder.decode(chunk, { stream: true }) as any;
    }
    return chunk;
  }

  /**
   * Wrap an old-style stream
   */
  wrap(stream: any): this {
    let paused = false;
    stream.on("data", (chunk: T) => {
      if (!this.push(chunk)) {
        stream.pause();
        paused = true;
      }
    });
    stream.on("end", () => this.push(null));
    stream.on("error", (err: Error) => this.destroy(err));
    stream.on("close", () => this.destroy());

    // When the Readable's buffer drains, _read is called — resume the source.
    // This matches Node.js wrap() behavior where _read triggers source.resume().
    this._hasReadImpl = true;
    this._read = () => {
      if (paused) {
        paused = false;
        stream.resume();
      }
    };

    return this;
  }

  /**
   * Pause the stream
   */
  pause(): this {
    if (this._flowing) {
      this._flowing = false;
      this.emit("pause");
    } else {
      // Node.js: pause() on a fresh stream transitions readableFlowing
      // from null to false and sets isPaused() to true.
      this._hasFlowed = true;
    }
    return this;
  }

  /**
   * Resume the stream
   */
  resume(): this {
    if (!this._flowing) {
      this._flowing = true;
      this._hasFlowed = true;

      // Emit asynchronously to match Node.js process.nextTick timing
      // Use _resumeScheduled guard to prevent duplicate emissions when
      // resume() is called multiple times synchronously (e.g. on("data") + resume())
      if (!this._resumeScheduled) {
        this._resumeScheduled = true;
        deferTask(() => {
          this._resumeScheduled = false;
          this.emit("resume");
        });
      }
    }

    // Drain buffered data asynchronously (matches Node.js process.nextTick behavior).
    // Node.js does NOT drain synchronously on resume() — it defers to nextTick
    // so that multiple resume()/pipe() calls can register before data flows.
    if (this._buf.length > 0) {
      deferTask(() => {
        while (this._buf.length > 0 && this._flowing) {
          const chunk = this._buf.shift();
          this._didRead = true;
          this.emit("data", this._applyEncoding(chunk));
        }
        // After draining, check for end
        if (this._ended && this._buf.length === 0) {
          this._emitEndOnce();
        }

        // If we were paused due to backpressure while reading from a Web Stream,
        // resume reading once the buffered chunks have been drained.
        if (
          this._flowing &&
          this._buf.length === 0 &&
          this._webStreamMode &&
          !this._pushMode &&
          !this._ended &&
          !this._destroyed
        ) {
          this._startReading();
        }

        // After draining buffered data, call _read() so the source can refill.
        // This is needed for wrap() where _read resumes the wrapped stream.
        if (
          this._hasReadImpl &&
          this._flowing &&
          !this._ended &&
          !this._destroyed &&
          this._constructed
        ) {
          const bufSize = this._objectMode ? this._buf.length : this._buf.byteSize;
          if (bufSize < this._highWaterMark || bufSize === 0) {
            this._callRead(this._highWaterMark);
          }
        }
      });
    } else if (this._ended && this._buf.length === 0) {
      deferTask(() => this._emitEndOnce());
    } else if (this._webStreamMode && !this._pushMode) {
      // Start reading from underlying Web Stream when resuming.
      this._startReading();
    } else if (this._hasReadImpl) {
      // Call user-provided read function asynchronously
      // This allows multiple pipe() calls to register before data flows
      deferTask(() => {
        if (this._flowing && !this._ended && !this._destroyed && this._constructed) {
          this._callRead(this._highWaterMark);
          // After _callRead, check if push(null) was called synchronously.
          // If so, emit 'end' here rather than letting push(null)'s own
          // deferTask schedule it — otherwise a user Promise.then
          // queued before this microtask could observe readableEnded=false,
          // diverging from Node.js where process.nextTick drains greedily.
          if (this._ended && this._buf.length === 0) {
            this._emitEndOnce();
          }
        }
      });
    }

    return this;
  }

  /**
   * Override on() to automatically resume when 'data' listener is added
   * This matches Node.js behavior where adding a 'data' listener puts
   * the stream into flowing mode.
   */
  override on(event: string | symbol, listener: (...args: any[]) => void): this {
    super.on(event, listener);

    // When a 'data' listener is added, switch to flowing mode
    if (event === "data") {
      this.resume();
    } else if (event === "readable") {
      // Node.js: adding a 'readable' listener always sets readableFlowing to
      // false, even if the stream is currently flowing. This pauses the stream
      // so consumers can use the pull-based readable interface.
      this._hasFlowed = true;
      this._flowing = false;
      if (this._buf.length > 0 || this._ended) {
        deferTask(() => this.emit("readable"));
      }
    }

    return this;
  }

  /**
   * Check if paused
   */
  isPaused(): boolean {
    // Node.js: isPaused() returns true only when readableFlowing is strictly false.
    // When readableFlowing is null (initial state), isPaused() returns false.
    return this.readableFlowing === false;
  }

  /**
   * Pipe to a writable stream, transform stream, or duplex stream.
   * Accepts any writable-like object (duck-typed, matching Node.js behavior).
   */
  pipe<W extends WritableLike>(destination: W, options?: { end?: boolean }): W {
    return this._pipes.pipe(destination, options) as W;
  }

  /**
   * Unpipe from destination
   */
  unpipe(destination?: WritableLike): this {
    this._pipes.unpipe(destination);
    return this;
  }

  /**
   * Destroy the stream
   */
  destroy(error?: Error): this {
    if (this._destroyed) {
      return this;
    }

    this._destroyed = true;
    // Node.js: destroy() resets readable override so getter reflects true state
    this._readableOverride = undefined;

    // Note: Node.js does NOT clear the readable buffer synchronously on destroy.
    // readableLength retains its value even after destroy() returns.

    // Ensure we detach from destinations to avoid leaking listeners.
    this.unpipe();

    if (this._reader) {
      const reader = this._reader;
      this._reader = null;
      reader
        .cancel()
        .catch(() => {})
        .finally(() => {
          try {
            reader.releaseLock();
          } catch {
            // Ignore if a read is still pending
          }
        });
    }

    // If subclass overrides _destroy, call it and wait for callback before
    // emitting error/close (matches Node.js behavior).
    const afterDestroy = (finalError?: Error | null): void => {
      // Node.js: _destroy's callback determines whether an error is emitted.
      // - cb(null) or cb() or cb(undefined): suppress the original error
      // - cb(new Error(...)): replace with the new error, emit it
      // Node.js checks `if (err)` on the callback argument, so null/undefined/no-arg
      // all suppress the error. Only a truthy error value triggers the error event.
      const err = finalError || null;
      if (err) {
        this._errored = err;
      }
      this._closed = true;
      const doEmit = (): void => {
        if (err) {
          this.emit("error", err);
        }
        if (this._emitClose) {
          this.emit("close");
        }
      };
      // When already inside a deferTask callback (e.g. autoDestroy after
      // 'end') and there is no async _destroy hook, emit synchronously.
      // This matches Node.js where nested process.nextTick runs before
      // interleaved Promise.then callbacks.
      // Only apply when the stream ended naturally (autoDestroy path) or
      // was destroyed internally (e.g. _callRead error) — user-initiated
      // destroy() from within a listener must still defer so that async
      // cleanup (e.g. iterator.return()) can complete.
      const isInternalPath = this._endEmitted || this._internalDestroy;
      this._internalDestroy = false;
      if (inDeferredContext() && !this._hasDestroyHook() && isInternalPath) {
        doEmit();
      } else {
        deferTask(doEmit);
      }
    };

    if (this._hasDestroyHook()) {
      try {
        this._destroy(error ?? null, afterDestroy);
      } catch (err) {
        afterDestroy(err instanceof Error ? err : new Error(String(err)));
      }
    } else {
      afterDestroy(error);
    }
    return this;
  }

  /**
   * Override in subclass to customise destroy behaviour.
   * Call `callback(err)` when cleanup is complete.
   */
  _destroy(error: Error | null, callback: (error?: Error | null) => void): void {
    callback(error);
  }

  /**
   * Reverse the effects of destroy() so the stream can potentially be reused.
   * Matches Node.js _undestroy() which resets destroyed and closed flags.
   */
  _undestroy(): void {
    this._destroyed = false;
    this._closed = false;
    this._errored = null;
  }

  /** Check if _destroy has been overridden by a subclass or constructor option. */
  private _hasDestroyHook(): boolean {
    return (
      Object.prototype.hasOwnProperty.call(this, "_destroy") ||
      Object.getPrototypeOf(this)._destroy !== Readable.prototype._destroy
    );
  }

  /**
   * Get a Web ReadableStream view of this stream (internal).
   *
   * For external Web Streams (_webStreamMode), returns the original stream.
   * For controllable streams, creates a pull-based ReadableStream that reads
   * from the Node-style Readable on demand — matching Node.js `Readable.toWeb()`
   * which does NOT force the stream into flowing mode.
   * @internal
   */
  private get _webStream(): ReadableStream<T> {
    if (this._stream) {
      return this._stream;
    }

    // Build a pull-based Web ReadableStream that only reads from the
    // Node-style Readable when the consumer requests data (via pull()).
    // This avoids forcing the stream into flowing mode immediately.
    // oxlint-disable-next-line no-this-alias -- needed for ReadableStream callback closures
    const readable = this;
    let ended = false;

    const ws = new ReadableStream<T>({
      pull(controller) {
        return new Promise<void>((resolve, reject) => {
          if (ended) {
            return resolve();
          }

          const tryRead = (): void => {
            // If the readable was destroyed or errored while we were waiting,
            // check before attempting read.
            if (readable._destroyed) {
              if (readable._errored) {
                ended = true;
                try {
                  controller.error(readable._errored);
                } catch {
                  /* controller may already be errored/closed */
                }
                return reject(readable._errored);
              }
              ended = true;
              try {
                controller.close();
              } catch {
                /* controller may already be closed */
              }
              return resolve();
            }

            const chunk = readable.read() as T | null;
            if (chunk !== null) {
              try {
                controller.enqueue(chunk);
              } catch {
                /* controller may be closed */
              }
              return resolve();
            }

            // No data available yet — check if the stream has ended.
            if (readable._endEmitted) {
              ended = true;
              try {
                controller.close();
              } catch {
                /* controller may already be closed */
              }
              return resolve();
            }

            // Wait for more data or end/error.
            const onReadable = (): void => {
              cleanup();
              tryRead();
            };
            const onEnd = (): void => {
              cleanup();
              ended = true;
              try {
                controller.close();
              } catch {
                /* controller may already be closed */
              }
              resolve();
            };
            const onError = (err: Error): void => {
              cleanup();
              ended = true;
              try {
                controller.error(err);
              } catch {
                /* controller may already be errored/closed */
              }
              reject(err);
            };
            const onClose = (): void => {
              cleanup();
              if (!ended) {
                ended = true;
                try {
                  controller.close();
                } catch {
                  /* controller may already be closed */
                }
              }
              resolve();
            };

            const cleanup = (): void => {
              readable.off("readable", onReadable);
              readable.off("end", onEnd);
              readable.off("error", onError);
              readable.off("close", onClose);
            };

            readable.once("readable", onReadable);
            readable.once("end", onEnd);
            readable.once("error", onError);
            readable.once("close", onClose);
          };

          tryRead();
        });
      },
      cancel() {
        readable.destroy();
      }
    });
    this._stream = ws;
    return ws;
  }

  private _readableOverride: boolean | undefined;

  get readable(): boolean {
    if (this._readableOverride !== undefined) {
      return this._readableOverride;
    }
    // Node.js: readable stays true while buffer has data, even after push(null).
    // It only becomes false after the 'end' event has been emitted (or on destroy).
    return !this._destroyed && !this._endEmitted;
  }

  set readable(val: boolean) {
    this._readableOverride = val;
  }

  get readableEnded(): boolean {
    // Node.js: readableEnded only becomes true after the 'end' event has been emitted,
    // not when push(null) is called.
    return this._endEmitted;
  }

  get readableLength(): number {
    return this._objectMode ? this._buf.length : this._buf.byteSize;
  }

  /** Whether the stream has been destroyed */
  get destroyed(): boolean {
    return this._destroyed;
  }

  set destroyed(val: boolean) {
    this._destroyed = val;
  }

  /** The error that destroyed the stream, or null */
  get errored(): Error | null {
    return this._errored;
  }

  /** Whether the stream has been closed */
  get closed(): boolean {
    return this._closed;
  }

  /** Whether the stream is in flowing mode */
  get readableFlowing(): boolean | null {
    if (this._flowing) {
      return true;
    }
    // Distinguish between "never flowed" (null) and "paused after flowing" (false)
    return this._hasFlowed ? false : null;
  }

  set readableFlowing(value: boolean | null) {
    if (value === true) {
      this._flowing = true;
      this._hasFlowed = true;
    } else if (value === false) {
      this._flowing = false;
      this._hasFlowed = true;
    } else {
      // null: reset to initial state
      this._flowing = false;
      this._hasFlowed = false;
    }
  }

  /**
   * Internal method to set readableFlowing state.
   * Used by Transform/Duplex to control the internal Readable's flow state.
   * @internal
   */
  _setReadableFlowing(value: boolean | null): void {
    this.readableFlowing = value;
  }

  /** Whether the stream was aborted (destroyed or errored before 'end' was emitted) */
  get readableAborted(): boolean {
    // Node.js: readable !== false && (destroyed || errored) && !endEmitted
    // The 'readable !== false' check is for when user explicitly sets stream.readable = false.
    // We check _readableOverride (not the computed getter which returns false after destroy).
    if (this._readableOverride === false) {
      return false;
    }
    return (this._destroyed || !!this._errored) && !this._endEmitted;
  }

  /** Whether read() has ever been called */
  get readableDidRead(): boolean {
    return this._didRead;
  }

  /** Current encoding or null */
  get readableEncoding(): string | null {
    return this._encoding;
  }

  /** Returns array of objects containing info about buffered data */
  get readableObjectMode(): boolean {
    return this._objectMode;
  }

  get readableHighWaterMark(): number {
    return this._highWaterMark;
  }

  /**
   * Get the internal buffer contents as an array (matches Node.js BufferList behavior)
   */
  get readableBuffer(): T[] {
    return this._buf.toArray();
  }

  /**
   * Release the internal reader lock, clearing _reader.
   * Safe to call even when no reader is held.
   */
  private _releaseReader(): void {
    if (this._reader) {
      const reader = this._reader;
      this._reader = null;
      try {
        reader.releaseLock();
      } catch {
        // Ignore if a read is still pending
      }
    }
  }

  private async _startReading(): Promise<void> {
    if (this._reading || this._destroyed || !this._flowing) {
      return;
    }

    this._reading = true;

    try {
      if (!this._reader) {
        this._reader = this._stream!.getReader();
      }

      while (this._flowing && !this._destroyed && !this._pushMode) {
        const { done, value } = await this._reader.read();

        // Check _pushMode again after async read - if push() was called, stop reading
        if (this._pushMode) {
          this._releaseReader();
          break;
        }

        if (done) {
          this._ended = true;
          if (this._buf.length === 0) {
            // Defer end to match Node.js nextTick-ish timing.
            deferTask(() => this._emitEndOnce());
          }
          this._releaseReader();
          break;
        }

        if (value !== undefined) {
          // Always buffer Web Stream chunks (Node.js Readable buffers internally).
          // If flowing, immediately drain the buffer into 'data' events.
          const wasEmpty = this._buf.length === 0;
          this._buf.push(value);
          if (wasEmpty) {
            deferTask(() => this.emit("readable"));
          }

          while (this._buf.length > 0 && this._flowing) {
            const chunk = this._buf.shift();
            this._didRead = true;
            this.emit("data", this._applyEncoding(chunk));
          }

          // Backpressure: if paused, stop reading from the underlying Web Stream
          // until resume() drains buffered chunks.
          if (!this._flowing) {
            break;
          }
        }
      }
    } catch (err) {
      this.emit("error", err);
      this._releaseReader();
    } finally {
      this._reading = false;
    }
  }

  /**
   * Async iterator support.
   *
   * Node.js parity: multiple calls to `[Symbol.asyncIterator]()` return
   * independent iterators that *compete* for chunks from the same underlying
   * stream.  Each chunk is delivered to exactly one waiting iterator (FIFO).
   * A single shared `data` listener is attached on the first call; subsequent
   * calls reuse the same shared state.
   */
  async *[Symbol.asyncIterator](): AsyncIterableIterator<T> {
    // Lazily initialise the shared async-iterator state on first call.
    if (!this._asyncIterState) {
      this._asyncIterState = this._initAsyncIterState();
    }
    const state = this._asyncIterState;
    state.activeCount++;

    try {
      while (true) {
        if (state.streamError) {
          throw state.streamError;
        }

        // Try to consume a chunk from the shared queue (competitive).
        if (state.dataQueueIndex < state.dataQueue.length) {
          const chunk = state.dataQueue[state.dataQueueIndex]!;
          state.dataQueue[state.dataQueueIndex] = undefined; // release ref
          state.dataQueueIndex++;
          state.queuedSize -= this._chunkSizeForBackpressure(chunk);

          // Compact the queue periodically.
          if (state.dataQueueIndex >= 1024 && state.dataQueueIndex * 2 >= state.dataQueue.length) {
            state.dataQueue.splice(0, state.dataQueueIndex);
            state.dataQueueIndex = 0;
          }

          this._maybeResumeIterState(state);
          yield chunk as T;
          continue;
        }

        if (state.done) {
          break;
        }

        // No data available — park this iterator until a chunk arrives.
        const chunk = await new Promise<any>((resolve, reject) => {
          state.resolverQueue.push({ resolve, reject });
        });

        if (chunk === null) {
          // Stream ended / closed while we were waiting.
          break;
        }

        // Chunk was delivered directly by dataHandler (bypassed the queue),
        // so queuedSize was never incremented — no need to decrement here.
        yield chunk as T;
      }
    } finally {
      state.activeCount--;
      // Node.js destroys the stream when ANY async iterator exits early
      // (break/return/throw), even if other iterators are still active.
      // This causes the remaining iterators' pending next() calls to resolve
      // with done:true (via the 'close' handler).
      if (state.activeCount === 0) {
        this._teardownAsyncIterState();
      }
      if (!this._destroyed) {
        this.destroy();
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Async iterator helpers (private)
  // ---------------------------------------------------------------------------

  private _chunkSizeForBackpressure(chunk: any): number {
    if (this._objectMode) {
      return 1;
    }
    if (chunk instanceof Uint8Array) {
      return chunk.byteLength;
    }
    if (typeof chunk === "string") {
      return chunk.length;
    }
    return 1;
  }

  private _maybeResumeIterState(state: _AsyncIterState<T>): void {
    if (
      state.pausedByIterator &&
      state.queuedSize < this._highWaterMark &&
      !state.done &&
      !this._destroyed
    ) {
      state.pausedByIterator = false;
      this.resume();
    }
  }

  /**
   * Create and attach the shared async-iterator state.
   * A single `data` listener routes chunks to the first waiting resolver
   * (FIFO) or enqueues them.  This matches Node.js semantics where multiple
   * iterators compete for the same data stream.
   */
  private _initAsyncIterState(): _AsyncIterState<T> {
    const hwm = this._highWaterMark;

    const state: _AsyncIterState<T> = {
      dataQueue: [],
      dataQueueIndex: 0,
      queuedSize: 0,
      resolverQueue: [],
      done: false,
      streamError: null,
      activeCount: 0,
      listenersAttached: false,
      pausedByIterator: false,
      // Handlers are assigned below; typed as `any` first, then replaced.
      dataHandler: null as any,
      doneHandler: null as any,
      errorHandler: null as any
    };

    // -- data handler ---------------------------------------------------------
    state.dataHandler = (chunk: any): void => {
      // Deliver directly to the first waiting resolver (FIFO competitive).
      // Chunks delivered directly do not contribute to queuedSize because
      // they never sit in the buffer — this avoids a transient inflation
      // that could trigger an unnecessary pause().
      if (state.resolverQueue.length > 0) {
        const { resolve } = state.resolverQueue.shift()!;
        resolve(chunk);
        return;
      }

      state.dataQueue.push(chunk);
      state.queuedSize += this._chunkSizeForBackpressure(chunk);
      if (!state.pausedByIterator && state.queuedSize >= hwm) {
        state.pausedByIterator = true;
        this.pause();
      }
    };

    // -- end / close handler --------------------------------------------------
    state.doneHandler = (): void => {
      state.done = true;
      // Resolve all waiting iterators with null (signals completion).
      for (const { resolve } of state.resolverQueue) {
        resolve(null);
      }
      state.resolverQueue.length = 0;
    };

    // -- error handler --------------------------------------------------------
    state.errorHandler = (err: Error): void => {
      state.done = true;
      state.streamError = err;
      // Reject all waiting iterators.
      for (const { reject } of state.resolverQueue) {
        reject(err);
      }
      state.resolverQueue.length = 0;
    };

    // Drain any already-buffered data into the shared queue before attaching
    // the `data` listener so that existing chunks are not lost.
    while (this._buf.length > 0) {
      const chunk = this._applyEncoding(this._buf.shift());
      state.dataQueue.push(chunk);
      state.queuedSize += this._chunkSizeForBackpressure(chunk);
    }

    if (this._ended) {
      // Flush the stream decoder's remainder (e.g. partial base64 group)
      // before marking the iterator as done. This mirrors the flush that
      // _emitEndOnce performs for the event-based consumption path.
      const flushed = this._flushDecoder();
      if (flushed) {
        state.dataQueue.push(flushed as any);
        state.queuedSize += this._chunkSizeForBackpressure(flushed);
      }
      state.done = true;
    } else {
      this.on("data", state.dataHandler);
      this.on("end", state.doneHandler);
      this.on("error", state.errorHandler);
      this.on("close", state.doneHandler);
      state.listenersAttached = true;

      // Drive the stream into flowing mode.
      this.resume();
    }

    return state;
  }

  /** Remove shared listeners and discard state. */
  private _teardownAsyncIterState(): void {
    const state = this._asyncIterState;
    if (!state) {
      return;
    }
    if (state.listenersAttached) {
      this.off("data", state.dataHandler);
      this.off("end", state.doneHandler);
      this.off("error", state.errorHandler);
      this.off("close", state.doneHandler);
    }
    this._asyncIterState = null;
  }

  /**
   * Synchronous dispose — destroys the stream.
   * Matches Node.js Symbol.dispose support (v20+, experimental).
   */
  [Symbol.dispose](): void {
    if (!this._destroyed) {
      this.destroy();
    }
  }

  /**
   * Async dispose hook for `await using`.
   * Destroys the stream and resolves after the 'close' event.
   */
  async [Symbol.asyncDispose](): Promise<void> {
    const selfInitiated = !this._destroyed;
    if (selfInitiated) {
      this.destroy();
    }
    return new Promise<void>((resolve, reject) => {
      const settle = (): void => {
        if (selfInitiated || this._endEmitted) {
          resolve();
        } else {
          reject(new Error("Premature close"));
        }
      };
      if (this._closed) {
        settle();
      } else {
        this.once("close", settle);
      }
    });
  }

  /**
   * Explicit iterator method (same as Symbol.asyncIterator)
   */
  iterator(options?: { destroyOnReturn?: boolean }): AsyncIterableIterator<T> {
    const destroyOnReturn = options?.destroyOnReturn ?? true;

    if (destroyOnReturn) {
      return this[Symbol.asyncIterator]();
    }

    // Build a standalone async iterator that does NOT destroy the stream
    // when the consumer breaks/returns early.
    //
    // Key difference from the default [Symbol.asyncIterator]():
    //   - We do NOT call resume() / flowing mode. Instead we use
    //     'readable' events and explicit read() calls (pull-mode).
    //     This ensures the stream keeps buffered data when the consumer
    //     breaks early, so 'end' (and therefore autoDestroy) never fires
    //     prematurely — matching Node.js behavior.
    //   - On return() we just clean up listeners without destroying.
    // oxlint-disable-next-line no-this-alias -- needed for closures in iterator object
    const stream = this;
    let resolveNext: ((value: IteratorResult<T>) => void) | null = null;
    let rejectNext: ((err: Error) => void) | null = null;
    let done = false;
    let streamError: Error | null = null;

    const tryRead = (): void => {
      if (!resolveNext) {
        return;
      }
      const chunk = stream.read() as T | null;
      if (chunk !== null) {
        const r = resolveNext;
        resolveNext = null;
        rejectNext = null;
        r({ done: false, value: chunk });
      }
    };

    const readableHandler = (): void => {
      tryRead();
    };

    const endHandler = (): void => {
      done = true;
      if (resolveNext) {
        const r = resolveNext;
        resolveNext = null;
        rejectNext = null;
        r({ done: true, value: undefined as any });
      }
    };

    const errorHandler = (err: Error): void => {
      done = true;
      streamError = err;
      if (rejectNext) {
        const r = rejectNext;
        resolveNext = null;
        rejectNext = null;
        r(err);
      }
    };

    stream.on("readable", readableHandler);
    stream.on("end", endHandler);
    stream.on("error", errorHandler);

    const cleanup = (): void => {
      stream.off("readable", readableHandler);
      stream.off("end", endHandler);
      stream.off("error", errorHandler);
    };

    const iter: AsyncIterableIterator<T> = {
      next: async () => {
        if (streamError) {
          throw toError(streamError);
        }

        // Try synchronous read first
        const chunk = stream.read() as T | null;
        if (chunk !== null) {
          return { done: false, value: chunk };
        }
        if (done) {
          cleanup();
          return { done: true, value: undefined as any };
        }

        // Wait for 'readable' or 'end'
        return new Promise<IteratorResult<T>>((resolve, reject) => {
          resolveNext = resolve;
          rejectNext = reject;
          // One more try in case data arrived between our read() and listener setup
          tryRead();
        });
      },
      return: async (value?: any) => {
        // Do NOT destroy the stream — just clean up listeners
        cleanup();
        return { done: true as const, value };
      },
      throw: async (err?: any) => {
        // On throw, we DO destroy (matching Node.js behavior)
        cleanup();
        stream.destroy(err);
        return { done: true as const, value: undefined as any };
      },
      [Symbol.asyncIterator]() {
        return this;
      }
    };

    return iter;
  }

  // =============================================================================
  // Functional / Higher-order Methods (Node.js Readable compatibility)
  // =============================================================================

  /**
   * Map each chunk through a function, returning a new Readable.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  map<U>(
    fn: (data: T, options: { signal: AbortSignal }) => U | Promise<U>,
    options?: {
      concurrency?: number;
      highWaterMark?: number;
      signal?: AbortSignal;
    }
  ): Readable<U> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    const result = new Readable<U>({ objectMode: true });

    pumpAsyncIterableToReadable(
      result,
      (async function* (src: Readable<T>) {
        try {
          yield* _mapWithConcurrency<T, U>(
            src,
            chunk => fn(chunk, { signal: innerSignal }),
            concurrency,
            signal
          );
        } finally {
          ac.abort();
        }
      })(this)
    );

    result.once("close", () => {
      if (!this.destroyed) {
        this.destroy();
      }
    });

    if (signal) {
      const onAbort = () => {
        result.destroy(createAbortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
    }

    return result;
  }

  /**
   * Filter chunks by a predicate, returning a new Readable.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  filter(
    fn: (data: T, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: {
      concurrency?: number;
      highWaterMark?: number;
      signal?: AbortSignal;
    }
  ): Readable<T> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    const result = new Readable<T>({ objectMode: true });
    const SKIP = Symbol("skip");

    pumpAsyncIterableToReadable(
      result,
      (async function* (src: Readable<T>) {
        try {
          if (concurrency <= 1) {
            for await (const chunk of src) {
              _throwIfAborted(signal);
              const keep = await fn(chunk, { signal: innerSignal });
              if (keep) {
                yield chunk;
              }
            }
          } else {
            for await (const item of _mapWithConcurrency<T, T | symbol>(
              src,
              async chunk => {
                const keep = await fn(chunk, { signal: innerSignal });
                return keep ? chunk : SKIP;
              },
              concurrency,
              signal
            )) {
              if (item !== SKIP) {
                yield item as T;
              }
            }
          }
        } finally {
          ac.abort();
        }
      })(this)
    );

    result.once("close", () => {
      if (!this.destroyed) {
        this.destroy();
      }
    });

    if (signal) {
      const onAbort = () => {
        result.destroy(createAbortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
    }

    return result;
  }

  /**
   * Iterate all chunks. Returns a promise that resolves when the stream ends.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  async forEach(
    fn: (data: T, options: { signal: AbortSignal }) => void | Promise<void>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<undefined> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    try {
      if (concurrency <= 1) {
        for await (const chunk of this) {
          _throwIfAborted(signal);
          await fn(chunk, { signal: innerSignal });
        }
      } else {
        // Consume all mapped results (side effects only) to drive concurrency
        for await (const _item of _mapWithConcurrency<T, void>(
          this,
          async chunk => {
            await fn(chunk, { signal: innerSignal });
          },
          concurrency,
          signal
        )) {
          // discard results; forEach is side-effect only
        }
      }
    } finally {
      ac.abort();
    }
    return undefined;
  }

  /**
   * Collect all chunks into an array.
   */
  async toArray(options?: { signal?: AbortSignal }): Promise<T[]> {
    const signal = options?.signal;
    _validateAbortSignal(signal);

    const result: T[] = [];
    for await (const chunk of this) {
      _throwIfAborted(signal);
      result.push(chunk);
    }
    return result;
  }

  /**
   * Returns true if any chunk passes the predicate. Short-circuits on first match.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  async some(
    fn: (data: T, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    try {
      if (concurrency <= 1) {
        for await (const chunk of this) {
          _throwIfAborted(signal);
          const result = await fn(chunk, { signal: innerSignal });
          if (result) {
            this.destroy();
            return true;
          }
        }
      } else {
        for await (const result of _mapWithConcurrency<T, boolean>(
          this,
          chunk => fn(chunk, { signal: innerSignal }),
          concurrency,
          signal
        )) {
          if (result) {
            this.destroy();
            return true;
          }
        }
      }
      return false;
    } finally {
      ac.abort();
    }
  }

  /**
   * Find the first chunk matching the predicate. Short-circuits on match.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  async find(
    fn: (data: T, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<T | undefined> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    try {
      if (concurrency <= 1) {
        for await (const chunk of this) {
          _throwIfAborted(signal);
          const result = await fn(chunk, { signal: innerSignal });
          if (result) {
            this.destroy();
            return chunk;
          }
        }
      } else {
        for await (const item of _mapWithConcurrency<T, { chunk: T; match: boolean }>(
          this,
          async chunk => ({
            chunk,
            match: await fn(chunk, { signal: innerSignal })
          }),
          concurrency,
          signal
        )) {
          if (item.match) {
            this.destroy();
            return item.chunk;
          }
        }
      }
      return undefined;
    } finally {
      ac.abort();
    }
  }

  /**
   * Returns true if all chunks pass the predicate. Short-circuits on first failure.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  async every(
    fn: (data: T, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    try {
      if (concurrency <= 1) {
        for await (const chunk of this) {
          _throwIfAborted(signal);
          const result = await fn(chunk, { signal: innerSignal });
          if (!result) {
            this.destroy();
            return false;
          }
        }
      } else {
        for await (const result of _mapWithConcurrency<T, boolean>(
          this,
          chunk => fn(chunk, { signal: innerSignal }),
          concurrency,
          signal
        )) {
          if (!result) {
            this.destroy();
            return false;
          }
        }
      }
      return true;
    } finally {
      ac.abort();
    }
  }

  /**
   * Map each chunk to multiple outputs (flattening), returning a new Readable.
   * Supports concurrent execution via `options.concurrency` (matching Node.js).
   */
  flatMap<U>(
    fn: (
      data: T,
      options: { signal: AbortSignal }
    ) => Iterable<U> | AsyncIterable<U> | Readable<U> | Promise<Iterable<U> | AsyncIterable<U>>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Readable<U> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const concurrency = options?.concurrency ?? 1;
    const ac = new AbortController();
    const innerSignal = ac.signal;

    const result = new Readable<U>({ objectMode: true });

    pumpAsyncIterableToReadable(
      result,
      (async function* (src: Readable<T>) {
        try {
          if (concurrency <= 1) {
            for await (const chunk of src) {
              _throwIfAborted(signal);
              const mapped = await fn(chunk, { signal: innerSignal });

              if (mapped && typeof (mapped as any)[Symbol.asyncIterator] === "function") {
                for await (const item of mapped as AsyncIterable<U>) {
                  yield item;
                }
              } else if (mapped && typeof (mapped as any)[Symbol.iterator] === "function") {
                for (const item of mapped as Iterable<U>) {
                  yield item;
                }
              }
            }
          } else {
            // With concurrency, map chunks to their flattened arrays in order,
            // then yield each item from each array in sequence.
            for await (const items of _mapWithConcurrency<T, U[]>(
              src,
              async chunk => {
                const mapped = await fn(chunk, { signal: innerSignal });
                const collected: U[] = [];
                if (mapped && typeof (mapped as any)[Symbol.asyncIterator] === "function") {
                  for await (const item of mapped as AsyncIterable<U>) {
                    collected.push(item);
                  }
                } else if (mapped && typeof (mapped as any)[Symbol.iterator] === "function") {
                  for (const item of mapped as Iterable<U>) {
                    collected.push(item);
                  }
                }
                return collected;
              },
              concurrency,
              signal
            )) {
              for (const item of items) {
                yield item;
              }
            }
          }
        } finally {
          ac.abort();
        }
      })(this)
    );

    result.once("close", () => {
      if (!this.destroyed) {
        this.destroy();
      }
    });

    if (signal) {
      const onAbort = () => {
        result.destroy(createAbortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
    }

    return result;
  }

  /**
   * Skip the first `limit` chunks, returning a new Readable.
   */
  drop(limit: number, options?: { signal?: AbortSignal }): Readable<T> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    _validateNonNegativeInteger(limit, "limit");

    const result = new Readable<T>({ objectMode: true });

    pumpAsyncIterableToReadable(
      result,
      (async function* (src: Readable<T>) {
        let count = 0;
        for await (const chunk of src) {
          _throwIfAborted(signal);
          if (count >= limit) {
            yield chunk;
          }
          count++;
        }
      })(this)
    );

    // Propagate destruction from result back to source (matches Node.js)
    result.once("close", () => {
      if (!this.destroyed) {
        this.destroy();
      }
    });

    if (signal) {
      const onAbort = () => {
        result.destroy(createAbortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
    }

    return result;
  }

  /**
   * Take only the first `limit` chunks, returning a new Readable.
   */
  take(limit: number, options?: { signal?: AbortSignal }): Readable<T> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    _validateNonNegativeInteger(limit, "limit");

    const result = new Readable<T>({ objectMode: true });

    pumpAsyncIterableToReadable(
      result,
      (async function* (src: Readable<T>) {
        let count = 0;
        for await (const chunk of src) {
          _throwIfAborted(signal);
          if (count >= limit) {
            break;
          }
          yield chunk;
          count++;
        }
      })(this)
    );

    // Propagate destruction from result back to source (matches Node.js)
    result.once("close", () => {
      if (!this.destroyed) {
        this.destroy();
      }
    });

    if (signal) {
      const onAbort = () => {
        result.destroy(createAbortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
    }

    return result;
  }

  /**
   * Reduce the stream to a single value.
   */
  async reduce<U = T>(
    fn: (previous: U, data: T, options: { signal: AbortSignal }) => U | Promise<U>,
    initial?: U,
    options?: { signal?: AbortSignal }
  ): Promise<U> {
    const signal = options?.signal;
    _validateAbortSignal(signal);
    const ac = new AbortController();
    const innerSignal = ac.signal;

    let accumulator: U;
    const hasInitial = arguments.length >= 2;
    let first = true;

    try {
      for await (const chunk of this) {
        _throwIfAborted(signal);
        if (first && !hasInitial) {
          accumulator = chunk as any as U;
          first = false;
          continue;
        }
        if (first) {
          accumulator = initial!;
          first = false;
        }
        accumulator = await fn(accumulator!, chunk, { signal: innerSignal });
      }

      if (first && !hasInitial) {
        throw new TypeError("Reduce of an empty stream requires an initial value");
      }
      if (first) {
        return initial!;
      }
      return accumulator!;
    } finally {
      ac.abort();
    }
  }

  /**
   * Returns a new readable containing [index, chunk] pairs.
   * Matches Node.js Readable.prototype.asIndexedPairs() (experimental).
   */
  asIndexedPairs(options?: { signal?: AbortSignal }): Readable<[number, T]> {
    let index = 0;
    return this.map(data => [index++, data] as [number, T], options);
  }

  /**
   * Compose this readable with a stream/transform, returning a Duplex.
   * Matches Node.js behavior where compose() returns a Duplex.
   */
  compose<U>(
    stream: WritableLike | ((source: AsyncIterable<T>) => AsyncIterable<U>),
    _options?: { signal?: AbortSignal }
  ): IDuplex<U, T> {
    // If it's an async generator function, pipe through it
    if (typeof stream === "function") {
      const result = new Readable<U>({ objectMode: true });
      pumpAsyncIterableToReadable(result, stream(this) as AsyncIterable<U>);

      // Propagate destruction from result back to source (matching map/filter/flatMap)
      result.once("close", () => {
        if (!this.destroyed) {
          this.destroy();
        }
      });

      const duplexFrom = getDuplexFrom();
      if (duplexFrom) {
        return duplexFrom(result) as IDuplex<U, T>;
      }
      // Fallback when DuplexFrom is not registered (should not happen at runtime)
      return result as unknown as IDuplex<U, T>;
    }

    // If it's a transform/duplex with pipe support, pipe this into it and
    // return its readable side wrapped as a Duplex.
    const target = stream as any;
    this.pipe(target);
    // If the target is already a Duplex-like, return it directly.
    if (target._readable && target._writable) {
      return target as IDuplex<U, T>;
    }
    // If the target has a _readable (Transform/Duplex), wrap it.
    if (target._readable) {
      const duplexFrom2 = getDuplexFrom();
      if (duplexFrom2) {
        return duplexFrom2(target._readable) as IDuplex<U, T>;
      }
      return target._readable as unknown as IDuplex<U, T>;
    }
    const duplexFrom3 = getDuplexFrom();
    if (duplexFrom3) {
      return duplexFrom3(target) as IDuplex<U, T>;
    }
    return target as IDuplex<U, T>;
  }
}

// Node.js: `Readable.prototype.addListener === Readable.prototype.on` (same function).
// Readable overrides `on` from EventEmitter, so we must re-alias `addListener`.
Readable.prototype.addListener = Readable.prototype.on;

// Node.js: `Readable.prototype._read` throws ERR_METHOD_NOT_IMPLEMENTED when called without
// a subclass override (Node.js 18+).  We replicate that behavior here so that subclasses that
// forget to implement `_read()` see the same error event as on Node.js.
// Factory-created readables (e.g., from async iterables) rely on the internal `_read` check
// being falsy, so the instance field `_read?: ...` shadows this prototype method when set.
(Readable.prototype as any)._read = function _read(_size?: number): void {
  const err = new Error("The _read() method is not implemented") as Error & {
    code: string;
  };
  err.code = "ERR_METHOD_NOT_IMPLEMENTED";
  throw err;
};

// =============================================================================
// Internal helpers – Functional method utilities
// =============================================================================

function _validateAbortSignal(signal: AbortSignal | undefined): void {
  if (signal !== undefined && !(signal instanceof AbortSignal)) {
    throw new TypeError("options.signal must be an AbortSignal");
  }
}

function _throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw createAbortError((signal as any).reason);
  }
}

function _validateNonNegativeInteger(value: number, name: string): void {
  if (
    typeof value !== "number" ||
    Number.isNaN(value) ||
    value < 0 ||
    Math.floor(value) !== value
  ) {
    throw new RangeError(
      `The value of "${name}" must be a non-negative integer. Received ${value}`
    );
  }
}

// =============================================================================
// Internal helpers – Concurrency support for higher-order methods
// =============================================================================

/**
 * Run an async mapper over an async iterable with bounded concurrency,
 * yielding results in input order (matching Node.js Readable.prototype.map).
 *
 * When concurrency === 1, falls back to a simple serial loop for efficiency.
 */
const _EOF = Symbol("EOF");

async function* _mapWithConcurrency<T, U>(
  source: AsyncIterable<T>,
  fn: (data: T) => U | Promise<U>,
  concurrency: number,
  signal: AbortSignal | undefined
): AsyncGenerator<U> {
  if (concurrency <= 1) {
    for await (const chunk of source) {
      _throwIfAborted(signal);
      yield await fn(chunk);
    }
    return;
  }

  // Ordered concurrent execution using a sliding window of promises.
  // Each slot holds a promise that resolves to the mapped value or _EOF sentinel.
  const queue: Promise<U | typeof _EOF>[] = [];
  let error: Error | null = null;
  let sourceDone = false;
  const iterator = source[Symbol.asyncIterator]();

  // Pull one chunk from the source and start processing it.
  async function pullOne(): Promise<U | typeof _EOF> {
    const next = await iterator.next();
    if (next.done) {
      sourceDone = true;
      return _EOF;
    }
    _throwIfAborted(signal);
    return fn(next.value);
  }

  function enqueue(): void {
    queue.push(
      pullOne().catch(err => {
        error = err;
        throw err;
      })
    );
  }

  try {
    // Fill the initial window
    while (queue.length < concurrency && !sourceDone && !error) {
      _throwIfAborted(signal);
      enqueue();
    }

    while (queue.length > 0) {
      _throwIfAborted(signal);
      const result = await queue.shift()!;
      if (error) {
        throw toError(error);
      }
      if (result === _EOF) {
        // Source exhausted — drain any remaining real results
        continue;
      }
      // Yield the result in order
      yield result;

      // Refill the window with one more
      if (!sourceDone && !error) {
        enqueue();
      }
    }
  } finally {
    // Clean up: return the iterator if not done
    if (!sourceDone) {
      await iterator.return?.();
    }
  }
}

// =============================================================================
// Internal helpers
// =============================================================================

export function toAsyncIterable<T>(iterable: Iterable<T> | AsyncIterable<T>): AsyncIterable<T> {
  if (iterable && typeof iterable[Symbol.asyncIterator] === "function") {
    return iterable as AsyncIterable<T>;
  }
  return (async function* (): AsyncIterable<T> {
    for (const item of iterable as Iterable<T>) {
      yield item;
    }
  })();
}

export function pumpAsyncIterableToReadable<T>(
  readable: Readable<T>,
  iterable: AsyncIterable<T>
): void {
  const iterator = iterable[Symbol.asyncIterator]();
  let reading = false;
  let iteratorDone = false;

  // Pull-based: advance the iterator one chunk per _read() call,
  // matching Node.js Readable.from() behavior.
  (readable as any)._hasReadImpl = true;
  (readable as any)._read = function () {
    if (reading || iteratorDone) {
      return;
    }
    reading = true;

    (async () => {
      try {
        if (readable.destroyed) {
          if (!iteratorDone) {
            iteratorDone = true;
            const p = iterator.return?.();
            if (p && typeof p.then === "function") {
              p.catch(() => {});
            }
          }
          return;
        }
        const { value, done } = await iterator.next();
        if (done) {
          iteratorDone = true;
          readable.push(null);
          return;
        }
        readable.push(value);
      } catch (err) {
        readable.destroy(err as Error);
      } finally {
        reading = false;
      }
    })();
  };
  (readable as any)._pushMode = true;
}
