/**
 * Browser Stream - Duplex
 */

import { addEmitterListener, createListenerRegistry } from "@stream/browser/helpers";
import { deferTask, inDeferredContext } from "@stream/browser/microtask-context";
import { Readable } from "@stream/browser/readable";
import { Writable } from "@stream/browser/writable";
import { parseEndArgs } from "@stream/common/end-args";
import { StreamTypeError } from "@stream/errors";
import type { DuplexStreamOptions, IDuplex, WritableLike } from "@stream/types";
import { createAbortError } from "@utils/errors";
import { EventEmitter } from "@utils/event-emitter";

// =============================================================================
// Duplex Stream
// =============================================================================

/**
 * A duplex stream that combines readable and writable
 */
export class Duplex<TRead = Uint8Array, TWrite = Uint8Array> extends EventEmitter {
  /**
   * Allow duck-typed instanceof checks.
   * Browser Duplex composes Readable + Writable, so we use Symbol.hasInstance
   * to check for key Duplex-like methods/properties (both readable and writable).
   * This makes `transform instanceof Duplex` return true.
   */
  static [Symbol.hasInstance](instance: any): boolean {
    if (instance == null || typeof instance !== "object") {
      return false;
    }
    // Fast path: actual Duplex prototype
    if (Object.prototype.isPrototypeOf.call(Duplex.prototype, instance)) {
      return true;
    }
    // Duck-type: must have both Readable and Writable characteristics + stream brand
    return (
      instance.__documonster_stream === true &&
      typeof instance.read === "function" &&
      typeof instance.pipe === "function" &&
      typeof instance.write === "function" &&
      typeof instance.end === "function" &&
      typeof instance.on === "function" &&
      "readableFlowing" in instance &&
      "writableFinished" in instance
    );
  }

  /** @internal */
  private readonly _readable: Readable<TRead>;
  /** @internal */
  private readonly _writable: Writable<TWrite>;
  allowHalfOpen: boolean;

  /**
   * Check if a stream has been disturbed (data read or piped).
   * Delegates to Readable.isDisturbed, checking internal _readable for Duplex/Transform.
   */
  static isDisturbed(stream: any): boolean {
    if (stream && stream._readable instanceof Readable) {
      return Readable.isDisturbed(stream._readable);
    }
    return Readable.isDisturbed(stream);
  }

  /**
   * Create a Duplex stream from various sources
   */
  static from<R = Uint8Array, W = Uint8Array>(
    source:
      | Duplex<R, W>
      | Readable<R>
      | Writable<W>
      | AsyncIterable<R>
      | Iterable<R>
      | string
      | Blob
      | Promise<any>
      | ReadableStream<R>
      | WritableStream<W>
      | {
          readable?: Readable<R>;
          writable?: Writable<W>;
        }
  ): Duplex<R, W> {
    // If it's already a Duplex, return as-is
    if (source instanceof Duplex) {
      return source;
    }

    const forwardReadableToDuplex = (readable: Readable<R>, duplex: Duplex<R, W>): void => {
      const sink = new Writable<R>({
        objectMode: duplex.readableObjectMode,
        write(chunk, _encoding, callback) {
          duplex.push(chunk);
          callback();
        },
        final(callback) {
          duplex.push(null);
          callback();
        }
      });

      const onError = (err: Error): void => {
        duplex.emit("error", err);
      };
      const cleanupError = addEmitterListener(readable, "error", onError);
      addEmitterListener(readable, "end", cleanupError, { once: true });
      addEmitterListener(readable, "close", cleanupError, { once: true });
      addEmitterListener(sink, "finish", cleanupError, { once: true });
      readable.pipe(sink);
    };

    // Promise source — resolve and recursively call from()
    if (source instanceof Promise) {
      const duplex = new Duplex<R, W>({ objectMode: true });
      source
        .then(value => {
          const inner = Duplex.from<R, W>(value);
          forwardReadableToDuplex(inner as unknown as Readable<R>, duplex);
        })
        .catch(err => {
          duplex.destroy(err instanceof Error ? err : new Error(String(err)));
        });
      return duplex;
    }

    // String source — wrap as a single-chunk readable
    if (typeof source === "string") {
      const readable = Readable.from([source] as Iterable<any>);
      const duplex = new Duplex<R, W>({ objectMode: true });
      forwardReadableToDuplex(readable as unknown as Readable<R>, duplex);
      return duplex;
    }

    // Blob source — convert to ReadableStream then to Readable
    if (typeof Blob !== "undefined" && source instanceof Blob) {
      const readable = Readable.fromWeb(source.stream() as ReadableStream);
      const duplex = new Duplex<R, W>();
      forwardReadableToDuplex(readable as unknown as Readable<R>, duplex);
      return duplex;
    }

    // If it has readable and/or writable properties
    if (
      typeof source === "object" &&
      source !== null &&
      "readable" in source &&
      "writable" in source
    ) {
      const pair = source as { readable?: Readable<R>; writable?: Writable<W> };

      // Create one duplex that can bridge both sides.
      // (Previous behavior returned a new writable-only Duplex and dropped the readable side.)
      const duplex = new Duplex<R, W>({
        readableObjectMode: pair.readable?.readableObjectMode,
        writableObjectMode: pair.writable?.writableObjectMode,
        write: pair.writable
          ? (chunk, encoding, callback) => {
              pair.writable!.write(chunk, encoding, callback);
            }
          : undefined,
        final: pair.writable
          ? callback => {
              pair.writable!.end(callback);
            }
          : undefined,
        destroy: (error, callback) => {
          // Propagate destroy to the original source streams so they are cleaned
          // up when the wrapping Duplex is destroyed. Node.js does this too.
          if (pair.readable && !pair.readable.destroyed) {
            pair.readable.destroy(error ?? undefined);
          }
          if (pair.writable && !pair.writable.destroyed) {
            pair.writable.destroy(error ?? undefined);
          }
          callback(error);
        }
      });

      if (pair.readable) {
        forwardReadableToDuplex(pair.readable, duplex);
      }
      return duplex;
    }

    // Web ReadableStream — wrap as readable-only Duplex (matches Node.js Duplex.from(ReadableStream))
    if (
      typeof source === "object" &&
      source !== null &&
      typeof (source as any).getReader === "function" &&
      typeof (source as any).cancel === "function"
    ) {
      const readable = Readable.fromWeb(source as ReadableStream<R>);
      const duplex = new Duplex<R, W>({
        objectMode: readable.readableObjectMode
      });
      forwardReadableToDuplex(readable, duplex);
      return duplex;
    }

    // Web WritableStream — wrap as writable-only Duplex (matches Node.js Duplex.from(WritableStream))
    if (
      typeof source === "object" &&
      source !== null &&
      typeof (source as any).getWriter === "function" &&
      typeof (source as any).close === "function"
    ) {
      const writable = Writable.fromWeb(source as WritableStream<W>);
      return new Duplex<R, W>({
        objectMode: writable.writableObjectMode,
        write(chunk, encoding, callback) {
          writable.write(chunk as W, encoding, callback);
        },
        final(callback) {
          writable.end(callback);
        },
        destroy(error, callback) {
          if (!writable.destroyed) {
            writable.destroy(error ?? undefined);
          }
          callback(error);
        }
      });
    }

    // If it's an iterable
    if (
      typeof source === "object" &&
      source !== null &&
      (Symbol.asyncIterator in (source as object) || Symbol.iterator in (source as object))
    ) {
      const readable = Readable.from(source as AsyncIterable<R> | Iterable<R>);
      const duplex = new Duplex<R, W>({
        objectMode: readable.readableObjectMode
      });
      forwardReadableToDuplex(readable, duplex);
      return duplex;
    }

    // If it's a Readable
    if (source instanceof Readable) {
      const duplex = new Duplex<R, W>({
        objectMode: source.readableObjectMode
      });
      forwardReadableToDuplex(source, duplex);
      return duplex;
    }

    // If it's a Writable
    if (source instanceof Writable) {
      return new Duplex<R, W>({
        objectMode: source.writableObjectMode,
        write(chunk, encoding, callback) {
          source.write(chunk as W, encoding, callback);
        },
        final(callback) {
          source.end(callback);
        }
      });
    }

    throw new StreamTypeError("Duplex-compatible source", typeof source);
  }

  /**
   * Create a Duplex from a Web ReadableWritablePair
   */
  static fromWeb<R = Uint8Array, W = Uint8Array>(
    pair: { readable: ReadableStream<R>; writable: WritableStream<W> },
    options?: DuplexStreamOptions
  ): Duplex<R, W> {
    const duplex = new Duplex<R, W>(options);

    const newReadable = new Readable<R>({
      stream: pair.readable,
      objectMode: duplex.readableObjectMode
    });
    const newWritable = new Writable<W>({
      stream: pair.writable,
      objectMode: duplex.writableObjectMode
    });

    if (duplex._sideForwardingCleanup) {
      duplex._sideForwardingCleanup();
      duplex._sideForwardingCleanup = null;
    }

    (duplex as any)._readable = newReadable;
    (duplex as any)._writable = newWritable;

    // Re-wire event forwarding (data forwarding remains lazy via Duplex.on)
    duplex._setupSideForwarding();

    return duplex;
  }

  /**
   * Convert a Node.js Duplex to Web ReadableWritablePair
   */
  static toWeb<R = Uint8Array, W = Uint8Array>(
    duplex: Duplex<R, W>
  ): { readable: ReadableStream<R>; writable: WritableStream<W> } {
    return {
      readable: Readable.toWeb(duplex._readable),
      writable: Writable.toWeb(duplex._writable)
    };
  }

  // Track if we've already set up data forwarding
  private _dataForwardingSetup: boolean = false;
  private _destroyed: boolean = false;
  private _emitClose: boolean;
  private _errored: Error | null = null;
  private _closed: boolean = false;
  private _autoDestroy: boolean;

  private _sideForwardingCleanup: (() => void) | null = null;
  // User-provided construct function (Node.js compatibility)
  private _constructFunc?: (callback: (error?: Error | null) => void) => void;
  private _constructed: boolean = true;

  /**
   * Detect a subclass-defined implementation hook (e.g. _read/_write/_writev/_final).
   *
   * Node.js allows implementing Duplex by subclassing and defining _read/_write on the
   * subclass prototype. Since the browser Duplex composes internal Readable/Writable
   * instances, we must explicitly forward these prototype hooks.
   */
  private _getSubclassHook(name: string): ((...args: any[]) => any) | undefined {
    let proto = Object.getPrototypeOf(this);
    while (proto && proto !== Duplex.prototype && proto !== Object.prototype) {
      if (Object.prototype.hasOwnProperty.call(proto, name)) {
        const fn = (this as any)[name];
        if (typeof fn === "function") {
          return fn.bind(this);
        }
        return undefined;
      }
      proto = Object.getPrototypeOf(proto);
    }
    return undefined;
  }

  constructor(
    options?: DuplexStreamOptions & {
      allowHalfOpen?: boolean;
      objectMode?: boolean;
      emitClose?: boolean;
      autoDestroy?: boolean;
      encoding?: string;
      decodeStrings?: boolean;
      defaultEncoding?: string;
      signal?: AbortSignal;
      read?: (this: Duplex<TRead, TWrite>, size?: number) => void;
      write?: (
        this: Duplex<TRead, TWrite>,
        chunk: TWrite,
        encoding: string,
        callback: (error?: Error | null) => void
      ) => void;
      writev?: (
        this: Duplex<TRead, TWrite>,
        chunks: Array<{ chunk: TWrite; encoding: string }>,
        callback: (error?: Error | null) => void
      ) => void;
      final?: (this: Duplex<TRead, TWrite>, callback: (error?: Error | null) => void) => void;
      destroy?: (
        this: Duplex<TRead, TWrite>,
        error: Error | null,
        callback: (error?: Error | null) => void
      ) => void;
      construct?: (this: Duplex<TRead, TWrite>, callback: (error?: Error | null) => void) => void;
    }
  ) {
    super();
    (this as any).__documonster_stream = true;

    this.allowHalfOpen = options?.allowHalfOpen ?? true;
    this._emitClose = options?.emitClose ?? true;
    this._autoDestroy = options?.autoDestroy ?? true;
    // Support shorthand objectMode option
    const objectMode = options?.objectMode ?? false;
    const readableObjMode = options?.readableObjectMode ?? objectMode;
    const writableObjMode = options?.writableObjectMode ?? objectMode;

    // HWM: if highWaterMark is explicitly provided it overrides per-side (matching Node)
    const hasGeneralHwm =
      options != null && Object.prototype.hasOwnProperty.call(options, "highWaterMark");
    const readableHwm = hasGeneralHwm ? options!.highWaterMark : options?.readableHighWaterMark;
    const writableHwm = hasGeneralHwm ? options!.highWaterMark : options?.writableHighWaterMark;

    // Store user-provided destroy function
    if (options?.destroy) {
      this._destroy = options.destroy.bind(this);
    }

    // Store user-provided construct function
    if (options?.construct) {
      this._constructFunc = options.construct.bind(this);
    }

    // When Duplex has a construct hook, propagate delay to child streams
    // so that reads/writes are queued until the Duplex-level construct fires.
    let readableConstructCb: ((error?: Error | null) => void) | undefined;
    let writableConstructCb: ((error?: Error | null) => void) | undefined;
    const hasConstruct = this._hasConstructHook();

    const readHook = options?.read ? options.read.bind(this) : this._getSubclassHook("_read");

    // Prefer constructor options over subclass hooks (matches Node's behavior where
    // options.{read,write,final,writev} override prototype hooks).
    const writeHook = options?.write
      ? options.write.bind(this)
      : (this._getSubclassHook("_write") as
          | ((chunk: TWrite, encoding: string, callback: (error?: Error | null) => void) => void)
          | undefined);
    const writevHook = options?.writev
      ? options.writev.bind(this)
      : (this._getSubclassHook("_writev") as
          | ((
              chunks: Array<{ chunk: TWrite; encoding: string }>,
              callback: (error?: Error | null) => void
            ) => void)
          | undefined);
    const finalHook = options?.final
      ? options.final.bind(this)
      : (this._getSubclassHook("_final") as
          | ((callback: (error?: Error | null) => void) => void)
          | undefined);

    this._readable = new Readable<TRead>({
      highWaterMark: readableHwm,
      objectMode: readableObjMode,
      read: readHook as any,
      encoding: options?.encoding,
      // Suppress child-level close/error — Duplex itself is the authority
      emitClose: false,
      autoDestroy: false,
      // Propagate construct delay to child readable
      construct: hasConstruct
        ? cb => {
            readableConstructCb = cb;
          }
        : undefined
    });

    this._writable = new Writable<TWrite>({
      highWaterMark: writableHwm,
      objectMode: writableObjMode,
      write: writeHook as any,
      writev: writevHook as any,
      final: finalHook as any,
      decodeStrings: options?.decodeStrings,
      defaultEncoding: options?.defaultEncoding,
      // Suppress child-level close/error — Duplex itself is the authority
      emitClose: false,
      autoDestroy: false,
      // Propagate construct delay to child writable
      construct: hasConstruct
        ? cb => {
            writableConstructCb = cb;
          }
        : undefined
    });

    // Prevent unhandled error throws on child streams.
    // Errors are forwarded to the Duplex via _setupSideForwarding; these
    // noop listeners act as safety nets after forwarding cleanup.
    const noop = (): void => {};
    this._readable.on("error", noop);
    this._writable.on("error", noop);

    this._setupSideForwarding();

    // signal option — destroy the Duplex when signal aborts (matching Node)
    if (options?.signal) {
      this._setupAbortSignal(options.signal);
    }

    // R5-1: _construct hook — if provided, delay reads/writes until constructed
    if (hasConstruct) {
      this._constructed = false;
      deferTask(() => {
        const fn = this._constructFunc ?? (this as any)._construct.bind(this);
        fn(err => {
          if (err) {
            readableConstructCb?.(err);
            writableConstructCb?.(err);
            this.destroy(err);
            return;
          }
          this._constructed = true;
          // Unblock child streams by firing their construct callbacks
          readableConstructCb?.();
          writableConstructCb?.();
        });
      });
    }
  }

  private _setupAbortSignal(signal: AbortSignal): void {
    if (signal.aborted) {
      this.destroy(createAbortError((signal as any).reason));
      return;
    }

    const onAbort = (): void => {
      this.destroy(createAbortError((signal as any).reason));
    };

    const cleanup = (): void => {
      signal.removeEventListener("abort", onAbort);
    };

    signal.addEventListener("abort", onAbort, { once: true });
    this.once("close", cleanup);
  }

  private _setupSideForwarding(): void {
    if (this._sideForwardingCleanup) {
      this._sideForwardingCleanup();
      this._sideForwardingCleanup = null;
    }

    // Ensure the pipe source identity is always the outer Duplex,
    // including after fromWeb() replaces the internal _readable.
    (this._readable as any)._pipes?.setSource(this);

    const registry = createListenerRegistry();

    // Deduplicate error forwarding: when destroy(err) tears down both sides,
    // both _readable and _writable emit "error". Node.js Duplex only emits
    // one "error" event, so we guard against duplicate forwarding.
    let errorForwarded = false;
    const forwardError = (err: any): void => {
      if (errorForwarded) {
        return;
      }
      errorForwarded = true;
      this.emit("error", err);
    };

    // Auto-destroy: when both sides finish, destroy the Duplex (matching Node.js).
    let readableEnded = false;
    let writableFinished = false;
    const maybeAutoDestroy = (): void => {
      if (this._autoDestroy && readableEnded && writableFinished && !this._destroyed) {
        this.destroy();
      }
    };

    // Forward non-data events (data forwarding is lazy to avoid premature flowing)
    registry.once(this._readable, "end", () => {
      this.emit("end");
      readableEnded = true;
      if (!this.allowHalfOpen) {
        this._writable.end();
      }
      maybeAutoDestroy();
    });
    registry.add(this._readable, "error", forwardError);
    // Use EventEmitter.prototype.on directly to register "readable" forwarding,
    // bypassing Readable's on() override which sets readableFlowing = false.
    const readableForwarder = (): void => {
      this.emit("readable");
    };
    EventEmitter.prototype.on.call(this._readable, "readable", readableForwarder);
    registry.add(this._readable, "pause", () => this.emit("pause"));
    registry.add(this._readable, "resume", () => this.emit("resume"));

    registry.add(this._writable, "error", forwardError);
    registry.once(this._writable, "prefinish", () => this.emit("prefinish"));
    registry.once(this._writable, "finish", () => {
      this.emit("finish");
      writableFinished = true;
      maybeAutoDestroy();
    });
    registry.add(this._writable, "drain", () => this.emit("drain"));
    // Forward "pipe"/"unpipe" from the internal writable so that
    // source.pipe(duplex) triggers "pipe" on the Duplex itself.
    registry.add(this._writable, "pipe", source => this.emit("pipe", source));
    registry.add(this._writable, "unpipe", source => this.emit("unpipe", source));
    // Node.js: when allowHalfOpen is false and the writable side finishes,
    // gracefully end the readable side with push(null) — NOT destroy().
    // The check is dynamic (inside the handler) so runtime changes to
    // allowHalfOpen are respected, matching Node.js behaviour.
    registry.once(this._writable, "finish", () => {
      if (!this.allowHalfOpen && !this._readable.readableEnded) {
        this._readable.push(null);
      }
    });

    // Node.js parity: when the internal _readable is destroyed independently
    // (e.g. by a higher-order method like map/filter/forEach that uses the
    // async iterator and calls _readable.destroy() on error or early exit),
    // propagate the destruction to the outer Duplex.  Without this, only the
    // readable side would be torn down, leaving the writable side open — a
    // resource leak that doesn't happen in Node.js where Duplex IS the
    // Readable (same object, so destroy() hits both sides).
    const origReadableDestroy = this._readable.destroy.bind(this._readable);
    (this._readable as any).destroy = (err?: Error): any => {
      const result = origReadableDestroy(err);
      if (!this._destroyed) {
        this.destroy(err);
      }
      return result;
    };

    this._sideForwardingCleanup = () => {
      registry.cleanup();
      this._readable.off("readable", readableForwarder);
      // Restore original destroy to avoid leaking the closure after cleanup.
      (this._readable as any).destroy = origReadableDestroy;
    };
  }

  /**
   * Override on() to set up data forwarding lazily
   */
  override on(event: string | symbol, listener: (...args: any[]) => void): this {
    // Register the listener FIRST so that when _readable.on("data") triggers
    // resume() and synchronously drains buffered data, the forwarding handler
    // can find the listener already in place on this Duplex.
    super.on(event, listener);

    // Set up data forwarding when first external data listener is added
    if (event === "data" && !this._dataForwardingSetup) {
      this._dataForwardingSetup = true;
      this._readable.on("data", chunk => this.emit("data", chunk));
    } else if (event === "readable") {
      // Node.js: adding a 'readable' listener sets readableFlowing to false
      this._readable._setReadableFlowing(false);
    }
    return this;
  }

  /**
   * Push data to readable side
   */
  push(chunk: TRead | null, encoding?: string): boolean {
    return this._readable.push(chunk, encoding);
  }

  /**
   * Put a chunk back at the front of the buffer (readable side)
   */
  unshift(chunk: TRead, encoding?: string): void {
    this._readable.unshift(chunk, encoding);
  }

  /**
   * Read from readable side
   */
  read(size?: number): TRead | null {
    return this._readable.read(size);
  }

  /**
   * Write to writable side
   */
  write(chunk: TWrite, callback?: (error?: Error | null) => void): boolean;
  write(chunk: TWrite, encoding?: string, callback?: (error?: Error | null) => void): boolean;
  write(
    chunk: TWrite,
    encodingOrCallback?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void
  ): boolean {
    const encoding = typeof encodingOrCallback === "string" ? encodingOrCallback : undefined;
    const cb = typeof encodingOrCallback === "function" ? encodingOrCallback : callback;
    return encoding !== undefined
      ? this._writable.write(chunk, encoding, cb)
      : this._writable.write(chunk, cb);
  }

  /**
   * End writable side
   */
  end(callback?: () => void): this;
  end(chunk: TWrite, callback?: () => void): this;
  end(chunk: TWrite, encoding?: string, callback?: () => void): this;
  end(
    chunkOrCallback?: TWrite | (() => void),
    encodingOrCallback?: string | (() => void),
    callback?: () => void
  ): this {
    const { chunk, encoding, cb } = parseEndArgs<TWrite>(
      chunkOrCallback,
      encodingOrCallback,
      callback
    );

    // Repeated end() protection (matches Node.js Writable.end() behavior)
    if (this._writable.writableEnded) {
      // If a chunk was provided, this is a write-after-end error.
      if (chunk !== undefined) {
        this.write(chunk, encoding, err => {
          (cb as any)?.(err ?? null);
        });
        return this;
      }

      // If we've already finished, callback receives ERR_STREAM_ALREADY_FINISHED.
      if (this.writableFinished && cb) {
        const err = new Error("Cannot call end after a stream was finished") as Error & {
          code: string;
        };
        err.code = "ERR_STREAM_ALREADY_FINISHED";
        deferTask(() => (cb as any)(err));
      } else if (cb) {
        // Redundant end() is a no-op; callback called with no error.
        deferTask(() => (cb as any)(null));
      }
      return this;
    }

    if (cb) {
      // Node.js fires the end() callback on 'finish', but also on 'close'
      // (which is emitted after destroy). This ensures the callback fires
      // even when the stream is destroyed before finishing.
      let called = false;
      const onceCb = (): void => {
        if (!called) {
          called = true;
          cb();
        }
      };
      this.once("finish", onceCb);
      this.once("close", onceCb);
    }

    if (chunk !== undefined) {
      // Propagate write errors through destroy (matching Node.js)
      const onWriteError = (err?: Error | null): void => {
        if (err && !this.destroyed) {
          this.destroy(err);
        }
      };
      if (encoding !== undefined) {
        this._writable.write(chunk, encoding, onWriteError);
      } else {
        this._writable.write(chunk, onWriteError as any);
      }
    }
    this._writable.end();
    return this;
  }

  /**
   * Cork the writable side
   */
  cork(): void {
    this._writable.cork();
  }

  /**
   * Uncork the writable side
   */
  uncork(): void {
    this._writable.uncork();
  }

  /**
   * Set encoding for readable side
   */
  setEncoding(encoding: string): this {
    this._readable.setEncoding(encoding);
    return this;
  }

  /**
   * Set default encoding for writable side
   */
  setDefaultEncoding(encoding: string): this {
    this._writable.setDefaultEncoding(encoding);
    return this;
  }

  /**
   * Pipe readable side to destination.
   * Accepts any writable-like object (duck-typed, matching Node.js behavior).
   */
  pipe<W extends WritableLike>(destination: W, options?: { end?: boolean }): W {
    this._readable.pipe(destination, options);
    return destination;
  }

  /**
   * Unpipe from destination
   */
  unpipe(destination?: Writable<TRead>): this {
    this._readable.unpipe(destination);
    return this;
  }

  /**
   * Pause the readable side
   */
  pause(): this {
    this._readable.pause();
    return this;
  }

  /**
   * Resume the readable side
   */
  resume(): this {
    this._readable.resume();
    return this;
  }

  /**
   * Check if paused
   */
  isPaused(): boolean {
    return this._readable.isPaused();
  }

  /**
   * Destroy both sides
   */
  destroy(error?: Error): this {
    if (this._destroyed) {
      return this;
    }
    this._destroyed = true;
    if (this._sideForwardingCleanup) {
      this._sideForwardingCleanup();
      this._sideForwardingCleanup = null;
    }

    const afterDestroy = (finalError?: Error | null): void => {
      // Node.js: _destroy's callback determines whether an error is emitted.
      // cb(null)/cb()/cb(undefined) all suppress the original error.
      // Only cb(new Error(...)) replaces and emits the error.
      const err = finalError || null;
      if (err) {
        this._errored = err;
      }
      this._closed = true;
      // Destroy internal streams without their own error/close emission — the
      // Duplex itself is the authority for those events.
      this._readable.destroy();
      this._writable.destroy();
      const doEmit = (): void => {
        if (err) {
          this.emit("error", err);
        }
        if (this._emitClose) {
          this.emit("close");
        }
      };
      if (
        inDeferredContext() &&
        !this._hasDestroyHook() &&
        this.readableEnded &&
        this.writableFinished
      ) {
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
    this._readable._undestroy();
    this._writable._undestroy();
    this._destroyed = false;
    this._closed = false;
    this._errored = null;
    this._setupSideForwarding();
  }

  /** Check if _destroy has been overridden by a subclass or constructor option. */
  private _hasDestroyHook(): boolean {
    return (
      Object.prototype.hasOwnProperty.call(this, "_destroy") ||
      Object.getPrototypeOf(this)._destroy !== Duplex.prototype._destroy
    );
  }

  /**
   * Check if a subclass defines _construct on its own prototype.
   * Node.js does NOT have _construct on any stream prototype — it only exists
   * when provided via constructor options or defined by a subclass.
   */
  private _hasConstructHook(): boolean {
    if (this._constructFunc) {
      return true;
    }
    let proto = Object.getPrototypeOf(this);
    while (proto && proto !== Duplex.prototype && proto !== Object.prototype) {
      if (Object.prototype.hasOwnProperty.call(proto, "_construct")) {
        return true;
      }
      proto = Object.getPrototypeOf(proto);
    }
    return false;
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
   * Async dispose support (using await).
   * Destroys the stream and resolves after the 'close' event.
   */
  async [Symbol.asyncDispose](): Promise<void> {
    // Match Node.js behavior:
    // 1. If not yet destroyed, destroy the stream (always resolves)
    // 2. If already destroyed by someone else, reject with "Premature close"
    //    unless the stream ended gracefully (writableFinished)
    const selfInitiated = !this._destroyed;
    if (selfInitiated) {
      this.destroy();
    }
    return new Promise<void>((resolve, reject) => {
      const settle = (): void => {
        if (selfInitiated || this.writableFinished) {
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

  get readable(): boolean {
    return this._readable.readable;
  }

  set readable(val: boolean) {
    this._readable.readable = val;
  }

  get writable(): boolean {
    return this._writable.writable;
  }

  set writable(val: boolean) {
    this._writable.writable = val;
  }

  get readableEnded(): boolean {
    return this._readable.readableEnded;
  }

  get writableEnded(): boolean {
    return this._writable.writableEnded;
  }

  get writableFinished(): boolean {
    return this._writable.writableFinished;
  }

  get readableLength(): number {
    return this._readable.readableLength;
  }

  get writableLength(): number {
    return this._writable.writableLength;
  }

  get readableHighWaterMark(): number {
    return this._readable.readableHighWaterMark;
  }

  get writableHighWaterMark(): number {
    return this._writable.writableHighWaterMark;
  }

  get destroyed(): boolean {
    return this._destroyed;
  }

  set destroyed(val: boolean) {
    this._destroyed = val;
    // Propagate to internal streams so their state stays consistent with
    // the Duplex — matches Node.js where Duplex/Readable/Writable share
    // a single destroyed flag via the prototype chain.
    this._readable.destroyed = val;
    this._writable.destroyed = val;
  }

  get writableCorked(): number {
    return this._writable.writableCorked;
  }

  get writableNeedDrain(): boolean {
    return this._writable.writableNeedDrain;
  }

  get readableObjectMode(): boolean {
    return this._readable.readableObjectMode;
  }

  get writableObjectMode(): boolean {
    return this._writable.writableObjectMode;
  }

  get readableFlowing(): boolean | null {
    return this._readable.readableFlowing;
  }

  get readableAborted(): boolean {
    return this._readable.readableAborted;
  }

  get readableDidRead(): boolean {
    return this._readable.readableDidRead;
  }

  get readableEncoding(): string | null {
    return this._readable.readableEncoding;
  }

  get errored(): Error | null {
    return this._errored ?? this._readable.errored ?? this._writable.errored;
  }

  get closed(): boolean {
    return this._closed;
  }

  get readableBuffer(): TRead[] {
    return this._readable.readableBuffer;
  }

  get writableBuffer(): TWrite[] {
    return this._writable.writableBuffer;
  }

  /**
   * Wrap a legacy stream
   */
  wrap(stream: any): this {
    this._readable.wrap(stream);
    return this;
  }

  /**
   * Create an async iterator with options.
   *
   * Node.js parity: when `destroyOnReturn` is true (the default) and the
   * iterator exits early (break/return), the entire Duplex is destroyed —
   * not just the readable side.  The inner `_readable`'s iterator would only
   * destroy `_readable` (which has `autoDestroy: false`, `emitClose: false`),
   * leaving the writable side open.  We wrap the iterator so that `return()`
   * propagates destruction to the outer Duplex.
   */
  iterator(options?: { destroyOnReturn?: boolean }): AsyncIterableIterator<TRead> {
    const destroyOnReturn = options?.destroyOnReturn !== false;
    if (!destroyOnReturn) {
      // No destruction needed on return — delegate directly.
      return this._readable.iterator(options);
    }
    return this._wrapAsyncIterator(this._readable[Symbol.asyncIterator]());
  }

  /**
   * Async iterator support.
   * Wraps the inner readable's iterator so that early exit destroys the
   * entire Duplex (both readable and writable sides), matching Node.js.
   */
  [Symbol.asyncIterator](): AsyncIterableIterator<TRead> {
    return this._wrapAsyncIterator(this._readable[Symbol.asyncIterator]());
  }

  /**
   * Wrap an inner readable iterator so that return()/throw() destroy the
   * outer Duplex, not just the inner readable.
   */
  private _wrapAsyncIterator(inner: AsyncIterableIterator<TRead>): AsyncIterableIterator<TRead> {
    // oxlint-disable-next-line no-this-alias -- needed for object literal method closures
    const duplex = this;
    return {
      next() {
        return inner.next();
      },
      async return() {
        const result = await inner.return?.();
        if (!duplex.destroyed) {
          duplex.destroy();
        }
        return result ?? { value: undefined as any, done: true as const };
      },
      async throw(err?: any) {
        const result = await inner.throw?.(err);
        if (!duplex.destroyed) {
          duplex.destroy(err instanceof Error ? err : undefined);
        }
        return result ?? { value: undefined as any, done: true as const };
      },
      [Symbol.asyncIterator]() {
        return this;
      }
    };
  }

  // =============================================================================
  // Functional / Higher-order Methods (forwarded to readable side)
  // =============================================================================

  map<U>(
    fn: (data: TRead, options: { signal: AbortSignal }) => U | Promise<U>,
    options?: { concurrency?: number; highWaterMark?: number; signal?: AbortSignal }
  ): Readable<U> {
    return this._readable.map(fn, options);
  }

  filter(
    fn: (data: TRead, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; highWaterMark?: number; signal?: AbortSignal }
  ): Readable<TRead> {
    return this._readable.filter(fn, options);
  }

  async forEach(
    fn: (data: TRead, options: { signal: AbortSignal }) => void | Promise<void>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<undefined> {
    return this._readable.forEach(fn, options);
  }

  async toArray(options?: { signal?: AbortSignal }): Promise<TRead[]> {
    return this._readable.toArray(options);
  }

  async some(
    fn: (data: TRead, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    return this._readable.some(fn, options);
  }

  async find(
    fn: (data: TRead, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<TRead | undefined> {
    return this._readable.find(fn, options);
  }

  async every(
    fn: (data: TRead, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    return this._readable.every(fn, options);
  }

  flatMap<U>(
    fn: (
      data: TRead,
      options: { signal: AbortSignal }
    ) => Iterable<U> | AsyncIterable<U> | Readable<U> | Promise<Iterable<U> | AsyncIterable<U>>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Readable<U> {
    return this._readable.flatMap(fn, options);
  }

  drop(limit: number, options?: { signal?: AbortSignal }): Readable<TRead> {
    return this._readable.drop(limit, options);
  }

  take(limit: number, options?: { signal?: AbortSignal }): Readable<TRead> {
    return this._readable.take(limit, options);
  }

  async reduce<U = TRead>(
    fn: (previous: U, data: TRead, options: { signal: AbortSignal }) => U | Promise<U>,
    initial?: U,
    options?: { signal?: AbortSignal }
  ): Promise<U> {
    if (arguments.length >= 2) {
      return this._readable.reduce(fn, initial, options);
    }
    return this._readable.reduce(fn);
  }

  compose<U>(
    stream: WritableLike | ((source: AsyncIterable<TRead>) => AsyncIterable<U>),
    options?: { signal?: AbortSignal }
  ): IDuplex<U, TRead> {
    return this._readable.compose(stream, options);
  }
}

// Node.js: `Duplex.prototype.addListener === Duplex.prototype.on` (same function).
// Duplex overrides `on` from EventEmitter, so we must re-alias `addListener`.
Duplex.prototype.addListener = Duplex.prototype.on;

// Node.js: Duplex.prototype._writev === null (copied from Writable).
(Duplex.prototype as any)._writev = null;
