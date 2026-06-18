/**
 * Browser Stream - Transform
 */

import { Duplex } from "@stream/browser/duplex";
import { createListenerRegistry } from "@stream/browser/helpers";
import { deferTask, inDeferredContext } from "@stream/browser/microtask-context";
import { Readable } from "@stream/browser/readable";
import { Writable } from "@stream/browser/writable";
import { parseEndArgs } from "@stream/common/end-args";
import type { DuplexStreamOptions, IDuplex, WritableLike } from "@stream/types";
import { createAbortError, toError } from "@utils/errors";
import { EventEmitter } from "@utils/event-emitter";

// =============================================================================
// Transform Stream Wrapper
// =============================================================================

/**
 * A wrapper around Web TransformStream that provides Node.js-like API
 */
export class Transform<TInput = Uint8Array, TOutput = Uint8Array> extends EventEmitter {
  /**
   * Allow duck-typed instanceof checks.
   * Makes `transform instanceof Transform` return true, and also
   * `transform instanceof Duplex` via Duplex's own Symbol.hasInstance.
   */
  static [Symbol.hasInstance](instance: any): boolean {
    if (instance == null || typeof instance !== "object") {
      return false;
    }
    // Fast path: actual Transform prototype
    if (Object.prototype.isPrototypeOf.call(Transform.prototype, instance)) {
      return true;
    }
    // Duck-type: must have Duplex characteristics + _transform method
    return (
      instance.__excelts_stream === true &&
      typeof instance.read === "function" &&
      typeof instance.pipe === "function" &&
      typeof instance.write === "function" &&
      typeof instance.end === "function" &&
      typeof instance.on === "function" &&
      typeof instance._transform === "function" &&
      "readableFlowing" in instance &&
      "writableFinished" in instance
    );
  }

  /** @internal */
  private readonly _readable: Readable<TOutput>;
  /** @internal */
  private readonly _writable: Writable<TInput>;
  allowHalfOpen: boolean;

  private _destroyed: boolean = false;
  private _closed: boolean = false;
  private _ended: boolean = false;
  private _errored: Error | null = null;
  private _dataForwardingSetup: boolean = false;
  private _emitClose: boolean;
  private _autoDestroy: boolean;
  /** @internal Set by _scheduleEnd to let _finalHandler complete synchronously */
  private _syncFinal: boolean = false;
  /** @internal Track whether end() was called from sync user code */
  private _endCalledFromSync: boolean = false;

  private _endGeneration: number = 0;
  private _endCallback: (() => void) | null = null;

  private _webStream: TransformStream<TInput, TOutput> | null = null;

  private _sideForwardingCleanup: (() => void) | null = null;
  // User-provided construct function (Node.js compatibility)
  private _constructFunc?: (callback: (error?: Error | null) => void) => void;

  /** Cached result of _hasSubclassTransform (called per-chunk, so worth caching) */
  private _isSubclassTransform: boolean | undefined;

  /**
   * Deferred write callback, stored when the readable buffer is full (backpressure).
   * Released when the internal readable's _read() is called (i.e. when the consumer
   * pulls data), matching Node.js Transform's kCallback mechanism.
   */
  private _afterTransformCallback: ((error?: Error | null) => void) | null = null;

  private _transformImpl:
    | ((chunk: TInput) => TOutput | Promise<TOutput>)
    | ((
        this: Transform<TInput, TOutput>,
        chunk: TInput,
        encoding: string,
        callback: (error?: Error | null, data?: TOutput) => void
      ) => void)
    | undefined;

  private _flushImpl:
    | (() => TOutput | void | Promise<TOutput | void>)
    | ((
        this: Transform<TInput, TOutput>,
        callback: (error?: Error | null, data?: TOutput) => void
      ) => void)
    | undefined;

  /**
   * Push data to the readable side (Node.js compatibility).
   * Intended to be called from within transform/flush.
   */
  push(chunk: TOutput | null, encoding?: string): boolean {
    return this._readable.push(chunk, encoding);
  }

  constructor(
    options?: DuplexStreamOptions & {
      emitClose?: boolean;
      autoDestroy?: boolean;
      encoding?: string;
      decodeStrings?: boolean;
      defaultEncoding?: string;
      signal?: AbortSignal;
      transform?:
        | ((chunk: TInput) => TOutput | Promise<TOutput>)
        | ((
            this: Transform<TInput, TOutput>,
            chunk: TInput,
            encoding: string,
            callback: (error?: Error | null, data?: TOutput) => void
          ) => void);
      flush?:
        | (() => TOutput | void | Promise<TOutput | void>)
        | ((
            this: Transform<TInput, TOutput>,
            callback: (error?: Error | null, data?: TOutput) => void
          ) => void);
      write?: (
        this: Transform<TInput, TOutput>,
        chunk: TInput,
        encoding: string,
        callback: (error?: Error | null) => void
      ) => void;
      writev?: (
        this: Transform<TInput, TOutput>,
        chunks: Array<{ chunk: TInput; encoding: string }>,
        callback: (error?: Error | null) => void
      ) => void;
      final?: (this: Transform<TInput, TOutput>, callback: (error?: Error | null) => void) => void;
      destroy?: (
        this: Transform<TInput, TOutput>,
        error: Error | null,
        callback: (error?: Error | null) => void
      ) => void;
      construct?: (
        this: Transform<TInput, TOutput>,
        callback: (error?: Error | null) => void
      ) => void;
    }
  ) {
    super();
    (this as any).__excelts_stream = true;

    // ObjectMode: per-side overrides general (matching Node)
    const objectMode = options?.objectMode ?? false;
    const readableObjMode = options?.readableObjectMode ?? objectMode;
    const writableObjMode = options?.writableObjectMode ?? objectMode;
    this.allowHalfOpen = options?.allowHalfOpen ?? true;
    this._emitClose = options?.emitClose ?? true;
    this._autoDestroy = options?.autoDestroy ?? true;
    this._transformImpl = options?.transform;
    this._flushImpl = options?.flush;

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

    // When Transform has a construct hook, propagate delay to child streams
    // so that reads/writes are queued until the Transform-level construct fires.
    let readableConstructCb: ((error?: Error | null) => void) | undefined;
    let writableConstructCb: ((error?: Error | null) => void) | undefined;
    const hasConstruct = this._hasConstructHook();

    this._readable = new Readable<TOutput>({
      highWaterMark: readableHwm,
      objectMode: readableObjMode,
      encoding: options?.encoding,
      // Suppress child-level close/error — Transform itself is the authority
      emitClose: false,
      autoDestroy: false,
      // Propagate construct delay to child readable
      construct: hasConstruct
        ? cb => {
            readableConstructCb = cb;
          }
        : undefined
    });

    // Node.js Transform uses _read() as the backpressure release mechanism:
    // when the consumer pulls data (read() or flowing-mode drain), the Readable
    // internals call _read(), which fires the deferred write callback.
    (this._readable as any)._hasReadImpl = true;
    (this._readable as any)._read = () => {
      if (this._afterTransformCallback) {
        const cb = this._afterTransformCallback;
        this._afterTransformCallback = null;
        cb(null);
      }
    };

    // Override pipe source identity so destinations see the outer Transform
    // (not the internal Readable) in "pipe"/"unpipe" events.
    (this._readable as any)._pipes.setSource(this);

    // Determine write/final handlers.
    // If a `write` option is provided, it replaces the transform-based write (matching Node).
    const writeHandler = options?.write
      ? (chunk: TInput, encoding: string, callback: (error?: Error | null) => void) => {
          options.write!.call(this, chunk, encoding, callback);
        }
      : (chunk: TInput, encoding: string, callback: (error?: Error | null) => void) => {
          // Capture readable length before transform to detect if data was pushed.
          // This mirrors Node.js Transform._write which captures rState.length
          // before calling _transform, then uses it in the afterTransform check.
          const lengthBefore = this._readable.readableLength;

          // Helper: call the write callback only when the readable side has room.
          // Matches Node.js afterTransform logic exactly:
          //   if (wState.ended || length === rState.length || rState.length < rState.highWaterMark)
          //     callback();
          //   else
          //     this[kCallback] = callback;  // defer
          const afterTransform = (): void => {
            const rLen = this._readable.readableLength;
            const rHwm = this._readable.readableHighWaterMark;
            if (
              this._writable.writableEnded ||
              lengthBefore === rLen || // No data pushed (filtered) — don't defer
              rLen < rHwm
            ) {
              callback(null);
            } else {
              // Readable backpressure — store callback to be released when
              // the readable's _read() is called by the consumer.
              this._afterTransformCallback = callback;
            }
          };

          // Try synchronous transform first.  If the transform completes
          // synchronously we MUST call the callback synchronously so that
          // the Writable write-queue drains in the same microtask, preventing
          // _scheduleEnd from racing ahead of dynamically-added writes.
          const maybePromise = this._runTransformSync(chunk, encoding);
          if (maybePromise === undefined) {
            // Completed synchronously
            afterTransform();
          } else {
            // Async – wait for the promise
            maybePromise.then(
              () => afterTransform(),
              err => callback(err)
            );
          }
        };

    const finalHandler = options?.final
      ? (callback: (error?: Error | null) => void) => {
          options.final!.call(this, (err?: Error | null) => {
            if (err) {
              callback(err);
              return;
            }
            // Even with a custom _final, the readable side must end (push null).
            // Node.js Transform always ends the readable side when the writable
            // side finishes, regardless of custom _final.
            this._readable.push(null);
            callback(null);
          });
        }
      : (callback: (error?: Error | null) => void) => {
          // Check if a subclass overrides _final on the prototype chain.
          // This handles all inheritance depths: A extends Transform, B extends A, etc.
          const hasFinalOverride =
            typeof (this as any)._final === "function" &&
            (this as any)._final !== Transform.prototype._final;

          if (hasFinalOverride) {
            (this as any)._final.call(this, (err?: Error | null) => {
              if (err) {
                callback(err);
                return;
              }
              // _final already handles push(null) via the prototype method
              callback(null);
            });
          } else if (this._hasSubclassFlush() || this._flushImpl) {
            // Has a flush function — must run it (potentially async)
            this._runFlush()
              .then(() => {
                this._readable.push(null);
                callback(null);
              })
              .catch(err => callback(err));
          } else if (this._syncFinal) {
            // _scheduleEnd determined it's safe to complete synchronously
            // (no flush, no _final override, no pipe destinations on readable).
            // Execute push(null) + callback(null) inline to keep end/finish/close
            // ahead of user Promises, matching Node.js process.nextTick nesting.
            this._readable.push(null);
            // push(null) schedules _emitEndOnce via deferTask. Call it
            // synchronously so 'end' fires before 'finish' (Node.js order).
            (this._readable as any)._emitEndOnce();
            callback(null);
          } else {
            // No flush, no _final override, but the readable side has pipe
            // destinations (e.g. PassThrough used in archive pipe chains).
            // Must defer to avoid deadlocking chained pipe scenarios that
            // need a microtask window for drain events to clear buffers.
            deferTask(() => {
              this._readable.push(null);
              callback(null);
            });
          }
        };

    this._writable = new Writable<TInput>({
      highWaterMark: writableHwm,
      objectMode: writableObjMode,
      // Suppress child-level close/error — Transform itself is the authority
      emitClose: false,
      autoDestroy: false,
      write: writeHandler,
      writev: options?.writev?.bind(this),
      final: finalHandler,
      decodeStrings: options?.decodeStrings,
      defaultEncoding: options?.defaultEncoding,
      // Propagate construct delay to child writable
      construct: hasConstruct
        ? cb => {
            writableConstructCb = cb;
          }
        : undefined
    });

    // Prevent unhandled error throws on child streams.
    // Errors are forwarded to the Transform via _setupSideForwarding; these
    // noop listeners act as safety nets after forwarding cleanup.
    const noop = (): void => {};
    this._readable.on("error", noop);
    this._writable.on("error", noop);

    this._setupSideForwarding();

    // signal option — destroy the Transform when signal aborts (matching Node)
    if (options?.signal) {
      this._setupAbortSignal(options.signal);
    }

    // R7-3: _construct hook — if provided, delay reads/writes until constructed
    if (hasConstruct) {
      deferTask(() => {
        const fn = this._constructFunc ?? (this as any)._construct.bind(this);
        fn((err?: Error | null) => {
          if (err) {
            readableConstructCb?.(err);
            writableConstructCb?.(err);
            this.destroy(err);
            return;
          }
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

    // Ensure the pipe source identity is always the outer Transform,
    // including after fromWeb() replaces the internal _readable.
    (this._readable as any)._pipes?.setSource(this);

    const registry = createListenerRegistry();

    // Auto-destroy: when both sides finish, destroy the Transform (matching Node.js).
    let readableEnded = false;
    let writableFinished = false;
    const maybeAutoDestroy = (): void => {
      if (this._autoDestroy && readableEnded && writableFinished && !this._destroyed) {
        this.destroy();
      }
    };

    registry.once(this._readable, "end", () => {
      this.emit("end");
      readableEnded = true;
      if (!this.allowHalfOpen) {
        this._writable.end();
      }
      maybeAutoDestroy();
    });
    registry.add(this._readable, "error", err => this._emitErrorOnce(err));
    // Use EventEmitter.prototype.on directly to register "readable" forwarding,
    // bypassing Readable's on() override which sets readableFlowing = false.
    const readableForwarder = (): void => {
      this.emit("readable");
    };
    EventEmitter.prototype.on.call(this._readable, "readable", readableForwarder);
    registry.add(this._readable, "pause", () => this.emit("pause"));
    registry.add(this._readable, "resume", () => this.emit("resume"));

    registry.once(this._writable, "prefinish", () => this.emit("prefinish"));
    registry.once(this._writable, "finish", () => {
      this.emit("finish");
      writableFinished = true;
      maybeAutoDestroy();
    });
    registry.add(this._writable, "drain", () => this.emit("drain"));
    registry.add(this._writable, "error", err => this._emitErrorOnce(err));
    // Node.js: when allowHalfOpen is false and the writable side finishes,
    // gracefully end the readable side with push(null) — NOT destroy().
    // Listen on "finish" (not "close") to match Node.js timing.
    registry.once(this._writable, "finish", () => {
      if (!this.allowHalfOpen && !this._readable.readableEnded) {
        this._readable.push(null);
      }
    });

    this._sideForwardingCleanup = () => {
      registry.cleanup();
      this._readable.off("readable", readableForwarder);
    };
  }

  private _scheduleEnd(): void {
    if (this._destroyed || this._errored) {
      return;
    }
    if (this._writable.writableEnded) {
      return;
    }

    const gen = ++this._endGeneration;

    // Defer to the next microtask so that synchronous code following the
    // readable push(null) can still register listeners or write data.
    // Node.js uses process.nextTick here; deferTask is the closest
    // browser equivalent.
    deferTask(() => {
      if (gen !== this._endGeneration) {
        return;
      }
      if (this._destroyed || this._errored || this._writable.writableEnded) {
        return;
      }
      // Signal to _finalHandler that it can complete synchronously when safe
      // (no pipe destinations on the readable side). Without this, the
      // _finalHandler's deferTask puts push(null) + callback(null) behind
      // user Promises, diverging from Node.js process.nextTick nesting.
      if (
        this._endCalledFromSync &&
        !this._hasSubclassFlush() &&
        !this._flushImpl &&
        !(
          typeof (this as any)._final === "function" &&
          (this as any)._final !== Transform.prototype._final
        ) &&
        (this._readable as any)._pipes._destinations.length === 0
      ) {
        this._syncFinal = true;
        this._writable._syncFinish = true;
      }
      this._writable.end();
      this._syncFinal = false;
      // Reset _syncFinish in case _finalFunc took an error path and
      // didn't consume the flag (the success path resets it itself).
      this._writable._syncFinish = false;
    });
  }

  private _emitErrorOnce(err: any): void {
    if (this._errored) {
      return;
    }
    const error = err instanceof Error ? err : new Error(String(err));
    this._errored = error;
    // Use destroy() so that _destroy hooks, cleanup, and close emission
    // all go through the standard path (matching Node.js behavior).
    // destroy() will emit both the error and the close event.
    if (!this._destroyed) {
      this.destroy(error);
    }
  }

  private _hasSubclassTransform(): boolean {
    if (this._isSubclassTransform !== undefined) {
      return this._isSubclassTransform;
    }
    // When options.transform was provided, it takes priority over prototype
    // overrides (matching Node.js behavior where the constructor stores
    // options.transform and uses it directly).
    if (this._transformImpl) {
      this._isSubclassTransform = false;
      return false;
    }
    const proto = Object.getPrototypeOf(this);
    this._isSubclassTransform = proto._transform !== Transform.prototype._transform;
    return this._isSubclassTransform;
  }

  private _hasSubclassFlush(): boolean {
    // When options.flush was provided, it takes priority (matching Node.js).
    if (this._flushImpl) {
      return false;
    }
    // Walk the prototype chain to find a subclass-defined _flush.
    // Node.js does NOT have _flush on Transform.prototype (it's undefined).
    let proto = Object.getPrototypeOf(this);
    while (proto && proto !== Transform.prototype && proto !== Object.prototype) {
      if (Object.prototype.hasOwnProperty.call(proto, "_flush")) {
        return true;
      }
      proto = Object.getPrototypeOf(proto);
    }
    return false;
  }

  /**
   * Run the transform function.  Returns `undefined` when the transform
   * completed synchronously, or a `Promise<void>` when it is async.
   * Keeping the sync path truly synchronous is critical so that the Writable
   * write-queue callback fires synchronously and _scheduleEnd cannot race
   * ahead of writes added during 'data' callbacks.
   */
  private _runTransformSync(chunk: TInput, encoding: string): Promise<void> | undefined {
    // Node.js silently drops writes on destroyed/errored streams (the internal
    // Writable state machine intercepts them before _transform is called).
    // Match that behavior by returning immediately.
    if (this._destroyed || this._errored) {
      return undefined;
    }

    if (this._hasSubclassTransform()) {
      // Use the same sync-detection pattern as _transformImpl below so that
      // subclasses whose _transform calls the callback synchronously (e.g.
      // PassThrough) stay on the synchronous fast path instead of always
      // paying a microtask-delay through Promise.
      let sync = true;
      let syncDone = false;
      let syncErr: Error | null = null;
      let syncData: TOutput | undefined;
      let callbackFired = false;

      let resolveAsync: (() => void) | null = null;
      let rejectAsync: ((err: any) => void) | null = null;

      this._transform(chunk, encoding, (err?: Error | null, data?: TOutput) => {
        if (callbackFired) {
          // Node.js throws ERR_MULTIPLE_CALLBACK on double invocation.
          const multiErr = new Error("Callback called multiple times") as Error & {
            code: string;
          };
          multiErr.code = "ERR_MULTIPLE_CALLBACK";
          this.destroy(multiErr);
          return;
        }
        callbackFired = true;
        if (sync) {
          syncDone = true;
          syncErr = err ?? null;
          syncData = data;
          return;
        }
        if (err) {
          rejectAsync?.(err);
          return;
        }
        if (data !== undefined) {
          this.push(data);
        }
        resolveAsync?.();
      });

      sync = false;

      if (syncDone) {
        if (syncErr) {
          throw toError(syncErr);
        }
        if (syncData !== undefined) {
          this.push(syncData);
        }
        return undefined; // sync completion — no microtask
      }

      // Callback was not called synchronously — wait for async callback.
      return new Promise<void>((resolve, reject) => {
        resolveAsync = resolve;
        rejectAsync = reject;
      });
    }

    const userTransform = this._transformImpl;
    if (!userTransform) {
      // No user transform AND no subclass override: call the prototype
      // _transform which throws ERR_METHOD_NOT_IMPLEMENTED (matching Node.js).
      // This is NOT the passthrough behavior — that requires explicitly setting
      // a transform function (as PassThrough does).
      let sync = true;
      let syncDone = false;
      let syncErr: Error | null = null;
      let syncData: TOutput | undefined;

      let resolveAsync: (() => void) | null = null;
      let rejectAsync: ((err: any) => void) | null = null;

      this._transform(chunk, encoding, (err?: Error | null, data?: TOutput) => {
        if (sync) {
          syncDone = true;
          syncErr = err ?? null;
          syncData = data;
          return;
        }
        if (err) {
          rejectAsync?.(err);
          return;
        }
        if (data !== undefined) {
          this.push(data);
        }
        resolveAsync?.();
      });

      sync = false;

      if (syncDone) {
        if (syncErr) {
          throw toError(syncErr);
        }
        if (syncData !== undefined) {
          this.push(syncData);
        }
        return undefined;
      }

      return new Promise<void>((resolve, reject) => {
        resolveAsync = resolve;
        rejectAsync = reject;
      });
    }

    // Node.js always invokes the user transform as (chunk, encoding, callback),
    // regardless of declared parameter count. Users may access the callback via
    // `arguments[2]` even when the function signature omits it.
    let sync = true;
    let syncDone = false;
    let syncErr: Error | null = null;
    let syncData: TOutput | undefined;
    let callbackFired = false;

    let resolveAsync: (() => void) | null = null;
    let rejectAsync: ((err: any) => void) | null = null;
    const promise = new Promise<void>((resolve, reject) => {
      resolveAsync = resolve;
      rejectAsync = reject;
    });

    (
      userTransform as (
        this: Transform<TInput, TOutput>,
        chunk: TInput,
        encoding: string,
        callback: (error?: Error | null, data?: TOutput) => void
      ) => any
    ).call(this, chunk, encoding, (err?: Error | null, data?: TOutput) => {
      if (callbackFired) {
        // Node.js throws ERR_MULTIPLE_CALLBACK on double invocation.
        const multiErr = new Error("Callback called multiple times") as Error & {
          code: string;
        };
        multiErr.code = "ERR_MULTIPLE_CALLBACK";
        this.destroy(multiErr);
        return;
      }
      callbackFired = true;
      if (sync) {
        syncDone = true;
        syncErr = err ?? null;
        syncData = data;
        return;
      }

      if (err) {
        rejectAsync?.(err);
        return;
      }
      if (data !== undefined) {
        this.push(data);
      }
      resolveAsync?.();
    });

    sync = false;

    if (syncDone) {
      // Callback was called synchronously — ignore return value (Node.js behavior).
      if (syncErr) {
        throw toError(syncErr);
      }
      if (syncData !== undefined) {
        this.push(syncData);
      }
      return undefined;
    }

    // Callback was NOT called synchronously.
    // Node.js ignores the return value of _transform — it always waits for the
    // callback. Match that behavior: wait for the async callback promise.
    return promise;
  }

  private async _runFlush(): Promise<void> {
    if (this._destroyed || this._errored) {
      return;
    }

    if (this._hasSubclassFlush()) {
      await new Promise<void>((resolve, reject) => {
        (this as any)._flush((err?: Error | null, data?: TOutput) => {
          if (err) {
            reject(err);
            return;
          }
          if (data !== undefined) {
            this.push(data);
          }
          resolve();
        });
      });
      return;
    }

    const userFlush = this._flushImpl;
    if (!userFlush) {
      return;
    }

    // Node.js always invokes flush as (callback), regardless of declared
    // parameter count. Node.js ignores the return value of _flush —
    // it always waits for the callback. Match that behavior.
    await new Promise<void>((resolve, reject) => {
      (
        userFlush as (
          this: Transform<TInput, TOutput>,
          callback: (error?: Error | null, data?: TOutput) => void
        ) => any
      ).call(this, (err?: Error | null, data?: TOutput) => {
        if (err) {
          reject(err);
          return;
        }
        if (data !== undefined) {
          this.push(data);
        }
        resolve();
      });
      // If callback is not called synchronously, the promise remains pending
      // until the user code calls it — matching Node.js behavior.
    });
  }

  /**
   * Override on() to lazily forward readable 'data' events.
   * Avoids starting flowing mode unless requested.
   */
  override on(event: string | symbol, listener: (...args: any[]) => void): this {
    // Register the listener FIRST so that when _readable.on("data") triggers
    // resume() and synchronously drains buffered data, the forwarding handler
    // can find the listener already in place on this Transform.
    super.on(event, listener);

    if (event === "data" && !this._dataForwardingSetup) {
      this._dataForwardingSetup = true;
      this._readable.on("data", chunk => {
        this.emit("data", chunk);
      });
    } else if (event === "readable") {
      // Node.js: adding a 'readable' listener sets readableFlowing to false
      this._readable._setReadableFlowing(false);
    }
    return this;
  }

  /**
   * Write to the writable side
   */
  write(chunk: TInput, callback?: (error?: Error | null) => void): boolean;
  write(chunk: TInput, encoding?: string, callback?: (error?: Error | null) => void): boolean;
  write(
    chunk: TInput,
    encodingOrCallback?: string | ((error?: Error | null) => void),
    callback?: (error?: Error | null) => void
  ): boolean {
    const encoding = typeof encodingOrCallback === "string" ? encodingOrCallback : undefined;
    const cb = typeof encodingOrCallback === "function" ? encodingOrCallback : callback;

    // Reject writes after end() — matches Node.js behavior.
    if (this._ended) {
      const err = new Error("write after end") as Error & { code: string };
      err.code = "ERR_STREAM_WRITE_AFTER_END";
      deferTask(() => this.emit("error", err));
      if (cb) {
        deferTask(() => cb(err));
      }
      return false;
    }

    return encoding !== undefined
      ? this._writable.write(chunk, encoding, cb)
      : this._writable.write(chunk, cb);
  }

  /**
   * End the transform stream.
   * Defers closing via _scheduleEnd to allow writes triggered during
   * 'data' callbacks to complete before the writable side is ended.
   */
  end(callback?: () => void): this;
  end(chunk: TInput, callback?: () => void): this;
  end(chunk: TInput, encoding?: string, callback?: () => void): this;
  end(
    chunkOrCallback?: TInput | (() => void),
    encodingOrCallback?: string | (() => void),
    callback?: () => void
  ): this {
    if (this._ended) {
      const {
        chunk,
        encoding,
        cb: endCb
      } = parseEndArgs<TInput>(chunkOrCallback, encodingOrCallback, callback);

      // If a chunk was provided, this is a write-after-end error (Node.js behavior).
      if (chunk !== undefined) {
        this.write(chunk, encoding, err => {
          (endCb as any)?.(err ?? null);
        });
        return this;
      }

      // If we've already finished, Node.js calls the callback with
      // ERR_STREAM_ALREADY_FINISHED (but does not emit an error event).
      if (this.writableFinished && endCb) {
        const err = new Error("Cannot call end after a stream was finished") as Error & {
          code: string;
        };
        err.code = "ERR_STREAM_ALREADY_FINISHED";
        deferTask(() => (endCb as any)(err));
      } else if (endCb) {
        // Redundant end() is a no-op; callback called with no error.
        deferTask(() => (endCb as any)(null));
      }
      return this;
    }

    const { chunk, encoding, cb } = parseEndArgs<TInput>(
      chunkOrCallback,
      encodingOrCallback,
      callback
    );

    if (cb) {
      this._endCallback = cb;
      this.once("finish", () => {
        const ecb = this._endCallback;
        if (ecb) {
          this._endCallback = null;
          ecb();
        }
      });
    }

    // Write the end-chunk BEFORE setting _ended so that synchronous writes
    // from data handlers (triggered during transform processing) are still
    // accepted — matching Node.js behaviour where writableEnded is false
    // during the transform callback for the end() chunk.
    if (chunk !== undefined) {
      // Propagate write errors through destroy (matching Node.js)
      const onWriteError = (err?: Error | null): void => {
        if (err && !this._destroyed) {
          this.destroy(err);
        }
      };
      if (encoding !== undefined) {
        this._writable.write(chunk, encoding, onWriteError);
      } else {
        this._writable.write(chunk, onWriteError as any);
      }
    }

    this._ended = true;
    // Track whether end() was called from synchronous user code (not from
    // inside a deferred callback like a pipe end-listener). This determines
    // whether _scheduleEnd can safely synchronize the final sequence.
    this._endCalledFromSync = !inDeferredContext();
    this._scheduleEnd();
    return this;
  }

  /**
   * Read from the transform stream.
   * Backpressure release is handled by _read() on the internal readable,
   * which is called automatically by the Readable when it needs more data.
   */
  read(size?: number): TOutput | null {
    return this._readable.read(size);
  }

  /**
   * Pipe readable side to destination.
   * Accepts any writable-like object (duck-typed, matching Node.js behavior).
   */
  pipe<W extends WritableLike>(destination: W, options?: { end?: boolean }): W {
    return this._readable.pipe(destination, options) as W;
  }

  /**
   * Unpipe from destination
   */
  unpipe(destination?: any): this {
    this._readable.unpipe(destination);
    return this;
  }

  /**
   * Check if paused
   */
  isPaused(): boolean {
    return this._readable.isPaused();
  }

  /**
   * Resume reading from the readable side
   */
  resume(): this {
    this._readable.resume();
    return this;
  }

  /**
   * Pause reading from the readable side
   */
  pause(): this {
    this._readable.pause();
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

    // Invalidate any pending _scheduleEnd
    this._endGeneration++;

    // Release any pending write callback held for backpressure.
    // Node.js invokes pending write callbacks with the destroy error.
    if (this._afterTransformCallback) {
      const cb = this._afterTransformCallback;
      this._afterTransformCallback = null;
      cb(error ?? null);
    }

    if (this._sideForwardingCleanup) {
      this._sideForwardingCleanup();
      this._sideForwardingCleanup = null;
    }

    const afterDestroy = (finalError?: Error | null): void => {
      // Node.js: _destroy's callback determines whether an error is emitted.
      // cb(null)/cb()/cb(undefined) all suppress the original error.
      // Only cb(new Error(...)) replaces and emits the error.
      const err = finalError || null;
      this._readable.destroy();
      this._writable.destroy();
      // Fire the pending end() callback before discarding it.
      // Node.js fires the end() callback on destroy (it listens on both
      // 'finish' and 'close'; destroy emits 'close').
      const endCb = this._endCallback;
      this._endCallback = null;
      this._closed = true;
      const doEmit = (): void => {
        if (endCb) {
          endCb();
        }
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
      Object.getPrototypeOf(this)._destroy !== Transform.prototype._destroy
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
    while (proto && proto !== Transform.prototype && proto !== Object.prototype) {
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
      if (this.closed) {
        settle();
      } else {
        this.once("close", settle);
      }
    });
  }

  /**
   * Get the underlying Web TransformStream (internal).
   * @internal
   */
  private _getWebStream(): TransformStream<TInput, TOutput> {
    if (this._webStream) {
      return this._webStream;
    }

    // Web Streams interop layer.
    const iterator = this[Symbol.asyncIterator]();

    const readable = new ReadableStream<TOutput>({
      pull: async controller => {
        const { done, value } = await iterator.next();
        if (done) {
          controller.close();
          return;
        }
        controller.enqueue(value);
      },
      cancel: reason => {
        this.destroy(reason instanceof Error ? reason : new Error(String(reason)));
      }
    });

    const writable = new WritableStream<TInput>({
      write: chunk =>
        new Promise<void>((resolve, reject) => {
          this.write(chunk, err => {
            if (err) {
              reject(err);
            } else {
              resolve();
            }
          });
        }),
      close: () =>
        new Promise<void>(resolve => {
          this.end(() => resolve());
        }),
      abort: reason => {
        this.destroy(reason instanceof Error ? reason : new Error(String(reason)));
      }
    });

    this._webStream = { readable, writable };
    return this._webStream;
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
    // Node.js: writableEnded becomes true immediately when Transform.end() is
    // called, even though the internal Writable's end() is deferred via
    // _scheduleEnd. We use _ended (set synchronously in end()) to match this.
    return this._ended || this._writable.writableEnded;
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

  get readableObjectMode(): boolean {
    return this._readable.readableObjectMode;
  }

  get readableFlowing(): boolean | null {
    return (this as any)._readable.readableFlowing;
  }

  get destroyed(): boolean {
    return this._destroyed;
  }

  set destroyed(val: boolean) {
    this._destroyed = val;
    // Propagate to internal streams so their state stays consistent with
    // the Transform — matches Node.js where destroy state is shared, and
    // mirrors the same propagation in Duplex.destroyed setter.
    this._readable.destroyed = val;
    this._writable.destroyed = val;
  }

  // =========================================================================
  // Delegated methods (Node.js Transform compatibility)
  // =========================================================================

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
   * Set encoding for the readable side
   */
  setEncoding(encoding: string): this {
    this._readable.setEncoding(encoding);
    return this;
  }

  /**
   * Set default encoding for the writable side
   */
  setDefaultEncoding(encoding: string): this {
    this._writable.setDefaultEncoding(encoding);
    return this;
  }

  /**
   * Put a chunk back at the front of the readable buffer
   */
  unshift(chunk: TOutput, encoding?: string): void {
    this._readable.unshift(chunk, encoding);
  }

  /**
   * Wrap a legacy stream
   */
  wrap(stream: any): this {
    this._readable.wrap(stream);
    return this;
  }

  /**
   * Create an async iterator with options
   */
  iterator(options?: { destroyOnReturn?: boolean }): AsyncIterableIterator<TOutput> {
    return this._readable.iterator(options);
  }

  // =========================================================================
  // Delegated getters (Node.js Transform compatibility)
  // =========================================================================

  get writableCorked(): number {
    return this._writable.writableCorked;
  }

  get writableNeedDrain(): boolean {
    return this._writable.writableNeedDrain;
  }

  get writableObjectMode(): boolean {
    return this._writable.writableObjectMode;
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

  get readableBuffer(): TOutput[] {
    return this._readable.readableBuffer;
  }

  get writableBuffer(): TInput[] {
    return this._writable.writableBuffer;
  }

  /**
   * Async iterator support
   */
  async *[Symbol.asyncIterator](): AsyncIterableIterator<TOutput> {
    yield* this._readable[Symbol.asyncIterator]();
  }

  // =============================================================================
  // Functional / Higher-order Methods (forwarded to readable side)
  // =============================================================================

  map<U>(
    fn: (data: TOutput, options: { signal: AbortSignal }) => U | Promise<U>,
    options?: { concurrency?: number; highWaterMark?: number; signal?: AbortSignal }
  ): Readable<U> {
    return this._readable.map(fn, options);
  }

  filter(
    fn: (data: TOutput, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; highWaterMark?: number; signal?: AbortSignal }
  ): Readable<TOutput> {
    return this._readable.filter(fn, options);
  }

  async forEach(
    fn: (data: TOutput, options: { signal: AbortSignal }) => void | Promise<void>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<undefined> {
    return this._readable.forEach(fn, options);
  }

  async toArray(options?: { signal?: AbortSignal }): Promise<TOutput[]> {
    return this._readable.toArray(options);
  }

  async some(
    fn: (data: TOutput, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    return this._readable.some(fn, options);
  }

  async find(
    fn: (data: TOutput, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<TOutput | undefined> {
    return this._readable.find(fn, options);
  }

  async every(
    fn: (data: TOutput, options: { signal: AbortSignal }) => boolean | Promise<boolean>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Promise<boolean> {
    return this._readable.every(fn, options);
  }

  flatMap<U>(
    fn: (
      data: TOutput,
      options: { signal: AbortSignal }
    ) => Iterable<U> | AsyncIterable<U> | Readable<U> | Promise<Iterable<U> | AsyncIterable<U>>,
    options?: { concurrency?: number; signal?: AbortSignal }
  ): Readable<U> {
    return this._readable.flatMap(fn, options);
  }

  drop(limit: number, options?: { signal?: AbortSignal }): Readable<TOutput> {
    return this._readable.drop(limit, options);
  }

  take(limit: number, options?: { signal?: AbortSignal }): Readable<TOutput> {
    return this._readable.take(limit, options);
  }

  async reduce<U = TOutput>(
    fn: (previous: U, data: TOutput, options: { signal: AbortSignal }) => U | Promise<U>,
    initial?: U,
    options?: { signal?: AbortSignal }
  ): Promise<U> {
    if (arguments.length >= 2) {
      return this._readable.reduce(fn, initial, options);
    }
    return this._readable.reduce(fn);
  }

  compose<U>(
    stream: WritableLike | ((source: AsyncIterable<TOutput>) => AsyncIterable<U>),
    options?: { signal?: AbortSignal }
  ): IDuplex<U, TOutput> {
    return this._readable.compose(stream, options);
  }

  // =========================================================================
  // Static Methods (Node.js compatibility)
  // =========================================================================

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
   * Create a Transform from various sources (delegates to Duplex.from).
   * Matches Node.js where Transform inherits static from() from Duplex.
   */
  static from<TIn = Uint8Array, TOut = Uint8Array>(
    source:
      | AsyncIterable<TIn>
      | Iterable<TIn>
      | {
          readable?: Readable<TIn>;
          writable?: Writable<TOut>;
        }
  ): Duplex<TIn, TOut> {
    return Duplex.from(source);
  }

  /**
   * Convert a Web TransformStream to Node.js Transform
   */
  static fromWeb<TIn = Uint8Array, TOut = Uint8Array>(
    webStream: TransformStream<TIn, TOut>,
    options?: DuplexStreamOptions
  ): Transform<TIn, TOut> {
    const transform = new Transform<TIn, TOut>(options);
    transform._webStream = webStream;

    // Replace internal streams with the ones from the web stream
    const newReadable = Readable.fromWeb(webStream.readable, { objectMode: options?.objectMode });
    const newWritable = Writable.fromWeb(webStream.writable, { objectMode: options?.objectMode });

    if (transform._sideForwardingCleanup) {
      transform._sideForwardingCleanup();
      transform._sideForwardingCleanup = null;
    }

    (transform as any)._readable = newReadable;
    (transform as any)._writable = newWritable;

    // Re-connect event forwarding (data forwarding remains lazy via Transform.on)
    transform._setupSideForwarding();

    return transform;
  }

  /**
   * Convert a Node.js Transform to Web TransformStream
   */
  static toWeb<TIn = Uint8Array, TOut = Uint8Array>(
    nodeStream: Transform<TIn, TOut>
  ): TransformStream<TIn, TOut> {
    return nodeStream._getWebStream();
  }

  // =========================================================================
  // Base Class Methods (for subclass override detection)
  // =========================================================================

  /**
   * Base transform method - can be overridden by subclasses.
   * Default behavior: throw ERR_METHOD_NOT_IMPLEMENTED (matches Node.js).
   * Node.js throws synchronously rather than calling callback(err).
   */
  _transform(
    _chunk: TInput,
    _encoding: string,
    _callback: (error?: Error | null, data?: TOutput) => void
  ): void {
    const err = new Error("The _transform() method is not implemented") as Error & { code: string };
    err.code = "ERR_METHOD_NOT_IMPLEMENTED";
    throw err;
  }

  /**
   * Base final method - matches Node.js Transform.prototype._final.
   * In Node.js this calls _flush (if defined), pushes null, and calls cb.
   * In our browser implementation, the actual final logic is handled by
   * the finalHandler passed to the internal Writable, so this method
   * exists primarily for API surface parity and subclass override detection.
   */
  _final(callback: (error?: Error | null) => void): void {
    if (typeof (this as any)._flush === "function" && !this.destroyed) {
      (this as any)._flush((err?: Error | null, data?: TOutput) => {
        if (err) {
          callback(err);
          return;
        }
        if (data != null) {
          this.push(data);
        }
        this.push(null);
        callback();
      });
    } else {
      this.push(null);
      callback();
    }
  }
}

// Node.js: `Transform.prototype.addListener === Transform.prototype.on` (same function).
// Transform overrides `on` from EventEmitter, so we must re-alias `addListener`.
Transform.prototype.addListener = Transform.prototype.on;

// Node.js: Transform.prototype._writev === null (inherited from Duplex/Writable chain).
// Browser Transform doesn't extend Duplex, so we set it explicitly.
(Transform.prototype as any)._writev = null;
