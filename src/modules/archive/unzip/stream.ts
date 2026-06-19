import zlib from "zlib";

import type { CrxHeader, ParseDriverState, ParseOptions } from "@archive/unzip/parser-core";
import { DATA_DESCRIPTOR_SIGNATURE_BYTES } from "@archive/unzip/parser-core";
import type {
  PullStreamPublicApi,
  InflateFactory,
  ParseEmitter,
  ParseIO,
  ZipEntry
} from "@archive/unzip/stream.base";
import {
  PullStream,
  runParseLoop,
  streamUntilValidatedDataDescriptor
} from "@archive/unzip/stream.base";
import type { Duplex, PassThrough, Transform } from "@stream";

/**
 * Creates an InflateRaw stream using Node.js native zlib.
 */
function createInflateRaw(): Transform {
  return zlib.createInflateRaw();
}

export type { CrxHeader, EntryProps, EntryVars, ParseOptions } from "@archive/unzip/parser-core";

export type { ZipEntry } from "@archive/unzip/stream.base";

const dataDescriptorSignature = DATA_DESCRIPTOR_SIGNATURE_BYTES;

export type ParseStream = Duplex & {
  promise(): Promise<void>;
} & PullStreamPublicApi & {
    crxHeader?: CrxHeader;
  };

export function createParseClass(createInflateRawFn: InflateFactory): {
  new (opts?: ParseOptions): ParseStream;
} {
  return class Parse extends PullStream<ZipEntry> {
    private _opts: ParseOptions;
    private _driverState: ParseDriverState = {};

    // ---------------------------------------------------------------
    // Parser completion — explicit deferred, independent of stream
    // lifecycle events (close / end). This avoids races between
    // push(null) and close that cause ERR_STREAM_PREMATURE_CLOSE in
    // Node.js's default Readable async iterator.
    // ---------------------------------------------------------------
    private _parserDone = false;
    private _parserError: Error | null = null;
    private _parserDeferred: {
      resolve: () => void;
      reject: (err: Error) => void;
    } | null = null;
    private _parserDonePromise: Promise<void> | null = null;

    // ---------------------------------------------------------------
    // Entry queue — custom [Symbol.asyncIterator] reads from here
    // instead of relying on Readable's default objectMode iterator
    // (which uses finished() internally and races with close).
    // ---------------------------------------------------------------
    private _entryQueue: ZipEntry[] = [];
    private _entryWaiter: {
      resolve: (result: IteratorResult<ZipEntry>) => void;
      reject: (err: unknown) => void;
    } | null = null;
    /** True once the parser has finished producing entries. */
    private _entriesDone = false;

    crxHeader?: CrxHeader;

    constructor(opts: ParseOptions = {}) {
      super(opts);
      this._opts = opts;

      // Always listen for error events to prevent Node.js from treating
      // them as uncaught exceptions. Route them to the parser deferred.
      this.on("error", (err: Error) => {
        this._rejectParser(err);
        this._closeEntryQueue(err);
      });

      const io: ParseIO = {
        pull: async (length: number) => this.pull(length),
        pullUntil: async (pattern: Uint8Array, includeEof?: boolean) =>
          this.pull(pattern, includeEof),
        stream: (length: number) => this.stream(length),
        streamUntilDataDescriptor: () => this._streamUntilValidatedDataDescriptor(),
        setDone: () => {
          this._maybeReleaseWriteCallback();
          this.end();
          this.push(null);
        }
      };

      const emitter: ParseEmitter = {
        emitEntry: (entry: ZipEntry) => {
          this.emit("entry", entry);
        },
        pushEntry: (entry: ZipEntry) => {
          // Feed the legacy Readable objectMode side (for pipe / data consumers).
          this.push(entry);
          // Also feed the custom entry queue (for our async iterator).
          this._enqueueEntry(entry);
        },
        pushEntryIfPiped: (entry: ZipEntry) => {
          const state = (this as any)._readableState;
          if (state.pipesCount || (state.pipes && state.pipes.length)) {
            this.push(entry);
          }
          // Always feed the entry queue regardless of pipe state.
          this._enqueueEntry(entry);
        },
        emitCrxHeader: header => {
          (this as any).crxHeader = header;
          this.emit("crx-header", header);
        },
        emitError: err => {
          this.__emittedError = err;
          this.emit("error", err);
        },
        emitClose: () => {
          this.emit("close");
        }
      };

      // Parse records as data arrives.
      runParseLoop(
        this._opts,
        io,
        emitter,
        createInflateRawFn,
        this._driverState,
        (data: Uint8Array) => zlib.inflateRawSync(data)
      ).then(
        () => {
          // If an error was emitted during parsing (e.g. invalid signature),
          // the parse loop returns normally but we should reject.
          if (this.__emittedError) {
            this._rejectParser(this.__emittedError);
            this._closeEntryQueue(this.__emittedError);
          } else {
            this._resolveParser();
            this._closeEntryQueue();
          }
        },
        (e: Error) => {
          if (!this.__emittedError || this.__emittedError !== e) {
            this.emit("error", e);
          }
          this._maybeReleaseWriteCallback();
          this._rejectParser(e);
          this._closeEntryQueue(e);
          this.emit("close");
        }
      );
    }

    // ---------------------------------------------------------------
    // Entry queue management
    // ---------------------------------------------------------------

    private _enqueueEntry(entry: ZipEntry): void {
      if (this._entryWaiter) {
        // A consumer is already waiting — deliver immediately.
        const { resolve } = this._entryWaiter;
        this._entryWaiter = null;
        resolve({ value: entry, done: false });
      } else {
        this._entryQueue.push(entry);
      }
    }

    private _closeEntryQueue(err?: Error): void {
      this._entriesDone = true;

      if (this._entryWaiter) {
        const waiter = this._entryWaiter;
        this._entryWaiter = null;
        if (err) {
          waiter.reject(err);
        } else {
          waiter.resolve({ value: undefined as any, done: true });
        }
      }
    }

    // ---------------------------------------------------------------
    // Custom async iterator — bypasses Node Readable's default
    // iterator which uses finished() and races with close.
    // ---------------------------------------------------------------

    // Override the default Readable async iterator with our custom entry-queue
    // based iterator. This avoids Node.js's Readable default iterator which uses
    // finished() internally and races with the close event.
    //
    // We cast through `any` because ES2024+ AsyncIterator requires
    // [Symbol.asyncDispose] which AsyncIterableIterator doesn't include,
    // and we don't need disposal semantics here.
    override [Symbol.asyncIterator](): any {
      const iterator = {
        next: (): Promise<IteratorResult<ZipEntry>> => {
          if (this._entryQueue.length > 0) {
            return Promise.resolve({ value: this._entryQueue.shift()!, done: false });
          }

          if (this._entriesDone) {
            if (this._parserError) {
              return Promise.reject(this._parserError);
            }
            return Promise.resolve({ value: undefined as any, done: true });
          }

          return new Promise<IteratorResult<ZipEntry>>((resolve, reject) => {
            this._entryWaiter = { resolve, reject };
          });
        },

        return: (): Promise<IteratorResult<ZipEntry>> => {
          this._entriesDone = true;
          this._entryQueue.length = 0;
          this._entryWaiter = null;
          return Promise.resolve({ value: undefined as any, done: true });
        },

        [Symbol.asyncIterator]() {
          return iterator;
        }
      };

      return iterator;
    }

    /**
     * Stream file data until we reach a DATA_DESCRIPTOR record boundary.
     */
    private _streamUntilValidatedDataDescriptor(): PassThrough {
      return streamUntilValidatedDataDescriptor({
        source: {
          getLength: () => this._queue.length,
          read: (length: number) => this._queue.read(length),
          peekChunks: (length: number) => this._queue.peekChunks(length),
          discard: (length: number) => this._queue.discard(length),
          indexOfPattern: (pattern: Uint8Array, startIndex: number) =>
            this._queue.indexOfPattern(pattern, startIndex),
          peekUint32LE: (offset: number) => this._queue.peekUint32LE(offset),
          isFinished: () => this.finished,
          onDataAvailable: (cb: () => void) => {
            this.on("chunk", cb);
            return () => this.removeListener("chunk", cb);
          },
          maybeReleaseWriteCallback: () => this._maybeReleaseWriteCallback()
        },
        dataDescriptorSignature
      });
    }

    // ---------------------------------------------------------------
    // Parser completion deferred
    // ---------------------------------------------------------------

    private _resolveParser(): void {
      if (this._parserDone) {
        return;
      }
      this._parserDone = true;
      if (this._parserDeferred) {
        const { resolve } = this._parserDeferred;
        this._parserDeferred = null;
        resolve();
      }
    }

    private _rejectParser(err: Error): void {
      if (this._parserDone) {
        return;
      }
      this._parserDone = true;
      this._parserError = err;
      if (this._parserDeferred) {
        const { reject } = this._parserDeferred;
        this._parserDeferred = null;
        reject(err);
      }
    }

    /**
     * Returns a promise that resolves when the parser has finished
     * processing all ZIP records, or rejects on parse error.
     *
     * This is driven by an internal deferred that is resolved/rejected
     * directly by the parse loop — it does NOT depend on stream
     * lifecycle events (close / end), avoiding the
     * ERR_STREAM_PREMATURE_CLOSE race.
     */
    promise(): Promise<void> {
      if (this._parserDone) {
        return this._parserError ? Promise.reject(this._parserError) : Promise.resolve();
      }

      if (this._parserDonePromise) {
        return this._parserDonePromise;
      }

      this._parserDonePromise = new Promise<void>((resolve, reject) => {
        this._parserDeferred = { resolve, reject };
      });
      return this._parserDonePromise;
    }
  };
}

const BaseParse = /* @__PURE__ */ createParseClass(createInflateRaw);

export class Parse extends BaseParse {}

export function createParse(opts?: ParseOptions): ParseStream {
  return new Parse(opts);
}
