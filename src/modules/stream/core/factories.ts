/**
 * Stream factory functions shared by the Node and browser entry points.
 *
 * The bodies are identical on both platforms; only the stream classes differ,
 * so each entry point builds them with its own.
 */

import { BufferedStream } from "@stream/buffered-stream";
import { isAsyncIterable } from "@stream/core/type-guards";
import { StreamArgumentError } from "@stream/errors";
import { PullStream } from "@stream/pull-stream";
import type {
  BufferedStreamOptions,
  DuplexStreamOptions,
  IPassThrough,
  IPushableDuplex,
  IPushableReadable,
  IReadable,
  ITransform,
  IWritable,
  PullStreamOptions,
  ReadableStreamOptions,
  TransformStreamOptions,
  WritableStreamOptions,
  WriteCallback
} from "@stream/types";
import { stringToUint8Array } from "@utils/binary";

type Ctor = new (options?: any) => any;

interface StreamClasses {
  Readable: Ctor & { from(iterable: unknown, options?: unknown): any };
  Writable: Ctor;
  Duplex: Ctor;
  Transform: Ctor;
  PassThrough: Ctor;
}

/** What a byte-mode stream accepts as a chunk; it reads back as `Uint8Array`. */
export type StreamByteChunk = string | ArrayBufferView;

/** Options that leave the factory's default object mode on: chunks pass through as given. */
export type StreamObjectModeOptions = ReadableStreamOptions & { objectMode?: true };

/** Options that switch to byte mode: every chunk reads back as a `Uint8Array`. */
export type StreamByteModeOptions = ReadableStreamOptions & { objectMode: false };

/**
 * Options whose mode may be either — a `boolean`, or options held in a
 * variable. The mode may turn out to be bytes, so the input must be what byte
 * mode carries, and the stream yields either the chunks themselves or bytes.
 */
type EitherModeOptions = ReadableStreamOptions;

/** A `destroy` hook, with the stream it destroys as `this`. */
export type StreamDestroyHook<S> = (
  this: S,
  error: Error | null,
  callback: (error?: Error | null) => void
) => void;

/**
 * The factories, typed by what each one actually yields.
 *
 * Where a factory defaults to object mode, `objectMode: false` is its own
 * overload: byte mode turns every chunk into a `Uint8Array`, so the element
 * type is not the input's. An option whose mode is only known at run time gets
 * the union of both. Hooks see the stream as `this`: a producer's `push` takes
 * the stream's element type, and a writer's stream is the writable itself.
 */
export interface StreamFactories {
  createPullStream(options?: PullStreamOptions): PullStream;
  createBufferedStream(options?: BufferedStreamOptions): BufferedStream;
  createReadable<T = Uint8Array>(
    options?: ReadableStreamOptions & {
      read?: (this: IPushableReadable<T>, size: number) => void;
      destroy?: StreamDestroyHook<IPushableReadable<T>>;
    }
  ): IPushableReadable<T>;
  /** Values are awaited, as `Readable.from` awaits them. */
  createReadableFromAsyncIterable<T>(
    iterable: AsyncIterable<T>,
    options?: StreamObjectModeOptions
  ): IReadable<Awaited<T>>;
  createReadableFromAsyncIterable(
    iterable: AsyncIterable<StreamByteChunk>,
    options: StreamByteModeOptions
  ): IReadable<Uint8Array>;
  createReadableFromAsyncIterable<T extends StreamByteChunk>(
    iterable: AsyncIterable<T>,
    options?: EitherModeOptions
  ): IReadable<Awaited<T> | Uint8Array>;
  /** `undefined` entries are skipped. */
  createReadableFromArray<T>(
    data: readonly T[],
    options?: StreamObjectModeOptions
  ): IReadable<Exclude<T, undefined>>;
  createReadableFromArray(
    data: readonly StreamByteChunk[],
    options: StreamByteModeOptions
  ): IReadable<Uint8Array>;
  createReadableFromArray<T extends StreamByteChunk>(
    data: readonly T[],
    options?: EitherModeOptions
  ): IReadable<Exclude<T, undefined> | Uint8Array>;
  createReadableFromGenerator<T>(
    generator: () => AsyncGenerator<T, void, unknown>,
    options?: StreamObjectModeOptions
  ): IReadable<Awaited<T>>;
  createReadableFromGenerator(
    generator: () => AsyncGenerator<StreamByteChunk, void, unknown>,
    options: StreamByteModeOptions
  ): IReadable<Uint8Array>;
  createReadableFromGenerator<T extends StreamByteChunk>(
    generator: () => AsyncGenerator<T, void, unknown>,
    options?: EitherModeOptions
  ): IReadable<Awaited<T> | Uint8Array>;
  /** Any thenable, adopted as `await` would; the stream yields what it resolves to. */
  createReadableFromPromise<T>(
    promise: PromiseLike<T>,
    options?: StreamObjectModeOptions
  ): IReadable<Awaited<T>>;
  createReadableFromPromise(
    promise: PromiseLike<StreamByteChunk>,
    options: StreamByteModeOptions
  ): IReadable<Uint8Array>;
  createReadableFromPromise<T extends StreamByteChunk>(
    promise: PromiseLike<T>,
    options?: EitherModeOptions
  ): IReadable<Awaited<T> | Uint8Array>;
  createEmptyReadable<T = Uint8Array>(options?: ReadableStreamOptions): IReadable<T>;
  createWritable<T = Uint8Array>(
    options?: WritableStreamOptions & {
      write?: (this: IWritable<T>, chunk: T, encoding: string, callback: WriteCallback) => void;
      final?: (this: IWritable<T>, callback: WriteCallback) => void;
      destroy?: StreamDestroyHook<IWritable<T>>;
    }
  ): IWritable<T>;
  createNullWritable<T = unknown>(options?: WritableStreamOptions): IWritable<T>;
  /**
   * A Transform from a function returning each output, or a thenable of it;
   * `undefined` pushes nothing.
   */
  createTransform<TInput = Uint8Array, TOutput = Uint8Array>(
    transformFn: (
      chunk: TInput,
      encoding?: string
    ) => TOutput | undefined | PromiseLike<TOutput | undefined>,
    options?: TransformStreamOptions & {
      flush?: () => TOutput | undefined | PromiseLike<TOutput | undefined>;
    }
  ): ITransform<TInput, TOutput>;
  createPassThrough<T = unknown>(options?: TransformStreamOptions): IPassThrough<T>;
  /** A byte stream of exactly `data`. */
  fromBytes(data: Uint8Array): IReadable<Uint8Array>;
  /** A byte stream of `str`, UTF-8 encoded. */
  fromString(str: string): IReadable<Uint8Array>;
  /** A byte stream of `data` as JSON text. */
  fromJSON(data: unknown): IReadable<Uint8Array>;
  /** A Transform from a function returning each output, or a thenable of it. */
  transform<TIn = Uint8Array, TOut = TIn>(
    fn: (chunk: TIn) => TOut | PromiseLike<TOut>
  ): ITransform<TIn, TOut>;
  /** An object-mode Transform keeping the chunks `predicate` accepts. */
  filter<T>(predicate: (chunk: T) => boolean | PromiseLike<boolean>): ITransform<T, T>;
  createDuplex<TRead = Uint8Array, TWrite = Uint8Array>(
    options?: DuplexStreamOptions & {
      allowHalfOpen?: boolean;
      read?: (this: IPushableDuplex<TRead, TWrite>, size: number) => void;
      write?: (
        this: IPushableDuplex<TRead, TWrite>,
        chunk: TWrite,
        encoding: string,
        callback: WriteCallback
      ) => void;
      final?: (this: IPushableDuplex<TRead, TWrite>, callback: WriteCallback) => void;
      destroy?: StreamDestroyHook<IPushableDuplex<TRead, TWrite>>;
    }
  ): IPushableDuplex<TRead, TWrite>;
}

/**
 * Deliver a value that may be a promise to a Node-style callback; `undefined`
 * means "nothing to push".
 *
 * A promise is recognised as `await` recognises one — anything with a `then`
 * method — not with `instanceof Promise`, which is false for a promise from
 * another realm (an iframe, a `vm` context) and for any other thenable, and
 * pushed the promise itself as the chunk. `Promise.resolve` then adopts it as
 * `await` would, so a thenable that calls back twice or throws still settles
 * once. A plain value keeps the synchronous path.
 */
function settle<T>(
  run: () => T | undefined | PromiseLike<T | undefined>,
  callback: (error?: Error | null, data?: T) => void
): void {
  let result: T | undefined | PromiseLike<T | undefined>;
  try {
    result = run();
  } catch (err) {
    callback(err as Error);
    return;
  }
  if (typeof (result as { then?: unknown } | null | undefined)?.then === "function") {
    Promise.resolve(result as PromiseLike<T | undefined>).then(
      data => (data !== undefined ? callback(null, data) : callback()),
      err => callback(err as Error)
    );
  } else if (result !== undefined) {
    callback(null, result as T);
  } else {
    callback();
  }
}

function emptyRead(this: { push(chunk: unknown): boolean }): void {
  this.push(null);
}

function nullWrite(_chunk: unknown, _encoding: string, callback: WriteCallback): void {
  callback();
}

/** `createPullStream`; platform-independent. */
export function createPullStream(options?: PullStreamOptions): PullStream {
  return new PullStream(options);
}

/** `createBufferedStream`; platform-independent. */
export function createBufferedStream(options?: BufferedStreamOptions): BufferedStream {
  return new BufferedStream(options);
}

// One maker per factory, each called once per platform with the one class it
// needs. A shared object of every class would make each factory keep all of
// them — the whole Readable/compose kernel behind a lone `createWritable`.

export function makeCreateReadable(
  Readable: StreamClasses["Readable"]
): StreamFactories["createReadable"] {
  return options => new Readable(options);
}

export function makeCreateReadableFromAsyncIterable(
  Readable: StreamClasses["Readable"]
): StreamFactories["createReadableFromAsyncIterable"] {
  return (iterable, options) => {
    // What `Readable.from` takes: an async or a plain iterable.
    if (
      !isAsyncIterable(iterable) &&
      typeof (iterable as { [Symbol.iterator]?: unknown } | null)?.[Symbol.iterator] !==
        "function" &&
      typeof iterable !== "string"
    ) {
      throw new StreamArgumentError("an async iterable", describeValue(iterable), {
        field: "iterable",
        where: "Streams.createReadableFromAsyncIterable"
      });
    }
    return Readable.from(iterable, { ...options, objectMode: options?.objectMode ?? true });
  };
}

/**
 * `undefined` entries are skipped: a stream's end-of-data signal is a nullish
 * push, so there is no in-band way to carry an undefined chunk, while `0`, `""`
 * and `false` are ordinary values that must survive.
 */
export function makeCreateReadableFromArray(
  Readable: StreamClasses["Readable"]
): StreamFactories["createReadableFromArray"] {
  return <T>(data: T[], options?: ReadableStreamOptions): IReadable<T> => {
    if (!Array.isArray(data)) {
      throw new StreamArgumentError("an array", typeof data, {
        field: "data",
        where: "Streams.createReadableFromArray"
      });
    }
    // A pushed null ends a stream, so it cannot be carried as a chunk: refused here rather
    // than ending the stream there and dropping everything after it.
    const nullAt = data.indexOf(null as T);
    if (nullAt >= 0) {
      throw new StreamArgumentError(
        "entries other than null (null ends a stream)",
        `null at index ${nullAt}`,
        { field: `data[${nullAt}]`, where: "Streams.createReadableFromArray" }
      );
    }
    let index = 0;
    return new Readable({
      ...options,
      objectMode: options?.objectMode ?? true,
      read(this: { push(chunk: unknown): boolean }) {
        while (index < data.length) {
          const chunk = data[index++];
          if (chunk === undefined) {
            continue;
          }
          if (!this.push(chunk)) {
            return;
          }
        }
        this.push(null);
      }
    });
  };
}

export function makeCreateReadableFromGenerator(
  Readable: StreamClasses["Readable"]
): StreamFactories["createReadableFromGenerator"] {
  return (generator, options) => {
    checkFunction("createReadableFromGenerator", "generator", generator);
    return Readable.from(generator(), { ...options, objectMode: options?.objectMode ?? true });
  };
}

export function makeCreateReadableFromPromise(
  Readable: StreamClasses["Readable"]
): StreamFactories["createReadableFromPromise"] {
  return <T>(promise: Promise<T>, options?: ReadableStreamOptions): IReadable<T> => {
    if (typeof (promise as { then?: unknown } | null)?.then !== "function") {
      throw new StreamArgumentError("a promise or thenable", describeValue(promise), {
        field: "promise",
        where: "Streams.createReadableFromPromise"
      });
    }
    const readable = new Readable({
      ...options,
      objectMode: options?.objectMode ?? true,
      read() {}
    });
    // Adopted as `await` would: a promise from another realm or any thenable
    // works, one that calls back twice settles once, and a `then` that throws
    // fails the stream rather than the call.
    Promise.resolve()
      .then(() => promise)
      .then(
        value => {
          readable.push(value);
          readable.push(null);
        },
        err => readable.destroy(err)
      );
    return readable;
  };
}

export function makeCreateEmptyReadable(
  Readable: StreamClasses["Readable"]
): StreamFactories["createEmptyReadable"] {
  return options => new Readable({ ...options, read: emptyRead });
}

export function makeCreateWritable(
  Writable: StreamClasses["Writable"]
): StreamFactories["createWritable"] {
  return options => new Writable(options);
}

export function makeCreateNullWritable(
  Writable: StreamClasses["Writable"]
): StreamFactories["createNullWritable"] {
  return options => new Writable({ ...options, write: nullWrite });
}

/** The shorthand `transformFn` returns its result; `undefined` pushes nothing. */
export function makeCreateTransform(
  Transform: StreamClasses["Transform"]
): StreamFactories["createTransform"] {
  return <TInput, TOutput>(
    transformFn: (
      chunk: TInput,
      encoding?: string
    ) => TOutput | undefined | PromiseLike<TOutput | undefined>,
    options?: TransformStreamOptions & {
      flush?: () => TOutput | undefined | PromiseLike<TOutput | undefined>;
    }
  ): ITransform<TInput, TOutput> => {
    checkFunction("createTransform", "transformFn", transformFn);
    const flush = options?.flush;
    return new Transform({
      ...options,
      transform(
        chunk: TInput,
        encoding: string,
        callback: (e?: Error | null, d?: TOutput) => void
      ) {
        settle(() => transformFn(chunk, encoding), callback);
      },
      flush: flush
        ? (callback: (e?: Error | null, d?: TOutput) => void) => settle(flush, callback)
        : undefined
    });
  };
}

export function makeCreatePassThrough(
  PassThrough: StreamClasses["PassThrough"]
): StreamFactories["createPassThrough"] {
  return options => new PassThrough(options);
}

export function makeCreateDuplex(Duplex: StreamClasses["Duplex"]): StreamFactories["createDuplex"] {
  return options => new Duplex(options);
}

/** How a refused argument is named: its type, or "null". */
function describeValue(value: unknown): string {
  return value === null ? "null" : typeof value;
}

function checkFunction(where: string, field: string, value: unknown): void {
  if (typeof value !== "function") {
    throw new StreamArgumentError("a function", describeValue(value), {
      field,
      where: `Streams.${where}`
    });
  }
}

/** A byte array, recognised structurally so one from another realm is accepted too. */
function isByteArray(value: unknown): value is Uint8Array {
  return (
    ArrayBuffer.isView(value) && (value as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  );
}

export function makeFromBytes(Readable: StreamClasses["Readable"]): StreamFactories["fromBytes"] {
  return data => {
    if (!isByteArray(data)) {
      throw new StreamArgumentError("a Uint8Array", data === null ? "null" : typeof data, {
        field: "data",
        where: "Streams.fromBytes"
      });
    }
    return new Readable({
      read(this: { push(chunk: unknown): boolean }) {
        this.push(data);
        this.push(null);
      }
    });
  };
}

export function makeFromString(Readable: StreamClasses["Readable"]): StreamFactories["fromString"] {
  const fromBytes = makeFromBytes(Readable);
  return str => {
    if (typeof str !== "string") {
      throw new StreamArgumentError("a string", str === null ? "null" : typeof str, {
        field: "str",
        where: "Streams.fromString"
      });
    }
    return fromBytes(stringToUint8Array(str));
  };
}

export function makeFromJSON(Readable: StreamClasses["Readable"]): StreamFactories["fromJSON"] {
  const fromString = makeFromString(Readable);
  return data => {
    let json: string | undefined;
    try {
      json = JSON.stringify(data);
    } catch (cause) {
      throw new StreamArgumentError(
        "JSON-serialisable data",
        `${describeValue(data)} (${(cause as Error).message})`,
        { cause, field: "data", where: "Streams.fromJson" }
      );
    }
    if (json === undefined) {
      // JSON.stringify answers `undefined` for a function, a symbol and `undefined` itself.
      throw new StreamArgumentError("JSON-serialisable data", describeValue(data), {
        field: "data",
        where: "Streams.fromJson"
      });
    }
    return fromString(json);
  };
}

export function makeTransform(Transform: StreamClasses["Transform"]): StreamFactories["transform"] {
  const createTransform = makeCreateTransform(Transform);
  return fn => createTransform(fn);
}

export function makeFilter(Transform: StreamClasses["Transform"]): StreamFactories["filter"] {
  const createTransform = makeCreateTransform(Transform);
  return <T>(predicate: (chunk: T) => boolean | Promise<boolean>) => {
    checkFunction("filter", "predicate", predicate);
    return createTransform<T, T>(async chunk => ((await predicate(chunk)) ? chunk : undefined), {
      objectMode: true
    });
  };
}
