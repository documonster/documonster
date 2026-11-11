/**
 * Reading an {@link XmlByteSource} chunk by chunk, for `Xml.readRecords` and `Xml.readRoot`: one
 * `Uint8Array`, or an iterable or async iterable of them. A read that stops early closes the source
 * (its iterator's `return()`), so a file handle or a network body is released, and an abort signal
 * interrupts whatever the read is waiting for — the source's next chunk, or a promise the caller's
 * callback returned — rather than waiting for it. Aborting rejects with this package's
 * `AbortError`, the signal's reason as its `cause`.
 */

import { isUint8Array } from "@utils/binary";
import { createAbortError, throwIfAborted } from "@utils/errors";
import type { CheckedByteSource } from "@xml/arguments";
import { XmlArgumentError } from "@xml/errors";

type ChunkIterator = Iterator<Uint8Array> | AsyncIterator<Uint8Array>;

export class ByteReader {
  private iterator: ChunkIterator | undefined;
  /** The iterator's `next`, taken once as the protocol does. */
  private nextMethod: ((this: ChunkIterator) => unknown) | undefined;
  private single: Uint8Array | undefined;
  private done = false;
  /** The one listener on the signal, and how it fails the wait in progress. */
  private onAbort: (() => void) | undefined;
  private failWait: ((reason: unknown) => void) | undefined;

  constructor(
    source: CheckedByteSource,
    private readonly where: string,
    private readonly signal: AbortSignal | undefined
  ) {
    // A signal that has fired already: the source is not even opened.
    throwIfAborted(signal);
    if ("bytes" in source) {
      this.single = source.bytes;
    } else {
      const iterator: unknown = source.open.call(source.source);
      let next: unknown;
      try {
        next = isObject(iterator) ? (iterator as { next?: unknown }).next : undefined;
        if (typeof next !== "function") {
          throw new XmlArgumentError(
            "source",
            `${where}: the source must give an iterator with a next() method`
          );
        }
      } catch (err) {
        // Opened, and so possibly holding something — a file handle, say: released before the
        // failure is reported, as a read that fails later releases it.
        releaseQuietly(iterator);
        throw err;
      }
      this.iterator = iterator as ChunkIterator;
      this.nextMethod = next as (this: ChunkIterator) => unknown;
    }
  }

  /** Throw an `AbortError` when the signal has fired. */
  checkAborted(): void {
    throwIfAborted(this.signal);
  }

  /**
   * What `start` returns, awaited — or an `AbortError` as soon as the signal fires, whether before
   * `start` is called, while it runs, or while what it returned is pending. The read waits for one
   * thing at a time, so one listener on the signal, added on the first wait and removed on
   * {@link close}, serves every wait: it fails the wait in progress. The wait is in place before
   * `start` runs, so a signal `start` fires itself fails it too, and what `start` began is taken
   * over at once: its failure after an abort is nobody's to report, not an unhandled rejection.
   */
  async abortable<T>(start: () => T | PromiseLike<T>): Promise<T> {
    const signal = this.signal;
    if (signal === undefined) {
      return start();
    }
    this.checkAborted();
    if (this.onAbort === undefined) {
      this.onAbort = () => this.failWait?.(createAbortError(signal.reason));
      signal.addEventListener("abort", this.onAbort, { once: true });
    }
    try {
      return await new Promise<T>((resolve, reject) => {
        this.failWait = reject;
        Promise.resolve(start()).then(resolve, reject);
      });
    } finally {
      this.failWait = undefined;
    }
  }

  /** The next chunk, or `undefined` at the end of the source. */
  async next(): Promise<Uint8Array | undefined> {
    this.checkAborted();
    if (this.done) {
      return undefined;
    }
    const iterator = this.iterator;
    if (iterator === undefined) {
      this.done = true;
      const single = this.single;
      this.single = undefined;
      return single;
    }
    const next = this.nextMethod!;
    const result: unknown = await this.abortable(() => next.call(iterator));
    if (!isObject(result)) {
      throw new XmlArgumentError(
        "source",
        `${this.where}: the source's iterator must give { done, value } results, got ${result === null ? "null" : typeof result}`
      );
    }
    if ((result as IteratorResult<unknown>).done) {
      this.done = true;
      return undefined;
    }
    const chunk: unknown = (result as IteratorResult<unknown>).value;
    if (!isUint8Array(chunk)) {
      throw new XmlArgumentError(
        "source",
        `${this.where}: source must yield Uint8Array chunks, got ${chunk === null ? "null" : typeof chunk}`
      );
    }
    return chunk;
  }

  /**
   * Release the source: its iterator's `return()`, unless it has ended.
   *
   * After a read that succeeded, the release is awaited — an abort still interrupts it — and its
   * failure is the read's. After a read that failed, it is asked for and not awaited: the reason the
   * read failed is what is reported, and a source that has gone quiet — an async generator finishes
   * `return()` only once the `next()` it is in has settled, which may be never — cannot keep the
   * failure from being reported.
   */
  async close(failed: boolean): Promise<void> {
    if (this.done) {
      this.removeListener();
      return;
    }
    this.done = true;
    this.single = undefined;
    const iterator = this.iterator;
    try {
      if (iterator === undefined) {
        return;
      }
      // Asked for whatever happens next: an abort may stop the wait for it, never the request
      // itself. Looking `return` up runs the source's code too, so it fails as calling it does.
      let released: Promise<unknown>;
      let release: ((this: ChunkIterator) => unknown) | undefined;
      try {
        release = iterator.return as ((this: ChunkIterator) => unknown) | undefined;
        if (release != null && typeof release !== "function") {
          // The protocol has a `return` that is there must be a method: one that is not is the source's mistake.
          throw new XmlArgumentError(
            "source",
            `${this.where}: the source's iterator return must be a method, got ${typeof release}`
          );
        }
        released = Promise.resolve(release == null ? undefined : release.call(iterator));
      } catch (err) {
        released = Promise.reject(err);
      }
      // Nobody may be left to wait for it — the read failed, or the signal fires before the wait.
      released.catch(() => undefined);
      if (!failed) {
        const result = await this.abortable(() => released);
        // A `return()` that is there answers with an iterator result, as the protocol has it.
        if (release != null && !isObject(result)) {
          throw new XmlArgumentError(
            "source",
            `${this.where}: the source's iterator return() must give an object, got ${result === null ? "null" : typeof result}`
          );
        }
      }
    } finally {
      this.removeListener();
    }
  }

  private removeListener(): void {
    if (this.onAbort !== undefined) {
      this.signal!.removeEventListener("abort", this.onAbort);
      this.onAbort = undefined;
    }
  }
}

/** Ask `iterator` to release what it holds, if it can, for a read that failed: what that does is not reported. */
function releaseQuietly(iterator: unknown): void {
  if (!isObject(iterator)) {
    return;
  }
  try {
    const release: unknown = (iterator as { return?: unknown }).return;
    if (typeof release === "function") {
      Promise.resolve(release.call(iterator)).catch(() => undefined);
    }
  } catch {
    // The failure that is reported is the one before.
  }
}

/** An object in the language's sense — a function is one — as the iterator protocol takes it. */
function isObject(value: unknown): value is object {
  return (typeof value === "object" && value !== null) || typeof value === "function";
}
