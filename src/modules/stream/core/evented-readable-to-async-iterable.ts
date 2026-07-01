/**
 * Convert an evented, pause/resume-capable readable into an AsyncIterable without
 * destroying/ending the underlying stream when iteration stops early.
 *
 * Why:
 * - Some Node-style readables (and stream-like objects) implement AsyncIterable
 *   by destroying the stream when the consumer breaks early.
 * - In some scenarios (e.g. entry streaming), we want best-effort cleanup
 *   without implicitly destroying the underlying stream.
 */

/**
 * Listener accepted by {@link EventedReadableLike}. The events consumed here
 * carry differing payloads (`data` → chunk, `end`/`close` → none, `error` →
 * error), so the listener is the union of those concrete handler shapes rather
 * than a single chunk signature.
 */
type EventedListener<T> = ((chunk: T) => void) | (() => void) | ((error: unknown) => void);

export interface EventedReadableLike<T> {
  on?(event: string, listener: EventedListener<T>): unknown;
  off?(event: string, listener: EventedListener<T>): unknown;
  removeListener?(event: string, listener: EventedListener<T>): unknown;
  pause?(): unknown;
  resume?(): unknown;
}

export function eventedReadableToAsyncIterableNoDestroy<T>(
  stream: EventedReadableLike<T>
): AsyncIterable<T> {
  return {
    [Symbol.asyncIterator](): AsyncIterator<T, undefined> {
      const chunks: T[] = [];
      let head = 0;
      let done = false;
      let error: unknown = null;
      let cleanedUp = false;
      let pending: {
        resolve: (r: IteratorResult<T, undefined>) => void;
        reject: (e: unknown) => void;
      } | null = null;

      const take = (): T => {
        const chunk = chunks[head++]!;
        // Periodically compact to avoid unbounded head growth.
        if (head > 64 && head * 2 > chunks.length) {
          chunks.splice(0, head);
          head = 0;
        }
        return chunk;
      };

      const cleanup = (): void => {
        if (cleanedUp) {
          return;
        }
        cleanedUp = true;

        if (typeof stream.off === "function") {
          stream.off("data", onData);
          stream.off("end", onEnd);
          stream.off("close", onClose);
          stream.off("error", onError);
        } else if (typeof stream.removeListener === "function") {
          stream.removeListener("data", onData);
          stream.removeListener("end", onEnd);
          stream.removeListener("close", onClose);
          stream.removeListener("error", onError);
        }

        if (typeof stream.pause === "function") {
          stream.pause();
        }
      };

      const onData = (chunk: T): void => {
        chunks.push(chunk);
        if (typeof stream.pause === "function") {
          stream.pause();
        }
        if (pending) {
          const { resolve } = pending;
          pending = null;
          resolve({ value: take(), done: false });
        }
      };

      const onEnd = (): void => {
        done = true;
        cleanup();
        if (pending) {
          const { resolve } = pending;
          pending = null;
          resolve({ value: undefined, done: true });
        }
      };

      const onClose = (): void => {
        onEnd();
      };

      const onError = (e: unknown): void => {
        error = e;
        done = true;
        cleanup();
        if (pending) {
          const { reject } = pending;
          pending = null;
          reject(e);
        }
      };

      if (typeof stream.pause === "function") {
        stream.pause();
      }
      if (typeof stream.on === "function") {
        stream.on("data", onData);
        stream.on("end", onEnd);
        stream.on("close", onClose);
        stream.on("error", onError);
      }

      return {
        next(): Promise<IteratorResult<T, undefined>> {
          if (error) {
            return Promise.reject(error);
          }
          if (head < chunks.length) {
            return Promise.resolve({ value: take(), done: false });
          }
          if (done) {
            return Promise.resolve({ value: undefined, done: true });
          }

          return new Promise((resolve, reject) => {
            pending = { resolve, reject };
            if (typeof stream.resume === "function") {
              stream.resume();
            }
          });
        },
        return(): Promise<IteratorResult<T, undefined>> {
          done = true;
          cleanup();
          if (pending) {
            const { resolve } = pending;
            pending = null;
            resolve({ value: undefined, done: true });
          }
          return Promise.resolve({ value: undefined, done: true });
        },
        throw(e?: unknown): Promise<IteratorResult<T, undefined>> {
          done = true;
          cleanup();
          if (pending) {
            const { reject } = pending;
            pending = null;
            reject(e);
          }
          return Promise.reject(e);
        }
      };
    }
  };
}
