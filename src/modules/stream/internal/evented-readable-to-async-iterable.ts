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

export interface EventedReadableLike<T> {
  on?(event: string, listener: (chunk: T) => void): any;
  off?(event: string, listener: (chunk: T) => void): any;
  removeListener?(event: string, listener: (chunk: T) => void): any;
  pause?(): any;
  resume?(): any;
}

export function eventedReadableToAsyncIterableNoDestroy<T>(
  stream: EventedReadableLike<T>
): AsyncIterable<T> {
  return {
    [Symbol.asyncIterator](): AsyncIterator<T> {
      const chunks: T[] = [];
      let head = 0;
      let done = false;
      let error: unknown = null;
      let cleanedUp = false;
      let pending: {
        resolve: (r: IteratorResult<T>) => void;
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

        if (typeof (stream as any)?.off === "function") {
          (stream as any).off("data", onData);
          (stream as any).off("end", onEnd);
          (stream as any).off("close", onClose);
          (stream as any).off("error", onError);
        } else if (typeof (stream as any)?.removeListener === "function") {
          (stream as any).removeListener("data", onData);
          (stream as any).removeListener("end", onEnd);
          (stream as any).removeListener("close", onClose);
          (stream as any).removeListener("error", onError);
        }

        if (typeof (stream as any)?.pause === "function") {
          (stream as any).pause();
        }
      };

      const onData = (chunk: T): void => {
        chunks.push(chunk);
        if (typeof (stream as any)?.pause === "function") {
          (stream as any).pause();
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
          resolve({ value: undefined as any, done: true });
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

      if (typeof (stream as any)?.pause === "function") {
        (stream as any).pause();
      }
      if (typeof (stream as any)?.on === "function") {
        (stream as any).on("data", onData);
        (stream as any).on("end", onEnd);
        (stream as any).on("close", onClose);
        (stream as any).on("error", onError);
      }

      return {
        next(): Promise<IteratorResult<T>> {
          if (error) {
            return Promise.reject(error);
          }
          if (head < chunks.length) {
            return Promise.resolve({ value: take(), done: false });
          }
          if (done) {
            return Promise.resolve({ value: undefined as any, done: true });
          }

          return new Promise((resolve, reject) => {
            pending = { resolve, reject };
            if (typeof (stream as any)?.resume === "function") {
              (stream as any).resume();
            }
          });
        },
        return(): Promise<IteratorResult<T>> {
          done = true;
          cleanup();
          if (pending) {
            const { resolve } = pending;
            pending = null;
            resolve({ value: undefined as any, done: true });
          }
          return Promise.resolve({ value: undefined as any, done: true });
        },
        throw(e?: unknown): Promise<IteratorResult<T>> {
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
