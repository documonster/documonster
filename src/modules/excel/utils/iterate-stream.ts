import { toError } from "@utils/errors";

type Listener = (...args: any[]) => void;

interface EventEmitterLike {
  on(event: string, listener: Listener): this;
  removeListener(event: string, listener: Listener): this;
}

/**
 * Minimal readable-stream shape consumed by {@link iterateStream}.
 * Intentionally structural so it matches Node `Readable`, zip entry streams,
 * object-mode streams that yield zip entries, and any third-party emitter
 * that raises `data`/`end`/`error`.
 */
export interface IterableStreamLike<T = unknown> extends EventEmitterLike {
  resume(): void;
  pause(): void;
  on(event: "data", listener: (chunk: T) => void): this;
  on(event: "end", listener: () => void): this;
  on(event: "error", listener: (err: Error) => void): this;
}

async function* iterateStream<T = Uint8Array | string>(
  stream: IterableStreamLike<T>
): AsyncGenerator<T> {
  const contents: T[] = [];
  let resolveDataPromise: (() => void) | null = null;

  const onData = (data: T) => {
    contents.push(data);
    if (resolveDataPromise) {
      resolveDataPromise();
      resolveDataPromise = null;
    }
  };
  stream.on("data", onData);

  let ended = false;
  const onEnd = () => {
    ended = true;
    if (resolveDataPromise) {
      resolveDataPromise();
      resolveDataPromise = null;
    }
  };
  stream.on("end", onEnd);

  let error: Error | false = false;
  const onError = (err: Error) => {
    error = err;
    if (resolveDataPromise) {
      resolveDataPromise();
      resolveDataPromise = null;
    }
  };
  stream.on("error", onError);

  try {
    while (!ended || contents.length > 0) {
      if (contents.length === 0) {
        stream.resume();
        await new Promise<void>(resolve => {
          resolveDataPromise = resolve;
        });
      } else {
        stream.pause();
        const data = contents.shift();
        if (data !== undefined) {
          yield data;
        }
      }
      if (error) {
        throw toError(error);
      }
    }
  } finally {
    // Clean up listeners
    stream.removeListener("data", onData);
    stream.removeListener("end", onEnd);
    stream.removeListener("error", onError);
  }
}

export { iterateStream };
