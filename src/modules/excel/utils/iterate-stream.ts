import { toError } from "@utils/errors";

// Structural listener signature. The `any[]` is required for interop: Node's
// built-in stream types (Readable/ReadStream, ZipEntry, ParseStream, …) declare
// `removeListener(event, (...args: any[]) => void)`, and assigning those streams
// to `IterableStreamLike` checks this parameter contravariantly — `unknown[]`
// or `never[]` would reject `any[]` and break the interop. Keep as `any[]`.
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
 *
 * The default chunk type is `Uint8Array | string` — the overwhelmingly common
 * byte/text stream case — so that `Parameters<typeof iterateStream>[0]` and
 * other unparameterised references resolve to a concrete chunk type instead of
 * `unknown`. Object-mode streams (e.g. a zip parser yielding entries) supply
 * their own concrete `T` via the stream's own type, so they are unaffected.
 */
export interface IterableStreamLike<T = Uint8Array | string> extends EventEmitterLike {
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
