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
  on(event: "close", listener: () => void): this;
  /** Node streams' failure, if any — read so a stream that failed before it was iterated is not waited on. */
  readonly errored?: unknown;
  readonly destroyed?: boolean;
  readonly readableEnded?: boolean;
}

async function* iterateStream<T = Uint8Array | string>(
  stream: IterableStreamLike<T>
): AsyncGenerator<T> {
  const contents: T[] = [];
  let resolveDataPromise: (() => void) | null = null;
  const wake = () => {
    if (resolveDataPromise) {
      resolveDataPromise();
      resolveDataPromise = null;
    }
  };

  const onData = (data: T) => {
    contents.push(data);
    wake();
  };
  stream.on("data", onData);

  let ended = false;
  const onEnd = () => {
    ended = true;
    wake();
  };
  stream.on("end", onEnd);

  // A stream that failed before it was iterated has already emitted its `error`; one destroyed before it ended never
  // will emit `end`. Either way, waiting for an event would wait forever.
  let error: Error | false = stream.errored ? toError(stream.errored) : false;
  const onError = (err: Error) => {
    error = err;
    wake();
  };
  stream.on("error", onError);

  // Only a destroyed stream's `close` means no `end` is coming: the ZIP parser announces `close` itself, ahead of the
  // `end` its readable side still has to deliver.
  const onClose = () => {
    if (!ended && !error && stream.destroyed) {
      error = prematureCloseError();
    }
    wake();
  };
  stream.on("close", onClose);
  if (!error && stream.destroyed && !stream.readableEnded) {
    error = prematureCloseError();
  }

  try {
    while (!ended || contents.length > 0) {
      if (contents.length === 0) {
        if (error) {
          throw toError(error);
        }
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
    stream.removeListener("close", onClose);
  }
}

/** What Node's own stream iterator reports for a stream closed before it ended. */
function prematureCloseError(): Error {
  return Object.assign(new Error("Premature close"), { code: "ERR_STREAM_PREMATURE_CLOSE" });
}

export { iterateStream };
