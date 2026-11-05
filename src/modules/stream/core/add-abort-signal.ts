/**
 * Stream Module - Common Abort Signal Helper
 *
 * Platform-neutral abort signal handling, parameterized by
 * event listener management (Node.js uses `.on`/`.off`,
 * browser uses `addEmitterListener`/`removeEmitterListener`).
 */

import { isReadableStream, isWritableStream } from "@stream/core/type-guards";
import { StreamArgumentError } from "@stream/errors";
import type { EventEmitterLike, ReadableLike, WritableLike } from "@stream/types";
import { createAbortError } from "@utils/errors";

// =============================================================================
// Types
// =============================================================================

interface ListenerOps {
  add(emitter: EventEmitterLike, event: string, listener: (...args: any[]) => void): void;
  remove(emitter: EventEmitterLike, event: string, listener: (...args: any[]) => void): void;
}

// =============================================================================
// Factory
// =============================================================================

/**
 * Create a platform-specific `addAbortSignal` function.
 */
export function createAddAbortSignal(ops: ListenerOps) {
  return function addAbortSignal<
    T extends (ReadableLike | WritableLike) & { destroy(error?: Error): any }
  >(signal: AbortSignal, stream: T): T {
    if (
      typeof signal !== "object" ||
      signal === null ||
      typeof signal.addEventListener !== "function" ||
      typeof signal.aborted !== "boolean"
    ) {
      throw new StreamArgumentError("an AbortSignal", signal === null ? "null" : typeof signal, {
        field: "signal",
        where: "addAbortSignal"
      });
    }
    // A Web stream has no `destroy`: aborting cancels a ReadableStream and aborts a
    // WritableStream, the nearest a caller can come to Node's erroring it.
    if (isReadableStream(stream) || isWritableStream(stream)) {
      const web = stream as unknown;
      const stop = (): void => {
        const reason = createAbortError(signal.reason);
        const settled = isReadableStream(web)
          ? web.cancel(reason)
          : (web as WritableStream).abort(reason);
        settled.catch(() => {
          // A locked stream refuses; its reader or writer sees the abort instead
        });
      };
      if (signal.aborted) {
        stop();
      } else {
        signal.addEventListener("abort", stop, { once: true });
      }
      return stream;
    }
    if (
      typeof stream !== "object" ||
      stream === null ||
      typeof (stream as { destroy?: unknown }).destroy !== "function" ||
      typeof (stream as { once?: unknown; on?: unknown }).on !== "function"
    ) {
      throw new StreamArgumentError("a stream", stream === null ? "null" : typeof stream, {
        field: "stream",
        where: "addAbortSignal"
      });
    }
    if (signal.aborted) {
      stream.destroy(createAbortError(signal.reason));
      return stream;
    }

    const cleanup = (): void => {
      signal.removeEventListener("abort", onAbort);
      ops.remove(stream, "close", onDone);
      ops.remove(stream, "end", onDone);
      ops.remove(stream, "finish", onDone);
      ops.remove(stream, "error", onError);
    };

    const onAbort = (): void => {
      cleanup();
      stream.destroy(createAbortError(signal.reason));
    };

    const onDone = (): void => {
      cleanup();
    };

    const onError = (): void => {
      cleanup();
    };

    signal.addEventListener("abort", onAbort, { once: true });
    ops.add(stream, "close", onDone);
    ops.add(stream, "end", onDone);
    ops.add(stream, "finish", onDone);
    ops.add(stream, "error", onError);

    return stream;
  };
}
