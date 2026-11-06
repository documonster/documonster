/**
 * The destination `Workbook.writeStream` writes into, whatever the caller handed it.
 *
 * The serializers drive a Node-style writable: `write` (a boolean or a promise of one, for backpressure), `end`, and
 * the `drain` / `error` / `finish` events. A Web `WritableStream` — what a browser, Deno or a `Response` body gives — is
 * none of those, and crashed the XLSX writer with `this.pipedStream.write is not a function`. It is driven here through
 * `SinkAdapter`, the stream module's one adapter for exactly this, and presented as the writable the serializers expect.
 */
import type { XlsxStreamListener, XlsxWritable } from "@excel/core/xlsx-io-types";
import { ExcelArgumentError } from "@excel/errors";
import type { SinkAdapter as SinkAdapterClass } from "@stream/core/sink-adapter";
import { describeValue as describe } from "@utils/describe-value";

/** What `Workbook.writeStream` writes into: a Node-style writable, or a Web `WritableStream`. */
export type WorkbookWritable = XlsxWritable | WritableStream<Uint8Array>;

export function isWebWritable(stream: unknown): stream is WritableStream<Uint8Array> {
  return (
    typeof stream === "object" &&
    stream !== null &&
    typeof (stream as { getWriter?: unknown }).getWriter === "function"
  );
}

/** The writable the serializers drive, and — for a Web stream — when the stream has really taken everything. */
export interface WorkbookSink {
  readonly sink: XlsxWritable;
  /**
   * Settles once the Web stream is closed (or fails); `writeStream` resolves after it, so the caller's data is in the
   * stream when the promise does. `undefined` for a Node-style writable, which the serializers settle themselves.
   */
  readonly finished?: Promise<void>;
}

/**
 * `stream` as the writable the serializers drive, or an `ExcelArgumentError` on `stream` (`where` names the call) for
 * something that is neither a Node-style writable — `write` and `end` — nor a Web `WritableStream`.
 *
 * **Synchronous for a Node-style writable**, so the serializer attaches its listeners in the tick `writeStream` was
 * called in — a sink that errors or closes straight away is observed. Only a Web stream needs the adapter, which is
 * loaded then.
 */
export async function workbookSink(where: string, stream: unknown): Promise<WorkbookSink> {
  if (isWebWritable(stream)) {
    const { SinkAdapter } = await import("@stream/core/sink-adapter");
    return webSink(new SinkAdapter(stream));
  }
  return nodeSink(where, stream);
}

/** {@link workbookSink} for anything that is not a Web stream: the writable itself, checked. */
export function nodeSink(where: string, stream: unknown): WorkbookSink {
  const candidate = stream as { write?: unknown; end?: unknown } | null;
  if (
    typeof candidate !== "object" ||
    candidate === null ||
    typeof candidate.write !== "function" ||
    typeof candidate.end !== "function"
  ) {
    throw new ExcelArgumentError(
      "stream",
      `${where}: stream must be a writable stream — a Node Writable or a Web WritableStream — got ${describe(stream)}.`
    );
  }
  return { sink: stream as XlsxWritable };
}

/**
 * A Web `WritableStream` as a Node-style writable: writes are queued in order and each resolves once the stream has
 * taken the chunk (its backpressure), `end` closes it after the last and then emits `finish` and `close`, and a failure of either is emitted as `error`
 * — which the serializers turn into a rejected `writeStream`. The stream's lock is released when it is closed or fails.
 */
function webSink(adapter: SinkAdapterClass): WorkbookSink {
  const listeners = new Map<string, Set<XlsxStreamListener>>();
  let failed = false;
  let queue: Promise<void> = Promise.resolve();
  let settle!: (error?: unknown) => void;
  const finished = new Promise<void>((resolve, reject) => {
    settle = error => (error === undefined ? resolve() : reject(error));
  });
  // Awaited by `writeStream`, which may already be rejecting for the same failure.
  finished.catch(() => {});
  const emit = (event: string, ...args: unknown[]): void => {
    for (const listener of [...(listeners.get(event) ?? [])]) {
      listener(...args);
    }
  };
  const fail = (error: unknown): void => {
    if (failed) {
      return;
    }
    failed = true;
    const reason = error instanceof Error ? error : new Error(String(error));
    settle(reason);
    emit("error", reason);
    emit("close");
  };
  const sink: XlsxWritable = {
    write(data) {
      const chunk = typeof data === "string" ? new TextEncoder().encode(data) : data;
      // In order, one after another: the serializers do not wait for a write before the next, nor before `end`.
      const written = queue.then(() => adapter.write(chunk));
      queue = written.catch(() => {});
      return written.then(
        () => true,
        error => {
          fail(error);
          return true;
        }
      );
    },
    end() {
      queue
        .then(() => adapter.end())
        .then(
          () => {
            settle();
            emit("finish");
            emit("close");
          },
          error => {
            void adapter.abort(error).catch(() => {});
            fail(error);
          }
        );
    },
    on(event, listener) {
      const set = listeners.get(event) ?? new Set();
      set.add(listener);
      listeners.set(event, set);
      return sink;
    },
    once(event, listener) {
      const wrapped: XlsxStreamListener = (...args) => {
        sink.off(event, wrapped);
        listener(...args);
      };
      return sink.on(event, wrapped);
    },
    off(event, listener) {
      listeners.get(event)?.delete(listener);
      return sink;
    }
  };
  return { sink, finished };
}
