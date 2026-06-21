import { onceEvent } from "@stream/core/event-utils";
import { isWritableStream } from "@stream/core/type-guards";
import { concatUint8Arrays } from "@utils/binary";

export type ArchiveSink =
  | WritableStream<Uint8Array>
  | {
      write(chunk: Uint8Array): boolean | void | Promise<unknown>;
      end?(cb?: () => void): void;
      // EventEmitter-style hooks (Node Writable). `...args: any[]` is the
      // standard emitter listener signature.
      on?(event: string, listener: (...args: any[]) => void): unknown;
      once?(event: string, listener: (...args: any[]) => void): unknown;
    };

export async function pipeIterableToSink(
  iterable: AsyncIterable<Uint8Array>,
  sink: ArchiveSink
): Promise<void> {
  if (isWritableStream(sink)) {
    const writer = sink.getWriter();
    try {
      for await (const chunk of iterable) {
        await writer.write(chunk);
      }
      await writer.close();
    } finally {
      try {
        writer.releaseLock();
      } catch {
        // Ignore
      }
    }
    return;
  }

  // Node-style Writable
  for await (const chunk of iterable) {
    const ok = sink.write(chunk);
    const hasEvents = typeof sink.once === "function" || typeof sink.on === "function";
    if (ok === false && hasEvents) {
      await onceEvent(sink, "drain");
    }
  }

  if (typeof sink.end === "function") {
    sink.end();
  }

  if (typeof sink.once === "function" || typeof sink.on === "function") {
    await Promise.race([onceEvent(sink, "finish"), onceEvent(sink, "close")]);
  }
}

export async function collect(iterable: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of iterable) {
    chunks.push(chunk);
    total += chunk.length;
  }
  return concatUint8Arrays(chunks, total);
}
