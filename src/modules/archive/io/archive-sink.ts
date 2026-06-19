import { onceEvent } from "@stream/core/event-utils";
import { isWritableStream } from "@stream/core/type-guards";
import { concatUint8Arrays } from "@utils/binary";

export type ArchiveSink =
  | WritableStream<Uint8Array>
  | {
      write(chunk: Uint8Array): any;
      end?(cb?: any): any;
      on?(event: string, listener: (...args: any[]) => void): any;
      once?(event: string, listener: (...args: any[]) => void): any;
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
    const hasEvents =
      typeof (sink as any).once === "function" || typeof (sink as any).on === "function";
    if (ok === false && hasEvents) {
      await onceEvent(sink as any, "drain");
    }
  }

  if (typeof sink.end === "function") {
    sink.end();
  }

  if (typeof (sink as any).once === "function" || typeof (sink as any).on === "function") {
    await Promise.race([onceEvent(sink as any, "finish"), onceEvent(sink as any, "close")]);
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
