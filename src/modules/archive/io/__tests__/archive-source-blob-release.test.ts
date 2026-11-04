/**
 * A `Blob` source is read through a stream `toAsyncIterable` opens itself, so that stream is
 * the reader's, not the caller's: leaving the iteration early must cancel it (a caller's Web
 * stream is only released, since it stays theirs). Reading to the end leaves nothing to cancel.
 */
import { Archive } from "@archive";
import { toAsyncIterable } from "@archive/io/archive-source";
import { describe, expect, it } from "vitest";

/** A Blob whose `stream()` reports what was opened and whether it was cancelled. */
function trackedBlob(data: Uint8Array): {
  blob: Blob;
  opened: ReadableStream<Uint8Array>[];
  cancelled: () => number;
} {
  const blob = new Blob([data as BlobPart]);
  const real = blob.stream.bind(blob);
  const opened: ReadableStream<Uint8Array>[] = [];
  let cancelled = 0;
  Object.defineProperty(blob, "stream", {
    value: () => {
      const inner = real().getReader();
      const stream = new ReadableStream<Uint8Array>(
        {
          async pull(controller) {
            const { done, value } = await inner.read();
            if (done) {
              controller.close();
            } else {
              controller.enqueue(value);
            }
          },
          cancel(reason) {
            cancelled++;
            return inner.cancel(reason);
          }
        },
        { highWaterMark: 0 }
      );
      opened.push(stream);
      return stream;
    }
  });
  return { blob, opened, cancelled: () => cancelled };
}

function bytes(length: number): Uint8Array {
  const out = new Uint8Array(length);
  for (let i = 0; i < length; i++) {
    out[i] = (i * 31) & 0xff;
  }
  return out;
}

describe("toAsyncIterable(Blob) owns the stream it opens", () => {
  it("cancels it when the loop is left early", async () => {
    const t = trackedBlob(bytes(512 * 1024));
    for await (const _chunk of toAsyncIterable(t.blob)) {
      break;
    }
    expect(t.opened).toHaveLength(1);
    expect(t.cancelled()).toBe(1);
    expect(t.opened[0]!.locked).toBe(false);
  });

  it("cancels it when the loop body throws", async () => {
    const t = trackedBlob(bytes(512 * 1024));
    await expect(
      (async () => {
        for await (const _chunk of toAsyncIterable(t.blob)) {
          throw new Error("stop");
        }
      })()
    ).rejects.toThrow("stop");
    expect(t.cancelled()).toBe(1);
    expect(t.opened[0]!.locked).toBe(false);
  });

  it("cancels it once when an abort already cancelled it", async () => {
    const t = trackedBlob(bytes(512 * 1024));
    const controller = new AbortController();
    await expect(
      (async () => {
        for await (const _chunk of toAsyncIterable(t.blob, { signal: controller.signal })) {
          controller.abort();
        }
      })()
    ).rejects.toThrow();
    expect(t.cancelled()).toBe(1);
    expect(t.opened[0]!.locked).toBe(false);
  });

  it("does not cancel a stream read to its end", async () => {
    const data = bytes(300 * 1024);
    const t = trackedBlob(data);
    let length = 0;
    for await (const chunk of toAsyncIterable(t.blob)) {
      length += chunk.length;
    }
    expect(length).toBe(data.length);
    expect(t.cancelled()).toBe(0);
    expect(t.opened[0]!.locked).toBe(false);
  });

  it("through parseTarStream: leaving at the first entry cancels it", async () => {
    const tar = await Archive.packTar([
      { name: "a.bin", source: bytes(400 * 1024) },
      { name: "b.bin", source: bytes(400 * 1024) }
    ]);
    const t = trackedBlob(tar);
    for await (const entry of Archive.parseTarStream(t.blob)) {
      expect(entry.info.path).toBe("a.bin");
      break;
    }
    expect(t.cancelled()).toBe(1);
    expect(t.opened[0]!.locked).toBe(false);
  });
});
