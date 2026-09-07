/**
 * `transformWithStream` must never abandon its read side.
 *
 * When a `DecompressionStream` rejects its input, *both* halves of the transform reject: the pending
 * `writer.write()` and the read loop draining `stream.readable`. The original code awaited the write
 * first and returned the read promise last, so the throw from `write()` left the read rejection with
 * no handler attached — an unhandled rejection, raised even though every caller of this function
 * catches the failure and recovers (`processWithStrategy` falls back to the pure-JS inflater).
 *
 * That is how a green Excel suite still failed CI: two `extractAll` tests in
 * `excel/xlsb/__tests__/protection-and-names.test.ts` passed while Vitest reported
 * `TypeError: The compressed data was not valid: invalid code lengths set.` as an unhandled
 * rejection, attributed to whichever test happened to be running when the microtask surfaced.
 *
 * This file is Node-only because catching an unhandled rejection needs `process`; the code under
 * test is the shared `compress.base` module, and the same write-rejects-first ordering reproduces
 * here, so the browser-only symptom has a deterministic Node regression test.
 */
import { transformWithStream } from "@archive/compression/compress.base";
import { afterEach, describe, expect, it } from "vitest";

/** Raw DEFLATE that decodes to nothing: a dynamic-Huffman block with an unusable code-length table. */
const INVALID_DEFLATE = new Uint8Array([
  0x05, 0xc0, 0x81, 0x00, 0x00, 0x00, 0x00, 0x00, 0x90, 0xff, 0x6b, 0x01, 0x02, 0x03, 0x04, 0x05
]);

const captured: unknown[] = [];
const onUnhandled = (reason: unknown): void => {
  captured.push(reason);
};

/** Let queued microtasks and one macrotask turn run, which is when an abandoned rejection reports. */
async function settle(): Promise<void> {
  await new Promise(resolve => setTimeout(resolve, 10));
}

afterEach(() => {
  process.off("unhandledRejection", onUnhandled);
  captured.length = 0;
});

describe("transformWithStream", () => {
  it("reports a failed decode without leaving an unhandled rejection", async () => {
    process.on("unhandledRejection", onUnhandled);

    await expect(
      transformWithStream(INVALID_DEFLATE, new DecompressionStream("deflate-raw"))
    ).rejects.toThrow();

    await settle();
    expect(captured).toEqual([]);
  });

  it("surfaces the read side's error, which is the one that names the defect", async () => {
    // A synthetic transform: the writable rejects with a bland error while the readable rejects with
    // the descriptive one. Awaiting the write first would surface the wrong error *and* drop the
    // right one, so this pins both halves of the fix without depending on a platform codec.
    const readError = new Error("read-side");
    const stream = {
      writable: {
        getWriter: () => ({
          write: () => Promise.reject(new Error("write-side")),
          close: () => Promise.resolve(),
          releaseLock: () => {}
        })
      },
      readable: {
        getReader: () => ({
          read: () => Promise.reject(readError),
          cancel: () => Promise.resolve(),
          releaseLock: () => {}
        })
      }
    } as unknown as DecompressionStream;

    await expect(transformWithStream(new Uint8Array([1, 2, 3]), stream)).rejects.toBe(readError);
  });

  it("does not hang when the write fails but the readable stays open", async () => {
    // `reader.cancel()` is what terminates the read loop here. Without it the read promise would
    // never settle and awaiting it would deadlock instead of reporting the write failure.
    const writeError = new Error("write-side");
    let cancelled = false;
    const stream = {
      writable: {
        getWriter: () => ({
          write: () => Promise.reject(writeError),
          close: () => Promise.resolve(),
          releaseLock: () => {}
        })
      },
      readable: {
        getReader: () => ({
          read: () =>
            new Promise<ReadableStreamReadResult<Uint8Array>>(resolve => {
              const poll = (): void => {
                if (cancelled) {
                  resolve({ done: true, value: undefined });
                } else {
                  setTimeout(poll, 1);
                }
              };
              poll();
            }),
          cancel: () => {
            cancelled = true;
            return Promise.resolve();
          },
          releaseLock: () => {}
        })
      }
    } as unknown as DecompressionStream;

    await expect(transformWithStream(new Uint8Array([1, 2, 3]), stream)).rejects.toBe(writeError);
  });

  it("still round-trips valid data", async () => {
    const original = new TextEncoder().encode("documonster ".repeat(64));
    const compressed = await transformWithStream(original, new CompressionStream("deflate-raw"));
    const restored = await transformWithStream(compressed, new DecompressionStream("deflate-raw"));
    expect(restored).toEqual(original);
  });
});
