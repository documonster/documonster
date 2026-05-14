/**
 * Backpressure regression for browser-side streaming codecs.
 *
 * Verifies `AsyncStreamCodec.write()` returns false at HWM and emits
 * 'drain' afterwards, so `pipeline(...)` and similar consumers can throttle
 * a fast producer.
 *
 * This mirrors the standard Node `Writable` contract.
 *
 * Only meaningful when the runtime supports native `CompressionStream("deflate-raw")`
 * — that's the path that constructs `AsyncStreamCodec`. On older browsers
 * (no native CompressionStream), `createDeflateStream` returns a
 * `BufferedCodec` whose `write()` always returns `true` by design, so the
 * backpressure assertions wouldn't apply. We skip the test in that case
 * rather than fail it.
 */
import { hasDeflateRawCompressionStream } from "@archive/compression/compress.base";
import { createDeflateStream } from "@archive/index.browser";
import { describe, it, expect } from "vitest";

describe("AsyncStreamCodec backpressure (browser path)", () => {
  it("write() returns false above HWM and emits 'drain' after settling", async ({ skip }) => {
    if (!hasDeflateRawCompressionStream()) {
      skip(
        "Runtime lacks native CompressionStream('deflate-raw'); falls back to BufferedCodec which has no async backpressure."
      );
      return;
    }

    const codec = createDeflateStream({ level: 6 });

    // Soak up data so backend doesn't refuse early.
    const sink: Uint8Array[] = [];
    codec.on("data", chunk => sink.push(chunk));

    // Pump a lot of small chunks back-to-back. The codec's `writeChain`
    // should fill up and `write()` should return false.
    let returnedFalse = 0;
    let returnedTrue = 0;
    for (let i = 0; i < 50; i++) {
      const ok = codec.write(new Uint8Array(1024).fill(i & 0xff));
      if (!ok) {
        returnedFalse++;
      } else {
        returnedTrue++;
      }
    }

    expect(returnedFalse).toBeGreaterThan(0); // we did hit HWM
    expect(returnedTrue).toBeGreaterThan(0); // first few writes were accepted

    // After draining, 'drain' must fire.
    const drained = new Promise<void>((resolve, reject) => {
      const t = setTimeout(() => reject(new Error("drain never fired")), 5_000);
      codec.once("drain", () => {
        clearTimeout(t);
        resolve();
      });
    });

    await drained;

    // Finish off the stream.
    await new Promise<void>((resolve, reject) => {
      codec.once("end", () => resolve());
      codec.once("error", reject);
      codec.end();
    });
  }, 15_000);
});
