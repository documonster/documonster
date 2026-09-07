import { compress, decompress, decompressSync } from "@archive/compression/compress";
import { hasDeflateRawDecompressionStream } from "@archive/compression/compress.base";
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

/**
 * Regression: native `DecompressionStream` (and `CompressionStream`) can
 * intermittently reject input that is in fact valid deflate — observed in
 * Chromium under heavy concurrent native-stream creation, surfacing as a
 * spurious "invalid literal/lengths set" / "invalid distances set" error on a
 * payload that the pure-JS inflater decodes correctly.
 *
 * `processWithStrategy` now falls back to the deterministic pure-JS codec when
 * the native stream throws, so a transient native failure can never corrupt a
 * read. These tests pin that behaviour:
 *
 *  1. A large batch of concurrent compress→decompress round-trips must all
 *     succeed and return byte-identical data (would flake ~2% before the fix).
 *  2. The async `decompress` must agree with `decompressSync` for the same
 *     bytes (the fallback guarantees they can never disagree on valid data).
 */

function randomText(byteLen: number): Uint8Array {
  const alphabet = "abcdefghijklmnopqrstuvwxyz0123456789 <w:p><w:r><w:t></w:t></w:r></w:p>";
  let s = "";
  for (let i = 0; i < byteLen; i++) {
    s += alphabet[(Math.random() * alphabet.length) | 0];
  }
  return new TextEncoder().encode(s);
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

describe("decompress robustness under concurrency", () => {
  it("200 concurrent compress→decompress round-trips are all byte-exact", async () => {
    let ok = 0;
    let bad = 0;

    for (let batch = 0; batch < 25; batch++) {
      // ~1.6 MB each crosses the size where the native flake was observed and
      // is large enough to produce multi-block deflate output.
      const inputs = Array.from({ length: 8 }, () => randomText(1_600_000));
      const results = await Promise.allSettled(
        inputs.map(async data => {
          const compressed = await compress(data, { level: 6 });
          const restored = await decompress(compressed);
          return bytesEqual(restored, data);
        })
      );
      for (const r of results) {
        if (r.status === "fulfilled" && r.value) {
          ok++;
        } else {
          bad++;
        }
      }
    }

    expect(bad).toBe(0);
    expect(ok).toBe(200);
  }, 300_000);

  it("async decompress always agrees with sync decompress on valid deflate", async () => {
    expect(hasDeflateRawDecompressionStream()).toBe(true);

    for (let i = 0; i < 20; i++) {
      const data = randomText(1_600_000);
      const compressed = await compress(data, { level: 6 });

      const viaAsync = await decompress(compressed);
      const viaSync = decompressSync(compressed);

      expect(bytesEqual(viaAsync, data)).toBe(true);
      expect(bytesEqual(viaSync, data)).toBe(true);
    }
  }, 120_000);
});

/**
 * The recovery above must also be *silent*, and for a long time it was not.
 *
 * When the native stream rejects, both halves of the transform reject: the pending
 * `writer.write()` and the read loop draining `stream.readable`. `transformWithStream`
 * awaited the write first and returned the read promise last, so the throw from `write()`
 * abandoned the read rejection with no handler attached. The fallback then produced correct
 * bytes and every assertion passed — while the runtime reported
 * `TypeError: The compressed data was not valid: invalid code lengths set.` as an unhandled
 * rejection, blamed on whichever test happened to be running when the microtask surfaced.
 *
 * That is how the `Browser` job failed with `411 passed / 2 errors`, pointing at
 * `excel/xlsb/__tests__/protection-and-names.test.ts` — a file that reads an XLSB package,
 * i.e. does nothing but call `extractAll`.
 *
 * Reproducing the native flake itself takes a loaded machine and luck, so this one forces
 * it: a `DecompressionStream` replacement whose transform always throws puts every decode on
 * the fallback path deterministically. Measured against the unfixed code, `extractAll`
 * leaked one unhandled rejection and an XLSB `Workbook.read` leaked ten.
 */
describe("decompress fallback does not leak unhandled rejections", () => {
  /** Swap in a `DecompressionStream` whose transform always errors, as Chromium's flake does. */
  function installFailingDecompressionStream(): () => void {
    const original = globalThis.DecompressionStream;
    class Failing {
      readable: ReadableStream<Uint8Array>;
      writable: WritableStream<Uint8Array>;
      constructor(_format: string) {
        // Construction must succeed: `hasDeflateRawDecompressionStream()` probes it, and a
        // constructor that threw would route around the native path instead of failing on it.
        const transform = new TransformStream<Uint8Array, Uint8Array>({
          transform() {
            throw new TypeError("The compressed data was not valid: invalid code lengths set.");
          }
        });
        this.readable = transform.readable;
        this.writable = transform.writable;
      }
    }
    (globalThis as unknown as { DecompressionStream: unknown }).DecompressionStream = Failing;
    return () => {
      (globalThis as unknown as { DecompressionStream: unknown }).DecompressionStream = original;
    };
  }

  /** Collect unhandled rejections, keeping them away from the runner's own reporter. */
  function captureUnhandledRejections(): { reasons: unknown[]; stop: () => void } {
    const reasons: unknown[] = [];
    const onRejection = (event: PromiseRejectionEvent): void => {
      reasons.push(event.reason);
      event.preventDefault();
    };
    globalThis.addEventListener("unhandledrejection", onRejection);
    return {
      reasons,
      stop: () => globalThis.removeEventListener("unhandledrejection", onRejection)
    };
  }

  it("recovers a deflated ZIP entry without reporting one", async () => {
    const body = randomText(200_000);
    const zip = await createZip([{ name: "a.txt", data: body }], { level: 6 });

    const restore = installFailingDecompressionStream();
    const captured = captureUnhandledRejections();
    try {
      const files = await extractAll(zip);
      // The fallback still produced the right bytes …
      expect(bytesEqual(files.get("a.txt")!.data, body)).toBe(true);
      // … and did so quietly. A rejection surfaces a macrotask after the throw.
      await new Promise(resolve => setTimeout(resolve, 300));
      expect(captured.reasons).toEqual([]);
    } finally {
      captured.stop();
      restore();
    }
  }, 60_000);
});
