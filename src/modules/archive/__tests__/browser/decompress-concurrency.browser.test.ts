import { compress, decompress, decompressSync } from "@archive/compression/compress";
import { hasDeflateRawDecompressionStream } from "@archive/compression/compress.base";
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
