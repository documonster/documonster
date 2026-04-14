import { compress, decompress, hasCompressionStream } from "@archive/compression/compress.browser";
import { describe, it } from "vitest";

function makeData(size: number): Uint8Array {
  const out = new Uint8Array(size);
  for (let i = 0; i < out.length; i++) {
    out[i] = i & 0xff;
  }
  return out;
}

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T | null> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<null>(resolve => {
        timeoutId = setTimeout(() => resolve(null), timeoutMs);
      })
    ]);
  } finally {
    if (timeoutId !== undefined) {
      clearTimeout(timeoutId);
    }
  }
}

function assertBytesEqual(actual: Uint8Array, expected: Uint8Array, label: string): void {
  if (actual.length !== expected.length) {
    throw new Error(`${label}: length mismatch: ${actual.length} != ${expected.length}`);
  }
  for (let i = 0; i < actual.length; i++) {
    if (actual[i] !== expected[i]) {
      throw new Error(`${label}: byte mismatch at ${i}: ${actual[i]} != ${expected[i]}`);
    }
  }
}

const itNative = hasCompressionStream() ? it : it.skip;

describe("browser compress benchmark (non-assertive)", () => {
  it("roundtrips via JS fallback branch", async () => {
    // Force JS fallback branch (very large threshold).
    const sizes = [256];
    for (const size of sizes) {
      const data = makeData(size);
      const compressed = await compress(data, { level: 6, thresholdBytes: 1024 * 1024 * 1024 });
      const decompressed = await decompress(compressed);
      assertBytesEqual(decompressed, data, "js-fallback");
    }
  });

  itNative("roundtrips via native CompressionStream branch", async () => {
    // Prefer native streams for almost everything (threshold=0).
    // This test is expected to be stable in environments where deflate-raw
    // CompressionStream is truly supported.
    const sizes = [256];
    const nativeTimeoutMs = 5000;

    for (const size of sizes) {
      const data = makeData(size);
      const compressed = await withTimeout(
        compress(data, { level: 6, thresholdBytes: 0 }),
        nativeTimeoutMs
      );
      if (!compressed) {
        throw new Error(
          `native CompressionStream compress timed out after ${nativeTimeoutMs}ms (deflate-raw)`
        );
      }

      const decompressed = await withTimeout(decompress(compressed), nativeTimeoutMs);
      if (!decompressed) {
        throw new Error(
          `native DecompressionStream decompress timed out after ${nativeTimeoutMs}ms (deflate-raw)`
        );
      }

      assertBytesEqual(decompressed, data, "native-stream");
    }
  });

  it("roundtrips when CompressionStream is unavailable (simulated old browser)", async () => {
    const originalCompressionStream = (globalThis as any).CompressionStream;
    const originalDecompressionStream = (globalThis as any).DecompressionStream;

    try {
      // Simulate an older browser without CompressionStream/DecompressionStream.
      (globalThis as any).CompressionStream = undefined;
      (globalThis as any).DecompressionStream = undefined;

      const sizes = [256];
      for (const size of sizes) {
        const data = makeData(size);

        // thresholdBytes=0 would normally prefer native streams, but those are disabled.
        const compressed = await compress(data, { level: 6, thresholdBytes: 0 });
        const decompressed = await decompress(compressed);
        assertBytesEqual(decompressed, data, "simulated-old-browser");
      }
    } finally {
      (globalThis as any).CompressionStream = originalCompressionStream;
      (globalThis as any).DecompressionStream = originalDecompressionStream;
    }
  });
});
