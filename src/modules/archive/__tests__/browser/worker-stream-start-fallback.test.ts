/**
 * A streaming session whose worker does not start — it misses its deadline on a busy
 * device, never answers, or fails to load — continues on the main thread instead of
 * failing the stream. Nothing was sent to the worker, so the output is complete.
 */
import { decompressSync as inflateRawSync } from "@archive/compression/compress.browser";
import { deflateRawCompressed } from "@archive/compression/deflate-fallback";
import {
  createDeflateStream,
  createInflateStream,
  hasWorkerSupport
} from "@archive/compression/streaming-compress.browser";
import { WorkerPool } from "@archive/compression/worker-pool/index.browser";
import { concatUint8Arrays } from "@utils/binary";
import { describe, expect, it } from "vitest";

const original = new TextEncoder().encode("worker start fallback ".repeat(5000));

function run(
  stream: ReturnType<typeof createDeflateStream>,
  input: Uint8Array
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    stream.on("data", (chunk: Uint8Array) => chunks.push(chunk));
    stream.on("end", () => resolve(concatUint8Arrays(chunks)));
    stream.on("error", reject);
    for (let i = 0; i < input.length; i += 4096) {
      stream.write(input.subarray(i, i + 4096));
    }
    stream.end();
  });
}

function scriptUrl(source: string): string {
  return URL.createObjectURL(new Blob([source], { type: "text/javascript" }));
}

const pools: Record<string, () => WorkerPool> = {
  // The real worker, given no time at all: the deadline passes before it can answer
  "a worker that misses its deadline": () => new WorkerPool({ streamStartTimeout: 0 }),
  "a worker that never answers": () =>
    new WorkerPool({ workerUrl: scriptUrl("self.onmessage = () => {};"), streamStartTimeout: 50 }),
  "a worker that fails to load": () =>
    new WorkerPool({ workerUrl: scriptUrl("throw new Error('boot failure');") })
};

describe("worker stream start fallback", () => {
  for (const [name, makePool] of Object.entries(pools)) {
    it(`deflate completes on the main thread with ${name}`, async () => {
      if (!hasWorkerSupport()) {
        return;
      }
      const pool = makePool();
      try {
        const compressed = await run(
          createDeflateStream({ useWorker: true, workerPool: pool }),
          original
        );
        expect(inflateRawSync(compressed)).toEqual(original);
      } finally {
        pool.terminate();
      }
    });

    it(`inflate completes on the main thread with ${name}`, async () => {
      if (!hasWorkerSupport()) {
        return;
      }
      const pool = makePool();
      try {
        const compressed = deflateRawCompressed(original, 6);
        const restored = await run(
          createInflateStream({ useWorker: true, workerPool: pool }),
          compressed
        );
        expect(restored).toEqual(original);
      } finally {
        pool.terminate();
      }
    });
  }

  it("a pool whose worker missed its deadline still serves later streams", async () => {
    if (!hasWorkerSupport()) {
      return;
    }
    const pool = new WorkerPool({ maxWorkers: 1, streamStartTimeout: 0 });
    try {
      for (let i = 0; i < 3; i++) {
        const compressed = await run(
          createDeflateStream({ useWorker: true, workerPool: pool }),
          original
        );
        expect(inflateRawSync(compressed)).toEqual(original);
      }
      // The unresponsive workers were let go rather than left bound to dead sessions
      expect(pool.getStats().activeWorkers).toBe(0);
    } finally {
      pool.terminate();
    }
  });
});
