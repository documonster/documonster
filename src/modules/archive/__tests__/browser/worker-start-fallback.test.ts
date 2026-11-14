/**
 * `useWorker` unset sends a large call to a worker, so it does not hold the page's thread. A worker
 * that cannot start — its script refused by a Content Security Policy, which throws from
 * `new Worker` in some browsers and fires `error` on it in others — ran nothing, so the call
 * completes on this thread instead of failing.
 */
import {
  compress,
  decompress,
  decompressSync,
  gunzipSync,
  gzip,
  hasWorkerSupport,
  unzlibSync,
  zlib
} from "@archive/compression/compress.browser";
import { deflateRawCompressed } from "@archive/compression/deflate-fallback";
import {
  WorkerPool,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool
} from "@archive/compression/worker-pool/index.browser";
import { ArchiveError } from "@archive/core/errors";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

const LARGE = 2 * 1024 * 1024; // above the automatic worker threshold

function compressible(size: number): Uint8Array {
  const pattern = new TextEncoder().encode("worker start fallback, repeated. ");
  const out = new Uint8Array(size);
  for (let i = 0; i < size; i++) {
    out[i] = pattern[i % pattern.length]!;
  }
  return out;
}

/** Bytes that do not compress, so their DEFLATE form is above the threshold too. */
function incompressible(size: number): Uint8Array {
  const out = new Uint8Array(size);
  for (let i = 0; i < size; i += 65536) {
    crypto.getRandomValues(out.subarray(i, i + 65536));
  }
  return out;
}

const OriginalWorker = globalThis.Worker;

/** A `Worker` whose construction throws, as a CSP refusal does in Chromium. */
function RefusedWorker(): never {
  throw new DOMException(
    "Failed to construct 'Worker': refused by the document's Content Security Policy",
    "SecurityError"
  );
}

const bootFailureUrl = URL.createObjectURL(
  new Blob(["throw new Error('boot failure');"], { type: "text/javascript" })
);

/** A `Worker` that is created and then fails to load, as a CSP refusal does in Firefox. */
class FailingWorker extends OriginalWorker {
  constructor(_url: string | URL) {
    super(bootFailureUrl);
  }
}

function installWorker(replacement: unknown): () => void {
  (globalThis as { Worker: unknown }).Worker = replacement;
  return () => {
    (globalThis as { Worker: unknown }).Worker = OriginalWorker;
  };
}

const failures: Record<string, unknown> = {
  "a Worker constructor that throws": RefusedWorker,
  "a worker that fails to load": FailingWorker
};

describe("automatic worker", () => {
  beforeEach(() => {
    terminateDefaultWorkerPool();
  });

  afterEach(() => {
    terminateDefaultWorkerPool();
  });

  it("runs a large call on a worker and a small one on this thread", async () => {
    if (!hasWorkerSupport()) {
      return;
    }
    const original = compressible(LARGE);
    const compressed = await compress(original);
    expect(getDefaultWorkerPool().getStats().completedTasks).toBe(1);
    expect(await decompress(compressed, { autoWorkerThreshold: 0 })).toEqual(original);
    expect(getDefaultWorkerPool().getStats().completedTasks).toBe(2);

    // The DEFLATE form of a repeated pattern is far below the threshold
    expect(await decompress(compressed)).toEqual(original);
    expect(getDefaultWorkerPool().getStats().completedTasks).toBe(2);
  });

  it("stays on this thread, and succeeds, where the platform has no native deflate-raw", async () => {
    if (!hasWorkerSupport()) {
      return;
    }
    // The worker codes with the native streams, so it could only fail here
    const cs = globalThis.CompressionStream;
    const ds = globalThis.DecompressionStream;
    (globalThis as { CompressionStream: unknown }).CompressionStream = undefined;
    (globalThis as { DecompressionStream: unknown }).DecompressionStream = undefined;
    try {
      const original = incompressible(LARGE);
      const compressed = await compress(original, { useWorker: true });
      expect(await decompress(compressed)).toEqual(original);
      expect(getDefaultWorkerPool().getStats().totalWorkers).toBe(0);
    } finally {
      globalThis.CompressionStream = cs;
      globalThis.DecompressionStream = ds;
    }
  });

  for (const [name, replacement] of Object.entries(failures)) {
    it(`compress completes on this thread with ${name}`, async () => {
      if (!hasWorkerSupport()) {
        return;
      }
      const original = compressible(LARGE);
      const restore = installWorker(replacement);
      try {
        const compressed = await compress(original);
        expect(decompressSync(compressed)).toEqual(original);
        expect(getDefaultWorkerPool().getStats().failedTasks).toBe(1);
      } finally {
        restore();
      }
    });

    it(`decompress completes on this thread with ${name}`, async () => {
      if (!hasWorkerSupport()) {
        return;
      }
      const original = incompressible(LARGE);
      const compressed = deflateRawCompressed(original, 6);
      const restore = installWorker(replacement);
      try {
        expect(await decompress(compressed)).toEqual(original);
        // The task was tried and failed, and nothing is left queued on a pool with no worker
        expect(getDefaultWorkerPool().getStats()).toMatchObject({
          failedTasks: 1,
          pendingTasks: 0
        });
      } finally {
        restore();
      }
    });
  }

  // The trailer's checksum and length are of the input, which `allowTransfer` hands to the worker:
  // read from it afterwards they were an empty view's, and the file failed its own check.
  it.each([
    ["gzip", gzip, gunzipSync],
    ["zlib", zlib, unzlibSync]
  ] as const)(
    "%s on a worker with allowTransfer writes a valid trailer",
    async (_name, encode, decode) => {
      if (!hasWorkerSupport()) {
        return;
      }
      const whole = compressible(LARGE);
      const expected = whole.slice();
      expect(decode(await encode(whole, { useWorker: true, allowTransfer: true }))).toEqual(
        expected
      );
      expect(whole.byteLength).toBe(0); // handed over, as allowTransfer allows

      // A view is copied before it is handed over, so the caller's buffer is left alone
      const backing = compressible(LARGE + 16);
      const view = backing.subarray(8, 8 + LARGE);
      expect(decode(await encode(view, { useWorker: true, allowTransfer: true }))).toEqual(
        backing.slice(8, 8 + LARGE)
      );
      expect(view.byteLength).toBe(LARGE);
    }
  );

  it("refuses, naming why, an input whose buffer was already transferred", async () => {
    if (!hasWorkerSupport()) {
      return;
    }
    const original = compressible(LARGE);
    const restore = installWorker(FailingWorker);
    try {
      const failure = compress(original, { useWorker: true, allowTransfer: true });
      await expect(failure).rejects.toThrow(ArchiveError);
      await expect(failure).rejects.toThrow(/allowTransfer had moved the input/);
      // The bytes went to the worker, so there was nothing left to compress here
      expect(original.byteLength).toBe(0);
    } finally {
      restore();
    }
  });
});

describe("WorkerPool with a worker that cannot start", () => {
  it("rejects every queued task with the archive's error and keeps none", async () => {
    if (!hasWorkerSupport()) {
      return;
    }
    const restore = installWorker(RefusedWorker);
    const pool = new WorkerPool({ minWorkers: 1 });
    try {
      const results = await Promise.allSettled([
        pool.execute("deflate", compressible(1024)),
        pool.execute("inflate", deflateRawCompressed(compressible(1024), 6))
      ]);
      for (const result of results) {
        expect(result.status).toBe("rejected");
        expect((result as PromiseRejectedResult).reason).toBeInstanceOf(ArchiveError);
        expect(((result as PromiseRejectedResult).reason as Error).cause).toBeInstanceOf(
          DOMException
        );
      }
      expect(pool.getStats()).toMatchObject({ pendingTasks: 0, failedTasks: 2 });
    } finally {
      pool.terminate();
      restore();
    }
  });
});
