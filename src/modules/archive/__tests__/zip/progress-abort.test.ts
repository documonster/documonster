import { zip } from "@archive/create-archive";
import { isAbortError } from "@archive/shared/errors";
import { describe, expect, it } from "vitest";

function delay(ms = 0): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function* slowChunks(totalChunks: number, chunkSize: number): AsyncIterable<Uint8Array> {
  for (let i = 0; i < totalChunks; i++) {
    await delay(0);
    yield new Uint8Array(chunkSize);
  }
}

describe("zip progress + abort", () => {
  it("should report progress and pointer during streaming", async () => {
    const events: Array<{ bytesIn: number; bytesOut: number; entriesDone: number; phase: string }> =
      [];

    const z = zip({ level: 0 })
      .add("a.bin", slowChunks(10, 1024))
      .add("b.bin", new Uint8Array(2048));

    const op = z.operation({
      progressIntervalMs: 0,
      onProgress: p => {
        events.push({
          bytesIn: p.bytesIn,
          bytesOut: p.bytesOut,
          entriesDone: p.entriesDone,
          phase: p.phase
        });
      }
    });

    let out = 0;
    for await (const chunk of op.iterable) {
      out += chunk.length;
    }

    expect(out).toBeGreaterThan(0);
    expect(op.pointer()).toBe(out);

    expect(events.length).toBeGreaterThan(0);
    for (let i = 1; i < events.length; i++) {
      expect(events[i]!.bytesIn).toBeGreaterThanOrEqual(events[i - 1]!.bytesIn);
      expect(events[i]!.bytesOut).toBeGreaterThanOrEqual(events[i - 1]!.bytesOut);
    }

    const last = events[events.length - 1]!;
    expect(last.phase).toBe("done");
    expect(last.entriesDone).toBe(2);
  });

  it("should abort an in-flight zip operation", async () => {
    const z = zip({ level: 0 }).add("big.bin", slowChunks(1000, 1024));

    const op = z.operation({ progressIntervalMs: 0 });

    const consume = (async () => {
      let total = 0;
      for await (const chunk of op.iterable) {
        total += chunk.length;
      }
      return total;
    })();

    // Let it start producing.
    await delay(5);
    const reason = new Error("stop");
    op.abort(reason);

    await expect(consume).rejects.toSatisfy((e: unknown) => {
      return (
        isAbortError(e) &&
        (e as any).cause === reason &&
        (e as any).message === "The operation was aborted"
      );
    });
    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe(reason);
    expect(op.progress().phase).toBe("aborted");
  });

  it("should abort even if aborted before consumption", async () => {
    const z = zip({ level: 0 }).add("big.bin", slowChunks(1000, 1024));
    const op = z.operation({ progressIntervalMs: 0 });

    op.abort("stop");

    const consume = (async () => {
      for await (const _ of op.iterable) {
        // no-op
      }
    })();

    await expect(consume).rejects.toSatisfy((e: unknown) => isAbortError(e));
    expect(op.signal.aborted).toBe(true);
    expect(op.progress().phase).toBe("aborted");
  });

  it("should cancel upstream work when consumer stops early", async () => {
    const z = zip({ level: 0 }).add("big.bin", slowChunks(10_000, 1024));
    const op = z.operation({ progressIntervalMs: 0 });

    let seen = 0;
    for await (const _ of op.iterable) {
      seen++;
      break;
    }

    expect(seen).toBe(1);

    // Give the async producer a tick to observe cancellation.
    await delay(5);

    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe("cancelled");
    expect(op.progress().phase).toBe("aborted");
  });

  it("abort() should be idempotent and keep the first reason", async () => {
    const z = zip({ level: 0 }).add("a.bin", slowChunks(10, 1024));
    const op = z.operation({ progressIntervalMs: 0 });

    op.abort("first");
    op.abort("second");

    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe("first");

    await expect(
      (async () => {
        for await (const _ of op.iterable) {
          // drain
        }
      })()
    ).rejects.toSatisfy((e: unknown) => isAbortError(e));
  });
});
