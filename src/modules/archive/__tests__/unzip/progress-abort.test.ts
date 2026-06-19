/**
 * Tests for unzip progress tracking and abort/cancel functionality.
 *
 * Stream lifecycle, premature close prevention, data descriptor, and
 * stress tests are in stream-lifecycle.test.ts and stream-robustness.test.ts.
 */

import { isAbortError } from "@archive/core/errors";
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { delay, chunkBytes } from "./test-helpers";

describe("unzip progress + abort", () => {
  it("should report bytesIn while parsing a streaming source", async () => {
    const z = zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("hello"))
      .add("b.txt", new TextEncoder().encode("world"));

    const zipBytes = await z.bytes();

    const events: Array<{ bytesIn: number; entriesEmitted: number; phase: string }> = [];

    const reader = unzip(chunkBytes(zipBytes, 64));
    const op = reader.operation({
      progressIntervalMs: 0,
      onProgress: p => {
        events.push({ bytesIn: p.bytesIn, entriesEmitted: p.entriesEmitted, phase: p.phase });
      }
    });

    const seen = new Map<string, Uint8Array>();
    for await (const entry of op.iterable) {
      const data = await entry.bytes();
      seen.set(entry.path, data);
    }

    expect(seen.get("a.txt")?.length).toBe(5);
    expect(seen.get("b.txt")?.length).toBe(5);

    expect(events.length).toBeGreaterThan(0);
    expect(events[events.length - 1]!.phase).toBe("done");
    expect(op.pointer()).toBe(zipBytes.length);
  });

  it("should abort while parsing a streaming source", async () => {
    const z = zip({ level: 0 }).add("big.bin", new Uint8Array(1024 * 1024));
    const zipBytes = await z.bytes();

    const reader = unzip(chunkBytes(zipBytes, 128));
    const op = reader.operation({ progressIntervalMs: 0 });

    const consume = (async () => {
      let count = 0;
      for await (const entry of op.iterable) {
        for await (const _ of entry.stream()) {
          count++;
          break;
        }
      }
      return count;
    })();

    await delay(2);
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

  it("should cancel upstream parsing when consumer stops early (streaming)", async () => {
    const z = zip({ level: 0 })
      .add("a.bin", new Uint8Array(1024 * 128))
      .add("b.bin", new Uint8Array(1024 * 128));
    const zipBytes = await z.bytes();

    const reader = unzip(chunkBytes(zipBytes, 4096));
    const op = reader.operation({ progressIntervalMs: 0 });

    let seen = 0;
    for await (const entry of op.iterable) {
      for await (const _ of entry.stream()) {
        break;
      }
      seen++;
      break;
    }

    expect(seen).toBe(1);
    await delay(5);
    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe("cancelled");
    expect(op.progress().phase).toBe("aborted");
  });

  it(
    "stopping an entry stream early should not hang and should allow parsing next entries",
    { timeout: 10_000 },
    async () => {
      const z = zip({ level: 0 })
        .add("a.bin", new Uint8Array(1024 * 256))
        .add("b.bin", new Uint8Array(1024 * 16));
      const zipBytes = await z.bytes();

      const reader = unzip(chunkBytes(zipBytes, 4096));
      const op = reader.operation({ progressIntervalMs: 0 });

      const seen: string[] = [];
      for await (const entry of op.iterable) {
        seen.push(entry.path);

        if (entry.path === "a.bin") {
          let touched = 0;
          for await (const chunk of entry.stream()) {
            touched += chunk.length;
            break;
          }
          expect(touched).toBeGreaterThan(0);
          continue;
        }

        if (entry.path === "b.bin") {
          const data = await entry.bytes();
          expect(data.length).toBe(1024 * 16);
        }
      }

      expect(seen).toEqual(["a.bin", "b.bin"]);
      expect(op.signal.aborted).toBe(false);
      expect(op.progress().phase).toBe("done");
    }
  );

  it("should mark aborted when consumer stops early (buffer)", async () => {
    const z = zip({ level: 0 })
      .add("a.txt", new TextEncoder().encode("hello"))
      .add("b.txt", new TextEncoder().encode("world"));
    const zipBytes = await z.bytes();

    const reader = unzip(zipBytes);
    const op = reader.operation({ progressIntervalMs: 0 });

    let seen = 0;
    for await (const _entry of op.iterable) {
      seen++;
      break;
    }

    expect(seen).toBe(1);
    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe("cancelled");
    expect(op.progress().phase).toBe("aborted");
  });

  it("abort() should be idempotent and keep the first reason", async () => {
    const z = zip({ level: 0 }).add("a.bin", new Uint8Array(1024 * 128));
    const zipBytes = await z.bytes();

    const reader = unzip(chunkBytes(zipBytes, 128));
    const op = reader.operation({ progressIntervalMs: 0 });

    op.abort("first");
    op.abort("second");

    expect(op.signal.aborted).toBe(true);
    expect((op.signal as any).reason).toBe("first");

    await expect(
      (async () => {
        for await (const _entry of op.iterable) {
          // drain
        }
      })()
    ).rejects.toSatisfy((e: unknown) => isAbortError(e));
  });
});
