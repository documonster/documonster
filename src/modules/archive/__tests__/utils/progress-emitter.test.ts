import { ProgressEmitter } from "@archive/core/progress";
import { afterEach, describe, expect, it, vi } from "vitest";

describe("ProgressEmitter", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("should batch updates in a microtask when intervalMs=0", async () => {
    const seen: number[] = [];
    const emitter = new ProgressEmitter(
      { n: 0 },
      s => {
        seen.push(s.n);
      },
      { intervalMs: 0 }
    );

    emitter.update({ n: 1 });
    emitter.update({ n: 2 });
    emitter.update({ n: 3 });

    expect(seen).toEqual([]);

    await new Promise<void>(resolve => queueMicrotask(resolve));
    expect(seen).toEqual([3]);

    emitter.update({ n: 4 });
    await new Promise<void>(resolve => queueMicrotask(resolve));
    expect(seen).toEqual([3, 4]);
  });

  it("emitNow should cancel pending microtask emit when intervalMs=0", async () => {
    let calls = 0;
    const emitter = new ProgressEmitter(
      { n: 0 },
      () => {
        calls++;
      },
      { intervalMs: 0 }
    );

    emitter.update({ n: 1 });
    emitter.emitNow();
    expect(calls).toBe(1);

    await new Promise<void>(resolve => queueMicrotask(resolve));
    expect(calls).toBe(1);
  });

  it("should throttle updates by intervalMs and always emit latest snapshot", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const seen: Array<{ n: number }> = [];
    const emitter = new ProgressEmitter(
      { n: 0 },
      s => {
        seen.push(s);
      },
      { intervalMs: 50 }
    );

    emitter.update({ n: 1 });
    emitter.update({ n: 2 });
    emitter.update({ n: 3 });

    expect(seen.length).toBe(0);

    await vi.advanceTimersByTimeAsync(49);
    expect(seen.length).toBe(0);

    await vi.advanceTimersByTimeAsync(1);
    expect(seen.length).toBe(1);
    expect(seen[0]!.n).toBe(3);

    emitter.update({ n: 4 });
    await vi.advanceTimersByTimeAsync(50);

    expect(seen.length).toBe(2);
    expect(seen[1]!.n).toBe(4);
  });

  it("emitNow should cancel pending scheduled emits", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    let calls = 0;
    const emitter = new ProgressEmitter(
      { n: 0 },
      () => {
        calls++;
      },
      { intervalMs: 50 }
    );

    emitter.update({ n: 1 });
    emitter.emitNow();

    expect(calls).toBe(1);

    await vi.advanceTimersByTimeAsync(500);
    expect(calls).toBe(1);
  });
});
