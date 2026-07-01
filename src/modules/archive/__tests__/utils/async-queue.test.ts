import { createAsyncQueue } from "@archive/core/async-queue";
import { describe, expect, it } from "vitest";

describe("async-queue", () => {
  it("return() should cancel, resolve pending next(), and call onCancel once", async () => {
    let cancels = 0;
    const q = createAsyncQueue<number>({
      onCancel: () => {
        cancels++;
      }
    });

    const itor = q.iterable[Symbol.asyncIterator]();

    const pendingNext = itor.next();
    const ret = await itor.return!();

    expect(ret.done).toBe(true);
    expect(cancels).toBe(1);

    await expect(pendingNext).resolves.toEqual({ value: undefined, done: true });
    await expect(itor.next()).resolves.toEqual({ value: undefined, done: true });

    // Should ignore pushes after cancellation.
    q.push(123);
    await expect(itor.next()).resolves.toEqual({ value: undefined, done: true });

    // Idempotent.
    await itor.return!();
    expect(cancels).toBe(1);
  });

  it("throw() should cancel, resolve pending next(), and reject with the provided error", async () => {
    let cancels = 0;
    const q = createAsyncQueue<number>({
      onCancel: () => {
        cancels++;
      }
    });

    const itor = q.iterable[Symbol.asyncIterator]();

    const pendingNext = itor.next();
    const err = new Error("boom");

    await expect(itor.throw!(err)).rejects.toBe(err);
    expect(cancels).toBe(1);

    await expect(pendingNext).resolves.toEqual({ value: undefined, done: true });
    await expect(itor.next()).resolves.toEqual({ value: undefined, done: true });
  });

  it("fail() should reject pending next() and all future next() calls", async () => {
    const q = createAsyncQueue<number>();
    const itor = q.iterable[Symbol.asyncIterator]();

    const pendingNext = itor.next();
    q.fail(new Error("x"));

    await expect(pendingNext).rejects.toThrow("x");
    await expect(itor.next()).rejects.toThrow("x");

    // close() after fail() should be ignored.
    q.close();
    await expect(itor.next()).rejects.toThrow("x");
  });
});
