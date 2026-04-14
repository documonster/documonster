import { isAbortError } from "@archive";
import { toAsyncIterable } from "@archive/io/archive-source";
import { describe, expect, it } from "vitest";

describe("archive-source", () => {
  it("toAsyncIterable(Uint8Array) should yield once and call onChunk", async () => {
    const src = new Uint8Array([1, 2, 3]);

    const seen: number[] = [];
    const out: Uint8Array[] = [];

    for await (const chunk of toAsyncIterable(src, {
      onChunk: c => {
        seen.push(c.length);
      }
    })) {
      out.push(chunk);
    }

    expect(seen).toEqual([3]);
    expect(out).toEqual([src]);
  });

  it("toAsyncIterable(Uint8Array) should throw immediately when already aborted", async () => {
    const ac = new AbortController();
    ac.abort("stop");

    const iter = toAsyncIterable(new Uint8Array([1]), { signal: ac.signal });
    const itor = iter[Symbol.asyncIterator]();

    await expect(itor.next()).rejects.toSatisfy((e: unknown) => {
      return isAbortError(e) && (e as any).cause === "stop";
    });
  });

  it("toAsyncIterable(ReadableStream) should cancel reader on abort and throw AbortError", async () => {
    let cancelCalled = 0;

    const rs = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([1]));
      },
      pull(controller) {
        controller.enqueue(new Uint8Array([2]));
      },
      cancel() {
        cancelCalled++;
      }
    });

    const ac = new AbortController();
    const iter = toAsyncIterable(rs, { signal: ac.signal });
    const itor = iter[Symbol.asyncIterator]();

    const first = await itor.next();
    expect(first.done).toBe(false);
    expect(first.value).toBeInstanceOf(Uint8Array);

    ac.abort("stop");
    // Let the abort handler run synchronously.
    await Promise.resolve();

    expect(cancelCalled).toBe(1);

    await expect(itor.next()).rejects.toSatisfy((e: unknown) => {
      return isAbortError(e) && (e as any).cause === "stop";
    });
  });

  it("toAsyncIterable(ReadableStream) should detach abort handler after completion", async () => {
    let cancelCalled = 0;

    const rs = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([1]));
        controller.close();
      },
      cancel() {
        cancelCalled++;
      }
    });

    const ac = new AbortController();
    const out: number[] = [];

    for await (const chunk of toAsyncIterable(rs, { signal: ac.signal })) {
      out.push(chunk[0]!);
    }

    expect(out).toEqual([1]);
    expect(cancelCalled).toBe(0);

    // After completion, abort should NOT call reader.cancel().
    ac.abort("stop");
    await Promise.resolve();
    expect(cancelCalled).toBe(0);
  });

  it("toAsyncIterable(AsyncIterable) should throw immediately when already aborted", async () => {
    const ac = new AbortController();
    ac.abort("stop");

    const src: AsyncIterable<Uint8Array> = {
      async *[Symbol.asyncIterator]() {
        yield new Uint8Array([1]);
      }
    };

    const iter = toAsyncIterable(src, { signal: ac.signal });
    const itor = iter[Symbol.asyncIterator]();

    await expect(itor.next()).rejects.toSatisfy((e: unknown) => {
      return isAbortError(e) && (e as any).cause === "stop";
    });
  });

  it("toAsyncIterable(AsyncIterable) should call return() on abort to clean up upstream", async () => {
    let returns = 0;
    let nexts = 0;

    const upstream: AsyncIterable<Uint8Array> = {
      [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
        return {
          async next() {
            nexts++;
            return { value: new Uint8Array([nexts]), done: false };
          },
          async return() {
            returns++;
            return { value: undefined as any, done: true };
          }
        };
      }
    };

    const ac = new AbortController();
    const iter = toAsyncIterable(upstream, { signal: ac.signal });
    const itor = iter[Symbol.asyncIterator]();

    const first = await itor.next();
    expect(first.done).toBe(false);
    expect(first.value).toBeInstanceOf(Uint8Array);

    ac.abort("stop");

    await expect(itor.next()).rejects.toSatisfy((e: unknown) => {
      return isAbortError(e) && (e as any).cause === "stop";
    });

    expect(returns).toBe(1);
    expect(nexts).toBeGreaterThanOrEqual(1);
  });
});
