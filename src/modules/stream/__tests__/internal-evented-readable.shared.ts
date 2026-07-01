/**
 * Shared tests for internal/evented-readable-to-async-iterable.ts
 *
 * Tests for eventedReadableToAsyncIterableNoDestroy() which converts
 * an evented readable into an AsyncIterable without destroying the stream.
 * Platform-agnostic — imported by both Node.js and browser test runners.
 */

import type { EventedReadableLike } from "@stream/core/evented-readable-to-async-iterable";
import { eventedReadableToAsyncIterableNoDestroy } from "@stream/core/evented-readable-to-async-iterable";
import { describe, it, expect } from "vitest";

/**
 * Minimal mock of an evented readable stream.
 */
function createMockReadable<T>(chunks: T[]): EventedReadableLike<T> & {
  emit(event: string, ...args: any[]): void;
  emitChunks(): void;
  isPaused: boolean;
} {
  const listeners = new Map<string, Set<(...args: any[]) => void>>();
  let paused = true;

  const mock = {
    get isPaused() {
      return paused;
    },
    on(event: string, listener: (...args: any[]) => void) {
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
      return mock;
    },
    off(event: string, listener: (...args: any[]) => void) {
      listeners.get(event)?.delete(listener);
      return mock;
    },
    pause() {
      paused = true;
      return mock;
    },
    resume() {
      paused = false;
      return mock;
    },
    emit(event: string, ...args: any[]) {
      const set = listeners.get(event);
      if (!set) {
        return;
      }
      for (const fn of [...set]) {
        fn(...args);
      }
    },
    emitChunks() {
      for (const chunk of chunks) {
        mock.emit("data", chunk);
      }
      mock.emit("end");
    }
  };

  return mock;
}

export function runInternalEventedReadableTests(): void {
  describe("eventedReadableToAsyncIterableNoDestroy", () => {
    it("should iterate all chunks from the stream", async () => {
      const readable = createMockReadable([1, 2, 3]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);

      // Start emitting after a microtask so the iterator has time to set up
      queueMicrotask(() => readable.emitChunks());

      const result: number[] = [];
      for await (const chunk of iterable) {
        result.push(chunk);
      }

      expect(result).toEqual([1, 2, 3]);
    });

    it("should handle empty stream", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);

      queueMicrotask(() => readable.emit("end"));

      const result: number[] = [];
      for await (const chunk of iterable) {
        result.push(chunk);
      }

      expect(result).toEqual([]);
    });

    it("should pause after each data event (backpressure)", async () => {
      const readable = createMockReadable<string>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<string>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      // Emit a chunk — should pause
      queueMicrotask(() => readable.emit("data", "a"));
      const r1 = await iter.next();
      expect(r1).toEqual({ value: "a", done: false });
      expect(readable.isPaused).toBe(true);

      // Emit end
      queueMicrotask(() => readable.emit("end"));
      const r2 = await iter.next();
      expect(r2.done).toBe(true);
    });

    it("should reject on error event", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      const error = new Error("stream error");
      queueMicrotask(() => readable.emit("error", error));

      await expect(iter.next()).rejects.toBe(error);
    });

    it("should reject subsequent next() calls after error", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      const error = new Error("stream error");
      queueMicrotask(() => readable.emit("error", error));
      await iter.next().catch(() => {});

      // Subsequent calls should also reject
      await expect(iter.next()).rejects.toBe(error);
    });

    it("should handle close event same as end", async () => {
      const readable = createMockReadable<string>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<string>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      queueMicrotask(() => readable.emit("close"));

      const result = await iter.next();
      expect(result.done).toBe(true);
    });

    it("should support early break (return) without destroying stream", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      // Emit some data
      queueMicrotask(() => readable.emit("data", 1));
      await iter.next();

      // Early return
      const returnResult = await iter.return!();
      expect(returnResult.done).toBe(true);

      // Stream should NOT be destroyed — just paused
      expect(readable.isPaused).toBe(true);
    });

    it("should handle throw", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      const error = new Error("thrown");
      await expect(iter.throw!(error)).rejects.toBe(error);
    });

    it("should buffer chunks and deliver them in order", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      // Emit multiple chunks before consuming
      readable.emit("data", 1);
      readable.emit("data", 2);
      readable.emit("data", 3);

      const r1 = await iter.next();
      const r2 = await iter.next();
      const r3 = await iter.next();

      expect(r1.value).toBe(1);
      expect(r2.value).toBe(2);
      expect(r3.value).toBe(3);
    });

    it("should compact buffer after head exceeds threshold", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      // Emit 100 chunks to trigger compaction (head > 64 && head * 2 > chunks.length)
      for (let i = 0; i < 100; i++) {
        readable.emit("data", i);
      }

      // Consume all 100
      for (let i = 0; i < 100; i++) {
        const r = await iter.next();
        expect(r.value).toBe(i);
      }

      // Emit end
      queueMicrotask(() => readable.emit("end"));
      const r = await iter.next();
      expect(r.done).toBe(true);
    });

    it("should resume on next() when no buffered data", async () => {
      const readable = createMockReadable<number>([]);
      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(readable);
      const iter = iterable[Symbol.asyncIterator]();

      // Start a next() call — should resume to get data
      const promise = iter.next();
      // Wait a tick for resume to be called
      await new Promise(r => setTimeout(r, 0));
      expect(readable.isPaused).toBe(false);

      // Now emit data to resolve it
      readable.emit("data", 42);
      const result = await promise;
      expect(result.value).toBe(42);
    });

    it("should work with removeListener-only emitter", async () => {
      const listeners = new Map<string, Set<(...args: any[]) => void>>();
      const emitter: EventedReadableLike<number> & { emit(e: string, ...a: any[]): void } = {
        on(event: string, listener: any) {
          if (!listeners.has(event)) {
            listeners.set(event, new Set());
          }
          listeners.get(event)!.add(listener);
        },
        removeListener(event: string, listener: any) {
          listeners.get(event)?.delete(listener);
        },
        pause() {},
        resume() {},
        emit(event: string, ...args: any[]) {
          const set = listeners.get(event);
          if (!set) {
            return;
          }
          for (const fn of [...set]) {
            fn(...args);
          }
        }
      };

      const iterable = eventedReadableToAsyncIterableNoDestroy<number>(emitter);
      const iter = iterable[Symbol.asyncIterator]();

      queueMicrotask(() => {
        emitter.emit("data", 10);
        emitter.emit("end");
      });

      const r1 = await iter.next();
      expect(r1.value).toBe(10);

      const r2 = await iter.next();
      expect(r2.done).toBe(true);
    });
  });
}
