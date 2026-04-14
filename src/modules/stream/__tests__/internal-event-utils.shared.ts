/**
 * Shared tests for internal/event-utils.ts
 *
 * Tests for onceEvent() promise-based event listener.
 * Platform-agnostic — imported by both Node.js and browser test runners.
 */

import { onceEvent } from "@stream/internal/event-utils";
import type { EventEmitterLike } from "@stream/types";
import { describe, it, expect } from "vitest";

/**
 * Minimal emitter with `once` support (like Node EventEmitter).
 */
function createEmitter(): EventEmitterLike & { emit(event: string, ...args: any[]): void } {
  const listeners = new Map<string, Set<(...args: any[]) => void>>();
  const onceListeners = new Set<(...args: any[]) => void>();

  return {
    on(event: string, listener: (...args: any[]) => void) {
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
    },
    once(event: string, listener: (...args: any[]) => void) {
      onceListeners.add(listener);
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
    },
    off(event: string, listener: (...args: any[]) => void) {
      listeners.get(event)?.delete(listener);
      onceListeners.delete(listener);
    },
    emit(event: string, ...args: any[]) {
      const set = listeners.get(event);
      if (!set) {
        return;
      }
      for (const fn of [...set]) {
        fn(...args);
        if (onceListeners.has(fn)) {
          set.delete(fn);
          onceListeners.delete(fn);
        }
      }
    }
  };
}

/**
 * Emitter without `once` — only `on` and `off`.
 */
function createSimpleEmitter(): EventEmitterLike & { emit(event: string, ...args: any[]): void } {
  const listeners = new Map<string, Set<(...args: any[]) => void>>();

  return {
    on(event: string, listener: (...args: any[]) => void) {
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
    },
    off(event: string, listener: (...args: any[]) => void) {
      listeners.get(event)?.delete(listener);
    },
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
}

/**
 * Emitter with `removeListener` instead of `off`.
 */
function createRemoveListenerEmitter(): EventEmitterLike & {
  emit(event: string, ...args: any[]): void;
} {
  const listeners = new Map<string, Set<(...args: any[]) => void>>();

  return {
    on(event: string, listener: (...args: any[]) => void) {
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
    },
    once(event: string, listener: (...args: any[]) => void) {
      if (!listeners.has(event)) {
        listeners.set(event, new Set());
      }
      listeners.get(event)!.add(listener);
    },
    removeListener(event: string, listener: (...args: any[]) => void) {
      listeners.get(event)?.delete(listener);
    },
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
}

export function runInternalEventUtilsTests(): void {
  describe("onceEvent", () => {
    it("should resolve when the event fires", async () => {
      const emitter = createEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("done");
      await expect(p).resolves.toBeUndefined();
    });

    it("should reject when error event fires", async () => {
      const emitter = createEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("error", new Error("boom"));
      await expect(p).rejects.toThrow("boom");
    });

    it("should reject with wrapped string error", async () => {
      const emitter = createEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("error", "string error");
      await expect(p).rejects.toThrow("string error");
    });

    it("should clean up listeners after event fires", async () => {
      const emitter = createEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("done");
      await p;
      // Emitting error after resolve should not throw
      emitter.emit("error", new Error("late error"));
    });

    it("should clean up listeners after error", async () => {
      const emitter = createEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("error", new Error("err"));
      await p.catch(() => {});
      // Emitting done after reject should not affect anything
      emitter.emit("done");
    });

    it("should work with emitter without once (fallback to on)", async () => {
      const emitter = createSimpleEmitter();
      const p = onceEvent(emitter, "finish");
      emitter.emit("finish");
      await expect(p).resolves.toBeUndefined();
    });

    it("should reject with emitter without once on error", async () => {
      const emitter = createSimpleEmitter();
      const p = onceEvent(emitter, "finish");
      emitter.emit("error", new Error("fail"));
      await expect(p).rejects.toThrow("fail");
    });

    it("should work with emitter using removeListener instead of off", async () => {
      const emitter = createRemoveListenerEmitter();
      const p = onceEvent(emitter, "done");
      emitter.emit("done");
      await expect(p).resolves.toBeUndefined();
    });
  });
}
