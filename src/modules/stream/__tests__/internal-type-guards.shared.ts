/**
 * Shared tests for internal/type-guards.ts
 *
 * Tests for Web Streams API type guard functions.
 * Platform-agnostic — imported by both Node.js and browser test runners.
 */

import {
  isReadableStream,
  isWritableStream,
  isAsyncIterable,
  isTransformStream
} from "@stream/core/type-guards";
import { describe, it, expect } from "vitest";

export function runInternalTypeGuardsTests(): void {
  // ===========================================================================
  // isReadableStream
  // ===========================================================================
  describe("isReadableStream", () => {
    it("should return true for a ReadableStream", () => {
      const rs = new ReadableStream();
      expect(isReadableStream(rs)).toBe(true);
    });

    it("should return true for an object with getReader method", () => {
      const fake = { getReader: () => ({}) };
      expect(isReadableStream(fake)).toBe(true);
    });

    it("should return false for null", () => {
      expect(isReadableStream(null)).toBe(false);
    });

    it("should return false for undefined", () => {
      expect(isReadableStream(undefined)).toBe(false);
    });

    it("should return false for a string", () => {
      expect(isReadableStream("hello")).toBe(false);
    });

    it("should return false for a number", () => {
      expect(isReadableStream(42)).toBe(false);
    });

    it("should return false for a plain object", () => {
      expect(isReadableStream({})).toBe(false);
    });

    it("should return false for an object with non-function getReader", () => {
      expect(isReadableStream({ getReader: true })).toBe(false);
    });
  });

  // ===========================================================================
  // isWritableStream
  // ===========================================================================
  describe("isWritableStream", () => {
    it("should return true for a WritableStream", () => {
      const ws = new WritableStream();
      expect(isWritableStream(ws)).toBe(true);
    });

    it("should return true for an object with getWriter method", () => {
      const fake = { getWriter: () => ({}) };
      expect(isWritableStream(fake)).toBe(true);
    });

    it("should return false for null", () => {
      expect(isWritableStream(null)).toBe(false);
    });

    it("should return false for undefined", () => {
      expect(isWritableStream(undefined)).toBe(false);
    });

    it("should return false for a plain object", () => {
      expect(isWritableStream({})).toBe(false);
    });

    it("should return false for an object with non-function getWriter", () => {
      expect(isWritableStream({ getWriter: "not a function" })).toBe(false);
    });
  });

  // ===========================================================================
  // isAsyncIterable
  // ===========================================================================
  describe("isAsyncIterable", () => {
    it("should return true for an async generator", () => {
      async function* gen() {
        yield 1;
      }
      expect(isAsyncIterable(gen())).toBe(true);
    });

    it("should return true for an object with Symbol.asyncIterator", () => {
      const fake = { [Symbol.asyncIterator]: () => ({}) };
      expect(isAsyncIterable(fake)).toBe(true);
    });

    it("should return true for a function with Symbol.asyncIterator", () => {
      const fn = () => {};
      (fn as any)[Symbol.asyncIterator] = () => ({});
      expect(isAsyncIterable(fn)).toBe(true);
    });

    it("should return false for null", () => {
      expect(isAsyncIterable(null)).toBe(false);
    });

    it("should return false for undefined", () => {
      expect(isAsyncIterable(undefined)).toBe(false);
    });

    it("should return false for a plain object", () => {
      expect(isAsyncIterable({})).toBe(false);
    });

    it("should return false for a regular iterable (sync)", () => {
      const syncIterable = { [Symbol.iterator]: () => ({}) };
      expect(isAsyncIterable(syncIterable)).toBe(false);
    });

    it("should return false for a number", () => {
      expect(isAsyncIterable(42)).toBe(false);
    });

    it("should return false for a string (no asyncIterator)", () => {
      expect(isAsyncIterable("hello")).toBe(false);
    });
  });

  // ===========================================================================
  // isTransformStream
  // ===========================================================================
  describe("isTransformStream", () => {
    it("should return true for a TransformStream", () => {
      const ts = new TransformStream();
      expect(isTransformStream(ts)).toBe(true);
    });

    it("should return true for an object with readable and writable streams", () => {
      const fake = {
        readable: new ReadableStream(),
        writable: new WritableStream()
      };
      expect(isTransformStream(fake)).toBe(true);
    });

    it("should return false for null", () => {
      expect(isTransformStream(null)).toBe(false);
    });

    it("should return false for undefined", () => {
      expect(isTransformStream(undefined)).toBe(false);
    });

    it("should return false for a plain object", () => {
      expect(isTransformStream({})).toBe(false);
    });

    it("should return false when readable is not a ReadableStream", () => {
      const fake = {
        readable: { notAStream: true },
        writable: new WritableStream()
      };
      expect(isTransformStream(fake)).toBe(false);
    });

    it("should return false when writable is not a WritableStream", () => {
      const fake = {
        readable: new ReadableStream(),
        writable: { notAStream: true }
      };
      expect(isTransformStream(fake)).toBe(false);
    });

    it("should return false when readable is falsy", () => {
      const fake = { readable: null, writable: new WritableStream() };
      expect(isTransformStream(fake)).toBe(false);
    });

    it("should return false when writable is falsy", () => {
      const fake = { readable: new ReadableStream(), writable: null };
      expect(isTransformStream(fake)).toBe(false);
    });
  });
}
