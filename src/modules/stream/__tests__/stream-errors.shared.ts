/**
 * Shared tests for stream/errors.ts
 *
 * Tests for stream-specific error classes.
 * Platform-agnostic — imported by both Node.js and browser test runners.
 */

import {
  StreamError,
  StreamStateError,
  StreamTypeError,
  UnsupportedStreamTypeError
} from "@stream/errors";
import { BaseError } from "@utils/errors";
import { describe, it, expect } from "vitest";

export function runStreamErrorsTests(): void {
  // ===========================================================================
  // StreamError
  // ===========================================================================
  describe("StreamError", () => {
    it("should have correct name", () => {
      const err = new StreamError("test");
      expect(err.name).toBe("StreamError");
    });

    it("should have correct message", () => {
      const err = new StreamError("something broke");
      expect(err.message).toBe("something broke");
    });

    it("should be instanceof Error", () => {
      const err = new StreamError("test");
      expect(err).toBeInstanceOf(Error);
    });

    it("should be instanceof BaseError", () => {
      const err = new StreamError("test");
      expect(err).toBeInstanceOf(BaseError);
    });

    it("should support cause option", () => {
      const cause = new Error("root cause");
      const err = new StreamError("wrapped", { cause });
      expect(err.cause).toBe(cause);
    });
  });

  // ===========================================================================
  // StreamStateError
  // ===========================================================================
  describe("StreamStateError", () => {
    it("should have correct name", () => {
      const err = new StreamStateError("write", "stream is closed");
      expect(err.name).toBe("StreamStateError");
    });

    it("should format message from operation and state", () => {
      const err = new StreamStateError("write", "stream is closed");
      expect(err.message).toBe("Cannot write: stream is closed");
    });

    it("should expose operation and state properties", () => {
      const err = new StreamStateError("read", "stream is not readable");
      expect(err.operation).toBe("read");
      expect(err.state).toBe("stream is not readable");
    });

    it("should be instanceof StreamError", () => {
      const err = new StreamStateError("op", "state");
      expect(err).toBeInstanceOf(StreamError);
    });

    it("should support cause option", () => {
      const cause = new Error("original");
      const err = new StreamStateError("push", "buffer full", { cause });
      expect(err.cause).toBe(cause);
    });
  });

  // ===========================================================================
  // StreamTypeError
  // ===========================================================================
  describe("StreamTypeError", () => {
    it("should have correct name", () => {
      const err = new StreamTypeError("Uint8Array", "string");
      expect(err.name).toBe("StreamTypeError");
    });

    it("should format message from expected and actual types", () => {
      const err = new StreamTypeError("Uint8Array", "string");
      expect(err.message).toBe("Expected Uint8Array, got string");
    });

    it("should expose expectedType and actualType properties", () => {
      const err = new StreamTypeError("Buffer", "number");
      expect(err.expectedType).toBe("Buffer");
      expect(err.actualType).toBe("number");
    });

    it("should be instanceof StreamError", () => {
      const err = new StreamTypeError("a", "b");
      expect(err).toBeInstanceOf(StreamError);
    });

    it("should support cause option", () => {
      const cause = new TypeError("original");
      const err = new StreamTypeError("a", "b", { cause });
      expect(err.cause).toBe(cause);
    });
  });

  // ===========================================================================
  // UnsupportedStreamTypeError
  // ===========================================================================
  describe("UnsupportedStreamTypeError", () => {
    it("should have correct name", () => {
      const err = new UnsupportedStreamTypeError("pipeline", "WebStream");
      expect(err.name).toBe("UnsupportedStreamTypeError");
    });

    it("should format message from operation and streamType", () => {
      const err = new UnsupportedStreamTypeError("pipeline", "WebStream");
      expect(err.message).toBe('pipeline: unsupported stream type "WebStream"');
    });

    it("should expose operation and streamType properties", () => {
      const err = new UnsupportedStreamTypeError("finished", "CustomStream");
      expect(err.operation).toBe("finished");
      expect(err.streamType).toBe("CustomStream");
    });

    it("should be instanceof StreamError", () => {
      const err = new UnsupportedStreamTypeError("op", "type");
      expect(err).toBeInstanceOf(StreamError);
    });

    it("should support cause option", () => {
      const cause = new Error("original");
      const err = new UnsupportedStreamTypeError("op", "type", { cause });
      expect(err.cause).toBe(cause);
    });
  });
}
