/**
 * Shared tests for stream/errors.ts
 *
 * Tests for stream-specific error classes.
 * Platform-agnostic — imported by both Node.js and browser test runners.
 */

import {
  StreamError,
  StreamStateError,
  StreamArgumentError,
  StreamUnsupportedTypeError
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
  // StreamArgumentError
  // ===========================================================================
  describe("StreamArgumentError", () => {
    it("should have correct name", () => {
      const err = new StreamArgumentError("Uint8Array", "string");
      expect(err.name).toBe("StreamArgumentError");
    });

    it("should format message from expected and actual types", () => {
      const err = new StreamArgumentError("Uint8Array", "string");
      expect(err.message).toBe("Expected Uint8Array, got string");
    });

    it("should expose expectedType and actualType properties", () => {
      const err = new StreamArgumentError("Buffer", "number");
      expect(err.expectedType).toBe("Buffer");
      expect(err.actualType).toBe("number");
    });

    it("should be instanceof StreamError", () => {
      const err = new StreamArgumentError("a", "b");
      expect(err).toBeInstanceOf(StreamError);
    });

    it("should support cause option", () => {
      const cause = new TypeError("original");
      const err = new StreamArgumentError("a", "b", { cause });
      expect(err.cause).toBe(cause);
    });
  });

  // ===========================================================================
  // StreamUnsupportedTypeError
  // ===========================================================================
  describe("StreamUnsupportedTypeError", () => {
    it("should have correct name", () => {
      const err = new StreamUnsupportedTypeError("pipeline", "WebStream");
      expect(err.name).toBe("StreamUnsupportedTypeError");
    });

    it("should format message from operation and streamType", () => {
      const err = new StreamUnsupportedTypeError("pipeline", "WebStream");
      expect(err.message).toBe('pipeline: unsupported stream type "WebStream"');
    });

    it("should expose operation and streamType properties", () => {
      const err = new StreamUnsupportedTypeError("finished", "CustomStream");
      expect(err.operation).toBe("finished");
      expect(err.streamType).toBe("CustomStream");
    });

    it("should be instanceof StreamError", () => {
      const err = new StreamUnsupportedTypeError("op", "type");
      expect(err).toBeInstanceOf(StreamError);
    });

    it("should support cause option", () => {
      const cause = new Error("original");
      const err = new StreamUnsupportedTypeError("op", "type", { cause });
      expect(err.cause).toBe(cause);
    });
  });
}
