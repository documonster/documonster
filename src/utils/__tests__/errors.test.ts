/**
 * Tests for error utilities and base error classes.
 */

import {
  BaseError,
  AbortError,
  createAbortError,
  isAbortError,
  throwIfAborted,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@utils/errors";
import { describe, it, expect } from "vitest";

describe("BaseError", () => {
  it("should create error with message", () => {
    const error = new BaseError("test message");
    expect(error.message).toBe("test message");
    expect(error.name).toBe("BaseError");
    expect(error).toBeInstanceOf(Error);
    expect(error).toBeInstanceOf(BaseError);
  });

  it("should support ES2022 error cause", () => {
    const cause = new Error("original error");
    const error = new BaseError("wrapped error", { cause });
    expect(error.message).toBe("wrapped error");
    expect(error.cause).toBe(cause);
  });

  it("should have correct prototype chain for instanceof", () => {
    class CustomError extends BaseError {
      constructor(message: string) {
        super(message);
        this.name = "CustomError";
      }
    }

    const error = new CustomError("test");
    expect(error).toBeInstanceOf(Error);
    expect(error).toBeInstanceOf(BaseError);
    expect(error).toBeInstanceOf(CustomError);
  });

  it("should serialize to JSON", () => {
    const error = new BaseError("test message");
    const json = error.toJSON();
    expect(json.name).toBe("BaseError");
    expect(json.message).toBe("test message");
    expect(json.stack).toBeDefined();
  });

  it("should serialize cause chain to JSON", () => {
    const cause = new BaseError("cause message");
    const error = new BaseError("main message", { cause });
    const json = error.toJSON();
    expect(json.cause).toBeDefined();
    expect((json.cause as any).message).toBe("cause message");
  });
});

describe("AbortError", () => {
  it("should create with no reason", () => {
    const error = new AbortError();
    expect(error.message).toBe("The operation was aborted");
    expect(error.name).toBe("AbortError");
    expect(error.code).toBe("ABORT_ERR");
    expect(error.cause).toBeUndefined();
  });

  it("should create with string reason as cause", () => {
    const error = new AbortError("user cancelled");
    expect(error.message).toBe("The operation was aborted");
    expect(error.cause).toBe("user cancelled");
  });

  it("should create with Error reason as cause", () => {
    const cause = new Error("timeout");
    const error = new AbortError(cause);
    expect(error.message).toBe("The operation was aborted");
    expect(error.cause).toBe(cause);
  });
});

describe("createAbortError", () => {
  it("should return same instance if already AbortError", () => {
    const original = new AbortError();
    const result = createAbortError(original);
    expect(result).toBe(original);
  });

  it("should wrap string reason as cause", () => {
    const error = createAbortError("cancelled");
    expect(error).toBeInstanceOf(AbortError);
    expect(error.message).toBe("The operation was aborted");
    expect(error.cause).toBe("cancelled");
  });

  it("should preserve Error as cause", () => {
    const original = new Error("original");
    const error = createAbortError(original);
    expect(error).toBeInstanceOf(AbortError);
    expect(error.message).toBe("The operation was aborted");
    expect(error.cause).toBe(original);
  });
});

describe("isAbortError", () => {
  it("should return true for AbortError", () => {
    expect(isAbortError(new AbortError())).toBe(true);
  });

  it("should return true for any error with name 'AbortError'", () => {
    const error = new Error("test");
    error.name = "AbortError";
    expect(isAbortError(error)).toBe(true);
  });

  it("should return false for other errors", () => {
    expect(isAbortError(new Error("test"))).toBe(false);
    expect(isAbortError(new BaseError("test"))).toBe(false);
  });

  it("should return false for non-errors", () => {
    expect(isAbortError(null)).toBe(false);
    expect(isAbortError(undefined)).toBe(false);
    expect(isAbortError("error")).toBe(false);
    expect(isAbortError({ message: "error" })).toBe(false);
  });
});

describe("throwIfAborted", () => {
  it("should not throw if signal is undefined", () => {
    expect(() => throwIfAborted(undefined)).not.toThrow();
  });

  it("should not throw if signal is not aborted", () => {
    const controller = new AbortController();
    expect(() => throwIfAborted(controller.signal)).not.toThrow();
  });

  it("should throw AbortError if signal is aborted", () => {
    const controller = new AbortController();
    controller.abort("user cancelled");
    expect(() => throwIfAborted(controller.signal)).toThrow(AbortError);
  });

  it("should use custom reason if provided", () => {
    const controller = new AbortController();
    controller.abort();
    try {
      throwIfAborted(controller.signal, "custom reason");
      expect.fail("should have thrown");
    } catch (e) {
      expect(e).toBeInstanceOf(AbortError);
      expect((e as AbortError).cause).toBe("custom reason");
    }
  });
});

describe("toError", () => {
  it("should return Error as-is", () => {
    const error = new Error("test");
    expect(toError(error)).toBe(error);
  });

  it("should wrap string in Error", () => {
    const error = toError("string message");
    expect(error).toBeInstanceOf(Error);
    expect(error.message).toBe("string message");
  });

  it("should wrap number in Error", () => {
    const error = toError(42);
    expect(error).toBeInstanceOf(Error);
    expect(error.message).toBe("42");
  });

  it("should wrap null/undefined", () => {
    expect(toError(null).message).toBe("null");
    expect(toError(undefined).message).toBe("undefined");
  });

  it("should wrap objects", () => {
    const error = toError({ foo: "bar" });
    expect(error).toBeInstanceOf(Error);
    expect(error.message).toBe("[object Object]");
  });
});

describe("errorToJSON", () => {
  it("should serialize BaseError", () => {
    const error = new BaseError("test");
    const json = errorToJSON(error);
    expect(json.name).toBe("BaseError");
    expect(json.message).toBe("test");
    expect(json.stack).toBeDefined();
  });

  it("should serialize native Error", () => {
    const error = new Error("native");
    const json = errorToJSON(error);
    expect(json.name).toBe("Error");
    expect(json.message).toBe("native");
  });

  it("should serialize nested cause chain", () => {
    const root = new Error("root");
    const middle = new BaseError("middle", { cause: root });
    const top = new BaseError("top", { cause: middle });

    const json = errorToJSON(top);
    expect(json.message).toBe("top");
    expect((json.cause as any).message).toBe("middle");
    expect((json.cause as any).cause.message).toBe("root");
  });
});

describe("getErrorChain", () => {
  it("should return single error for no cause", () => {
    const error = new BaseError("single");
    const chain = getErrorChain(error);
    expect(chain).toHaveLength(1);
    expect(chain[0]).toBe(error);
  });

  it("should return full cause chain", () => {
    const root = new Error("root");
    const middle = new BaseError("middle", { cause: root });
    const top = new BaseError("top", { cause: middle });

    const chain = getErrorChain(top);
    expect(chain).toHaveLength(3);
    expect(chain[0]).toBe(top);
    expect(chain[1]).toBe(middle);
    expect(chain[2]).toBe(root);
  });

  it("should stop at non-Error cause", () => {
    const error = new BaseError("test", { cause: "string cause" });
    const chain = getErrorChain(error);
    expect(chain).toHaveLength(1);
  });
});

describe("getRootCause", () => {
  it("should return self for no cause", () => {
    const error = new BaseError("single");
    expect(getRootCause(error)).toBe(error);
  });

  it("should return deepest error in chain", () => {
    const root = new Error("root");
    const middle = new BaseError("middle", { cause: root });
    const top = new BaseError("top", { cause: middle });

    expect(getRootCause(top)).toBe(root);
  });

  it("should stop at non-Error cause", () => {
    const middle = new BaseError("middle", { cause: "string" });
    const top = new BaseError("top", { cause: middle });

    expect(getRootCause(top)).toBe(middle);
  });
});
