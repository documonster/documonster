/**
 * Tests for archive error classes.
 */

import {
  ArchiveError,
  AbortError,
  ZipParseError,
  InvalidZipSignatureError,
  EocdNotFoundError,
  Crc32MismatchError,
  DecryptionError,
  PasswordRequiredError,
  RangeNotSupportedError,
  HttpRangeError,
  FileTooLargeError,
  UnsupportedCompressionError,
  toError,
  suppressUnhandledRejection
} from "@archive/shared/errors";
import { describe, it, expect } from "vitest";

describe("errors", () => {
  describe("ArchiveError", () => {
    it("should create error with message", () => {
      const error = new ArchiveError("test message");
      expect(error.message).toBe("test message");
      expect(error.name).toBe("ArchiveError");
      expect(error).toBeInstanceOf(Error);
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

    it("should create with string reason", () => {
      const error = new AbortError("user cancelled");
      expect(error.message).toBe("The operation was aborted");
      expect(error.cause).toBe("user cancelled");
    });

    it("should create with Error reason", () => {
      const cause = new Error("timeout");
      const error = new AbortError(cause);
      expect(error.message).toBe("The operation was aborted");
      expect(error.cause).toBe(cause);
    });
  });

  describe("ZipParseError", () => {
    it("should be instanceof ArchiveError", () => {
      const error = new ZipParseError("parse failed");
      expect(error).toBeInstanceOf(ArchiveError);
      expect(error.name).toBe("ZipParseError");
    });
  });

  describe("InvalidZipSignatureError", () => {
    it("should format signature mismatch without context", () => {
      const error = new InvalidZipSignatureError("0x04034b50", 0x12345678);
      expect(error.message).toBe("Invalid signature: expected 0x04034b50, got 0x12345678");
      expect(error.name).toBe("InvalidZipSignatureError");
    });

    it("should format signature mismatch with context", () => {
      const error = new InvalidZipSignatureError("0x04034b50", 0x12345678, "local file header");
      expect(error.message).toBe("Invalid local file header: expected 0x04034b50, got 0x12345678");
    });

    it("should pad signature to 8 hex digits", () => {
      const error = new InvalidZipSignatureError("0x04034b50", 0x00000001);
      expect(error.message).toContain("0x00000001");
    });
  });

  describe("EocdNotFoundError", () => {
    it("should have standard message", () => {
      const error = new EocdNotFoundError();
      expect(error.message).toBe("Invalid ZIP file: End of Central Directory not found");
      expect(error.name).toBe("EocdNotFoundError");
      expect(error).toBeInstanceOf(ZipParseError);
    });
  });

  describe("Crc32MismatchError", () => {
    it("should format CRC32 mismatch with path and values", () => {
      const error = new Crc32MismatchError("test.txt", 0xaabbccdd, 0x11223344);
      expect(error.message).toBe(
        'CRC32 mismatch for "test.txt": expected 0xaabbccdd, got 0x11223344'
      );
      expect(error.name).toBe("Crc32MismatchError");
      expect(error.path).toBe("test.txt");
      expect(error.expected).toBe(0xaabbccdd);
      expect(error.actual).toBe(0x11223344);
    });

    it("should pad CRC32 to 8 hex digits", () => {
      const error = new Crc32MismatchError("file.bin", 0x00000001, 0x00000002);
      expect(error.message).toContain("0x00000001");
      expect(error.message).toContain("0x00000002");
    });
  });

  describe("DecryptionError", () => {
    it("should format without details", () => {
      const error = new DecryptionError("secret.txt");
      expect(error.message).toBe(
        'Failed to decrypt "secret.txt": incorrect password or corrupted data'
      );
      expect(error.name).toBe("DecryptionError");
    });

    it("should format with details", () => {
      const error = new DecryptionError("secret.txt", "AES key derivation failed");
      expect(error.message).toBe('Failed to decrypt "secret.txt": AES key derivation failed');
    });
  });

  describe("PasswordRequiredError", () => {
    it("should format with path", () => {
      const error = new PasswordRequiredError("encrypted.zip");
      expect(error.message).toBe(
        'File "encrypted.zip" is encrypted. Please provide a password to extract.'
      );
      expect(error.name).toBe("PasswordRequiredError");
    });
  });

  describe("RangeNotSupportedError", () => {
    it("should format with URL", () => {
      const error = new RangeNotSupportedError("https://example.com/file.zip");
      expect(error.message).toBe(
        "Server does not support Range requests for: https://example.com/file.zip"
      );
      expect(error.name).toBe("RangeNotSupportedError");
    });
  });

  describe("HttpRangeError", () => {
    it("should format with URL and status", () => {
      const error = new HttpRangeError("https://example.com/file.zip", 404, "Not Found");
      expect(error.message).toBe("HTTP 404 Not Found for: https://example.com/file.zip");
      expect(error.name).toBe("HttpRangeError");
      expect(error.url).toBe("https://example.com/file.zip");
      expect(error.status).toBe(404);
      expect(error.statusText).toBe("Not Found");
    });
  });

  describe("FileTooLargeError", () => {
    it("should format with path and reason", () => {
      const error = new FileTooLargeError("huge.bin", "exceeds 100MB limit");
      expect(error.message).toBe(
        'File "huge.bin" is too large to extract into memory (exceeds 100MB limit)'
      );
      expect(error.name).toBe("FileTooLargeError");
    });
  });

  describe("UnsupportedCompressionError", () => {
    it("should format with method number", () => {
      const error = new UnsupportedCompressionError(99);
      expect(error.message).toBe("Unsupported compression method: 99");
      expect(error.name).toBe("UnsupportedCompressionError");
    });
  });

  describe("toError", () => {
    it("should return Error as-is", () => {
      const original = new Error("test");
      const result = toError(original);
      expect(result).toBe(original);
    });

    it("should wrap string in Error", () => {
      const result = toError("string error");
      expect(result).toBeInstanceOf(Error);
      expect(result.message).toBe("string error");
    });

    it("should wrap number in Error", () => {
      const result = toError(42);
      expect(result).toBeInstanceOf(Error);
      expect(result.message).toBe("42");
    });

    it("should wrap null in Error", () => {
      const result = toError(null);
      expect(result).toBeInstanceOf(Error);
      expect(result.message).toBe("null");
    });

    it("should wrap undefined in Error", () => {
      const result = toError(undefined);
      expect(result).toBeInstanceOf(Error);
      expect(result.message).toBe("undefined");
    });

    it("should wrap object in Error", () => {
      const result = toError({ code: "ERR" });
      expect(result).toBeInstanceOf(Error);
      expect(result.message).toBe("[object Object]");
    });
  });

  describe("toError (re-exported)", () => {
    it("should be the same toError from @utils/errors", () => {
      expect(toError).toBeTypeOf("function");
    });
  });

  describe("suppressUnhandledRejection", () => {
    it("should suppress rejection warning", async () => {
      const rejectedPromise = Promise.reject(new Error("test"));
      suppressUnhandledRejection(rejectedPromise);
      // If this doesn't throw unhandled rejection warning, the test passes
      await new Promise(resolve => setTimeout(resolve, 10));
    });

    it("should not affect resolved promises", async () => {
      const resolvedPromise = Promise.resolve("value");
      suppressUnhandledRejection(resolvedPromise);
      await expect(resolvedPromise).resolves.toBe("value");
    });
  });
});
