import {
  AbortError,
  Archive,
  ArchiveError,
  ArchiveLimitExceededError,
  ZipCrc32MismatchError,
  ZipDecryptionError,
  ZipEntrySizeMismatchError,
  ZipEocdNotFoundError,
  ArchiveFileTooLargeError,
  ArchiveHttpRangeError,
  ZipInvalidSignatureError,
  ZipPasswordRequiredError,
  ArchiveRangeNotSupportedError,
  ZipUnsupportedCompressionError,
  ZipParseError,
  isArchiveError
} from "@archive/index";
import * as browserEntry from "@archive/index.browser";
import { BaseError } from "@utils/errors";
import { describe, expect, it } from "vitest";

describe("isArchiveError", () => {
  it("is exported from both archive entries", () => {
    expect(browserEntry.isArchiveError).toBe(isArchiveError);
  });

  it("is true for ArchiveError and every subclass", () => {
    const errors: Error[] = [
      new ArchiveError("x"),
      new ZipParseError("x"),
      new ZipInvalidSignatureError("PK\\x03\\x04", 0, "local header"),
      new ZipEocdNotFoundError(),
      new ZipCrc32MismatchError("a.txt", 1, 2),
      new ZipDecryptionError("a.txt"),
      new ZipPasswordRequiredError("a.txt"),
      new ArchiveRangeNotSupportedError("https://example.com/a.zip"),
      new ArchiveHttpRangeError("https://example.com/a.zip", 500, "Server Error"),
      new ArchiveFileTooLargeError("a.txt", "too big"),
      new ZipUnsupportedCompressionError(99),
      new ZipEntrySizeMismatchError("a.txt", 1, 2, "too-many-bytes"),
      new ArchiveLimitExceededError("maxEntries", 1)
    ];
    for (const error of errors) {
      expect(isArchiveError(error), error.name).toBe(true);
    }
  });

  it("is true for an error a public archive function throws", () => {
    let thrown: unknown;
    try {
      Archive.unpackZipSync(new Uint8Array([1, 2, 3, 4]));
    } catch (err) {
      thrown = err;
    }
    expect(thrown).toBeDefined();
    expect(isArchiveError(thrown)).toBe(true);
  });

  it("is false for AbortError, unrelated errors and non-errors", () => {
    for (const value of [
      new AbortError(),
      new BaseError("x"),
      new Error("x"),
      new TypeError("x"),
      { name: "ArchiveError", message: "x" },
      "ArchiveError",
      null,
      undefined
    ]) {
      expect(isArchiveError(value)).toBe(false);
    }
  });
});
