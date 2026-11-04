import {
  AbortError,
  Archive,
  ArchiveError,
  ArchiveLimitError,
  Crc32MismatchError,
  DecryptionError,
  EntrySizeMismatchError,
  EocdNotFoundError,
  FileTooLargeError,
  HttpRangeError,
  InvalidZipSignatureError,
  PasswordRequiredError,
  RangeNotSupportedError,
  UnsupportedCompressionError,
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
      new InvalidZipSignatureError("PK\\x03\\x04", 0, "local header"),
      new EocdNotFoundError(),
      new Crc32MismatchError("a.txt", 1, 2),
      new DecryptionError("a.txt"),
      new PasswordRequiredError("a.txt"),
      new RangeNotSupportedError("https://example.com/a.zip"),
      new HttpRangeError("https://example.com/a.zip", 500, "Server Error"),
      new FileTooLargeError("a.txt", "too big"),
      new UnsupportedCompressionError(99),
      new EntrySizeMismatchError("a.txt", 1, 2, "too-many-bytes"),
      new ArchiveLimitError("maxEntries", 1)
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
