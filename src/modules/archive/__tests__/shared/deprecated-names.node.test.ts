/**
 * The archive names renamed or moved out of `Archive` stay reachable under their old
 * spelling until 1.0, as the very same values.
 */
import * as archive from "@archive";
import * as archiveBrowser from "@archive/index.browser";
import * as stream from "@stream/index";
import * as streamBrowser from "@stream/index.browser";
import { describe, expect, it } from "vitest";

describe("archive deprecated names", () => {
  it.each([
    [
      "EocdNotFoundError",
      archive.EocdNotFoundError,
      archive.ZipEocdNotFoundError,
      archiveBrowser.EocdNotFoundError,
      archiveBrowser.ZipEocdNotFoundError
    ],
    [
      "Crc32MismatchError",
      archive.Crc32MismatchError,
      archive.ZipCrc32MismatchError,
      archiveBrowser.Crc32MismatchError,
      archiveBrowser.ZipCrc32MismatchError
    ],
    [
      "DecryptionError",
      archive.DecryptionError,
      archive.ZipDecryptionError,
      archiveBrowser.DecryptionError,
      archiveBrowser.ZipDecryptionError
    ],
    [
      "PasswordRequiredError",
      archive.PasswordRequiredError,
      archive.ZipPasswordRequiredError,
      archiveBrowser.PasswordRequiredError,
      archiveBrowser.ZipPasswordRequiredError
    ],
    [
      "EntrySizeMismatchError",
      archive.EntrySizeMismatchError,
      archive.ZipEntrySizeMismatchError,
      archiveBrowser.EntrySizeMismatchError,
      archiveBrowser.ZipEntrySizeMismatchError
    ],
    [
      "UnsupportedCompressionError",
      archive.UnsupportedCompressionError,
      archive.ZipUnsupportedCompressionError,
      archiveBrowser.UnsupportedCompressionError,
      archiveBrowser.ZipUnsupportedCompressionError
    ],
    [
      "RangeNotSupportedError",
      archive.RangeNotSupportedError,
      archive.ArchiveRangeNotSupportedError,
      archiveBrowser.RangeNotSupportedError,
      archiveBrowser.ArchiveRangeNotSupportedError
    ],
    [
      "HttpRangeError",
      archive.HttpRangeError,
      archive.ArchiveHttpRangeError,
      archiveBrowser.HttpRangeError,
      archiveBrowser.ArchiveHttpRangeError
    ],
    [
      "FileTooLargeError",
      archive.FileTooLargeError,
      archive.ArchiveFileTooLargeError,
      archiveBrowser.FileTooLargeError,
      archiveBrowser.ArchiveFileTooLargeError
    ],
    [
      "WorkerStreamStartError",
      archive.WorkerStreamStartError,
      archive.ArchiveWorkerStreamStartError,
      archiveBrowser.WorkerStreamStartError,
      archiveBrowser.ArchiveWorkerStreamStartError
    ]
  ] as const)(
    "%s forwards to its new name on both entries",
    (_name, old, renamed, oldBrowser, renamedBrowser) => {
      expect(old).toBe(renamed);
      expect(oldBrowser).toBe(renamedBrowser);
    }
  );

  it("an error thrown under the new name is caught by the old one", () => {
    const error = new archive.ZipDecryptionError("a.txt");
    expect(error).toBeInstanceOf(archive.DecryptionError);
    expect(error.name).toBe("ZipDecryptionError");
    expect(archive.isArchiveError(error)).toBe(true);
  });

  it("Archive.TAR_TYPE forwards to Archive.TarType", () => {
    expect(archive.Archive.TAR_TYPE).toBe(archive.Archive.TarType);
    expect(archive.Archive.TarType.DIRECTORY).toBe("5");
  });

  it("the deprecated Archive abort helpers still work", () => {
    expect(archive.Archive.isAbortError).toBe(archive.isAbortError);
    const error = archive.Archive.createAbortError("why");
    expect(error).toBeInstanceOf(archive.AbortError);
    expect(error.cause).toBe("why");
    expect(archive.Archive.createAbortError(error)).toBe(error);
  });

  it("the deprecated Archive byte helpers still work", () => {
    const bytes = archive.Archive.stringToUint8Array("abc");
    expect(archive.Archive.uint8ArrayToString(bytes)).toBe("abc");
    expect(archive.Archive.uint8ArrayToBase64(bytes)).toBe("YWJj");
    expect([...archive.Archive.base64ToUint8Array("YWJj")]).toEqual([...bytes]);
    expect([...archive.Archive.concatUint8Arrays([bytes, bytes])]).toHaveLength(6);
  });
});

describe("flat isAbortError", () => {
  it("is the same guard on every archive and stream entry", () => {
    expect(archiveBrowser.isAbortError).toBe(archive.isAbortError);
    expect(stream.isAbortError).toBe(archive.isAbortError);
    expect(streamBrowser.isAbortError).toBe(archive.isAbortError);
    expect(stream.AbortError).toBe(archive.AbortError);
  });

  it("answers for this library's AbortError and a native one", () => {
    expect(archive.isAbortError(new archive.AbortError())).toBe(true);
    expect(archive.isAbortError(new DOMException("x", "AbortError"))).toBe(true);
    expect(archive.isAbortError(new archive.ArchiveError("x"))).toBe(false);
  });

  it("Archive.throwIfAborted throws the library's AbortError, unlike signal.throwIfAborted()", () => {
    const controller = new AbortController();
    controller.abort("why");
    let native: unknown;
    try {
      controller.signal.throwIfAborted();
    } catch (error) {
      native = error;
    }
    expect(native).toBe("why");
    let caught: unknown;
    try {
      archive.Archive.throwIfAborted(controller.signal);
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(archive.AbortError);
    expect((caught as Error).cause).toBe("why");
  });
});

describe("Archive.getEncryptionMethodName (deprecated)", () => {
  it("names what Crypto.getEncryptionMethodName names, and keeps answering undefined for the rest", () => {
    expect(archive.Archive.getEncryptionMethodName("aes-256")).toBe(
      archive.Archive.Crypto.getEncryptionMethodName("aes-256")
    );
    expect(archive.Archive.getEncryptionMethodName("bogus" as never)).toBeUndefined();
    expect(() => archive.Archive.Crypto.getEncryptionMethodName("bogus" as never)).toThrow(
      archive.ArchiveArgumentError
    );
  });
});
