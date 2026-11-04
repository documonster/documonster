/**
 * The random-access readers and the remote entry points check their arguments where they
 * are given: a source that is not bytes or a URL that is not a string is an `ArchiveError`.
 * A byte offset that is not one (`NaN`, a fraction, `Infinity`, not a number) is an
 * `ArchiveArgumentError` naming it, and an offset outside the source a `RangeError` naming it.
 */
import { Archive, ArchiveArgumentError, ArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

describe("createBufferReader", () => {
  it.each([null, 42, "text", {}])("refuses %s", data => {
    expect(() => Archive.createBufferReader(data as never)).toThrow(ArchiveError);
  });

  it.each<[number, number, string, RegExp]>([
    [-1, 5, "start", /start -1 is out of range: it must be >= 0 and < 10/],
    [10, 11, "start", /start 10 is out of range: it must be >= 0 and < 10/],
    [5, 2, "end", /end 2 is out of range: it must be > start \(5\) and <= 10/],
    [5, 5, "end", /end 5 is out of range/],
    [0, 1e9, "end", /end 1000000000 is out of range: it must be > start \(0\) and <= 10/]
  ])(
    "refuses the out-of-range [%s, %s) with a RangeError naming %s",
    async (start, end, _field, message) => {
      const reader = Archive.createBufferReader(new Uint8Array(10));
      // An out-of-range offset is a `RangeError`, as Node's and the Web's checked byte APIs throw.
      const error = await reader.read(start, end).catch((caught: unknown) => caught);
      expect(error).toBeInstanceOf(RangeError);
      expect((error as Error).message).toMatch(message);
    }
  );

  it.each<[unknown, unknown, string]>([
    [Number.NaN, 3, "start"],
    [0, Number.NaN, "end"],
    [1.5, 3, "start"],
    [0, 2.5, "end"],
    ["1", 3, "start"],
    [0, null, "end"],
    [Number.POSITIVE_INFINITY, 3, "start"]
  ])(
    "refuses the non-offset [%s, %s) with an ArchiveArgumentError naming %s",
    async (start, end, field) => {
      const reader = Archive.createBufferReader(new Uint8Array(10));
      const error = await reader
        .read(start as number, end as number)
        .catch((caught: unknown) => caught);
      expect(error).toBeInstanceOf(ArchiveArgumentError);
      expect((error as ArchiveArgumentError).field).toBe(field);
    }
  );

  it("checks HttpRangeReader.read the same way", async () => {
    const body = new Uint8Array(10);
    const fetchFn = (async (_url: string, init?: { method?: string }) =>
      new Response(init?.method === "HEAD" ? null : body, {
        status: init?.method === "HEAD" ? 200 : 206,
        headers: { "content-length": "10", "accept-ranges": "bytes" }
      })) as unknown as typeof fetch;
    const reader = await Archive.HttpRangeReader.open("https://example.test/x.zip", {
      fetch: fetchFn
    });
    await expect(reader.read(0, 11)).rejects.toThrow(/end 11 is out of range/);
    await expect(reader.read(0, 11)).rejects.toBeInstanceOf(RangeError);
    const error = await reader.read(Number.NaN, 3).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(ArchiveArgumentError);
    expect((error as ArchiveArgumentError).field).toBe("start");
  });

  it("reads a valid range", async () => {
    const reader = Archive.createBufferReader(new Uint8Array([1, 2, 3, 4]));
    expect([...(await reader.read(1, 3))]).toEqual([2, 3]);
  });
});

describe("remote entry points", () => {
  it.each([42, null, undefined])("refuses %s as a URL", async url => {
    await expect(Archive.HttpRangeReader.open(url as never)).rejects.toBeInstanceOf(ArchiveError);
    await expect(Archive.RemoteZipReader.open(url as never)).rejects.toBeInstanceOf(ArchiveError);
    await expect(Archive.editZipUrl(url as never)).rejects.toBeInstanceOf(ArchiveError);
  });
});

describe("Archive.Crypto", () => {
  it("refuses a key strength that is not 128, 192 or 256", async () => {
    await expect(
      Archive.Crypto.encryptAes(new Uint8Array(3), "pw", 999 as never)
    ).rejects.toBeInstanceOf(ArchiveError);
    expect(() => Archive.Crypto.getAesEncryptedSize(10, 999 as never)).toThrow(ArchiveError);
    expect(() => Archive.Crypto.createAesExtraField(2, 999 as never, 8)).toThrow(ArchiveError);
  });

  it("refuses a size that is not a non-negative integer", () => {
    expect(() => Archive.Crypto.getAesEncryptedSize(Number.NaN, 256)).toThrow(ArchiveError);
    expect(() => Archive.Crypto.randomBytes(-1)).toThrow(ArchiveError);
  });

  it("refuses an AES version other than AE-1 or AE-2", () => {
    expect(() => Archive.Crypto.createAesExtraField(7 as never, 256, 8)).toThrow(ArchiveError);
  });
});
