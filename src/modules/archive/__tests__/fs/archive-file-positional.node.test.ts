/**
 * `ArchiveFile`'s read methods and `RemoteZipReader`'s extractors keep their deprecated positional
 * password / encoding beside the options object. Anything but an options object is the positional
 * value: a `null` placeholder falls back to the archive's password, a `Uint8Array` is a password —
 * whichever realm made it — and `null` reaches `TextDecoder` as the encoding it always did.
 *
 * Expected values were observed against the last release (c3b2407c), except where a test says it
 * pins an improvement.
 */
import { runInNewContext } from "node:vm";

import { Archive, ArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

function loose<T>(value: unknown): T {
  return value as T;
}

function encrypted(): Uint8Array {
  const zip = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "zipcrypto" });
  zip.addText("old", "z.txt");
  return zip.toBufferSync();
}

const decode = (data: Uint8Array | undefined) =>
  data === undefined ? undefined : new TextDecoder().decode(data);

describe("ArchiveFile positional placeholders", () => {
  const bytes = encrypted();
  const withPassword = () => Archive.ArchiveFile.fromBuffer(bytes, { password: "pw" });
  const withoutPassword = () => Archive.ArchiveFile.fromBuffer(bytes);

  it("a null password falls back to the archive's", async () => {
    expect(decode(await withPassword().readEntry("z.txt", loose<string>(null)))).toBe("old");
    expect(decode(withPassword().readEntrySync("z.txt", loose<string>(null)))).toBe("old");
    expect(decode(await withPassword().readEntry("z.txt", undefined))).toBe("old");
  });

  it("an empty password does not fall back", async () => {
    await expect(withPassword().readEntry("z.txt", "")).rejects.toThrow(/provide a password/);
  });

  it("a Uint8Array is the password, not an options object", async () => {
    const password = new TextEncoder().encode("pw");
    expect(decode(await withoutPassword().readEntry("z.txt", password))).toBe("old");
    expect(decode(withoutPassword().readEntrySync("z.txt", password))).toBe("old");
  });

  // A null encoding still reaches TextDecoder rather than defaulting. TextDecoder refuses the
  // label "null", which used to escape as its bare `RangeError`; it is now an `ArchiveError`
  // naming `encoding`, like every other label TextDecoder does not know.
  it("a null encoding reaches TextDecoder, an undefined one defaults", async () => {
    await expect(withPassword().readAsText("z.txt", loose<string>(null))).rejects.toThrow(
      ArchiveError
    );
    expect(() => withPassword().readAsTextSync("z.txt", loose<string>(null))).toThrow(
      /encoding must be/
    );
    expect(await withPassword().readAsText("z.txt", undefined)).toBe("old");
    expect(withPassword().readAsTextSync("z.txt", undefined)).toBe("old");
  });
});

describe("RemoteZipReader positional password", () => {
  const bytes = encrypted();
  const open = () => Archive.RemoteZipReader.fromReader(Archive.createBufferReader(bytes));

  it("a string or Uint8Array password, and the options object, all decrypt", async () => {
    const reader = await open();
    expect(decode(await reader.extract("z.txt", "pw"))).toBe("old");
    expect(decode(await reader.extract("z.txt", new TextEncoder().encode("pw")))).toBe("old");
    expect(decode(await reader.extract("z.txt", { password: "pw" }))).toBe("old");
  });

  // An improvement, not parity: the last release tested `instanceof Uint8Array`, so a password made
  // in another realm was read as an options object without one.
  it("a Uint8Array from another realm is still a password", async () => {
    const foreign = runInNewContext("new Uint8Array([112, 119])") as Uint8Array;
    expect(foreign instanceof Uint8Array).toBe(false);
    const reader = await open();
    expect(decode(await reader.extract("z.txt", foreign))).toBe("old");
  });
});
