/**
 * `ArchiveFile` reads answer for both formats, check the name they are given, and honour the
 * entry-size limits the other ZIP readers take.
 */
import { Archive, ArchiveArgumentError, FileTooLargeError } from "@archive/index";
import { describe, expect, it } from "vitest";

describe("ArchiveFile lookups", () => {
  it("has / getEntry answer for an opened TAR, the later member of a name", () => {
    const tar = Archive.packTarSync([
      ["a.txt", "first"],
      ["a.txt", "second!"],
      ["d/", ""]
    ]);
    const file = Archive.ArchiveFile.fromBuffer(tar, { format: "tar" });
    expect(file.has("a.txt")).toBe(true);
    expect(file.has("missing")).toBe(false);
    expect(file.getEntry("a.txt")).toMatchObject({ path: "a.txt", size: 7, isDirectory: false });
    expect(file.getEntry("d/")?.isDirectory).toBe(true);
    expect(file.getEntry("missing")).toBeNull();
  });

  it.each([42, undefined, ""])("refuses %j as a name in every read", bad => {
    const name = bad as never;
    const zip = Archive.ArchiveFile.fromBuffer(Archive.packZipSync([["a.txt", "A"]]));
    for (const call of [
      () => zip.has(name),
      () => zip.getEntry(name),
      () => zip.readEntrySync(name),
      () => zip.getZipEntryComment(name)
    ]) {
      expect(call).toThrow(ArchiveArgumentError);
    }
    expect(() => zip.has(name)).toThrow(/^Archive\.ArchiveFile#has: entryPath must be/);
  });

  it("setZipComment is ZIP-only, as setPassword is", () => {
    const tar = new Archive.ArchiveFile({ format: "tar" });
    expect(() => tar.setZipComment("x")).toThrow(/only available for ZIP/);
  });
});

describe("ArchiveFile limits", () => {
  const zip = Archive.packZipSync([
    ["a.txt", "A"],
    ["big.bin", new Uint8Array(4096)]
  ]);

  it("maxEntrySize bounds what a read inflates", async () => {
    const file = Archive.ArchiveFile.fromBuffer(zip, { maxEntrySize: 100 });
    await expect(file.readEntry("big.bin")).rejects.toThrow(FileTooLargeError);
    expect(file.readAsTextSync("a.txt")).toBe("A");
  });

  it("maxEntries is checked when the archive is opened", () => {
    expect(() => Archive.ArchiveFile.fromBuffer(zip, { maxEntries: 1 })).toThrow(/maxEntries/);
  });

  it("fromBuffer takes any ArrayBufferView", () => {
    const file = Archive.ArchiveFile.fromBuffer(new DataView(zip.slice().buffer));
    expect(file.getEntryNames()).toEqual(["a.txt", "big.bin"]);
  });
});
