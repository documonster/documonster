/**
 * `ZipReader` answers what `TarReader` answers: `list()` names every entry, and each
 * `UnzipEntry` carries its record as `info` — from the central directory when the reader
 * has the whole archive, from the local header when it streams. `get()` / `bytes()` /
 * `list()` buffer a streamed source, as their documentation says, and a source that is
 * not one is refused as such.
 */
import { Archive, ArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

const modTime = new Date(2021, 2, 4, 5, 6, 8);
const zip = Archive.createZip()
  .add("a.txt", "hello", { modTime, comment: "first" })
  .addDirectory("d", { modTime })
  .add("d/b.txt", "B", { modTime })
  .toBufferSync();
const streamed = () =>
  (async function* () {
    yield zip;
  })();

describe("ZipReader.list", () => {
  it("lists every entry of an in-memory archive", async () => {
    expect(await Archive.openZip(zip).list()).toEqual(["a.txt", "d/", "d/b.txt"]);
  });

  it("lists every entry of a streamed archive", async () => {
    expect(await Archive.openZip(streamed()).list()).toEqual(["a.txt", "d/", "d/b.txt"]);
  });
});

describe("UnzipEntry.info", () => {
  it("is the central-directory record for an in-memory archive", async () => {
    const entry = (await Archive.openZip(zip).get("a.txt"))!;
    expect(entry.info).toMatchObject({
      path: "a.txt",
      type: "file",
      uncompressedSize: 5,
      comment: "first",
      lastModified: modTime
    });
    expect(entry.info.crc32).toBe(Archive.crc32(new TextEncoder().encode("hello")));
  });

  it("is the local-header record while streaming", async () => {
    const seen: Record<string, number> = {};
    for await (const entry of Archive.openZip(streamed()).entries()) {
      expect(entry.info.path).toBe(entry.path);
      expect(entry.info.lastModified.getTime()).toBe(modTime.getTime());
      seen[entry.path] = entry.info.lastModified.getTime();
      entry.discard();
    }
    expect(Object.keys(seen)).toEqual(["a.txt", "d/", "d/b.txt"]);
  });

  it("is a copy the caller cannot use to change the entry", async () => {
    const entry = (await Archive.openZip(zip).get("a.txt"))!;
    expect(() => {
      (entry.info as { path: string }).path = "other";
    }).toThrow(TypeError);
    expect(new TextDecoder().decode(await entry.bytes())).toBe("hello");
  });
});

describe("ZipReader random access", () => {
  it("buffers a streamed source for get() and bytes()", async () => {
    expect(new TextDecoder().decode((await Archive.openZip(streamed()).bytes("d/b.txt"))!)).toBe(
      "B"
    );
    expect((await Archive.openZip(new Blob([zip as BlobPart]).stream()).get("a.txt"))?.path).toBe(
      "a.txt"
    );
  });

  it.each([null, 42, {}])("refuses %s as a source when it is opened", source => {
    expect(() => Archive.openZip(source as never)).toThrow(ArchiveError);
    expect(() => Archive.openZip(source as never)).toThrow(/^Archive\.openZip: source must be/);
  });
});
