/**
 * Every header field a TAR writer is given is either stored so it reads back, or refused:
 * a modification time beyond the 11-digit octal field (after 2242) is written in GNU
 * base-256, a link target or a name too long for its field goes into a GNU long-name record
 * (measured in UTF-8 bytes, not characters), and a user/group name or entry type the header
 * cannot hold is refused instead of cut short.
 */
import { Archive, ArchiveError, Compression } from "@archive";
import { describe, expect, it } from "vitest";

const readBack = (archive: ReturnType<typeof Archive.createTar>) =>
  Archive.unpackTarSync(archive.toBufferSync());

describe("TAR header fields", () => {
  it.each([2200, 2300, 9999])("round-trips a modification time in %i", year => {
    const mtime = new Date(Date.UTC(year, 0, 1));
    const read = readBack(Archive.createTar().add("a", "x", { mtime }));
    expect(read.get("a")!.info.mtime.toISOString()).toBe(mtime.toISOString());
  });

  it("round-trips a symlink target longer than the 100-byte field", async () => {
    const target = "t/".repeat(120) + "end";
    const bytes = Archive.createTar().addSymlink("link", target).toBufferSync();
    expect(Archive.unpackTarSync(bytes).get("link")!.info.linkname).toBe(target);
    // The streaming parser and the tar.gz path read it too.
    const streamed = await Archive.unpackTar(
      (async function* () {
        yield bytes;
      })()
    );
    expect(streamed.get("link")!.info.linkname).toBe(target);
    const tgz = Compression.gzipSync(bytes);
    expect((await Archive.unpackTarGz(tgz)).get("link")!.info.linkname).toBe(target);
  });

  it("round-trips a long name and a long link target on the same entry", () => {
    const name = "n/".repeat(80) + "file";
    const target = "t/".repeat(80) + "end";
    const read = readBack(Archive.createTar().addSymlink(name, target));
    expect(read.get(name)!.info.linkname).toBe(target);
  });

  it("measures the name field in UTF-8 bytes", () => {
    const name = "名".repeat(60); // 60 characters, 180 bytes
    const nested = "d/" + "字".repeat(60) + ".txt";
    const read = readBack(Archive.createTar().add(name, "x").add(nested, "y"));
    expect([...read.keys()]).toEqual([name, nested]);
  });

  it("refuses a user or group name longer than the 31-byte field", () => {
    expect(() => Archive.createTar().add("a", "x", { uname: "u".repeat(32) })).toThrow(
      ArchiveError
    );
    expect(() => Archive.createTar().add("a", "x", { gname: "名".repeat(11) })).toThrow(
      ArchiveError
    );
    expect(
      readBack(Archive.createTar().add("a", "x", { uname: "u".repeat(31) })).get("a")!.info.uname
    ).toBe("u".repeat(31));
  });

  it("refuses an entry type that is not one a TAR entry can have", () => {
    expect(() => Archive.createTar().add("a", "x", { type: "bogus" as never })).toThrow(
      ArchiveError
    );
    expect(() =>
      Archive.createTar().add("a", "x", { type: Archive.TAR_TYPE.GNU_LONG_NAME })
    ).toThrow(ArchiveError);
    const fifo = readBack(Archive.createTar().add("f", "", { type: Archive.TAR_TYPE.FIFO }));
    expect(fifo.get("f")!.info.type).toBe(Archive.TAR_TYPE.FIFO);
  });
});

describe("TAR numeric fields beyond octal", () => {
  it.each([2 ** 21, 2 ** 40, Number.MAX_SAFE_INTEGER])(
    "round-trips uid and gid %i in GNU base-256",
    id => {
      const info = readBack(Archive.createTar().add("a", "x", { uid: id, gid: id })).get("a")!.info;
      expect([info.uid, info.gid]).toEqual([id, id]);
    }
  );

  it.each(["1960-05-06T07:08:09Z", "1969-12-31T23:59:59Z", "1901-01-01T00:00:00Z"])(
    "round-trips a modification time before 1970 (%s)",
    iso => {
      const mtime = new Date(iso);
      const read = readBack(Archive.createTar().add("a", "x", { mtime }));
      expect(read.get("a")!.info.mtime.toISOString()).toBe(mtime.toISOString());
    }
  );

  it("still writes octal for values that fit, as other tools expect", () => {
    const bytes = Archive.createTar().add("a", "x", { uid: 1000, gid: 1000 }).toBufferSync();
    expect(new TextDecoder().decode(bytes.subarray(108, 115))).toBe("0001750");
  });

  it("still refuses an id that is not a non-negative integer", () => {
    expect(() => Archive.createTar().add("a", "x", { uid: -1 })).toThrow(ArchiveError);
    expect(() => Archive.createTar().add("a", "x", { gid: 1.5 })).toThrow(ArchiveError);
  });
});
