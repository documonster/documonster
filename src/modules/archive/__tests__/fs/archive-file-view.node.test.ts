/**
 * Every `ArchiveFile` read answers for the archive a write would produce now: pending adds,
 * replacements, deletes, renames and directory expansions are visible to `getEntryNames`,
 * `entryCount`, `has`, `getEntry`, `getEntries`, `readEntry`, `readAsText` and `extractTo`,
 * and the metadata of a pending entry is what the write records. A pending stream can be read
 * once, so a read buffers it and the later write still has its data.
 */
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-archive-file-view-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const tree = path.join(dir, "tree");
mkdirSync(path.join(tree, "sub"), { recursive: true });
writeFileSync(path.join(tree, "b.txt"), "tree-b");
writeFileSync(path.join(tree, "sub", "c.txt"), "tree-c");
const fixedTime = new Date("2021-03-04T05:06:08Z");
utimesSync(path.join(tree, "b.txt"), fixedTime, fixedTime);
utimesSync(path.join(tree, "sub", "c.txt"), fixedTime, fixedTime);
const diskFile = path.join(dir, "disk.txt");
writeFileSync(diskFile, "from disk");
utimesSync(diskFile, fixedTime, fixedTime);

const decoder = new TextDecoder();
const text = (bytes: Uint8Array | null) => (bytes === null ? null : decoder.decode(bytes));

const zipSource = () =>
  Archive.ArchiveFile.fromBuffer(
    new Archive.ArchiveFile({ comment: "kept" })
      .addText("A", "a.txt", { modTime: fixedTime })
      .addText("B", "b.txt", { modTime: fixedTime })
      .addText("C", "d/c.txt", { modTime: fixedTime })
      .toBufferSync()
  );

/** The fields a write records, for comparing a pending entry with its written form. */
const recorded = (value: object) => {
  const info = value as { path: string } & Record<string, unknown>;
  return {
    path: info.path,
    isDirectory: info.isDirectory,
    size: info.size,
    compressedSize: info.compressedSize,
    crc32: info.crc32,
    lastModified: (info.lastModified as Date | undefined)?.getTime(),
    isEncrypted: info.isEncrypted,
    comment: info.comment
  };
};

async function* streamOf(...chunks: string[]): AsyncIterable<Uint8Array> {
  for (const chunk of chunks) {
    yield new TextEncoder().encode(chunk);
  }
}

describe("ArchiveFile ZIP reads reflect pending edits", () => {
  const edited = () => {
    const file = zipSource();
    file.addText("N", "new.txt", { modTime: fixedTime });
    file.set("a.txt", "A2");
    file.delete("b.txt");
    file.rename("d/c.txt", "d/z.txt");
    return file;
  };

  it("lists, counts and answers has() for the output", () => {
    const file = edited();
    expect(file.getEntryNames()).toEqual(["a.txt", "d/z.txt", "new.txt"]);
    expect(file.entryCount).toBe(3);
    expect(file.has("new.txt")).toBe(true);
    expect(file.has("b.txt")).toBe(false);
    expect(file.has("d/c.txt")).toBe(false);
  });

  it("reads the pending content", async () => {
    const file = edited();
    expect(text(await file.readEntry("new.txt"))).toBe("N");
    expect(text(file.readEntrySync("a.txt"))).toBe("A2");
    expect(text(await file.readEntry("d/z.txt"))).toBe("C");
    expect(await file.readEntry("b.txt")).toBeNull();
    expect(await file.readAsText("new.txt")).toBe("N");
    expect(file.readAsTextSync("d/z.txt")).toBe("C");
    expect(file.getEntry("b.txt")).toBeNull();
  });

  it("describes every entry as the write records it", async () => {
    const file = edited();
    const before = (await file.getEntries()).map(recorded);
    const beforeSync = file.getEntriesSync().map(recorded);
    const one = recorded(file.getEntry("new.txt")! as never);
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    const after = (await written.getEntries()).map(recorded);
    // A write re-compresses entries, so compare everything a write fixes before it runs.
    const stable = (list: ReturnType<typeof recorded>[]) =>
      list.map(({ compressedSize: _c, ...rest }) => rest);
    expect(stable(before)).toEqual(stable(after));
    expect(stable(beforeSync)).toEqual(stable(after));
    expect(one).toEqual(after.find(entry => entry.path === "new.txt"));
    expect(written.getEntryNames()).toEqual(["a.txt", "d/z.txt", "new.txt"]);
  });

  it("keeps the archive comment of an opened archive and shows a pending one", async () => {
    const file = zipSource();
    expect(file.getZipComment()).toBe("kept");
    expect(
      Archive.ArchiveFile.fromBuffer(await file.addText("x", "x").toBuffer()).getZipComment()
    ).toBe("kept");
    file.setZipComment("new comment");
    expect(file.getZipComment()).toBe("new comment");
  });

  it("reads a pending file from disk with the metadata the write records", async () => {
    const file = zipSource().addFile(diskFile);
    expect(text(await file.readEntry("disk.txt"))).toBe("from disk");
    expect(file.getEntry("disk.txt")!.lastModified.getTime()).toBe(fixedTime.getTime());
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(recorded(file.getEntry("disk.txt")! as never)).toEqual(
      recorded(written.getEntry("disk.txt")! as never)
    );
  });

  it("expands a pending directory as the write does, and can delete what it produces", async () => {
    const file = zipSource().addDirectory(tree);
    const names = file.getEntryNames();
    expect(names).toEqual(expect.arrayContaining(["tree/b.txt", "tree/sub/c.txt"]));
    expect(text(await file.readEntry("tree/sub/c.txt"))).toBe("tree-c");
    expect(file.delete("tree/b.txt")).toBe(true);
    expect(file.has("tree/b.txt")).toBe(false);
    const view = file.getEntryNames();
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(written.getEntryNames()).toEqual(view);
    expect(view).not.toContain("tree/b.txt");
  });

  it("lets the last write of a name win, in the view and in the output", async () => {
    const file = new Archive.ArchiveFile().addText("1", "x.txt").addText("2", "x.txt");
    expect(file.getEntryNames()).toEqual(["x.txt"]);
    expect(text(await file.readEntry("x.txt"))).toBe("2");
    const out = await file.toBuffer();
    expect(Archive.ArchiveFile.fromBuffer(out).getEntryNames()).toEqual(["x.txt"]);
    expect(text(await Archive.ArchiveFile.fromBuffer(out).readEntry("x.txt"))).toBe("2");
  });

  it("reads a new archive before it is written", async () => {
    const file = new Archive.ArchiveFile().addText("hello", "h.txt");
    expect(file.getEntryNames()).toEqual(["h.txt"]);
    expect(text(await file.readEntry("h.txt"))).toBe("hello");
    expect((await file.getEntries()).map(entry => entry.path)).toEqual(["h.txt"]);
  });

  it("stamps a pending entry when it is added, so the read and the write agree", async () => {
    const file = new Archive.ArchiveFile({ timestamps: "dos+utc" }).addText("x", "x.txt");
    const before = file.getEntry("x.txt")!.lastModified.getTime();
    await new Promise(resolve => setTimeout(resolve, 1100));
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(written.getEntry("x.txt")!.lastModified.getTime()).toBe(before);
  });

  it("uses the archive's reproducible stamp for a pending entry", async () => {
    const file = new Archive.ArchiveFile({ reproducible: true }).addText("x", "x.txt");
    const expected = new Date(1980, 0, 1).getTime();
    expect(file.getEntry("x.txt")!.lastModified.getTime()).toBe(expected);
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(written.getEntry("x.txt")!.lastModified.getTime()).toBe(expected);
  });

  it("extracts the output view", async () => {
    const target = path.join(dir, "extract-zip");
    await edited().extractTo(target);
    expect(readFileSync(path.join(target, "a.txt"), "utf8")).toBe("A2");
    expect(readFileSync(path.join(target, "new.txt"), "utf8")).toBe("N");
    expect(readFileSync(path.join(target, "d", "z.txt"), "utf8")).toBe("C");
    expect(existsSync(path.join(target, "b.txt"))).toBe(false);
    expect(existsSync(path.join(target, "d", "c.txt"))).toBe(false);
    expect(statSync(path.join(target, "new.txt")).mtime.getTime()).toBe(fixedTime.getTime());
    const syncTarget = path.join(dir, "extract-zip-sync");
    edited().extractToSync(syncTarget);
    expect(readFileSync(path.join(syncTarget, "a.txt"), "utf8")).toBe("A2");
    const one = path.join(dir, "one.txt");
    expect(await edited().extractEntryTo("new.txt", one)).toBe(true);
    expect(readFileSync(one, "utf8")).toBe("N");
    expect(edited().extractEntryToSync("b.txt", path.join(dir, "gone.txt"))).toBe(false);
  });
});

describe("ArchiveFile pending streams", () => {
  it("buffers a stream a read needs, so the write still has its data", async () => {
    const file = zipSource().appendStream(streamOf("str", "eam"), "s.txt");
    expect(file.getEntryNames()).toContain("s.txt");
    expect(text(await file.readEntry("s.txt"))).toBe("stream");
    expect(text(await file.readEntry("s.txt"))).toBe("stream");
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(text(await written.readEntry("s.txt"))).toBe("stream");
  });

  it("refuses a synchronous read that would need an unread stream, until it is buffered", async () => {
    const file = zipSource().appendStream(streamOf("x"), "s.txt");
    expect(() => file.readEntrySync("s.txt")).toThrow(ArchiveError);
    expect(() => file.getEntriesSync()).toThrow(/async/);
    await file.getEntries();
    expect(text(file.readEntrySync("s.txt"))).toBe("x");
  });

  it("does not touch a pending stream for a read that does not need it", async () => {
    const file = zipSource().appendStream(
      // eslint-disable-next-line require-yield
      (async function* () {
        throw new Error("consumed");
      })(),
      "s.txt"
    );
    expect(text(await file.readEntry("a.txt"))).toBe("A");
    expect(file.getEntryNames()).toContain("s.txt");
    expect(file.has("s.txt")).toBe(true);
  });
});

describe("ArchiveFile TAR reads reflect pending edits", () => {
  const tarSource = () =>
    Archive.ArchiveFile.fromBuffer(
      Archive.createTar().add("a.txt", "A", { mtime: fixedTime }).toBufferSync(),
      { format: "tar" }
    );

  it("lists and reads pending entries, opened or new", async () => {
    const file = tarSource().addText("N", "new.txt").addFile(diskFile);
    expect(file.getEntryNames()).toEqual(["a.txt", "new.txt", "disk.txt"]);
    expect(file.entryCount).toBe(3);
    expect(text(await file.readEntry("new.txt"))).toBe("N");
    expect(text(file.readEntrySync("disk.txt"))).toBe("from disk");
    const fresh = new Archive.ArchiveFile({ format: "tar" }).addText("x", "x.txt");
    expect(fresh.getEntryNames()).toEqual(["x.txt"]);
    expect(text(await fresh.readEntry("x.txt"))).toBe("x");
  });

  it("describes pending entries as the write records them", async () => {
    const file = tarSource().addText("N", "new.txt").addDirectory(tree);
    const pick = (value: object) => {
      const entry = value as Record<string, unknown>;
      return {
        path: entry.path,
        size: entry.size,
        mode: entry.mode,
        type: entry.type,
        isDirectory: entry.isDirectory,
        mtime: (entry.mtime as Date | undefined)?.getTime()
      };
    };
    const before = (await file.getEntries()).map(pick);
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer(), { format: "tar" });
    expect(before).toEqual((await written.getEntries()).map(pick));
  });

  it("extracts the output view", async () => {
    const target = path.join(dir, "extract-tar");
    await tarSource().addText("N", "new.txt").extractTo(target);
    expect(readFileSync(path.join(target, "a.txt"), "utf8")).toBe("A");
    expect(readFileSync(path.join(target, "new.txt"), "utf8")).toBe("N");
  });

  it("buffers a pending stream a read needs", async () => {
    const file = tarSource().appendStream(streamOf("ta", "r"), "s.txt");
    expect(text(await file.readEntry("s.txt"))).toBe("tar");
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer(), { format: "tar" });
    expect(text(await written.readEntry("s.txt"))).toBe("tar");
  });
});

describe("ArchiveFile writes agree with the view", () => {
  const collect = async (iterable: AsyncIterable<Uint8Array>) => {
    const parts: Uint8Array[] = [];
    for await (const part of iterable) {
      parts.push(part);
    }
    return Buffer.concat(parts);
  };

  it("streams the same entries, names and dates a buffered write records", async () => {
    const make = () =>
      zipSource()
        .addDirectory(tree, {
          transform: entry => ({ ...entry, name: entry.name.toUpperCase() })
        })
        .addFile(diskFile)
        .appendStream(streamOf("s"), "s.txt");
    const streamed = Archive.ArchiveFile.fromBuffer(await collect(make().stream()));
    const buffered = Archive.ArchiveFile.fromBuffer(await make().toBuffer());
    const stable = async (file: Archive.ArchiveFile) =>
      (await file.getEntries()).map(entry => ({
        ...recorded(entry),
        compressedSize: undefined
      }));
    expect(await stable(streamed)).toEqual(await stable(buffered));
    expect(streamed.getEntryNames()).toContain("tree/SUB/C.TXT");
    expect(streamed.getZipComment()).toBe("kept");
  });

  it("keeps a stream a streamed write consumed for the next write", async () => {
    const file = new Archive.ArchiveFile().appendStream(streamOf("once"), "s.txt");
    await collect(file.stream());
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer());
    expect(text(await written.readEntry("s.txt"))).toBe("once");
  });

  it("rewrites an opened archive whose only change is its comment", async () => {
    const file = zipSource();
    expect(file.hasPendingChanges()).toBe(false);
    file.setZipComment("changed");
    expect(file.hasPendingChanges()).toBe(true);
    expect(Archive.ArchiveFile.fromBuffer(await file.toBuffer()).getZipComment()).toBe("changed");
    expect(file.hasPendingChanges()).toBe(false);
  });

  it("describes a pending encrypted entry as the write records it", async () => {
    for (const encryptionMethod of ["zipcrypto", "aes-256"] as const) {
      const file = new Archive.ArchiveFile({ password: "pw", encryptionMethod }).addText(
        "secret secret secret",
        "s.txt",
        { modTime: fixedTime }
      );
      expect(file.hasEncryptedEntries()).toBe(true);
      const pending = { ...file.getEntry("s.txt")! };
      const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer(), { password: "pw" });
      expect(pending).toEqual(written.getEntry("s.txt"));
    }
  });

  it("refuses a streamed write that would have to encrypt", () => {
    const file = new Archive.ArchiveFile({ password: "pw", encryptionMethod: "zipcrypto" });
    file.addText("x", "x.txt");
    expect(() => file.stream()).toThrow(/cannot/);
  });

  it("replaces an entry of an opened TAR in place", async () => {
    const source = Archive.createTar()
      .add("a.txt", "A", { mtime: fixedTime })
      .add("b.txt", "B", { mtime: fixedTime })
      .toBufferSync();
    const file = Archive.ArchiveFile.fromBuffer(source, { format: "tar" }).addText("A2", "a.txt");
    expect(file.getEntryNames()).toEqual(["a.txt", "b.txt"]);
    const written = Archive.ArchiveFile.fromBuffer(await file.toBuffer(), { format: "tar" });
    expect(written.getEntryNames()).toEqual(["a.txt", "b.txt"]);
    expect(written.readAsTextSync("a.txt")).toBe("A2");
  });

  it("extracts a TAR synchronously", () => {
    const target = path.join(dir, "extract-tar-sync");
    new Archive.ArchiveFile({ format: "tar" }).addText("x", "x.txt").extractToSync(target);
    expect(readFileSync(path.join(target, "x.txt"), "utf8")).toBe("x");
  });
});
