/**
 * `ArchiveFile` writes the whole archive every time: an archive opened for reading keeps
 * its entries when it is written back (TAR as well as ZIP), building twice gives the same
 * archive, and adding after a build adds to what was built rather than replacing it.
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError } from "@archive";
import { concatUint8Arrays } from "@utils/binary";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-archive-file-rebuild-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const decoder = new TextDecoder();
const tarContents = (bytes: Uint8Array) =>
  Object.fromEntries(
    [...Archive.unpackTarSync(bytes)].map(([name, { data }]) => [name, decoder.decode(data)])
  );
const zipContents = (bytes: Uint8Array) =>
  Object.fromEntries(
    [...Archive.unpackZipSync(bytes)].map(([name, { data }]) => [name, decoder.decode(data)])
  );

describe("ArchiveFile TAR read-modify-write", () => {
  const source = Archive.packTarSync([
    ["a.txt", "A"],
    ["d/b.txt", "B"]
  ]);

  it("writes an unchanged archive back with every entry", async () => {
    const file = Archive.ArchiveFile.fromBuffer(source, { format: "tar" });
    expect(tarContents(await file.toBuffer())).toEqual({ "a.txt": "A", "d/b.txt": "B" });
    expect(tarContents(file.toBufferSync())).toEqual({ "a.txt": "A", "d/b.txt": "B" });
  });

  it("keeps the source entries when one is added, and replaces one added under its name", async () => {
    const file = Archive.ArchiveFile.fromBuffer(source, { format: "tar" });
    file.addText("N", "new.txt").addText("A2", "a.txt");
    expect(tarContents(await file.toBuffer())).toEqual({
      "a.txt": "A2",
      "d/b.txt": "B",
      "new.txt": "N"
    });
    // One `a.txt`, not the old one followed by the new one.
    expect(
      Archive.parseTar(await file.toBuffer()).filter(e => e.info.path === "a.txt")
    ).toHaveLength(1);
  });

  it("keeps the source entries in the streamed output", async () => {
    const file = Archive.ArchiveFile.fromBuffer(source, { format: "tar" });
    file.addText("N", "new.txt");
    const chunks: Uint8Array[] = [];
    for await (const chunk of file.stream()) {
      chunks.push(chunk);
    }
    expect(tarContents(concatUint8Arrays(chunks))).toEqual({
      "a.txt": "A",
      "d/b.txt": "B",
      "new.txt": "N"
    });
  });

  it("keeps entry metadata when it rewrites a source entry", async () => {
    const mtime = new Date("2001-02-03T04:05:06Z");
    const original = Archive.createTar()
      .add("x.sh", "#!/bin/sh", { mode: 0o755, uid: 7, gid: 8, uname: "u", gname: "g", mtime })
      .addSymlink("link", "x.sh")
      .toBufferSync();
    const out = await Archive.ArchiveFile.fromBuffer(original, { format: "tar" }).toBuffer();
    const info = Archive.unpackTarSync(out).get("x.sh")!.info;
    expect(info).toMatchObject({ mode: 0o755, uid: 7, gid: 8, uname: "u", gname: "g", mtime });
    expect(Archive.unpackTarSync(out).get("link")!.info).toMatchObject({
      type: Archive.TarType.SYMLINK,
      linkname: "x.sh"
    });
  });

  it("writes a .tgz it opened back as gzip", async () => {
    const tgzPath = path.join(dir, "in.tgz");
    writeFileSync(tgzPath, Archive.packTarGzSync([["a.txt", "A"]]));
    const file = await Archive.ArchiveFile.fromFile(tgzPath, { format: "tar" });
    file.addText("N", "new.txt");
    const outPath = path.join(dir, "out.tgz");
    await file.writeFile(outPath);
    const out = new Uint8Array(readFileSync(outPath));
    expect([out[0], out[1]]).toEqual([0x1f, 0x8b]);
    expect(
      Object.fromEntries(
        [...Archive.unpackTarGzSync(out)].map(([name, { data }]) => [name, decoder.decode(data)])
      )
    ).toEqual({ "a.txt": "A", "new.txt": "N" });
  });

  it("counts and lists the entries of an opened TAR synchronously", () => {
    const file = Archive.ArchiveFile.fromBuffer(source, { format: "tar" });
    expect(file.entryCount).toBe(2);
    expect(file.getEntryNames()).toEqual(["a.txt", "d/b.txt"]);
    expect(file.getEntriesSync().map(e => e.path)).toEqual(["a.txt", "d/b.txt"]);
  });
});

describe("ArchiveFile builds the same archive every time", () => {
  it.each(["zip", "tar"] as const)("%s: building twice gives the same entries", async format => {
    const file = new Archive.ArchiveFile({ format } as { format: "tar" }).addText("A", "a.txt");
    const read = format === "zip" ? zipContents : tarContents;
    expect(read(await file.toBuffer())).toEqual({ "a.txt": "A" });
    expect(read(await file.toBuffer())).toEqual({ "a.txt": "A" });
    expect(read(file.toBufferSync())).toEqual({ "a.txt": "A" });
  });

  it.each(["zip", "tar"] as const)(
    "%s: adding after a build adds to what was built",
    async format => {
      const file = new Archive.ArchiveFile({ format } as { format: "tar" }).addText("A", "a.txt");
      await file.toBuffer();
      file.addText("B", "b.txt");
      const read = format === "zip" ? zipContents : tarContents;
      expect(read(await file.toBuffer())).toEqual({ "a.txt": "A", "b.txt": "B" });
    }
  );

  it("a new ZIP can be read once it has been written", async () => {
    const file = new Archive.ArchiveFile().addText("A", "a.txt");
    await file.toBuffer();
    expect(file.getEntryNames()).toEqual(["a.txt"]);
    expect(decoder.decode((await file.readEntry("a.txt"))!)).toBe("A");
  });
});

describe("ArchiveFile ZIP edits reach every output", () => {
  const source = Archive.packZipSync([
    ["a.txt", "A"],
    ["b.txt", "B"],
    ["c.txt", "C"]
  ]);
  const edited = () =>
    Archive.ArchiveFile.fromBuffer(source)
      .addText("N", "new.txt")
      .set("b.txt", "B2") as InstanceType<typeof Archive.ArchiveFile>;
  const expected = { "z.txt": "A", "b.txt": "B2", "new.txt": "N" };
  const edit = (file: InstanceType<typeof Archive.ArchiveFile>) => {
    file.rename("a.txt", "z.txt");
    file.delete("c.txt");
    return file;
  };

  it("toBuffer / toBufferSync write a renamed entry under its new name", async () => {
    expect(zipContents(await edit(edited()).toBuffer())).toEqual(expected);
    expect(zipContents(edit(edited()).toBufferSync())).toEqual(expected);
  });

  it("stream() writes the opened archive's entries and its edits", async () => {
    const chunks: Uint8Array[] = [];
    for await (const chunk of edit(edited()).stream()) {
      chunks.push(chunk);
    }
    expect(zipContents(concatUint8Arrays(chunks))).toEqual(expected);
  });
});

describe("ArchiveFile checks its arguments", () => {
  it("refuses an unknown format", () => {
    expect(() => new Archive.ArchiveFile({ format: "rar" } as never)).toThrow(ArchiveError);
    expect(() =>
      Archive.ArchiveFile.fromBuffer(new Uint8Array(0), { format: "rar" } as never)
    ).toThrow(ArchiveError);
  });

  it.each(["zip", "tar"] as const)("%s: refuses entry data of the wrong type", format => {
    const file = () => new Archive.ArchiveFile({ format } as { format: "tar" });
    expect(() => file().addBuffer(null as never, "x")).toThrow(ArchiveError);
    expect(() => file().addBuffer("text" as never, "x")).toThrow(ArchiveError);
    expect(() => file().addText(42 as never, "x")).toThrow(ArchiveError);
  });

  it("refuses a comment that is not a string", () => {
    expect(() => new Archive.ArchiveFile().addZipComment(42 as never)).toThrow(ArchiveError);
  });

  it("refuses set() data of the wrong type", () => {
    expect(() => new Archive.ArchiveFile().set("x", 42 as never)).toThrow(ArchiveError);
  });

  it("writeToFile honours overwrite for TAR as it does for ZIP", async () => {
    const target = path.join(dir, "exists.tar");
    writeFileSync(target, "keep");
    const file = new Archive.ArchiveFile({ format: "tar" }).addText("A", "a.txt");
    await expect(file.writeFile(target)).rejects.toBeInstanceOf(ArchiveError);
    expect(() => file.writeFileSync(target)).toThrow(ArchiveError);
    expect(readFileSync(target, "utf8")).toBe("keep");
    await file.writeFile(target, { overwrite: "overwrite" });
    expect(tarContents(new Uint8Array(readFileSync(target)))).toEqual({ "a.txt": "A" });
  });
});

describe("ArchiveFile names in the shared vocabulary", () => {
  it("addSymlink takes the mode as an option, as the builders do", async () => {
    const zip = await new Archive.ArchiveFile().addSymlink("link", "a.txt").toBuffer();
    expect(Archive.unpackZipSync(zip).get("link")!.info.type).toBe("symlink");
    const tar = new Archive.ArchiveFile({ format: "tar" })
      .addSymlink("link", "a.txt", { mode: 0o755 })
      .toBufferSync();
    expect(Archive.unpackTarSync(tar).get("link")!.info).toMatchObject({
      type: Archive.TarType.SYMLINK,
      linkname: "a.txt",
      mode: 0o755
    });
  });

  it("setZipComment sets the comment addZipComment set", () => {
    const file = new Archive.ArchiveFile().addText("A", "a.txt").setZipComment("hello");
    expect(file.getZipComment()).toBe("hello");
    const bytes = file.toBufferSync();
    expect(Archive.ArchiveFile.fromBuffer(bytes).getZipComment()).toBe("hello");
  });

  it("writeFile / writeFileSync write what writeToFile wrote", async () => {
    const file = new Archive.ArchiveFile().addText("A", "a.txt");
    const a = path.join(dir, "names-a.zip");
    const b = path.join(dir, "names-b.zip");
    await file.writeFile(a);
    file.writeFileSync(b);
    expect(zipContents(new Uint8Array(readFileSync(a)))).toEqual({ "a.txt": "A" });
    expect(zipContents(new Uint8Array(readFileSync(b)))).toEqual({ "a.txt": "A" });
  });
});
