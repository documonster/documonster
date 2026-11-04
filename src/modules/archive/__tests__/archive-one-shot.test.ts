import { Archive, Compression } from "@archive";
import { TarGzArchive } from "@archive/tar/tar-gzip";
import { describe, expect, it } from "vitest";

const text = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);

describe("Archive format matrix", () => {
  const entries: [string, string][] = [
    ["a.txt", "hello"],
    ["dir/b.txt", "world"]
  ];

  it("packZip → unpackZip round-trips tuples, objects and maps", async () => {
    for (const input of [
      entries,
      entries.map(([name, source]) => ({ name, source })),
      new Map(entries)
    ]) {
      const files = await Archive.unpackZip(await Archive.packZip(input));
      expect([...files.keys()]).toEqual(["a.txt", "dir/b.txt"]);
      expect(text(files.get("dir/b.txt")?.data)).toBe("world");
      expect(files.get("a.txt")?.info.path).toBe("a.txt");
    }
  });

  it("sync and async ZIP forms are interchangeable", async () => {
    const syncZip = Archive.packZipSync(entries, { level: 0 });
    const asyncFiles = await Archive.unpackZip(syncZip);
    const syncFiles = Archive.unpackZipSync(await Archive.packZip(entries));
    expect(text(asyncFiles.get("a.txt")?.data)).toBe("hello");
    expect(text(syncFiles.get("a.txt")?.data)).toBe("hello");
  });

  it("unpackZip reads a streamed source", async () => {
    const zip = await Archive.packZip(entries);
    async function* chunks(): AsyncGenerator<Uint8Array> {
      yield zip.subarray(0, 10);
      yield zip.subarray(10);
    }
    const files = await Archive.unpackZip(chunks());
    expect(text(files.get("dir/b.txt")?.data)).toBe("world");
  });

  it("unpackZip matches what openZip reads", async () => {
    const zip = await Archive.packZip(entries);
    const reader = Archive.openZip(zip);
    const files = await Archive.unpackZip(zip);
    for (const path of files.keys()) {
      expect(await reader.bytes(path)).toEqual(files.get(path)?.data);
    }
  });

  it("unpackTarSync matches unpackTar", async () => {
    const tar = Archive.packTarSync(entries);
    const sync = Archive.unpackTarSync(tar);
    const async_ = await Archive.unpackTar(tar);
    expect([...sync.keys()]).toEqual([...async_.keys()]);
    expect(text(sync.get("dir/b.txt")?.data)).toBe("world");
    expect(sync.get("a.txt")?.info).toEqual(async_.get("a.txt")?.info);
  });

  it("tar.gz has the same four verbs, sync forms included", async () => {
    const tgz = Archive.packTarGzSync(entries);
    expect(Compression.isGzipData(tgz)).toBe(true);
    const sync = Archive.unpackTarGzSync(await Archive.packTarGz(entries));
    const async_ = await Archive.unpackTarGz(tgz);
    expect([...sync.keys()]).toEqual(["a.txt", "dir/b.txt"]);
    expect(text(async_.get("dir/b.txt")?.data)).toBe("world");
  });

  it("openTarGz reads entries like openTar, and can be read twice", async () => {
    const reader = Archive.openTarGz(Archive.packTarGzSync(entries));
    expect(reader).toBeInstanceOf(Archive.TarReader);
    for (let pass = 0; pass < 2; pass++) {
      const seen: string[] = [];
      for await (const entry of reader.entries()) {
        seen.push(`${entry.path}=${await entry.text()}`);
      }
      expect(seen).toEqual(["a.txt=hello", "dir/b.txt=world"]);
    }
  });

  it("builders and readers are instances of the public classes", () => {
    expect(Archive.createZip()).toBeInstanceOf(Archive.ZipArchive);
    expect(Archive.openZip(new Uint8Array())).toBeInstanceOf(Archive.ZipReader);
    expect(Archive.createTar()).toBeInstanceOf(Archive.TarArchive);
    expect(Archive.openTar(new Uint8Array())).toBeInstanceOf(Archive.TarReader);
    expect(Archive.TarGzArchive).toBe(TarGzArchive);
    expect(Archive.createTarGz()).toBeInstanceOf(Archive.TarGzArchive);
  });
});
