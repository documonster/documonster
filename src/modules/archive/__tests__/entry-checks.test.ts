/**
 * Archive writers check what they are asked to write when the entry is added: a zip-slip
 * name, a compression level outside 0–9, an invalid date or a mode the header cannot hold
 * is an `ArchiveError` naming the call and the field.
 */
import { Archive, ArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

function expectArchiveError(call: () => unknown, ...fragments: string[]): void {
  let caught: unknown;
  try {
    call();
  } catch (err) {
    caught = err;
  }
  expect(caught).toBeInstanceOf(ArchiveError);
  for (const fragment of fragments) {
    expect((caught as Error).message).toContain(fragment);
  }
}

const UNSAFE_NAMES = [
  "../evil.txt",
  "a/../../evil.txt",
  "..\\evil.txt",
  "/etc/passwd",
  "\\windows\\evil.txt",
  "C:/evil.txt",
  "c:evil.txt",
  ".."
];

describe("entry names", () => {
  it.each(UNSAFE_NAMES)("createZip().add refuses %j", name => {
    const zip = Archive.createZip();
    expectArchiveError(() => zip.add(name, "x"), "Archive.ZipArchive#add", "name");
  });

  it.each(UNSAFE_NAMES)("createTar().add refuses %j", name => {
    expectArchiveError(() => Archive.createTar().add(name, "x"), "Archive.TarArchive#add", "name");
  });

  it("refuses an unsafe name in the one-shot forms before writing anything", async () => {
    await expect(Archive.packZip([["../evil.txt", "x"]])).rejects.toBeInstanceOf(ArchiveError);
    expect(() => Archive.packTarSync([{ name: "/abs", source: "x" }])).toThrow(ArchiveError);
  });

  it("refuses an unsafe name in ZipEditor.set and rename", async () => {
    const editor = await Archive.editZip(await Archive.packZip([["a.txt", "a"]]));
    expectArchiveError(() => editor.set("../evil.txt", "x"), "Archive.ZipEditor#set");
    expectArchiveError(() => editor.rename("a.txt", "/evil.txt"), "Archive.ZipEditor#rename");
    expect(editor.has("a.txt")).toBe(true);
  });

  it("keeps names that stay inside the archive", async () => {
    const zip = Archive.createZip()
      .add("a/../b.txt", "b")
      .add("./c.txt", "c")
      .add("dir/sub/..x", "x")
      .addDirectory("d");
    const files = await Archive.unpackZip(await zip.toBuffer());
    expect([...files.keys()].sort()).toEqual(["./c.txt", "a/../b.txt", "d/", "dir/sub/..x"]);
  });

  it("lets an explicit path mode decide instead", async () => {
    // `posix` normalises but keeps a leading `..`: choosing a mode is choosing its rules.
    const zip = Archive.createZip({ path: { mode: "posix" } }).add("../kept.txt", "x");
    expect([...(await Archive.unpackZip(await zip.toBuffer())).keys()]).toEqual(["../kept.txt"]);
    // Refused when it is added, under the mode's rules, rather than when the archive is written.
    expect(() => Archive.createZip({ path: { mode: "safe" } }).add("../evil.txt", "x")).toThrow(
      /Unsafe ZIP path/
    );
  });

  it("applies the path mode to toBuffer and toBufferSync as stream() always did", async () => {
    const make = () => Archive.createZip({ path: { mode: "legacy" } }).add("\\dir\\a.txt", "a");
    const names = async (bytes: Uint8Array) => [...(await Archive.unpackZip(bytes)).keys()];
    expect(await names(await make().toBuffer())).toEqual(["dir/a.txt"]);
    expect(await names(make().toBufferSync())).toEqual(["dir/a.txt"]);
    const streamed = await new Response(
      Archive.toReadableStream(make().stream()) as ReadableStream<Uint8Array>
    ).arrayBuffer();
    expect(await names(new Uint8Array(streamed))).toEqual(["dir/a.txt"]);
  });
});

describe("builder options", () => {
  it.each<[string, () => unknown]>([
    ["level", () => Archive.createZip({ level: 10 })],
    ["level", () => Archive.createZip({ level: NaN })],
    ["level", () => Archive.createZip({ level: 1.5 })],
    ["modTime", () => Archive.createZip({ modTime: new Date(NaN) })],
    ["timestamps", () => Archive.createZip({ timestamps: "unix" as "dos" })],
    ["zip64", () => Archive.createZip({ zip64: "yes" as "auto" })],
    ["comment", () => Archive.createZip({ comment: 5 as unknown as string })],
    ["level", () => Archive.createZip().add("a", "x", { level: -1 })],
    ["mode", () => Archive.createZip().add("a", "x", { mode: 0x10000 })],
    ["atime", () => Archive.createZip().add("a", "x", { atime: new Date("x") })],
    ["level", () => Archive.createTarGz({ level: 11 })],
    // A date before 1970 and a uid past 7 octal digits are written in GNU base-256 now
    // (see tar-header-fields.test.ts); these pinned the old limits.
    ["modTime", () => Archive.createTar({ modTime: new Date("not a date") })],
    ["mode", () => Archive.createTar().add("a", "x", { mode: NaN })],
    ["uid", () => Archive.createTar().add("a", "x", { uid: 2 ** 53 })],
    ["gid", () => Archive.createTar().add("a", "x", { gid: -1 })],
    ["mtime", () => Archive.createTar().add("a", "x", { mtime: new Date(NaN) })]
  ])("refuses %s", (field, call) => {
    expectArchiveError(call, field);
  });

  it("refuses bad editor defaults when the editor is opened", async () => {
    const bytes = await Archive.packZip([["a.txt", "a"]]);
    await expect(Archive.editZip(bytes, { level: 12 })).rejects.toThrow(/level must be/);
  });

  it("keeps accepting the boundaries", async () => {
    const zip = Archive.createZip({ level: 0, zip64: false, timestamps: "dos+utc+ntfs" })
      .add("a", "x", { level: 9, mode: 0o100644, modTime: new Date(0) })
      .add("b", "y", { mode: 0xffff });
    expect((await Archive.unpackZip(await zip.toBuffer())).size).toBe(2);
    const tar = Archive.createTar({ modTime: new Date(0) }).add("a", "x", {
      mode: 0o7777777,
      uid: 0o7777777,
      gid: 0
    });
    expect((await Archive.unpackTar(await tar.toBuffer())).size).toBe(1);
    expect(
      (await Archive.createTarGz({ level: 0 }).add("a", "x").toBuffer()).length
    ).toBeGreaterThan(0);
  });
});
