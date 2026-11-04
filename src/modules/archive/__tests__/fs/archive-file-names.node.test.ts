/**
 * `ArchiveFile` names ZIP entries under the same rule as `ZipArchive.add`: with no `path`
 * option, an absolute name or a `..` that climbs above the archive root is refused when the
 * entry is added; an explicit `path` mode is how a caller writes one on purpose.
 */
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-archive-file-names-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
writeFileSync(path.join(dir, "a.txt"), "a");
mkdirSync(path.join(dir, "tree"));
writeFileSync(path.join(dir, "tree", "b.txt"), "b");

const names = async (bytes: Uint8Array) => [...(await Archive.unpackZip(bytes)).keys()].sort();

describe("ArchiveFile ZIP entry names", () => {
  it.each<[string, (zip: InstanceType<typeof Archive.ArchiveFile>) => unknown]>([
    // `addText` is `addBuffer` with encoded text.
    ["addBuffer", zip => zip.addText("x", "../../escaped.txt")],
    ["addBuffer", zip => zip.addBuffer(new Uint8Array(1), "/etc/passwd")],
    ["addFile", zip => zip.addFile(path.join(dir, "a.txt"), { name: "..\\evil.txt" })],
    ["addFile", zip => zip.addFile(path.join(dir, "a.txt"), { prefix: "../up" })],
    ["symlink", zip => zip.symlink("C:/evil", "a.txt")],
    ["set", zip => zip.set("../evil.txt", "x")],
    ["rename", zip => zip.addText("x", "ok.txt").rename("ok.txt", "../evil.txt")]
  ])("%s refuses a traversal name when it is called", (method, call) => {
    const zip = new Archive.ArchiveFile();
    expect(() => call(zip)).toThrow(ArchiveError);
    expect(() => call(zip)).toThrow(`Archive.ArchiveFile#${method}`);
  });

  it("refuses a directory prefix that climbs out, in the buffered and streaming builds", async () => {
    const zip = new Archive.ArchiveFile().addDirectory(path.join(dir, "tree"), { prefix: "../up" });
    await expect(zip.toBuffer()).rejects.toBeInstanceOf(ArchiveError);
    expect(() => zip.toBufferSync()).toThrow(ArchiveError);
  });

  it("writes the same names as before for everything that stays inside", async () => {
    const zip = new Archive.ArchiveFile()
      .addText("x", "a/../b.txt")
      .addFile(path.join(dir, "a.txt"), { prefix: "p" })
      .addDirectory(path.join(dir, "tree"));
    expect(await names(await zip.toBuffer())).toEqual(["a/../b.txt", "p/a.txt", "tree/b.txt"]);
  });

  it("writes a traversal name when a path mode is chosen explicitly", async () => {
    const zip = new Archive.ArchiveFile({ path: { mode: "legacy" } }).addText(
      "x",
      "../../escaped.txt"
    );
    expect(await names(await zip.toBuffer())).toEqual(["../../escaped.txt"]);
    expect(await names(zip.toBufferSync())).toEqual(["../../escaped.txt"]);
  });
});
