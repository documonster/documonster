/**
 * An `encoding` label `TextDecoder` does not know used to escape as a bare `TypeError` from the
 * shared decoder or a `RangeError` — and an unknown ZIP name encoding as a `TypeError` once the
 * archive was written. Each is an `ArchiveError` naming the option.
 */
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const BOGUS = "no-such-encoding";
const dir = mkdtempSync(path.join(tmpdir(), "documonster-encoding-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("archive encoding labels", () => {
  it("refuses an unknown ZIP name encoding when the archive is created", () => {
    expect(() => Archive.createZip({ encoding: BOGUS as "utf-8" })).toThrow(ArchiveError);
    expect(() => Archive.createZip({ encoding: BOGUS as "utf-8" })).toThrow(/encoding/);
    expect(() => Archive.createZip().add("a", "x", { encoding: 5 as unknown as "utf-8" })).toThrow(
      ArchiveError
    );
  });

  it("keeps the built-in and custom name codecs", async () => {
    const codec = {
      name: "upper",
      encode: (s: string) => new TextEncoder().encode(s.toUpperCase()),
      decode: (b: Uint8Array) => new TextDecoder().decode(b)
    };
    for (const encoding of ["utf-8", "cp437", codec] as const) {
      const bytes = await Archive.createZip({ encoding }).add("a.txt", "x").toBuffer();
      expect(bytes.length).toBeGreaterThan(0);
    }
  });

  it("refuses an unknown label when an entry is read as text", async () => {
    const zip = await Archive.packZip([["a.txt", "x"]]);
    const entry = await Archive.openZip(zip).get("a.txt");
    await expect(entry!.text(BOGUS)).rejects.toThrow(ArchiveError);
    const tar = await Archive.packTar([{ name: "a.txt", source: "x" }]);
    for await (const tarEntry of Archive.openTar(tar).entries()) {
      await expect(tarEntry.text(BOGUS)).rejects.toThrow(ArchiveError);
    }
    const file = path.join(dir, "a.zip");
    writeFileSync(file, zip);
    const archive = await Archive.ArchiveFile.fromFile(file);
    await expect(archive.readAsText("a.txt", { encoding: BOGUS })).rejects.toThrow(ArchiveError);
    expect(() => archive.readAsTextSync("a.txt", { encoding: BOGUS })).toThrow(ArchiveError);
    expect(await archive.readAsText("a.txt", { encoding: "utf-8" })).toBe("x");
  });
});
