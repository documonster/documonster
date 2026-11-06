/**
 * The extraction side of the zip-slip rule: an archive written elsewhere — or with an
 * explicit `path` mode that keeps a leading `..` — is refused when extracted, and nothing
 * lands outside the destination.
 */
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-zip-slip-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("extracting a traversal name", () => {
  it("ArchiveFile.extractTo refuses it and writes nothing outside", async () => {
    const hostile = await Archive.createZip({ path: { mode: "posix" } })
      .add("../escaped.txt", "pwned")
      .toBuffer();
    const target = path.join(dir, "out");
    await expect(Archive.ArchiveFile.fromBuffer(hostile).extractTo(target)).rejects.toBeInstanceOf(
      ArchiveError
    );
    expect(existsSync(path.join(dir, "escaped.txt"))).toBe(false);
  });
});
