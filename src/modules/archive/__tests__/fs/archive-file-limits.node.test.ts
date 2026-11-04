/**
 * `ArchiveFile.fromBuffer` / `fromFile` / `fromFileSync` check the ZIP limit options they take
 * when they are called, as `Archive.openZip` does (an `ArchiveArgumentError` naming the call and
 * `options.<limit>`), instead of opening the archive with no limit.
 */
import { writeFileSync } from "node:fs";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { Archive, ArchiveArgumentError } from "@archive/index";
import { ArchiveFile } from "@archive/surface/archive-node";
import { beforeAll, describe, expect, it } from "vitest";

const zip = Archive.packZipSync([["a.txt", "A"]]);
let file: string;

beforeAll(async () => {
  file = join(await mkdtemp(join(tmpdir(), "dm-archive-limits-")), "a.zip");
  writeFileSync(file, zip);
});

async function failure(run: () => unknown): Promise<unknown> {
  try {
    await run();
    return undefined;
  } catch (error) {
    return error;
  }
}

describe("ArchiveFile openers check their limit options", () => {
  const openers: Array<[string, (options: object) => unknown]> = [
    ["Archive.ArchiveFile.fromBuffer", o => ArchiveFile.fromBuffer(zip, o)],
    ["Archive.ArchiveFile.fromFile", o => ArchiveFile.fromFile(file, o)],
    ["Archive.ArchiveFile.fromFileSync", o => ArchiveFile.fromFileSync(file, o)]
  ];
  for (const [name, open] of openers) {
    for (const limit of ["maxEntrySize", "maxEntries"]) {
      it(`${name}: options.${limit}`, async () => {
        for (const bad of ["x", -1, Number.NaN]) {
          const error = await failure(() => open({ [limit]: bad }));
          expect(error).toBeInstanceOf(ArchiveArgumentError);
          expect((error as ArchiveArgumentError).field).toBe(`options.${limit}`);
          expect((error as Error).message.startsWith(`${name}: `)).toBe(true);
        }
        expect(await failure(() => open({ [limit]: Infinity }))).toBeUndefined();
      });
    }
  }
});
