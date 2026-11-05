/**
 * `Fs.*` check their arguments at the boundary: a wrong argument is an `ArchiveArgumentError`
 * naming the call and the parameter, where Node's own `ERR_INVALID_ARG_TYPE` used to escape.
 * What the file system then refuses is still the platform's error, unwrapped.
 */
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import * as entry from "@archive";
import { Archive, ArchiveArgumentError, Fs, isArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(join(tmpdir(), "fs-arguments-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

function caught(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  return undefined;
}

async function caughtAsync(call: () => unknown): Promise<unknown> {
  try {
    await call();
  } catch (error) {
    return error;
  }
  return undefined;
}

function expectArgumentError(error: unknown, field: string, where: string): void {
  expect(error).toBeInstanceOf(ArchiveArgumentError);
  expect((error as ArchiveArgumentError).field).toBe(field);
  expect((error as Error).message.startsWith(`${where}: `)).toBe(true);
}

type AnyFn = (...args: unknown[]) => unknown;
const fs = Fs as unknown as Record<string, AnyFn>;

describe("Fs argument checks", () => {
  it.each<[string, string, unknown[]]>([
    ["readFileBytes", "filePath", []],
    ["readFileBytesSync", "filePath", [42]],
    ["readFileText", "filePath", [null]],
    ["readFileTextSync", "encoding", ["a.txt", "bogus"]],
    ["writeFileBytes", "data", ["a.bin", "text"]],
    ["writeFileBytesSync", "filePath", [{}]],
    ["writeFileText", "content", ["a.txt", 42]],
    ["writeFileTextSync", "encoding", ["a.txt", "x", 42]],
    ["setFileTime", "mtime", ["a.txt"]],
    ["setFileTimeSync", "mtime", ["a.txt", new Date(Number.NaN)]],
    ["traverseDirectory", "dirPath", []],
    ["traverseDirectorySync", "options", [".", null]],
    ["traverseDirectorySync", "options.filter", [".", { filter: 1 }]],
    ["glob", "pattern", []],
    ["globSync", "options.ignore", ["*", { ignore: 42 }]],
    ["globSync", "options.ignore[1]", ["*", { ignore: ["a", 1] }]],
    ["globSync", "options.cwd", ["*", { cwd: 1 }]],
    ["globToRegex", "pattern", []],
    ["globToRegex", "options.dot", ["*", { dot: "yes" }]],
    ["matchGlob", "pattern", ["a.txt"]],
    ["matchGlobAny", "patterns", ["a.txt", "*"]],
    ["matchGlobAny", "patterns[0]", ["a.txt", [1]]],
    ["createTempDir", "prefix", [null]],
    ["createTempDirSync", "prefix", [42]],
    ["ensureDir", "dirPath", [undefined]],
    ["ensureDirSync", "dirPath", ["a\0b"]],
    ["fileExists", "filePath", [42]],
    ["fileExistsSync", "filePath", [42]],
    ["safeStats", "filePath", [{}]],
    ["safeStatsSync", "filePath", [{}]],
    ["remove", "targetPath", []],
    ["removeSync", "targetPath", [1]],
    ["copyFile", "dest", ["a.txt"]],
    ["copyFileSync", "src", [null, "b.txt"]]
  ])("Fs.%s refuses %s", async (name, field, args) => {
    const error =
      name.endsWith("Sync") || name === "globToRegex" || name.startsWith("matchGlob")
        ? caught(() => fs[name]!(...args))
        : await caughtAsync(() => fs[name]!(...args));
    expectArgumentError(error, field, `Fs.${name}`);
  });

  it("an async generator member refuses its arguments when called, not on the first next()", () => {
    expectArgumentError(
      caught(() => Fs.glob(42 as never)),
      "pattern",
      "Fs.glob"
    );
    expectArgumentError(
      caught(() => Fs.traverseDirectory(dir, { recursive: "no" } as never)),
      "options.recursive",
      "Fs.traverseDirectory"
    );
  });

  it("passes the file system's refusal on unwrapped", async () => {
    const missing = join(dir, "missing.txt");
    const error = (await caughtAsync(() => Fs.readFileBytes(missing))) as NodeJS.ErrnoException;
    expect(error.code).toBe("ENOENT");
    expect(isArchiveError(error)).toBe(false);
    expect((caught(() => Fs.readFileTextSync(missing)) as NodeJS.ErrnoException).code).toBe(
      "ENOENT"
    );
  });

  it("still works for valid arguments", async () => {
    const file = join(dir, "a.txt");
    await Fs.writeFileText(file, "hello");
    expect(await Fs.readFileText(file)).toBe("hello");
    expect(Fs.readFileBytesSync(file)).toEqual(new TextEncoder().encode("hello"));
    expect(Fs.globSync("*.txt", { cwd: dir }).map(e => e.relativePath)).toEqual(["a.txt"]);
    expect(Fs.matchGlobAny("a/b.txt", ["**/*.txt"])).toBe(true);
    await Fs.setFileTime(file, new Date(0));
    expect((await Fs.safeStats(file))?.mtime.getTime()).toBe(0);
  });

  it("does not publish file streams: those are node:fs's", () => {
    expect("createReadStream" in Fs).toBe(false);
    expect("createWriteStream" in Fs).toBe(false);
    // The deprecated flat forwards keep working until 1.0.
    expect(typeof entry.createReadStream).toBe("function");
    expect(typeof entry.createWriteStream).toBe("function");
  });
});

describe("Node-only Archive members check their options", () => {
  it("new Archive.ArchiveFile refuses a non-object options argument", () => {
    expectArgumentError(
      caught(() => new Archive.ArchiveFile(null as never)),
      "options",
      "Archive.ArchiveFile"
    );
  });

  it("Archive.toNodeReadable refuses a non-object options argument", () => {
    expectArgumentError(
      caught(() => Archive.toNodeReadable(new Uint8Array(1), null as never)),
      "options",
      "Archive.toNodeReadable"
    );
  });
});
