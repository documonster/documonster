import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { Archive, ArchiveError, ArchiveFileError, Fs, isArchiveError } from "@archive";
import { afterAll, describe, expect, it } from "vitest";

const dir = mkdtempSync(path.join(tmpdir(), "documonster-archive-file-errors-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

function expectFileError(
  error: unknown,
  filePath: string,
  operation: "read" | "write",
  code: string
): void {
  expect(error).toBeInstanceOf(ArchiveFileError);
  expect(error).toBeInstanceOf(ArchiveError);
  expect(isArchiveError(error)).toBe(true);
  const fileError = error as ArchiveFileError;
  expect(fileError.name).toBe("ArchiveFileError");
  expect(fileError.path).toBe(filePath);
  expect(fileError.operation).toBe(operation);
  expect((fileError.cause as { code?: string }).code).toBe(code);
  expect(fileError.message).toContain(`Failed to ${operation} `);
  expect(fileError.message).toContain(`file "${filePath}"`);
}

function caught(fn: () => unknown): unknown {
  try {
    fn();
  } catch (error) {
    return error;
  }
  throw new Error("expected a throw");
}

describe("ArchiveFile wraps the file system's refusals", () => {
  const missing = path.join(dir, "missing.zip");

  it("fromFile and fromFileSync report a missing file as ArchiveFileError", async () => {
    const asyncError = await Archive.ArchiveFile.fromFile(missing).catch((e: unknown) => e);
    expectFileError(asyncError, missing, "read", "ENOENT");
    expect((asyncError as Error).message).toContain("File not found");

    expectFileError(
      caught(() => Archive.ArchiveFile.fromFileSync(missing)),
      missing,
      "read",
      "ENOENT"
    );
  });

  it("writeFile and writeFileSync report an unwritable target as ArchiveFileError", async () => {
    // A regular file where a parent directory is needed: ensureDir cannot create it.
    const blocker = path.join(dir, "blocker");
    writeFileSync(blocker, "x");
    const target = path.join(blocker, "out.zip");
    const archive = new Archive.ArchiveFile();
    archive.addText("alpha", "a.txt");

    const asyncError = await archive.writeFile(target).catch((e: unknown) => e);
    expect(asyncError).toBeInstanceOf(ArchiveFileError);
    expect((asyncError as ArchiveFileError).path).toBe(target);
    expect((asyncError as ArchiveFileError).operation).toBe("write");
    expect((asyncError as ArchiveFileError).cause).toBeInstanceOf(Error);

    const syncError = caught(() => archive.writeFileSync(target));
    expect(syncError).toBeInstanceOf(ArchiveFileError);
    expect((syncError as ArchiveFileError).operation).toBe("write");
  });

  it("a directory as the target is a write error with the system's code", async () => {
    const asDir = path.join(dir, "a-dir.zip");
    mkdirSync(asDir);
    const archive = new Archive.ArchiveFile();
    archive.addText("alpha", "a.txt");
    const error = await archive
      .writeFile(asDir, { overwrite: "overwrite" })
      .catch((e: unknown) => e);
    expectFileError(error, asDir, "write", "EISDIR");
  });

  it("an existing file is still refused by the overwrite policy, not the file system", async () => {
    const existing = path.join(dir, "existing.zip");
    writeFileSync(existing, "x");
    const archive = new Archive.ArchiveFile();
    archive.addText("alpha", "a.txt");
    const error = await archive.writeFile(existing).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ArchiveError);
    expect(error).not.toBeInstanceOf(ArchiveFileError);
  });

  it("Fs helpers pass the platform's error on unwrapped", async () => {
    const error = await Fs.readFileBytes(missing).catch((e: unknown) => e);
    expect(error).not.toBeInstanceOf(ArchiveError);
    expect((error as { code?: string }).code).toBe("ENOENT");
  });

  it("a local source that has gone is a read error naming the source", async () => {
    const source = path.join(dir, "gone.txt");
    writeFileSync(source, "x");
    const archive = new Archive.ArchiveFile();
    archive.addFile(source);
    rmSync(source);

    const asyncError = await archive.toBuffer().catch((e: unknown) => e);
    expectFileError(asyncError, source, "read", "ENOENT");
    expect((asyncError as Error).message).toContain("Failed to read source file");
    expectFileError(
      caught(() => archive.toBufferSync()),
      source,
      "read",
      "ENOENT"
    );

    const tar = new Archive.ArchiveFile({ format: "tar" });
    tar.addFile(path.join(dir, "never-there.txt"));
    expectFileError(
      await tar.toBuffer().catch((e: unknown) => e),
      path.join(dir, "never-there.txt"),
      "read",
      "ENOENT"
    );
  });

  it("a source that fails while streaming surfaces as a read error", async () => {
    const source = path.join(dir, "stream-gone.txt");
    writeFileSync(source, "x");
    const archive = new Archive.ArchiveFile();
    archive.addFile(source);
    rmSync(source);
    const error = await (async () => {
      for await (const _chunk of archive.stream()) {
        // drain
      }
    })().catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ArchiveFileError);
    expect((error as ArchiveFileError).operation).toBe("read");
  });

  it("extracting under a file is a write error, sync and async, ZIP and TAR", async () => {
    const blocker = path.join(dir, "extract-blocker");
    writeFileSync(blocker, "x");
    const target = path.join(blocker, "out");

    for (const format of ["zip", "tar"] as const) {
      const archive = new Archive.ArchiveFile({ format } as never);
      archive.addText("alpha", "a.txt");
      const asyncError = await archive.extractTo(target).catch((e: unknown) => e);
      expect(asyncError).toBeInstanceOf(ArchiveFileError);
      expect((asyncError as ArchiveFileError).operation).toBe("write");
      expect((asyncError as Error).message).toContain("extracted file");
      const syncError = caught(() => archive.extractToSync(target));
      expect(syncError).toBeInstanceOf(ArchiveFileError);
    }

    const zip = new Archive.ArchiveFile();
    zip.addText("alpha", "a.txt");
    const entryError = await zip
      .extractEntryTo("a.txt", path.join(target, "a.txt"))
      .catch((e: unknown) => e);
    expect(entryError).toBeInstanceOf(ArchiveFileError);
    expect(
      caught(() => zip.extractEntryToSync("a.txt", path.join(target, "a.txt")))
    ).toBeInstanceOf(ArchiveFileError);
  });

  it("streamToFile under a file is a write error naming the target", async () => {
    const blocker = path.join(dir, "stream-blocker");
    writeFileSync(blocker, "x");
    const target = path.join(blocker, "out.zip");
    const archive = new Archive.ArchiveFile();
    archive.addText("alpha", "a.txt");
    const error = await archive.streamToFile(target).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ArchiveFileError);
    expect((error as ArchiveFileError).path).toBe(target);
    expect((error as ArchiveFileError).operation).toBe("write");
  });
});
