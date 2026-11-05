/**
 * DOCX Module - File-path IO (Node only) Tests
 *
 * `Io.readFile` / `Io.writeFile` mirror Excel's `Workbook.readFile` /
 * `Workbook.writeFile`. The browser build of `Io` does not publish them, so
 * these tests are Node-only and named `*.node.test.ts` to be excluded from the
 * browser suite.
 */

import { join } from "node:path";

import { createTempDir } from "@utils/fs";
import { describe, it, expect } from "vitest";

import {
  Document,
  Build,
  DocxArgumentError,
  DocxFileError,
  DocxParseError,
  Io,
  isDocxError
} from "../index";

describe("Io.writeFile / Io.readFile", () => {
  it("round-trips a document through the filesystem", async () => {
    const dir = await createTempDir("documonster-word-io-");
    const filePath = join(dir, "out.docx");

    const h = Document.create();
    Document.addHeading(h, "File IO Title", 1);
    Document.addContent(h, Build.textParagraph("hello from disk"));
    const doc = h;

    await Io.writeFile(doc, filePath);
    const parsed = await Io.read(await Io.toBuffer(doc));
    const roundTripped = await Io.readFile(filePath);

    // readFile result must equal reading the same bytes via read().
    expect(roundTripped.body.length).toBe(parsed.body.length);
    expect(JSON.stringify(roundTripped.body)).toBe(JSON.stringify(parsed.body));
  });

  it("refuse a path that is not a non-empty string with DocxArgumentError", async () => {
    const doc = Document.create();
    for (const path of [42, undefined, null, "", {}]) {
      const read = await Io.readFile(path as never).catch((e: unknown) => e);
      expect(read, String(path)).toBeInstanceOf(DocxArgumentError);
      expect((read as DocxArgumentError).field).toBe("path");
      expect((read as Error).message).toMatch(/^Io\.readFile: path /);
      const write = await Io.writeFile(doc, path as never).catch((e: unknown) => e);
      expect(write, String(path)).toBeInstanceOf(DocxArgumentError);
      expect((write as DocxArgumentError).field).toBe("path");
      expect((write as Error).message).toMatch(/^Io\.writeFile: path /);
    }
  });

  it("wrap a file-system error in DocxFileError, keeping the system's error as cause", async () => {
    const dir = await createTempDir("documonster-word-io-");
    const missing = join(dir, "missing.docx");
    const read = await Io.readFile(missing).catch((e: unknown) => e);
    expect(read).toBeInstanceOf(DocxFileError);
    expect(isDocxError(read)).toBe(true);
    expect((read as DocxFileError).path).toBe(missing);
    expect((read as DocxFileError).operation).toBe("read");
    expect((read as Error).message).toBe(`Failed to read Word file "${missing}": File not found`);
    expect(((read as Error).cause as NodeJS.ErrnoException).code).toBe("ENOENT");

    const target = join(dir, "no", "such", "dir.docx");
    const write = await Io.writeFile(Document.create(), target).catch((e: unknown) => e);
    expect(write).toBeInstanceOf(DocxFileError);
    expect((write as DocxFileError).path).toBe(target);
    expect((write as DocxFileError).operation).toBe("write");
    expect(((write as Error).cause as NodeJS.ErrnoException).code).toBe("ENOENT");
  });

  it("report a file that was read but is not a document as a parse error, not a file error", async () => {
    const dir = await createTempDir("documonster-word-io-");
    const path = join(dir, "not.docx");
    await Io.writeFile(Document.create(), path);
    const { writeFile } = await import("node:fs/promises");
    await writeFile(path, "not a zip");
    const error = await Io.readFile(path).catch((e: unknown) => e);
    expect(error).not.toBeInstanceOf(DocxFileError);
    expect(isDocxError(error)).toBe(true);
    void DocxParseError;
  });
});
