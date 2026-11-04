/**
 * DOCX Module - File-path IO (Node only) Tests
 *
 * `Io.readFile` / `Io.writeFile` mirror Excel's `Workbook.readFile` /
 * `Workbook.writeFile`. They use `@utils/fs` (whose `.browser` variant throws),
 * so these tests are Node-only and named `*.node.test.ts` to be excluded from
 * the browser suite.
 */

import { join } from "node:path";

import { createTempDir } from "@utils/fs";
import { describe, it, expect } from "vitest";

import { Document, Build, DocxArgumentError, Io } from "../index";

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

  it("let a file-system error through unchanged, code and all", async () => {
    const dir = await createTempDir("documonster-word-io-");
    const missing = join(dir, "missing.docx");
    const read = await Io.readFile(missing).catch((e: unknown) => e);
    expect((read as NodeJS.ErrnoException).code).toBe("ENOENT");
    const write = await Io.writeFile(Document.create(), join(dir, "no", "such", "dir.docx")).catch(
      (e: unknown) => e
    );
    expect((write as NodeJS.ErrnoException).code).toBe("ENOENT");
  });
});
