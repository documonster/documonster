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

import { Document, Build, Io } from "../index";

describe("Io.writeFile / Io.readFile", () => {
  it("round-trips a document through the filesystem", async () => {
    const dir = await createTempDir("documonster-word-io-");
    const filePath = join(dir, "out.docx");

    const h = Document.create();
    Document.addHeading(h, "File IO Title", 1);
    Document.addContent(h, Build.textParagraph("hello from disk"));
    const doc = Document.build(h);

    await Io.writeFile(doc, filePath);
    const parsed = await Io.read(await Io.toBuffer(doc));
    const roundTripped = await Io.readFile(filePath);

    // readFile result must equal reading the same bytes via read().
    expect(roundTripped.body.length).toBe(parsed.body.length);
    expect(JSON.stringify(roundTripped.body)).toBe(JSON.stringify(parsed.body));
  });
});
