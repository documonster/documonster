/**
 * `Stream.DocumentReader` reads every document in the fixture corpus to the body `Document.read` reads.
 *
 * The two share the code that turns a body element into content, but run it differently — the streaming reader a
 * piece at a time, building each element as its bytes arrive, the whole-document reader in one call over the parsed
 * part — and every difference in how they get there (where an element is cut, what state a step leaves behind) would
 * show here as a document that reads differently.
 */
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { Document, Stream } from "../index";
import type { BodyContent } from "../types";

const DATA = fileURLToPath(new URL("./data", import.meta.url));

function docxFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap(entry =>
    entry.isDirectory()
      ? docxFiles(join(dir, entry.name))
      : entry.name.endsWith(".docx")
        ? [join(dir, entry.name)]
        : []
  );
}

const files = docxFiles(DATA).filter(file => !file.includes("office-encrypted"));

describe("Stream.DocumentReader over the fixture corpus", () => {
  it("has documents to read", () => {
    expect(files.length).toBeGreaterThan(100);
  });

  it.each(files.map(file => [file.slice(DATA.length + 1), file]))(
    "reads %s to the body Document.read reads",
    async (_name, file) => {
      const bytes = new Uint8Array(readFileSync(file));
      const streamed: BodyContent[] = [];
      for await (const item of new Stream.DocumentReader(bytes)) {
        streamed.push(item);
      }
      expect(streamed).toEqual((await Document.read(bytes)).body);
    }
  );
});
