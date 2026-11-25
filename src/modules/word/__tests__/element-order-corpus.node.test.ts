/**
 * Every committed document, read and written back, has its children in the order the ECMA-376 content models
 * require (`writer/order-check.ts`).
 *
 * This is the broad gate for the writer's element order. A purpose-built document (`schema-order.test.ts`) can only
 * hold the constructs someone thought to put in it; these two hundred files hold what real documents contain —
 * numbered headings with spacing, East Asian line breaking, revisions, nested tables — and the order defect this
 * was written for (`w:numPr` after `w:spacing`) sat in 370 of their paragraphs. The inputs are not held to the
 * check: most were written by an earlier version of this library and keep its mistakes, and a few are minimal
 * hand-written fixtures. What is checked is what the writer makes of them now.
 */
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Document } from "@word/index";
import { checkDocxElementOrder } from "@word/writer/order-check";
import { describe, expect, it } from "vitest";

const ROOTS = ["src/modules/word/__tests__/data", "src/modules/pdf/__tests__/data/word-reference"];

function docxFiles(dir: string): string[] {
  return readdirSync(dir).flatMap(name => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) {
      // An encrypted package is not a ZIP until decrypted; it has its own tests.
      return name === "office-encrypted" ? [] : docxFiles(path);
    }
    return name.endsWith(".docx") ? [path] : [];
  });
}

// Sorted: `readdirSync`'s order is the runtime's and the file system's, and differs between Node and Bun.
const FILES = ROOTS.flatMap(docxFiles).sort();

describe("element order of rewritten documents", () => {
  it("finds the corpus", () => {
    expect(FILES.length).toBeGreaterThan(150);
  });

  it("writes every committed document back in schema order", async () => {
    const problems: string[] = [];
    for (const file of FILES) {
      const written = await Document.toBuffer(
        await Document.read(new Uint8Array(readFileSync(file)))
      );
      const parts = new Map<string, string>();
      for (const [path, entry] of await extractAll(written)) {
        if (path.endsWith(".xml")) {
          parts.set(path, new TextDecoder().decode(entry.data));
        }
      }
      for (const { part, message } of checkDocxElementOrder(parts)) {
        problems.push(`${file} ${part}: ${message}`);
      }
    }
    expect(problems).toEqual([]);
  }, 120_000);
});

describe("the corpus check is not vacuous", () => {
  it("reports a paragraph property moved out of order in a rewritten document", async () => {
    // A named document with styled paragraphs. "The first probe" was whichever `readdirSync` listed first, which under
    // Bun was one with no `w:pStyle`, so the edit below changed nothing.
    const file = join(ROOTS[0]!, "word-probes", "R1e-lists-borders.docx");
    const written = await Document.toBuffer(
      await Document.read(new Uint8Array(readFileSync(file)))
    );
    const parts = new Map<string, string>();
    for (const [path, entry] of await extractAll(written)) {
      if (path.endsWith(".xml")) {
        parts.set(path, new TextDecoder().decode(entry.data));
      }
    }
    const body = parts.get("word/document.xml")!;
    // Move the first styled paragraph's style to the end of its properties: `w:pStyle` must come first.
    const styled = /<w:pPr>(<w:pStyle [^>]*\/>)(.+?)<\/w:pPr>/;
    expect(
      styled.test(body),
      "the document has a styled paragraph with more properties than its style"
    ).toBe(true);
    const broken = body.replace(styled, "<w:pPr>$2$1</w:pPr>");
    parts.set("word/document.xml", broken);
    expect(checkDocxElementOrder(parts)).not.toEqual([]);
  });
});
