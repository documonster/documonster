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

const FILES = ROOTS.flatMap(docxFiles);

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
    const file = FILES.find(path => path.includes("word-probes"))!;
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
    // Move the first paragraph's style to the end of its properties: `w:pStyle` must come first.
    const broken = body.replace(/<w:pPr>(<w:pStyle [^>]*\/>)(.*?)<\/w:pPr>/, "<w:pPr>$2$1</w:pPr>");
    expect(broken).not.toBe(body);
    parts.set("word/document.xml", broken);
    expect(checkDocxElementOrder(parts)).not.toEqual([]);
  });
});
