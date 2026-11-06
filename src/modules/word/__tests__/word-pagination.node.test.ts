/**
 * Page numbers against Word's own.
 *
 * Each fixture in `data/word-pagination` is a document the library wrote with
 * `Transform.updateFields`: default styles and none, headings 1–3, wrapped
 * paragraphs, page breaks, next- and odd-page sections, A4 and Letter, tables
 * across pages, images, lists, keep-with-next headings, widow control, line
 * spacing rules, several faces and sizes, a tall header, landscape and roman
 * front matter, and tables of contents of up to two pages. `PG` bookmarks mark
 * every heading and every fourth paragraph.
 *
 * The `M` and `T` documents mix those: footnotes of several lines, nested
 * tables, rows that break across pages, sections of different paper, and a
 * default table style with paragraph and run properties (`T1`, which also
 * declares two default table styles, as `useDefaultStyles` followed by an
 * `addStyle({ isDefault: true })` used to write). `PG` marks a random third
 * of their paragraphs.
 *
 * `expected.json` holds what Microsoft Word for Mac 16.114 made of each one:
 * the page of every `PG` bookmark and the page count as the document was
 * written, and the page numbers its table of contents listed once Word had
 * updated it — which is what the library's own numbers must agree with.
 */

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

import { Io, Layout, Transform } from "../index";
import type { DocxDocument, Run } from "../types";

const DATA = resolve(__dirname, "data/word-pagination");

interface Expected {
  readonly pageCount: number;
  readonly pages: Readonly<Record<string, number>>;
  readonly toc?: readonly string[];
}

const { documents } = JSON.parse(readFileSync(resolve(DATA, "expected.json"), "utf8")) as {
  documents: Record<string, Expected>;
};

function markPages(doc: DocxDocument): Record<string, number> {
  const pages: Record<string, number> = {};
  for (const [name, page] of Layout.getPageMap(doc).bookmarkPages) {
    if (name.startsWith("PG")) {
      pages[name] = page;
    }
  }
  return pages;
}

/** The page number each entry of the document's table of contents shows. */
function tocNumbers(doc: DocxDocument): string[] | undefined {
  const toc = doc.body.find(item => item.type === "tableOfContents");
  if (!toc || toc.type !== "tableOfContents") {
    return undefined;
  }
  const numbers: string[] = [];
  for (const paragraph of toc.cachedParagraphs ?? []) {
    for (const child of paragraph.children) {
      const runs = "content" in child ? [child] : "children" in child ? child.children : [];
      for (const run of runs as readonly Run[]) {
        for (const content of run.content) {
          if (
            content.type === "field" &&
            /^\s*PAGEREF\b/.test(content.instruction) &&
            content.cachedValue
          ) {
            numbers.push(content.cachedValue);
          }
        }
      }
    }
  }
  return numbers;
}

describe("page numbers agree with Word", () => {
  it.each(Object.entries(documents))("%s", async (name, expected) => {
    const doc = await Io.read(readFileSync(resolve(DATA, `${name}.docx`)));
    expect(markPages(doc)).toEqual(expected.pages);
    expect(Layout.getPageMap(doc).pageCount).toBe(expected.pageCount);

    // Updated again, the fields come out as Word lists them, and the document
    // they are written into still paginates the same: the numbers are a fixed
    // point of their own layout.
    const updated = Transform.updateFields(doc);
    if (expected.toc) {
      expect(tocNumbers(updated)).toEqual(expected.toc);
    }
    expect(markPages(updated)).toEqual(expected.pages);
    expect(updated.settings?.updateFieldsOnOpen).toBeUndefined();
  });
});
