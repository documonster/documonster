/**
 * Word → PDF draws the document's final view, the same text the converters
 * produce: tracked insertions and moves shown, deletions and hidden text not.
 */

import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import type { DocxDocument, RevisionInfo, Run } from "@word/types";
import { describe, expect, it } from "vitest";

import { Pdf } from "../index";
import { docxToPdf } from "../word-bridge";

const rev = { id: 1, author: "t" } as RevisionInfo;
const run = (text: string, vanish?: boolean): Run =>
  ({ content: [{ type: "text", text }], ...(vanish ? { properties: { vanish } } : {}) }) as Run;

describe("docxToPdf — final view", () => {
  it("draws the text the Markdown converter renders", async () => {
    const doc = {
      body: [
        {
          type: "paragraph",
          children: [
            run("Alpha "),
            { type: "insertedRun", revision: rev, run: run("Bravo ") },
            { type: "deletedRun", revision: rev, run: run("Charlie ") },
            { type: "movedToRun", revision: rev, run: run("Delta ") },
            { type: "movedFromRun", revision: rev, run: run("Echo ") },
            run("Hidden ", true),
            { content: [{ type: "field", instruction: " =1 ", cachedValue: "Foxtrot" }] }
          ]
        }
      ],
      styles: []
    } as unknown as DocxDocument;
    const read = await Pdf.read(await docxToPdf(doc));
    const text = read.pages
      .map(p => p.text)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    expect(text).toBe("Alpha Bravo Delta Foxtrot");
    expect(text).toBe(renderToMarkdown(doc).replace(/\s+/g, " ").trim());
  });
});
