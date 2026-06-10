/**
 * DOCX Module - Field engine smoke tests.
 *
 * Covers `updateFields` and `updateTableOfContents` end-to-end:
 *   - empty document does not crash
 *   - PAGE / NUMPAGES cached values get populated after layout
 *   - TOC entries are generated from heading paragraphs
 *   - input document is never mutated
 *
 * The intent is a smoke test, not full coverage of every supported field
 * type. The 1400+ line field engine has many branches (REF / SEQ / STYLEREF /
 * INCLUDETEXT / formulas …); those should grow their own tests as bugs
 * surface.
 */

import { describe, it, expect } from "vitest";

import { updateFields, updateTableOfContents } from "../advanced/field-engine";
import { heading, paragraph, textParagraph } from "../builder/paragraph-builders";
import {
  indexEntryField,
  indexField,
  pageBreak,
  pageNumberField,
  text,
  tocField,
  totalPagesField
} from "../builder/run-builders";
import type { DocxDocument, FieldContent, Paragraph, Run } from "../types";

function getFirstField(doc: DocxDocument): FieldContent | undefined {
  for (const block of doc.body) {
    if (block.type !== "paragraph") {
      continue;
    }
    for (const child of block.children) {
      if (!("content" in child) || !Array.isArray((child as Run).content)) {
        continue;
      }
      for (const c of (child as Run).content) {
        if (c.type === "field") {
          return c;
        }
      }
    }
  }
  return undefined;
}

describe("updateFields", () => {
  it("does not crash on an empty document", () => {
    const doc: DocxDocument = { body: [] };
    expect(() => updateFields(doc)).not.toThrow();
  });

  it("does not mutate the input document", () => {
    const doc: DocxDocument = {
      body: [paragraph([pageNumberField("?"), text(" of "), totalPagesField("?")])]
    };
    const originalBody = doc.body;
    updateFields(doc);
    expect(doc.body).toBe(originalBody);
  });

  it("populates PAGE field cachedValue with a numeric string", () => {
    const doc: DocxDocument = {
      body: [paragraph([pageNumberField("?")])]
    };
    const updated = updateFields(doc);
    const f = getFirstField(updated);
    expect(f).toBeDefined();
    expect(f!.cachedValue).toBeDefined();
    // PAGE on a single-paragraph doc must be page 1.
    expect(f!.cachedValue).toBe("1");
  });

  it("populates NUMPAGES field cachedValue with a numeric string", () => {
    const doc: DocxDocument = {
      body: [paragraph([totalPagesField("?")])]
    };
    const updated = updateFields(doc);
    const f = getFirstField(updated);
    expect(f).toBeDefined();
    expect(f!.cachedValue).toBeDefined();
    // Single-paragraph doc fits on one page.
    expect(f!.cachedValue).toBe("1");
  });

  it("returns the same reference when there are no fields to update", () => {
    const doc: DocxDocument = {
      body: [textParagraph("plain text only")]
    };
    const updated = updateFields(doc);
    expect(updated).toBe(doc);
  });
});

describe("updateTableOfContents", () => {
  it("does not crash on an empty document", () => {
    const doc: DocxDocument = { body: [] };
    expect(() => updateTableOfContents(doc)).not.toThrow();
  });

  it("does not mutate the input document", () => {
    const doc: DocxDocument = {
      body: [
        paragraph([tocField({ headingLevels: "1-3" })]),
        heading("Chapter 1", 1),
        heading("Section 1.1", 2)
      ]
    };
    const originalBody = doc.body;
    updateTableOfContents(doc);
    expect(doc.body).toBe(originalBody);
  });

  it("returns the same document when there is no TOC to update", () => {
    const doc: DocxDocument = {
      body: [heading("Just a heading", 1), textParagraph("body")]
    };
    const updated = updateTableOfContents(doc);
    expect(updated).toBe(doc);
  });

  it("populates a tableOfContents block with cached entries from headings", () => {
    const doc: DocxDocument = {
      body: [
        { type: "tableOfContents", headingStyleRange: "1-3", hyperlink: true },
        heading("Chapter 1", 1),
        heading("Section 1.1", 2),
        heading("Chapter 2", 1)
      ]
    };
    const updated = updateTableOfContents(doc);
    // Find the TOC block in the output and check that it carries some
    // cachedParagraphs (the field engine fills these from the heading list).
    const toc = updated.body.find(b => b.type === "tableOfContents") as
      | (DocxDocument["body"][number] & { cachedParagraphs?: Paragraph[] })
      | undefined;
    expect(toc).toBeDefined();
    expect(toc!.cachedParagraphs).toBeDefined();
    // Expect at least one entry per heading we passed in.
    expect(toc!.cachedParagraphs!.length).toBeGreaterThanOrEqual(3);
  });
});

describe("updateFields — INDEX field", () => {
  /** Find the cachedValue of the first INDEX field in the document. */
  function findIndexCachedValue(doc: DocxDocument): string | undefined {
    for (const block of doc.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("content" in child) || !Array.isArray((child as Run).content)) {
          continue;
        }
        for (const c of (child as Run).content) {
          if (c.type === "field" && c.instruction.includes("INDEX")) {
            return c.cachedValue;
          }
        }
      }
    }
    return undefined;
  }

  it("merges repeated terms into one line and lists real page numbers", () => {
    // Page 1: widget. Page 2 (after a page break): gadget + widget again.
    const doc: DocxDocument = {
      body: [
        paragraph([text("alpha "), indexEntryField("widget")]),
        paragraph([pageBreak()]),
        paragraph([text("beta "), indexEntryField("gadget"), indexEntryField("widget")]),
        paragraph([pageBreak()]),
        paragraph([indexField()])
      ]
    };
    const updated = updateFields(doc);
    const value = findIndexCachedValue(updated);
    expect(value).toBeDefined();

    // Two distinct terms → two lines, sorted alphabetically (gadget < widget).
    const lines = value!.split("\n");
    expect(lines).toHaveLength(2);
    expect(lines[0].startsWith("gadget\t")).toBe(true);
    expect(lines[1].startsWith("widget\t")).toBe(true);

    // widget appears on pages 1 and 2 → merged into a single line "1, 2"
    // (no duplicate "widget" row). gadget only on page 2.
    const widgetPages = lines[1].split("\t")[1];
    expect(widgetPages).toBe("1, 2");
    const gadgetPages = lines[0].split("\t")[1];
    expect(gadgetPages).toBe("2");
  });

  it("produces no duplicate term rows when the same term is marked twice on one page", () => {
    const doc: DocxDocument = {
      body: [
        paragraph([indexEntryField("widget"), text(" "), indexEntryField("widget")]),
        paragraph([indexField()])
      ]
    };
    const updated = updateFields(doc);
    const value = findIndexCachedValue(updated);
    expect(value).toBe("widget\t1");
  });
});
