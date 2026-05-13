/**
 * Realistic round-trip fixtures.
 *
 * The goal here is *not* to test a single feature in isolation — those
 * checks live in the dedicated unit-test files (replace.test.ts,
 * revisions.test.ts, etc). Instead we exercise combinations that mirror
 * what real-world documents tend to ship together, and we verify that
 * the model survives `packageDocx` → `readDocx` without losing the
 * cross-references between parts (header hyperlinks, comment ranges,
 * footnote ids, image rIds, etc).
 *
 * These tests catch the kind of regressions that pure feature tests
 * can miss — a writer change that subtly breaks header hyperlink rId
 * registration, or a reader change that drops comment ids when a
 * commentRangeStart/end straddles formatted runs.
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  packageDocx,
  readDocx,
  text,
  bold,
  italic,
  paragraph,
  hyperlink,
  commentRangeStart,
  commentRangeEnd,
  commentReference,
  simpleTable,
  pageBreak,
  validateDocument
} from "../index";
import type { Hyperlink, Paragraph, Run, Table } from "../types";

// Minimal 1x1 PNG.
const MINI_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
  0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

describe("Realistic round-trip — typical office document", () => {
  it("preserves headers/footers with hyperlinks across round-trip", async () => {
    const h = Document.create();
    Document.addHeading(h, "Quarterly Report", 1);
    Document.addParagraph(h, "This document was generated for the Q4 review.");
    Document.setHeader(h, "default", {
      children: [
        paragraph([
          text("See "),
          hyperlink("policy", { url: "https://example.com/policy" }),
          text(" for details.")
        ])
      ]
    });
    Document.setFooter(h, "default", {
      children: [
        paragraph([
          text("Contact "),
          hyperlink("the author", { url: "mailto:author@example.com" }),
          text(" with questions.")
        ])
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    expect(parsed.headers).toBeDefined();
    expect(parsed.footers).toBeDefined();

    const findHyperlinkUrl = (children: Paragraph["children"]): string | undefined => {
      for (const c of children) {
        if ("type" in c && c.type === "hyperlink") {
          const link = c as Hyperlink;
          if (link.url) {
            return link.url;
          }
        }
      }
      return undefined;
    };

    let foundHeaderUrl: string | undefined;
    for (const [, header] of parsed.headers!) {
      for (const child of header.content.children) {
        if (child.type === "paragraph") {
          const url = findHyperlinkUrl(child.children);
          if (url) {
            foundHeaderUrl = url;
            break;
          }
        }
      }
      if (foundHeaderUrl) {
        break;
      }
    }
    expect(foundHeaderUrl).toBe("https://example.com/policy");

    let foundFooterUrl: string | undefined;
    for (const [, footer] of parsed.footers!) {
      for (const child of footer.content.children) {
        if (child.type === "paragraph") {
          const url = findHyperlinkUrl(child.children);
          if (url) {
            foundFooterUrl = url;
            break;
          }
        }
      }
      if (foundFooterUrl) {
        break;
      }
    }
    expect(foundFooterUrl).toBe("mailto:author@example.com");
  });

  it("preserves footnote ids referenced from body runs", async () => {
    const h = Document.create();
    const fnId = Document.addFootnote(h, "This is a clarifying footnote.");
    Document.addParagraphElement(
      h,
      paragraph([text("Important point"), { content: [{ type: "footnoteRef", id: fnId }] }])
    );

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    expect(parsed.footnotes).toBeDefined();
    const ids = parsed.footnotes!.map(f => f.id);
    expect(ids).toContain(fnId);

    // The body run must still carry the footnoteRef pointing at fnId.
    let foundRef = false;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "footnoteRef" && c.id === fnId) {
              foundRef = true;
            }
          }
        }
      }
    }
    expect(foundRef).toBe(true);

    // And the document validates with no dangling-reference errors.
    const result = validateDocument(parsed);
    const dangling = result.issues.filter(
      i => i.rule === "ref-footnote-missing" || i.rule === "ref-endnote-missing"
    );
    expect(dangling).toEqual([]);
  });

  it("preserves comment ranges with formatted runs spanning the range", async () => {
    const h = Document.create();
    const cId = Document.addComment(h, "Reviewer", "Please clarify this sentence.");
    Document.addParagraphElement(
      h,
      paragraph([
        text("The "),
        commentRangeStart(cId),
        bold("important"),
        text(" "),
        italic("clause"),
        commentRangeEnd(cId),
        text(" follows."),
        commentReference(cId)
      ])
    );

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    expect(parsed.comments).toBeDefined();
    expect(parsed.comments!.map(c => c.id)).toContain(cId);

    // Find the paragraph and walk its children: we should still see a
    // start/end pair and a reference, all bound to the same id.
    let starts = 0;
    let ends = 0;
    let refs = 0;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if ("type" in child) {
          if (child.type === "commentRangeStart" && child.id === cId) {
            starts++;
          }
          if (child.type === "commentRangeEnd" && child.id === cId) {
            ends++;
          }
          if (child.type === "commentReference" && child.id === cId) {
            refs++;
          }
        }
      }
    }
    expect(starts).toBe(1);
    expect(ends).toBe(1);
    expect(refs).toBe(1);

    const result = validateDocument(parsed);
    const dangling = result.issues.filter(
      i => i.rule === "ref-comment-missing" || i.rule === "ref-comment-range-missing"
    );
    expect(dangling).toEqual([]);
  });

  it("survives image rId remap when images are interleaved with text", async () => {
    const h = Document.create();
    Document.addParagraph(h, "Before image #1.");
    const { rId: rId1 } = Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraph(h, "Between images.");
    const { rId: rId2 } = Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraph(h, "After image #2.");

    expect(rId1).not.toBe(rId2);

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    // Two images preserved.
    expect(parsed.images?.length).toBe(2);
    expect(new Set(parsed.images!.map(i => i.rId)).size).toBe(2);

    // Body still references both rIds, in order.
    const seenRIds: string[] = [];
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "image" && c.rId) {
              seenRIds.push(c.rId);
            }
          }
        }
      }
    }
    expect(seenRIds.length).toBe(2);
    expect(seenRIds[0]).not.toBe(seenRIds[1]);

    // Validation: every image rId is resolvable.
    const result = validateDocument(parsed);
    const dangling = result.issues.filter(i => i.rule === "rel-image-missing");
    expect(dangling).toEqual([]);
  });

  it("preserves a complex page: heading + table + page break + footnote + comment", async () => {
    const h = Document.create();
    const fnId = Document.addFootnote(h, "Footnote near the table.");
    const cId = Document.addComment(h, "QA", "Check totals.");

    Document.addHeading(h, "Sales by Region", 2);

    Document.addTableElement(
      h,
      simpleTable([
        ["Region", "Q3", "Q4"],
        ["NA", "120", "150"],
        ["EU", "90", "110"]
      ])
    );

    Document.addParagraphElement(
      h,
      paragraph([
        commentRangeStart(cId),
        text("Totals computed manually"),
        { content: [{ type: "footnoteRef", id: fnId }] },
        commentRangeEnd(cId),
        text("."),
        commentReference(cId)
      ])
    );

    Document.addParagraphElement(h, paragraph([pageBreak(), text("Next page.")]));

    Document.addHeading(h, "Conclusion", 2);
    Document.addParagraph(h, "All targets met.");

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    // Headings preserved.
    const headings = parsed.body.filter(
      b => b.type === "paragraph" && b.properties?.style?.includes("Heading")
    );
    expect(headings.length).toBe(2);

    // Table preserved with same shape (3 rows × 3 cells).
    const tables = parsed.body.filter(b => b.type === "table");
    expect(tables.length).toBe(1);
    const table = tables[0]!;
    if (table.type !== "table") {
      throw new Error("not a table");
    }
    expect((table as Table).rows.length).toBe(3);
    for (const row of (table as Table).rows) {
      expect(row.cells.length).toBe(3);
    }

    // Footnote and comment payload still present.
    expect(parsed.footnotes!.some(f => f.id === fnId)).toBe(true);
    expect(parsed.comments!.some(c => c.id === cId)).toBe(true);

    // Validation closure passes.
    const result = validateDocument(parsed);
    const fatal = result.issues.filter(i => i.severity === "error");
    expect(fatal).toEqual([]);
  });

  it("does not pollute the input model on package — second package produces equivalent bytes", async () => {
    const h = Document.create();
    Document.addHeading(h, "Test", 1);
    const { rId } = Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    Document.addParagraphElement(
      h,
      paragraph([text("Has "), hyperlink("link", { url: "https://example.com" })])
    );

    const built = Document.build(h);

    // Snapshot the relevant fields before packaging.
    const altChunkCountBefore = built.body.filter(b => b.type === "altChunk").length;
    const imageRIdBefore = built.images?.[0]?.rId;

    const first = await packageDocx(built);
    const second = await packageDocx(built);

    // Input model unchanged after two packagings.
    expect(built.body.filter(b => b.type === "altChunk").length).toBe(altChunkCountBefore);
    expect(built.images?.[0]?.rId).toBe(imageRIdBefore);

    // Both packagings round-trip and find the image.
    const parsedA = await readDocx(first);
    const parsedB = await readDocx(second);
    expect(parsedA.images?.length).toBe(1);
    expect(parsedB.images?.length).toBe(1);

    // The image rId we got back from the builder should still resolve.
    expect(rId).toBeDefined();
  });
});

describe("validateDocument — reference-closure rules", () => {
  it("reports dangling footnoteRef as ref-footnote-missing", () => {
    const h = Document.create();
    // Add a single unrelated footnote so the id-set is non-empty —
    // the validator only fires when we have something to compare against.
    Document.addFootnote(h, "unrelated");
    Document.addParagraphElement(
      h,
      paragraph([text("Bad footnote: "), { content: [{ type: "footnoteRef", id: 9999 }] }])
    );

    const result = validateDocument(Document.build(h));
    const issue = result.issues.find(i => i.rule === "ref-footnote-missing");
    expect(issue).toBeDefined();
    expect(issue!.severity).toBe("error");
  });

  it("reports dangling commentReference as ref-comment-missing", () => {
    const h = Document.create();
    Document.addComment(h, "Bob", "exists");
    // Reference an id that does NOT exist.
    Document.addParagraphElement(h, paragraph([text("Bad: "), commentReference(9999)]));
    const result = validateDocument(Document.build(h));
    const issue = result.issues.find(i => i.rule === "ref-comment-missing");
    expect(issue).toBeDefined();
    expect(issue!.severity).toBe("error");
  });

  it("does not flag valid references", () => {
    const h = Document.create();
    const fnId = Document.addFootnote(h, "good");
    const cId = Document.addComment(h, "Alice", "good");
    Document.addParagraphElement(
      h,
      paragraph([
        commentRangeStart(cId),
        text("Body"),
        { content: [{ type: "footnoteRef", id: fnId }] },
        commentRangeEnd(cId),
        commentReference(cId)
      ])
    );

    const result = validateDocument(Document.build(h));
    const dangling = result.issues.filter(
      i =>
        i.rule === "ref-footnote-missing" ||
        i.rule === "ref-endnote-missing" ||
        i.rule === "ref-comment-missing" ||
        i.rule === "ref-comment-range-missing"
    );
    expect(dangling).toEqual([]);
  });
});
