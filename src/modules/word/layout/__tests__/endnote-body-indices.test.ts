import { Build, Document, Transform } from "@word/index";
import { layoutDocumentFull, paginateDocument } from "@word/layout/layout-full";
import type { DocxDocument } from "@word/types";
import { describe, expect, it } from "vitest";

/**
 * Endnotes are laid out as body text — after the text (`docEnd`) or before
 * each section break (`sectEnd`, pages2 R3j) — but they are not body items.
 * Every index a layout hands back must still name an item of the caller's
 * `doc.body`: the notes inserted mid-body must not shift the items after them.
 */

const LETTER = {
  pageSize: { width: 12240, height: 15840 },
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 }
};

function build(position: "sectEnd" | "docEnd"): DocxDocument {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  const note = Array.from({ length: 30 }, (_, i) =>
    Build.paragraph([Build.text(`Endnote line ${i} long enough to take a whole line of its own.`)])
  );
  for (let s = 0; s < 3; s++) {
    const id = Document.allocateBookmarkId(doc);
    Document.addParagraph(
      doc,
      Build.paragraph([
        Build.bookmarkStart(id, `S${s}`),
        Build.bookmarkEnd(id),
        Build.text(`Section ${s} text.`),
        {
          properties: { vertAlign: "superscript" },
          content: [{ type: "endnoteRef", id: Document.addEndnote(doc, note) }]
        } as never
      ])
    );
    const props = { ...LETTER, endnoteProperties: { position } } as never;
    if (s < 2) {
      Document.addSectionBreak(doc, { ...(props as object), breakType: "continuous" } as never);
    } else {
      Document.setSectionProperties(doc, props);
    }
  }
  Document.setSettings(doc, { endnoteProperties: { position } } as never);
  return doc;
}

/** The page each body item's first block is drawn on, by `sourceIndex`. */
function drawnPages(doc: DocxDocument): Map<number, number> {
  const pages = new Map<number, number>();
  for (const page of layoutDocumentFull(doc).pages) {
    for (const block of page.content) {
      if (block.sourceIndex >= 0 && !pages.has(block.sourceIndex)) {
        pages.set(block.sourceIndex, page.pageNumber);
      }
    }
  }
  return pages;
}

describe("endnotes laid out among the body", () => {
  for (const position of ["sectEnd", "docEnd"] as const) {
    it(`keeps every index on the caller's body (${position})`, () => {
      const doc = build(position);
      const drawn = drawnPages(doc);
      // Only body items are reported, each under its own index.
      for (const index of drawn.keys()) {
        expect(index).toBeLessThan(doc.body.length);
      }
      const map = paginateDocument(doc);
      expect(map.contentPages).toHaveLength(doc.body.length);
      expect(map.contentSections).toHaveLength(doc.body.length);
      for (const [index, page] of drawn) {
        expect(map.contentPages[index]).toBe(page);
      }
      // Each section's paragraph follows the previous section's notes.
      const textPages = doc.body
        .map((item, index) => (item.type === "paragraph" ? map.contentPages[index]! : 0))
        .filter(page => page > 0);
      if (position === "sectEnd") {
        expect(textPages[1]).toBeGreaterThan(textPages[0]!);
        expect(textPages[2]).toBeGreaterThan(textPages[1]!);
      }
      expect(map.bookmarkPages.get("S2")).toBe(textPages[2]);
    });
  }

  it("numbers a PAGEREF after a section's notes by the page it is on", () => {
    const doc = build("sectEnd");
    Document.addParagraph(doc, Build.paragraph([Build.field(" PAGEREF S2 ", "0")]));
    const map = paginateDocument(doc);
    const updated = Transform.updateFields(doc);
    const text = JSON.stringify(updated.body.at(-1));
    expect(text).toContain(`"${map.bookmarkPages.get("S2")}"`);
  });
});
