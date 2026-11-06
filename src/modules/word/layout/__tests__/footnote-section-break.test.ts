import { Build, Document } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import { describe, expect, it } from "vitest";
/**
 * A footnote still continuing when a section ends on a page break.
 *
 * Word finishes the note in the section it belongs to, on a page holding
 * nothing but its continuation, before the next section opens (final7 N6d,
 * Word for Mac 16.114: page 5 of the landscape section holds only the rest of
 * note 4, and the letter-size section starts on page 6).
 */

const LAND = {
  pageSize: { width: 15840, height: 12240, orientation: "landscape" as const },
  margins: { top: 1080, right: 1080, bottom: 1080, left: 1080, header: 540, footer: 540 }
};
const LETTER = {
  pageSize: { width: 12240, height: 15840 },
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 }
};

function build() {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  const line = "A footnote long enough to run past the bottom of a landscape page. ".repeat(6);
  const id = Document.addFootnote(
    doc,
    Array.from({ length: 12 }, () => Build.paragraph([Build.text(line)]))
  );
  Document.addParagraph(
    doc,
    Build.paragraph([
      Build.text("Body text referencing a long note."),
      { properties: { vertAlign: "superscript" }, content: [{ type: "footnoteRef", id }] }
    ])
  );
  Document.addSectionBreak(doc, { ...LAND, breakType: "nextPage" });
  Document.addParagraph(doc, Build.textParagraph("The next section."));
  Document.setSectionProperties(doc, { ...LETTER, breakType: "nextPage" });
  return layoutDocumentFull(doc);
}

describe("a footnote continuing at a section break", () => {
  it("finishes on a page of its own before the next section starts", () => {
    const layout = build();
    const [first, notes, next] = layout.pages;
    expect(layout.pages).toHaveLength(3);
    expect(first!.footnoteArea?.length).toBeGreaterThan(0);
    // The continuation page belongs to the landscape section and holds no body
    // text: at most the empty paragraph carrying the section break.
    expect(notes!.geometry.width).toBe(792);
    expect(JSON.stringify(notes!.content)).not.toContain("next section");
    expect(notes!.footnoteArea?.length).toBeGreaterThan(0);
    // The letter-size section opens on the page after it, with no notes left.
    expect(next!.geometry.width).toBe(612);
    expect(JSON.stringify(next!.content)).toContain("The next section.");
    expect(next!.footnoteArea ?? []).toHaveLength(0);
  });
});
