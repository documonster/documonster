/**
 * layoutDocumentFull — footnote area population.
 *
 * Verifies that body paragraphs which reference footnotes cause the
 * page's `footnoteArea` to be populated with the corresponding
 * footnote content, in document order, with rect.y placing the
 * stack's bottom edge at `pageHeight - pgMar.footer`.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import type { DocxDocument, FootnoteDef, Paragraph, Run } from "../types";

function para(...children: Paragraph["children"]): Paragraph {
  return { type: "paragraph", children };
}

function plainText(text: string): Run {
  return { content: [{ type: "text", text }] };
}

function refRun(id: number): Run {
  return { content: [{ type: "footnoteRef", id }] };
}

const baseDoc = (overrides: Partial<DocxDocument> = {}): DocxDocument => ({
  body: [],
  styles: [],
  abstractNumberings: [],
  numberingInstances: [],
  headers: new Map(),
  footers: new Map(),
  footnotes: [],
  endnotes: [],
  comments: [],
  images: [],
  fonts: [],
  embeddedFonts: [],
  customXmlParts: [],
  customProperties: [],
  opaqueParts: [],
  ...overrides
});

describe("layoutDocumentFull — footnotes", () => {
  it("populates footnoteArea for each referenced footnote", () => {
    const note1: FootnoteDef = {
      id: 1,
      content: [para(plainText("Footnote one body."))]
    };
    const note2: FootnoteDef = {
      id: 2,
      content: [para(plainText("Footnote two body."))]
    };
    const doc = baseDoc({
      footnotes: [note1, note2],
      body: [para(plainText("Intro"), refRun(1), plainText(" middle "), refRun(2))]
    });

    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    expect(page.footnoteArea).toBeDefined();
    expect(page.footnoteArea).toHaveLength(2);
  });

  it("ignores `separator` / `continuationSeparator` notes (they are chrome)", () => {
    const sep: FootnoteDef = {
      id: -1,
      type: "separator",
      content: [para(plainText("---"))]
    };
    const real: FootnoteDef = {
      id: 1,
      content: [para(plainText("real footnote"))]
    };
    const doc = baseDoc({
      footnotes: [sep, real],
      body: [para(refRun(1))]
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].footnoteArea).toHaveLength(1);
  });

  it("leaves footnoteArea undefined when the page references no footnotes", () => {
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("never referenced"))] }],
      body: [para(plainText("Body without any reference"))]
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].footnoteArea).toBeUndefined();
  });

  it("deduplicates references when the same footnote is cited twice", () => {
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("X"))] }],
      body: [para(refRun(1), plainText(" again "), refRun(1))]
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].footnoteArea).toHaveLength(1);
  });

  it("places the footnote stack so its bottom is at pageHeight - pgMar.footer", () => {
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("body"))] }],
      body: [para(refRun(1))],
      sectionProperties: {
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, footer: 1440 }
      }
    });
    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    expect(page.footnoteArea).toBeDefined();
    // The single laid paragraph should sit just above
    // `pageHeight - 72pt` (1 inch). Its rect.y + rect.height ≈ that.
    const last = page.footnoteArea![page.footnoteArea!.length - 1];
    const bottom = last.rect.y + last.rect.height;
    expect(bottom).toBeCloseTo(page.geometry.height - 72, 0);
  });

  it("collects refs from inside table cells", () => {
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("note"))] }],
      body: [
        {
          type: "table",
          rows: [{ cells: [{ content: [para(refRun(1))] }] }]
        }
      ]
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].footnoteArea).toHaveLength(1);
  });

  it("does not let the footnote stack overlap body content (defers overflowing notes)", () => {
    // Body that fills most of the page, then a paragraph at the
    // bottom referencing many tall footnotes. Layout should:
    //   - keep some footnotes on the page (whatever fits)
    //   - carry the rest forward to the next page
    const tallNote = (id: number): FootnoteDef => ({
      id,
      content: Array.from({ length: 10 }, (_, k) => para(plainText(`Note ${id} line ${k}`)))
    });
    const fillerLines = Array.from({ length: 50 }, (_, i) => para(plainText(`Body line ${i}`)));
    const doc = baseDoc({
      body: [
        ...fillerLines,
        // Single paragraph that cites every footnote — they all
        // belong to the same page in document terms.
        para(refRun(1), refRun(2), refRun(3), refRun(4), refRun(5))
      ],
      footnotes: [tallNote(1), tallNote(2), tallNote(3), tallNote(4), tallNote(5)]
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages.length).toBeGreaterThanOrEqual(2);
    // The footnote area on the referencing page must end above (or
    // exactly at) the bottom of the body content. Find the page
    // hosting the citing paragraph.
    const citingPage = layout.pages.find(p => (p.footnoteArea?.length ?? 0) > 0);
    expect(citingPage).toBeDefined();
    if (citingPage?.footnoteArea && citingPage.footnoteArea.length > 0) {
      // The first laid footnote paragraph's y is the stack's top.
      const stackTop = citingPage.footnoteArea[0].rect.y;
      // Body bottom = max(rect.y + rect.height) over all body items.
      let bodyBottom = 0;
      for (const item of citingPage.content) {
        const bottom = item.rect.y + item.rect.height;
        if (bottom > bodyBottom) {
          bodyBottom = bottom;
        }
      }
      const bodyBottomPageY = citingPage.geometry.marginTop + bodyBottom;
      expect(stackTop).toBeGreaterThanOrEqual(bodyBottomPageY - 0.5);
    }
    // Total footnote paragraphs across all pages must equal the sum
    // of all referenced footnotes' paragraphs (no data lost in the
    // deferral).
    const expectedTotalParagraphs = 5 * 10;
    let actualTotal = 0;
    for (const p of layout.pages) {
      actualTotal += p.footnoteArea?.length ?? 0;
    }
    expect(actualTotal).toBe(expectedTotalParagraphs);
  });

  it("first oversized footnote is forced (single-note page) rather than dropped", () => {
    // A single oversized footnote on a nearly-empty page must still
    // render — losing its content silently would be worse than
    // visually overflowing into the bottom margin.
    const giant: FootnoteDef = {
      id: 1,
      content: Array.from({ length: 200 }, (_, k) => para(plainText(`Giant ${k}`)))
    };
    const doc = baseDoc({
      footnotes: [giant],
      body: [para(refRun(1))]
    });
    const layout = layoutDocumentFull(doc);
    let total = 0;
    for (const p of layout.pages) {
      total += p.footnoteArea?.length ?? 0;
    }
    expect(total).toBeGreaterThan(0);
  });

  it("emits a `separator` rule on a page that introduces fresh footnotes", () => {
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("note"))] }],
      body: [para(refRun(1))]
    });
    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    expect(page.footnoteSeparator).toBeDefined();
    expect(page.footnoteSeparator?.kind).toBe("separator");
    // Separator y must sit slightly above the first footnote.
    expect(page.footnoteSeparator!.y).toBeLessThan(page.footnoteArea![0].rect.y);
  });

  it("emits a `continuationSeparator` on a page that only carries deferred footnotes", () => {
    // Three notes whose stack overflows the available footnote band
    // on the page that introduces them; the deferred ones land on
    // the synthetic overflow page added by `layoutDocumentFull`.
    const tall = (id: number): FootnoteDef => ({
      id,
      content: Array.from({ length: 50 }, (_, k) => para(plainText(`Note ${id} line ${k}`)))
    });
    const doc = baseDoc({
      footnotes: [tall(1), tall(2), tall(3)],
      // Single body paragraph at the top — the page has plenty of
      // room above, but the footnote stack alone is enormous so
      // notes 2 / 3 must spill to the next page.
      body: [para(refRun(1), refRun(2), refRun(3))]
    });
    const layout = layoutDocumentFull(doc);
    const continuationPage = layout.pages.find(
      p => p.footnoteSeparator?.kind === "continuationSeparator"
    );
    expect(continuationPage).toBeDefined();
  });

  it("renders the separator rule into the SVG output", async () => {
    const { renderPageFromLayout } = await import("../layout/render-page");
    const doc = baseDoc({
      footnotes: [{ id: 1, content: [para(plainText("note"))] }],
      body: [para(refRun(1))]
    });
    const layout = layoutDocumentFull(doc);
    const svg = renderPageFromLayout(layout, 1);
    // Single <line> element with stroke="black" stroke-width="0.5".
    expect(svg).toMatch(/<line[^/]*stroke="black"[^/]*stroke-width="0\.5"\/>/);
  });
});
