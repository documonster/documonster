/**
 * layoutDocumentFull — header / footer page resolution.
 *
 * Verifies the layout engine populates `LayoutPage.header` and
 * `LayoutPage.footer` from the document's section properties +
 * header / footer parts. Without this, every renderer downstream of
 * layout (SVG, PDF) silently loses the document's running titles,
 * page numbers, etc.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import type { DocxDocument, FooterDef, HeaderDef, Paragraph } from "../types";

function para(text: string): Paragraph {
  return { type: "paragraph", children: [{ content: [{ type: "text", text }] }] };
}

const baseDoc = (overrides: Partial<DocxDocument> = {}): DocxDocument => ({
  body: [para("body text")],
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

describe("layoutDocumentFull — header / footer", () => {
  it("populates LayoutPage.header from default header reference", () => {
    const headerDef: HeaderDef = {
      content: { children: [para("Confidential")] }
    };
    const doc = baseDoc({
      headers: new Map([["rIdH1", headerDef]]),
      sectionProperties: {
        headers: [{ type: "default", rId: "rIdH1" }]
      }
    });

    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    expect(page.header).toBeDefined();
    expect(page.header).toHaveLength(1);
    expect(page.header![0].type).toBe("paragraph");
  });

  it("populates LayoutPage.footer from default footer reference", () => {
    const footerDef: FooterDef = {
      content: { children: [para("Page 1")] }
    };
    const doc = baseDoc({
      footers: new Map([["rIdF1", footerDef]]),
      sectionProperties: {
        footers: [{ type: "default", rId: "rIdF1" }]
      }
    });

    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    expect(page.footer).toBeDefined();
    expect(page.footer).toHaveLength(1);
  });

  it("falls back to the first reference when no default is declared", () => {
    const headerDef: HeaderDef = {
      content: { children: [para("first-page only")] }
    };
    const doc = baseDoc({
      headers: new Map([["rIdH1", headerDef]]),
      sectionProperties: {
        // Only `first` declared — no `default`. Layout falls back to
        // the first reference rather than emitting an empty header.
        headers: [{ type: "first", rId: "rIdH1" }]
      }
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].header).toBeDefined();
  });

  it("omits the header field entirely when section has no header refs", () => {
    const doc = baseDoc();
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].header).toBeUndefined();
    expect(layout.pages[0].footer).toBeUndefined();
  });

  it("produces no header when the referenced part is missing", () => {
    const doc = baseDoc({
      // Reference rIdH1 but no entry in `headers` for it.
      sectionProperties: {
        headers: [{ type: "default", rId: "rIdH1" }]
      }
    });
    const layout = layoutDocumentFull(doc);
    expect(layout.pages[0].header).toBeUndefined();
  });

  it("renders header text into the SVG output", async () => {
    const { renderPageFromLayout } = await import("../layout/render-page");
    const headerDef: HeaderDef = {
      content: { children: [para("HEADER LINE")] }
    };
    const doc = baseDoc({
      headers: new Map([["rIdH1", headerDef]]),
      sectionProperties: {
        headers: [{ type: "default", rId: "rIdH1" }]
      }
    });
    const layout = layoutDocumentFull(doc);
    const svg = renderPageFromLayout(layout, 1);
    expect(svg).toContain("HEADER LINE");
  });

  it("offsets header content by pgMar.header (Header from top)", () => {
    // 1440 twips = 1 inch = 72pt — make it distinctively bigger than
    // the default of 720 (0.5 inch) so the test can prove the value
    // came from the section, not from the default.
    const headerDef: HeaderDef = {
      content: { children: [para("anchored")] }
    };
    const doc = baseDoc({
      headers: new Map([["rIdH1", headerDef]]),
      sectionProperties: {
        headers: [{ type: "default", rId: "rIdH1" }],
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 1440 }
      }
    });
    const layout = layoutDocumentFull(doc);
    const headerItem = layout.pages[0].header![0];
    expect(headerItem.type).toBe("paragraph");
    if (headerItem.type === "paragraph") {
      // First paragraph's rect.y should reflect the 72pt offset.
      expect(headerItem.rect.y).toBeCloseTo(72, 1);
    }
  });

  it("offsets footer content so its top sits at pageHeight - pgMar.footer", () => {
    const footerDef: FooterDef = {
      content: { children: [para("bottom")] }
    };
    const doc = baseDoc({
      footers: new Map([["rIdF1", footerDef]]),
      sectionProperties: {
        footers: [{ type: "default", rId: "rIdF1" }],
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, footer: 1440 }
      }
    });
    const layout = layoutDocumentFull(doc);
    const page = layout.pages[0];
    const footerItem = page.footer![0];
    expect(footerItem.type).toBe("paragraph");
    if (footerItem.type === "paragraph") {
      // pageHeight - 72pt (1 inch).
      expect(footerItem.rect.y).toBeCloseTo(page.geometry.height - 72, 1);
    }
  });

  it("uses the `first` header on page 1 when sectionProperties.titlePage is set", () => {
    const firstHeader: HeaderDef = {
      content: { children: [para("title-page header")] }
    };
    const defaultHeader: HeaderDef = {
      content: { children: [para("regular header")] }
    };
    const doc = baseDoc({
      headers: new Map([
        ["rIdHF", firstHeader],
        ["rIdHD", defaultHeader]
      ]),
      sectionProperties: {
        titlePage: true,
        headers: [
          { type: "first", rId: "rIdHF" },
          { type: "default", rId: "rIdHD" }
        ]
      }
    });
    const layout = layoutDocumentFull(doc);
    const headerItem = layout.pages[0].header![0];
    expect(headerItem.type).toBe("paragraph");
    if (headerItem.type === "paragraph") {
      // The text run should come from the first-page header.
      const firstItem = headerItem.lines[0]?.runs[0];
      const text = firstItem && firstItem.type !== "image" ? firstItem.text : undefined;
      expect(text).toContain("title-page");
    }
  });

  it("uses the `even` header when settings.evenAndOddHeaders is set", () => {
    const evenHeader: HeaderDef = {
      content: { children: [para("EVEN HEADER")] }
    };
    const defaultHeader: HeaderDef = {
      content: { children: [para("ODD HEADER")] }
    };
    // Force at least 2 pages so the even-page case can be asserted.
    const longBody = Array.from({ length: 60 }, (_, i) => para(`Body line ${i}`));
    const doc = baseDoc({
      body: longBody,
      headers: new Map([
        ["rIdHE", evenHeader],
        ["rIdHD", defaultHeader]
      ]),
      sectionProperties: {
        headers: [
          { type: "even", rId: "rIdHE" },
          { type: "default", rId: "rIdHD" }
        ]
      },
      settings: { evenAndOddHeaders: true }
    });

    const layout = layoutDocumentFull(doc);
    expect(layout.totalPages).toBeGreaterThanOrEqual(2);
    const evenPage = layout.pages.find(p => p.pageNumber === 2);
    expect(evenPage?.header).toBeDefined();
    if (evenPage?.header) {
      const headerItem = evenPage.header[0];
      if (headerItem.type === "paragraph") {
        const firstItem = headerItem.lines[0]?.runs[0];
        const text = firstItem && firstItem.type !== "image" ? firstItem.text : undefined;
        expect(text).toContain("EVEN");
      }
    }
  });

  it("lays out tables inside header content (not just paragraphs)", () => {
    const tableHeader: HeaderDef = {
      content: {
        children: [
          {
            type: "table",
            rows: [
              {
                cells: [{ content: [para("logo")] }, { content: [para("title")] }]
              }
            ]
          }
        ]
      }
    };
    const doc = baseDoc({
      headers: new Map([["rIdH1", tableHeader]]),
      sectionProperties: {
        headers: [{ type: "default", rId: "rIdH1" }]
      }
    });
    const layout = layoutDocumentFull(doc);
    const headerItems = layout.pages[0].header;
    expect(headerItems).toHaveLength(1);
    expect(headerItems![0].type).toBe("table");
  });
});
