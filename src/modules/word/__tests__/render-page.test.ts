/**
 * DOCX Module - Page Renderer (SVG Output) Tests
 */

import { describe, it, expect } from "vitest";

import { renderPageToSvg, renderDocumentToSvg } from "../layout/render-page";
import type { DocxDocument, Paragraph, Run, Table } from "../types";

// =============================================================================
// Test Helpers
// =============================================================================

/** Create a simple text run. */
function makeRun(text: string, props?: Run["properties"]): Run {
  return {
    ...(props ? { properties: props } : {}),
    content: [{ type: "text", text }]
  };
}

/** Create a simple paragraph with text. */
function makeParagraph(text: string, props?: Paragraph["properties"]): Paragraph {
  return {
    type: "paragraph",
    ...(props ? { properties: props } : {}),
    children: [makeRun(text)]
  };
}

/** Create a minimal DocxDocument with paragraphs. */
function makeDoc(body: DocxDocument["body"]): DocxDocument {
  return { body };
}

// =============================================================================
// Tests
// =============================================================================

describe("renderPageToSvg", () => {
  it("renders a simple single-paragraph document", () => {
    const doc = makeDoc([makeParagraph("Hello World")]);
    const svg = renderPageToSvg(doc, 1);

    expect(svg).toContain("<svg");
    expect(svg).toContain("<text");
    expect(svg).toContain("Hello World");
    expect(svg).toContain("</svg>");
  });

  it("renders heading text with larger font-size", () => {
    const normalDoc = makeDoc([makeParagraph("Normal text")]);
    const headingDoc = makeDoc([makeParagraph("Heading text", { style: "Heading1" })]);

    const normalSvg = renderPageToSvg(normalDoc, 1);
    const headingSvg = renderPageToSvg(headingDoc, 1);

    // Extract font-size values
    const normalFontSize = normalSvg.match(/font-size="([^"]+)"/);
    const headingFontSize = headingSvg.match(/font-size="([^"]+)"/);

    expect(normalFontSize).not.toBeNull();
    expect(headingFontSize).not.toBeNull();

    const normalSize = parseFloat(normalFontSize![1]);
    const headingSize = parseFloat(headingFontSize![1]);
    // Heading1 uses a scale of 2.0, so heading text should be larger
    expect(headingSize).toBeGreaterThan(normalSize);
  });

  it("renders multiple paragraphs with multiple <text> elements", () => {
    const doc = makeDoc([
      makeParagraph("First paragraph"),
      makeParagraph("Second paragraph"),
      makeParagraph("Third paragraph")
    ]);
    const svg = renderPageToSvg(doc, 1);

    expect(svg).toContain("First paragraph");
    expect(svg).toContain("Second paragraph");
    expect(svg).toContain("Third paragraph");

    // Count <text elements
    const textElements = svg.match(/<text /g);
    expect(textElements).not.toBeNull();
    expect(textElements!.length).toBeGreaterThanOrEqual(3);
  });

  it("renders bold text with font-weight='bold'", () => {
    const doc = makeDoc([
      {
        type: "paragraph",
        children: [makeRun("Bold text", { bold: true })]
      }
    ]);
    const svg = renderPageToSvg(doc, 1);

    expect(svg).toContain('font-weight="bold"');
    expect(svg).toContain("Bold text");
  });

  it("renders italic text with font-style='italic'", () => {
    const doc = makeDoc([
      {
        type: "paragraph",
        children: [makeRun("Italic text", { italic: true })]
      }
    ]);
    const svg = renderPageToSvg(doc, 1);

    expect(svg).toContain('font-style="italic"');
    expect(svg).toContain("Italic text");
  });

  it("renders colored text with fill attribute", () => {
    const doc = makeDoc([
      {
        type: "paragraph",
        children: [makeRun("Red text", { color: "FF0000" })]
      }
    ]);
    const svg = renderPageToSvg(doc, 1);

    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain("Red text");
  });

  it("renders a table with <rect> and <line> elements", () => {
    const table: Table = {
      type: "table",
      rows: [
        {
          cells: [{ content: [makeParagraph("Cell A1")] }, { content: [makeParagraph("Cell B1")] }]
        },
        {
          cells: [{ content: [makeParagraph("Cell A2")] }, { content: [makeParagraph("Cell B2")] }]
        }
      ]
    };
    const doc = makeDoc([table]);
    const svg = renderPageToSvg(doc, 1);

    // Table outer border is drawn as a <rect>
    expect(svg).toContain("<rect");
    // Row/column separators drawn as <line>
    expect(svg).toContain("<line");
    // Cell text should be present
    expect(svg).toContain("Cell A1");
    expect(svg).toContain("Cell B1");
  });

  it("throws RangeError for page number out of range", () => {
    const doc = makeDoc([makeParagraph("Single page")]);

    expect(() => renderPageToSvg(doc, 0)).toThrow(RangeError);
    expect(() => renderPageToSvg(doc, 999)).toThrow(RangeError);
  });

  it("applies custom background color", () => {
    const doc = makeDoc([makeParagraph("Hello")]);
    const svg = renderPageToSvg(doc, 1, { backgroundColor: "#EEEEEE" });

    expect(svg).toContain('fill="#EEEEEE"');
  });

  it("applies custom scale factor to output dimensions", () => {
    const doc = makeDoc([makeParagraph("Scaled")]);
    const svg1x = renderPageToSvg(doc, 1, { scale: 1.0 });
    const svg2x = renderPageToSvg(doc, 1, { scale: 2.0 });

    // Extract width from both SVGs
    const width1x = svg1x.match(/width="([^"]+)"/);
    const width2x = svg2x.match(/width="([^"]+)"/);

    expect(width1x).not.toBeNull();
    expect(width2x).not.toBeNull();

    const w1 = parseFloat(width1x![1]);
    const w2 = parseFloat(width2x![1]);
    // 2x scale should produce double the width
    expect(w2).toBeCloseTo(w1 * 2, 0);
  });
});

describe("renderDocumentToSvg", () => {
  it("returns an array of SVG strings", () => {
    const doc = makeDoc([makeParagraph("Page content")]);
    const pages = renderDocumentToSvg(doc);

    expect(Array.isArray(pages)).toBe(true);
    expect(pages.length).toBeGreaterThanOrEqual(1);
    expect(pages[0]).toContain("<svg");
    expect(pages[0]).toContain("Page content");
  });

  it("respects rendering options", () => {
    const doc = makeDoc([makeParagraph("Content")]);
    const pages = renderDocumentToSvg(doc, { backgroundColor: "lightblue" });

    expect(pages.length).toBeGreaterThanOrEqual(1);
    expect(pages[0]).toContain('fill="lightblue"');
  });

  it("renders an empty document body with at least one page", () => {
    const doc = makeDoc([]);
    const pages = renderDocumentToSvg(doc);

    expect(Array.isArray(pages)).toBe(true);
    // Even an empty document should produce at least one page
    expect(pages.length).toBeGreaterThanOrEqual(1);
    expect(pages[0]).toContain("<svg");
  });
});
