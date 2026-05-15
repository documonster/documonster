/**
 * layoutDocumentFull — full BodyContent coverage tests.
 *
 * Every `BodyContent` variant must produce a corresponding `PageContent`
 * with a positive-size rect and the correct `sourceIndex`. The
 * `never`-typed exhaustiveness guard in `layout-full.ts` is the
 * structural enforcement; these tests are the runtime safety net.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import { renderPageFromLayout } from "../layout/render-page";
import type {
  AltChunk,
  ChartContent,
  CheckBox,
  DocxDocument,
  DrawingShape,
  FloatingImage,
  MathBlock,
  OpaqueDrawing,
  StructuredDocumentTag,
  TableOfContents,
  TextBox
} from "../types";

const minimalDoc = (body: DocxDocument["body"]): DocxDocument => ({
  body,
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
  opaqueParts: []
});

describe("layoutDocumentFull — full BodyContent coverage", () => {
  it("lays out a FloatingImage as a `float` PageContent with positive-size rect", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400, // 1 inch in EMU
      height: 914_400
    };
    const layout = layoutDocumentFull(minimalDoc([fi]));
    expect(layout.pages).toHaveLength(1);
    const items = layout.pages[0].content;
    expect(items).toHaveLength(1);
    expect(items[0].type).toBe("float");
    expect(items[0].rect.width).toBeCloseTo(72); // 1 inch = 72pt
    expect(items[0].rect.height).toBeCloseTo(72);
  });

  it("respects horizontalPosition center / verticalPosition top alignment", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400, // 1 inch in EMU
      height: 914_400,
      horizontalPosition: { relativeTo: "margin", align: "center" },
      verticalPosition: { relativeTo: "margin", align: "top" }
    };
    const layout = layoutDocumentFull(minimalDoc([fi]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("float");
    if (item.type === "float") {
      const g = layout.pages[0].geometry;
      // Center horizontally inside the content area.
      expect(item.rect.x).toBeCloseTo((g.contentWidth - 72) / 2);
      // Top of the content area.
      expect(item.rect.y).toBeCloseTo(0);
    }
  });

  it("interprets simplePos as page-absolute coordinates", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      // (1in, 2in) from the page corner.
      simplePos: { x: 914_400, y: 1_828_800 }
    };
    const layout = layoutDocumentFull(minimalDoc([fi]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("float");
    if (item.type === "float") {
      const g = layout.pages[0].geometry;
      // Translate from page-absolute to content-area-relative.
      expect(item.rect.x).toBeCloseTo(72 - g.marginLeft);
      expect(item.rect.y).toBeCloseTo(144 - g.marginTop);
    }
  });

  it("anchor-positioned floats do not advance the in-flow cursor", () => {
    const anchored: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      horizontalPosition: { relativeTo: "page", offset: 100_000 },
      verticalPosition: { relativeTo: "page", offset: 100_000 }
    };
    // Place a paragraph immediately after the float; it should land at
    // y=0 in the content area, not pushed down by the float's height.
    const para = {
      type: "paragraph" as const,
      children: [{ content: [{ type: "text" as const, text: "after float" }] }]
    };
    const layout = layoutDocumentFull(minimalDoc([anchored, para]));
    const items = layout.pages[0].content;
    expect(items).toHaveLength(2);
    const paraItem = items[1];
    expect(paraItem.type).toBe("paragraph");
    expect(paraItem.rect.y).toBeCloseTo(0);
  });

  it("topAndBottom wrap pushes the body cursor below the float", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400, // 1 inch tall
      horizontalPosition: { relativeTo: "margin", align: "center" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "topAndBottom" }
    };
    const para = {
      type: "paragraph" as const,
      children: [{ content: [{ type: "text" as const, text: "below float" }] }]
    };
    const layout = layoutDocumentFull(minimalDoc([fi, para]));
    const items = layout.pages[0].content;
    expect(items).toHaveLength(2);
    const paraItem = items[1];
    expect(paraItem.type).toBe("paragraph");
    // Float top=0, height=72pt (1in), no margin: cursor must be ≥72.
    expect(paraItem.rect.y).toBeGreaterThanOrEqual(72);
  });

  it("topAndBottom wrap respects wrap.margins.bottom padding", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      horizontalPosition: { relativeTo: "margin", align: "left" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: {
        style: "topAndBottom",
        // 0.5 inch = 457_200 EMU below the float must be cleared.
        margins: { bottom: 457_200 }
      }
    };
    const para = {
      type: "paragraph" as const,
      children: [{ content: [{ type: "text" as const, text: "x" }] }]
    };
    const layout = layoutDocumentFull(minimalDoc([fi, para]));
    const paraItem = layout.pages[0].content[1];
    expect(paraItem.type).toBe("paragraph");
    // Float bottom = 72pt; +0.5in pad = 36pt → paragraph at y ≥ 108.
    expect(paraItem.rect.y).toBeGreaterThanOrEqual(108);
  });

  it("preserves wrap info on LayoutFloat for renderers (square style)", () => {
    const fi: FloatingImage = {
      type: "floatingImage",
      rId: "rId1",
      width: 914_400,
      height: 914_400,
      horizontalPosition: { relativeTo: "margin", align: "center" },
      verticalPosition: { relativeTo: "margin", align: "top" },
      wrap: { style: "square", side: "bothSides" }
    };
    const layout = layoutDocumentFull(minimalDoc([fi]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("float");
    if (item.type === "float") {
      expect(item.wrap?.style).toBe("square");
      expect(item.wrap?.side).toBe("bothSides");
    }
  });

  it("lays out a TextBox with inner content recursively", () => {
    const tb: TextBox = {
      type: "textBox",
      content: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "inside" }] }] }],
      width: 2880, // 2 inches in twips
      height: 1440
    };
    const layout = layoutDocumentFull(minimalDoc([tb]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("textBox");
    if (item.type === "textBox") {
      expect(item.content.length).toBeGreaterThan(0);
      expect(item.content[0].type).toBe("paragraph");
    }
  });

  it("lays out a DrawingShape with text content", () => {
    const sh: DrawingShape = {
      type: "drawingShape",
      shapeType: "rect",
      width: 914_400,
      height: 457_200,
      fillColor: "FFCC00",
      textContent: [
        { type: "paragraph", children: [{ content: [{ type: "text", text: "label" }] }] }
      ]
    };
    const layout = layoutDocumentFull(minimalDoc([sh]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("shape");
    if (item.type === "shape") {
      expect(item.preset).toBe("rect");
      expect(item.fillColor).toBe("FFCC00");
      expect(item.textContent?.length).toBeGreaterThan(0);
    }
  });

  it("lays out a Chart as a `chart` placeholder rect", () => {
    const ch: ChartContent = {
      type: "chart",
      chart: { type: "bar", title: "Q1", series: [] }
    };
    const layout = layoutDocumentFull(minimalDoc([ch]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("chart");
    if (item.type === "chart") {
      expect(item.chartKind).toBe("chart");
      expect(item.title).toBe("Q1");
      expect(item.rect.width).toBeGreaterThan(0);
    }
  });

  it("uses Chart.width/Chart.height when present (extent recovered by reader)", () => {
    // 2 inch × 1 inch in EMU written into the inner Chart model — that
    // is where the writer emits `<wp:extent>` from and where the
    // reader stores the recovered extent.
    const ch: ChartContent = {
      type: "chart",
      chart: {
        type: "bar",
        series: [],
        width: 1_828_800,
        height: 914_400
      }
    };
    const layout = layoutDocumentFull(minimalDoc([ch]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("chart");
    if (item.type === "chart") {
      expect(item.rect.width).toBeCloseTo(144); // 2 in
      expect(item.rect.height).toBeCloseTo(72); // 1 in
    }
  });

  it("falls back to a default 6×3.5 inch chart when extent is missing", () => {
    const ch: ChartContent = {
      type: "chart",
      chart: { type: "bar", series: [] }
    };
    const layout = layoutDocumentFull(minimalDoc([ch]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("chart");
    if (item.type === "chart") {
      // Default width is capped at content width on a Letter page; just
      // check the height (which is not capped) matches the documented
      // default of 3.5 inches → 252 pt (with rounding tolerance for
      // the 3_200_400 EMU constant used in the layout engine).
      expect(item.rect.height).toBeGreaterThan(220);
      expect(item.rect.height).toBeLessThan(260);
    }
  });

  it("lays out an SDT transparently", () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: { tag: "name" },
      content: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "Acme" }] }] }]
    };
    const layout = layoutDocumentFull(minimalDoc([sdt]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("sdt");
    if (item.type === "sdt") {
      expect(item.tag).toBe("name");
      expect(item.content.length).toBe(1);
    }
  });

  it("lays out a MathBlock with text fallback", () => {
    const mb: MathBlock = {
      type: "math",
      content: [
        {
          type: "mathFraction",
          numerator: [{ type: "mathRun", text: "1" }],
          denominator: [{ type: "mathRun", text: "2" }]
        }
      ]
    };
    const layout = layoutDocumentFull(minimalDoc([mb]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("math");
    if (item.type === "math") {
      expect(item.text.length).toBeGreaterThan(0);
    }
  });

  it("lays out a CheckBox with checked/unchecked glyph", () => {
    const cb: CheckBox = { type: "checkBox", checked: true };
    const layout = layoutDocumentFull(minimalDoc([cb]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("checkBox");
    if (item.type === "checkBox") {
      expect(item.checked).toBe(true);
      expect(item.glyph.length).toBeGreaterThan(0);
    }
  });

  it("lays out a TableOfContents from cachedParagraphs", () => {
    const toc: TableOfContents = {
      type: "tableOfContents",
      cachedParagraphs: [
        { type: "paragraph", children: [{ content: [{ type: "text", text: "1. Intro" }] }] },
        { type: "paragraph", children: [{ content: [{ type: "text", text: "2. Body" }] }] }
      ]
    };
    const layout = layoutDocumentFull(minimalDoc([toc]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("tableOfContents");
    if (item.type === "tableOfContents") {
      expect(item.entries).toHaveLength(2);
    }
  });

  it("lays out a TableOfContents stub when cached entries are absent", () => {
    const toc: TableOfContents = { type: "tableOfContents" };
    const layout = layoutDocumentFull(minimalDoc([toc]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("tableOfContents");
    if (item.type === "tableOfContents") {
      expect(item.entries).toHaveLength(1);
    }
  });

  it("lays out an AltChunk as a placeholder", () => {
    const ac: AltChunk = { type: "altChunk", rId: "x", contentType: "text/html" };
    const layout = layoutDocumentFull(minimalDoc([ac]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("altChunk");
    if (item.type === "altChunk") {
      expect(item.contentType).toBe("text/html");
    }
  });

  it("lays out an OpaqueDrawing as a placeholder", () => {
    const od: OpaqueDrawing = {
      type: "opaqueDrawing",
      rawXml: "<w:drawing/>",
      referencedRIds: []
    };
    const layout = layoutDocumentFull(minimalDoc([od]));
    const item = layout.pages[0].content[0];
    expect(item.type).toBe("opaqueDrawing");
    if (item.type === "opaqueDrawing") {
      expect(item.rawXml).toContain("<w:drawing");
    }
  });

  it("renderPageFromLayout handles every PageContent variant without throwing", () => {
    const body: DocxDocument["body"] = [
      // every variant
      { type: "paragraph", children: [{ content: [{ type: "text", text: "para" }] }] },
      { type: "table", rows: [{ cells: [{ content: [] }] }] },
      { type: "floatingImage", rId: "rId1", width: 914_400, height: 914_400 } as FloatingImage,
      {
        type: "textBox",
        content: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "tb" }] }] }],
        width: 1440,
        height: 720
      } as TextBox,
      {
        type: "drawingShape",
        shapeType: "ellipse",
        width: 457_200,
        height: 228_600
      } as DrawingShape,
      { type: "chart", chart: { type: "bar", series: [] } } as ChartContent,
      {
        type: "sdt",
        content: [{ type: "paragraph", children: [{ content: [{ type: "text", text: "sdt" }] }] }]
      } as StructuredDocumentTag,
      { type: "math", content: [{ type: "mathRun", text: "x" }] } as MathBlock,
      { type: "checkBox", checked: false } as CheckBox,
      { type: "tableOfContents" } as TableOfContents,
      { type: "altChunk", rId: "rId2" } as AltChunk,
      { type: "opaqueDrawing", rawXml: "<x/>", referencedRIds: [] } as OpaqueDrawing
    ];
    const layout = layoutDocumentFull(minimalDoc(body));
    const svg = renderPageFromLayout(layout, 1);
    expect(svg).toContain("<svg");
    expect(svg).toContain("</svg>");
  });
});
