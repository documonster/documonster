/**
 * docxToSemantic — exhaustive BodyContent coverage tests.
 *
 * Every BodyContent variant must produce a corresponding SemanticBlock
 * (no silent drops). The compile-time `never` guard in
 * `convert/docx-to-semantic.ts` is the structural enforcement; these
 * tests are the runtime safety net.
 */

import { describe, it, expect } from "vitest";

import { docxToSemantic } from "../convert/docx-to-semantic";
import type {
  AltChunk,
  ChartContent,
  CheckBox,
  DocxDocument,
  MathBlock,
  OpaqueDrawing,
  TableOfContents
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

describe("docxToSemantic — full BodyContent coverage", () => {
  it("converts MathBlock to a SemanticBlock with text + mathML", () => {
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
    const { document, context } = docxToSemantic(minimalDoc([mb]));
    const block = document.blocks.find(b => b.type === "math");
    expect(block).toBeDefined();
    expect(block!.type).toBe("math");
    if (block?.type === "math") {
      expect(block.text.length).toBeGreaterThan(0);
      // mathML may be undefined if conversion fails; if present must look like XML
      if (block.mathML) {
        expect(block.mathML).toContain("<");
      }
    }
    // No "unsupported-block" warning was emitted
    expect(context.warnings.find(w => w.code === "unsupported-block")).toBeUndefined();
  });

  it("converts ChartContent to a chart SemanticBlock", () => {
    const chart: ChartContent = {
      type: "chart",
      chart: {
        type: "bar",
        title: "Sales 2024",
        series: []
      },
      altText: "Bar chart of sales"
    };
    const { document } = docxToSemantic(minimalDoc([chart]));
    const block = document.blocks.find(b => b.type === "chart");
    expect(block).toBeDefined();
    if (block?.type === "chart") {
      expect(block.title).toBe("Sales 2024");
      expect(block.altText).toBe("Bar chart of sales");
    }
  });

  it("assigns distinct chartIds to two charts with the same title", () => {
    // Without position-based ids two charts with the same title would
    // collide on chartId, breaking renderers that look the chart up
    // by id (e.g. a markdown exporter writing `[chart-1] Sales`).
    const chart = (title: string): ChartContent => ({
      type: "chart",
      chart: { type: "bar", title, series: [] }
    });
    const { document } = docxToSemantic(minimalDoc([chart("Sales"), chart("Sales")]));
    const charts = document.blocks.filter(b => b.type === "chart");
    expect(charts).toHaveLength(2);
    if (charts[0].type === "chart" && charts[1].type === "chart") {
      expect(charts[0].chartId).not.toBe(charts[1].chartId);
    }
  });

  it("prefers source `name` when assigning chartId", () => {
    const c: ChartContent = {
      type: "chart",
      chart: { type: "bar", title: "Renamed in title", series: [] },
      name: "Chart 7"
    };
    const { document } = docxToSemantic(minimalDoc([c]));
    const block = document.blocks.find(b => b.type === "chart");
    expect(block).toBeDefined();
    if (block?.type === "chart") {
      expect(block.chartId).toBe("Chart 7");
    }
  });

  it("converts CheckBox to a checkBox SemanticBlock with checked state", () => {
    const cbTrue: CheckBox = { type: "checkBox", checked: true };
    const cbFalse: CheckBox = { type: "checkBox", checked: false };
    const cbUndef: CheckBox = { type: "checkBox" };
    const { document } = docxToSemantic(minimalDoc([cbTrue, cbFalse, cbUndef]));
    const blocks = document.blocks.filter(b => b.type === "checkBox");
    expect(blocks).toHaveLength(3);
    expect(blocks[0].type === "checkBox" && blocks[0].checked).toBe(true);
    expect(blocks[1].type === "checkBox" && blocks[1].checked).toBe(false);
    expect(blocks[2].type === "checkBox" && blocks[2].checked).toBe(false);
  });

  it("converts AltChunk to an embed SemanticBlock with contentType", () => {
    const ac: AltChunk = {
      type: "altChunk",
      rId: "rId99",
      contentType: "text/html",
      data: new TextEncoder().encode("<p>hi</p>"),
      fileName: "afchunk.html"
    };
    const { document } = docxToSemantic(minimalDoc([ac]));
    const block = document.blocks.find(b => b.type === "embed");
    expect(block).toBeDefined();
    if (block?.type === "embed") {
      expect(block.contentType).toBe("text/html");
      expect(block.fileName).toBe("afchunk.html");
      expect(block.data).toBeDefined();
    }
  });

  it("converts AltChunk without contentType to a default octet-stream embed", () => {
    const ac: AltChunk = { type: "altChunk", rId: "rIdX" };
    const { document, context } = docxToSemantic(minimalDoc([ac]));
    const block = document.blocks.find(b => b.type === "embed");
    expect(block).toBeDefined();
    if (block?.type === "embed") {
      expect(block.contentType).toBe("application/octet-stream");
    }
    // The fallback must be visible to the caller — silently consuming
    // an unknown content type would mask a malformed source document.
    expect(context.warnings.some(w => w.code === "altchunk-missing-content-type")).toBe(true);
  });

  it("converts OpaqueDrawing to a raw SemanticBlock preserving the source XML", () => {
    const od: OpaqueDrawing = {
      type: "opaqueDrawing",
      rawXml: "<w:drawing><wp:inline/></w:drawing>",
      referencedRIds: []
    };
    const { document } = docxToSemantic(minimalDoc([od]));
    const block = document.blocks.find(b => b.type === "raw");
    expect(block).toBeDefined();
    if (block?.type === "raw") {
      expect(block.format).toBe("ooxml-drawing");
      expect(block.xml).toContain("<w:drawing>");
    }
  });

  it("emits a TOC placeholder when cachedParagraphs is empty (no silent drop)", () => {
    const toc: TableOfContents = {
      type: "tableOfContents",
      cachedParagraphs: []
    };
    const { document, context } = docxToSemantic(minimalDoc([toc]));
    expect(document.blocks).toHaveLength(1);
    const block = document.blocks[0];
    expect(block.type).toBe("paragraph");
    expect(context.warnings.some(w => w.code === "toc-not-cached")).toBe(true);
  });

  it("emits no `unsupported-block` warning for any of the previously-dropped variants", () => {
    const items: DocxDocument["body"] = [
      { type: "math", content: [{ type: "mathRun", text: "x" }] } as MathBlock,
      { type: "checkBox", checked: true } as CheckBox,
      { type: "altChunk", rId: "x" } as AltChunk,
      {
        type: "opaqueDrawing",
        rawXml: "<x/>",
        referencedRIds: []
      } as OpaqueDrawing
    ];
    const { context } = docxToSemantic(minimalDoc(items));
    expect(context.warnings.filter(w => w.code === "unsupported-block")).toEqual([]);
  });
});
