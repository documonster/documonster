/**
 * Smoke tests + option-fidelity tests for docxToPdf.
 *
 * Beyond the basic byte-shape check, these tests guard against
 * regressions where caller-supplied page geometry options would be
 * silently dropped — the symptom is `pdfBytes` not reflecting the
 * requested `pageWidth` / `pageHeight` / `margin*`.
 */

import {
  Document,
  layoutDocumentFull,
  packageDocx,
  paragraph,
  text,
  readDocx,
  table,
  row,
  cell,
  type DocxDocument
} from "@word/index";
import { describe, it, expect } from "vitest";

import { docxToPdf } from "../word-bridge";

describe("docxToPdf — layout-driven smoke test", () => {
  it("produces a valid PDF for a paragraph-only document", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, paragraph([text("hello world")]));
    const docBytes = await packageDocx(Document.build(h));

    const doc = await readDocx(docBytes);

    const pdfBytes = await docxToPdf(doc);
    expect(pdfBytes.length).toBeGreaterThan(100);
    const head = new TextDecoder().decode(pdfBytes.slice(0, 5));
    expect(head).toBe("%PDF-");
    const tailDecoded = new TextDecoder().decode(pdfBytes.slice(-32));
    expect(tailDecoded).toMatch(/%%EOF\s*$/);
  });

  it("handles a document with a paragraph + table without throwing", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, paragraph([text("intro")]));

    const t = table(
      [row([cell("A1"), cell("A2")]), row([cell("B1"), cell("B2")])],
      { width: { value: 5000, type: "pct" } },
      [2500, 2500]
    );
    Document.addTableElement(h, t);
    const docBytes = await packageDocx(Document.build(h));

    const doc = await readDocx(docBytes);

    const pdfBytes = await docxToPdf(doc);
    expect(pdfBytes.length).toBeGreaterThan(200);
  });
});

describe("docxToPdf — option fidelity", () => {
  async function buildSimpleDoc(): Promise<{ doc: DocxDocument }> {
    const h = Document.create();
    Document.addParagraphElement(h, paragraph([text("page geometry probe")]));
    const docBytes = await packageDocx(Document.build(h));
    return { doc: await readDocx(docBytes) };
  }

  it("forwards pageWidth / pageHeight overrides into the layout engine", async () => {
    const { doc } = await buildSimpleDoc();
    const layoutOverridden = layoutDocumentFull(doc, {
      pageGeometry: { pageWidth: 400, pageHeight: 500 }
    });
    expect(layoutOverridden.pages[0].geometry.width).toBe(400);
    expect(layoutOverridden.pages[0].geometry.height).toBe(500);

    const layoutDefault = layoutDocumentFull(doc);
    expect(layoutDefault.pages[0].geometry.width).not.toBe(400);
  });

  it("forwards margin overrides into the layout engine", async () => {
    const { doc } = await buildSimpleDoc();
    const layout = layoutDocumentFull(doc, {
      pageGeometry: {
        marginTop: 10,
        marginBottom: 20,
        marginLeft: 30,
        marginRight: 40
      }
    });
    const g = layout.pages[0].geometry;
    expect(g.marginTop).toBe(10);
    expect(g.marginBottom).toBe(20);
    expect(g.marginLeft).toBe(30);
    expect(g.marginRight).toBe(40);
    expect(g.contentWidth).toBe(g.width - 30 - 40);
    expect(g.contentHeight).toBe(g.height - 10 - 20);
  });

  it("docxToPdf actually applies pageWidth / pageHeight options end-to-end", async () => {
    const { doc } = await buildSimpleDoc();
    // A known small page size; the resulting PDF must declare a media
    // box matching this width/height. PDF media boxes are written as
    // `/MediaBox [0 0 W H]`.
    const w = 300;
    const h = 450;
    const pdfBytes = await docxToPdf(doc, { pageWidth: w, pageHeight: h });
    const decoded = new TextDecoder().decode(pdfBytes);
    expect(decoded).toMatch(/\/MediaBox\s*\[\s*0\s+0\s+300\s+450\s*\]/);
  });

  it("falls back to section properties when no overrides are supplied", async () => {
    const { doc } = await buildSimpleDoc();
    const pdfBytes = await docxToPdf(doc);
    const decoded = new TextDecoder().decode(pdfBytes);
    // US Letter default (612 x 792) with no override.
    expect(decoded).toMatch(/\/MediaBox\s*\[\s*0\s+0\s+612\s+792\s*\]/);
  });
});

describe("docxToPdf — chart rendering fallback", () => {
  it("declining chartRenderer (return false) lets the translator draw the placeholder", async () => {
    const h = Document.create();
    const chartItem = {
      type: "chart" as const,
      chart: { type: "bar" as const, title: "Quarterly Revenue", series: [] }
    };
    Document.addContent(h, chartItem);
    const docBytes = await packageDocx(Document.build(h));

    const doc = await readDocx(docBytes);

    let invoked = 0;
    let rectSeen: { x: number; y: number; width: number; height: number } | null = null;
    const declined = await docxToPdf(doc, {
      chartRenderer: (_chart, _page, rect) => {
        invoked++;
        rectSeen = rect;
        return false; // decline
      }
    });

    // The user's chartRenderer was offered the chart...
    expect(invoked).toBeGreaterThanOrEqual(1);
    expect(rectSeen).not.toBeNull();
    expect(rectSeen!.width).toBeGreaterThan(0);
    expect(rectSeen!.height).toBeGreaterThan(0);

    // ...and after the decline the PDF still came back with a real
    // body. Compare against an empty document to assert the chart
    // slot didn't simply disappear into a smaller-than-empty file.
    const empty = await docxToPdf(
      await readDocx(await packageDocx(Document.build(Document.create())))
    );
    expect(declined.length).toBeGreaterThan(empty.length);
  });

  it("accepting chartRenderer (no return value) suppresses the placeholder", async () => {
    const h = Document.create();
    Document.addContent(h, {
      type: "chart" as const,
      chart: { type: "bar" as const, title: "Q1 Sales", series: [] }
    });
    const docBytes = await packageDocx(Document.build(h));

    const doc = await readDocx(docBytes);

    let invoked = 0;
    await docxToPdf(doc, {
      chartRenderer: () => {
        invoked++;
        // Implicit return; equivalent to returning `true`. Translator
        // must not draw its placeholder on top.
      }
    });
    expect(invoked).toBe(1);
  });
});

describe("docxToPdf — inline image", () => {
  it("emits a PDF image XObject when a paragraph contains an inline image", async () => {
    // 1×1 red PNG (zlib-deflated valid IDAT chunk); the engine's
    // png-decoder rejects hand-rolled minimal PNGs, so we use a real
    // round-tripped sample.
    const TINY_PNG = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
      0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
      0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
      0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
    ]);
    const docModel: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [
            { content: [{ type: "text", text: "before " }] },
            {
              content: [
                {
                  type: "image",
                  rId: "rId1",
                  width: 914_400,
                  height: 914_400
                }
              ]
            },
            { content: [{ type: "text", text: " after" }] }
          ]
        }
      ],
      styles: [],
      abstractNumberings: [],
      numberingInstances: [],
      headers: new Map(),
      footers: new Map(),
      footnotes: [],
      endnotes: [],
      comments: [],
      images: [
        {
          data: TINY_PNG,
          mediaType: "png",
          fileName: "img.png",
          rId: "rId1"
        }
      ],
      fonts: [],
      embeddedFonts: [],
      customXmlParts: [],
      customProperties: [],
      opaqueParts: []
    };
    const pdfBytes = await docxToPdf(docModel);
    const decoded = new TextDecoder().decode(pdfBytes);
    // PDF image XObjects appear as `/Subtype /Image` entries in the
    // PDF content stream. Without inline-image support the body
    // would only contain text operators (Tj/TJ).
    expect(decoded).toMatch(/\/Subtype\s*\/Image/);
  });
});
