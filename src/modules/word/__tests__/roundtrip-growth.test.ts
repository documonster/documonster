/**
 * A document written by this library and read back must not grow when it is
 * written again: every read → write round trip of the same package has to be
 * a fixed point. Each case below used to add something on every pass.
 */

import { describe, expect, it } from "vitest";

import { extractParagraphText } from "../core/text-utils";
import { Build, Document, Io } from "../index";
import type { DocxDocument, Watermark } from "../index";
import {
  differingParts,
  packageIssues,
  packageParts,
  partText
} from "./roundtrip-integrity.helpers";

/** The smallest valid PNG: one transparent pixel. */
const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

const SVG = new TextEncoder().encode(
  `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"><rect width="8" height="8"/></svg>`
);

/** Write `doc`, then read and write the result twice more; return all three packages. */
async function roundTrips(doc: DocxDocument): Promise<Map<string, Uint8Array>[]> {
  const first = await Io.toBuffer(doc);
  const second = await Io.toBuffer(await Io.read(first));
  const third = await Io.toBuffer(await Io.read(second));
  return Promise.all([first, second, third].map(packageParts));
}

function count(xml: string, pattern: RegExp): number {
  return xml.match(pattern)?.length ?? 0;
}

function imageWatermarkDocument(): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "Body");
  const img = Document.addImage(d, PNG, "png", 100, 100);
  Document.removeContent(d, Document.getContentCount(d) - 1);
  const wm: Watermark = { type: "image", rId: img.rId, washout: true, widthPt: 300, heightPt: 120 };
  Document.setWatermark(d, wm);
  return d;
}

describe("watermark round trip", () => {
  it("an image watermark is written once, however often the document is round-tripped", async () => {
    const [written, once, twice] = await roundTrips(imageWatermarkDocument());
    for (const parts of [written, once, twice]) {
      expect(packageIssues(parts)).toEqual([]);
      const doc = partText(parts, "word/document.xml");
      expect(count(doc, /<w:headerReference\b/g)).toBe(1);
      expect([...parts.keys()].filter(p => /^word\/header\d+\.xml$/.test(p))).toHaveLength(1);
    }
    // The first write gives the builder's `__img_<n>` key a relationship id;
    // the header's id is a fresh key once read, so only later passes are equal.
    expect([...once.keys()].sort()).toEqual([...written.keys()].sort());
    expect(differingParts(once, twice)).toEqual([]);
  });

  it("reads an image watermark's size back", async () => {
    const read = await Io.read(await Io.toBuffer(imageWatermarkDocument()));
    expect(read.watermark).toMatchObject({ type: "image", widthPt: 300, heightPt: 120 });
  });

  it("a text watermark is written once, however often the document is round-tripped", async () => {
    const d = Document.create();
    Document.addParagraph(d, "Body");
    Document.setWatermark(d, { type: "text", text: "DRAFT", color: "C0C0C0", rotation: -45 });
    const [written, once, twice] = await roundTrips(d);
    for (const parts of [written, once, twice]) {
      expect(count(partText(parts, "word/document.xml"), /<w:headerReference\b/g)).toBe(1);
    }
    expect(differingParts(written, once)).toEqual([]);
    expect(differingParts(once, twice)).toEqual([]);
    expect((await Io.read(await Io.toBuffer(d))).watermark).toMatchObject({
      type: "text",
      text: "DRAFT"
    });
  });

  it("joins a watermark to the section's own default header instead of adding a second one", async () => {
    const d = Document.create();
    Document.addParagraph(d, "Body");
    Document.setHeader(d, "default", { children: [Build.textParagraph("Header text")] });
    Document.setWatermark(d, { type: "text", text: "DRAFT" });
    const [written, once, twice] = await roundTrips(d);
    for (const parts of [written, once, twice]) {
      expect(packageIssues(parts)).toEqual([]);
      expect(count(partText(parts, "word/document.xml"), /<w:headerReference\b/g)).toBe(1);
      const header = partText(parts, "word/header1.xml");
      expect(header).toContain("Header text");
      expect(count(header, /PowerPlusWaterMarkObject/g)).toBe(1);
      expect(parts.has("word/header2.xml")).toBe(false);
    }
    expect(differingParts(once, twice)).toEqual([]);

    const read = await Io.read(await Io.toBuffer(d));
    expect(read.watermark).toMatchObject({ type: "text", text: "DRAFT" });
    expect(read.headers?.size).toBe(1);
  });

  it("an image watermark in the default header registers its picture in that header's rels", async () => {
    const d = imageWatermarkDocument();
    Document.setHeader(d, "default", { children: [Build.textParagraph("Header text")] });
    const [written, once, twice] = await roundTrips(d);
    for (const parts of [written, once, twice]) {
      expect(packageIssues(parts)).toEqual([]);
      const header = partText(parts, "word/header1.xml");
      const rId = header.match(/<v:imagedata r:id="([^"]+)"/)?.[1];
      expect(rId).toBeDefined();
      expect(partText(parts, "word/_rels/header1.xml.rels")).toMatch(
        new RegExp(`Id="${rId}"[^>]*Target="media/`)
      );
    }
    expect(differingParts(once, twice)).toEqual([]);
  });
});

describe("table of contents round trip", () => {
  it("keeps the cached entries as they were written", async () => {
    const d = Document.create();
    Document.addTableOfContents(d, {
      headingStyleRange: "1-3",
      hyperlink: true,
      cachedParagraphs: [Build.textParagraph("Chapter 1\t1"), Build.textParagraph("Chapter 2\t2")]
    });
    Document.addHeading(d, "Chapter 1", 1);
    const [written, once, twice] = await roundTrips(d);
    for (const parts of [written, once, twice]) {
      const doc = partText(parts, "word/document.xml");
      expect(count(doc, /<w:p\/>/g)).toBe(0);
      expect(count(doc, /fldCharType="begin"/g)).toBe(1);
      expect(doc).toContain(">Chapter 1\t1<");
      expect(doc).toContain(">Chapter 2\t2<");
    }
    expect(differingParts(written, once)).toEqual([]);
    expect(differingParts(once, twice)).toEqual([]);

    const read = await Io.read(await Io.toBuffer(d));
    const toc = read.body[0];
    expect(toc.type).toBe("tableOfContents");
    // Read back, the entries hold the field where it was: opened in the first
    // entry and closed in a paragraph of its own after the last, as Word
    // writes a table of contents and as the writer put it.
    expect(toc.type === "tableOfContents" && toc.field).toBe("entries");
    expect(
      toc.type === "tableOfContents" && toc.cachedParagraphs?.map(p => extractParagraphText(p))
    ).toEqual(["Chapter 1\t1", "Chapter 2\t2", ""]);
  });
});

describe("SVG image round trip", () => {
  it("writes an SVG and its raster fallback once, however often it is round-tripped", async () => {
    const d = Document.create();
    Document.addImage(d, SVG, "svg", 100, 100, { fallbackData: PNG });
    const [written, once, twice] = await roundTrips(d);
    for (const parts of [written, once, twice]) {
      expect(packageIssues(parts)).toEqual([]);
      expect([...parts.keys()].filter(p => p.startsWith("word/media/")).sort()).toEqual([
        "word/media/image1.svg",
        "word/media/image1_fallback.png"
      ]);
      const doc = partText(parts, "word/document.xml");
      const rels = partText(parts, "word/_rels/document.xml.rels");
      const blip = doc.match(/<a:blip r:embed="([^"]+)"/)![1];
      const svg = doc.match(/svgBlip[^>]*r:embed="([^"]+)"/)![1];
      expect(rels).toMatch(new RegExp(`Id="${blip}"[^>]*Target="media/image1_fallback\\.png"`));
      expect(rels).toMatch(new RegExp(`Id="${svg}"[^>]*Target="media/image1\\.svg"`));
    }
    expect(differingParts(written, once)).toEqual([]);
    expect(differingParts(once, twice)).toEqual([]);
  });
});
