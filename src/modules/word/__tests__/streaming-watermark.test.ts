/**
 * The streaming writer places a watermark the way the buffered writer does:
 * inside the default header when there is one (an image watermark's picture
 * in that header's own rels), and in a header part of its own only when there
 * is none. Both writers produce the same headers. Every other header a
 * section shows is covered in `watermark-every-header.test.ts`.
 */

import { describe, expect, it } from "vitest";

import { Build, Document, Io, Stream } from "../index";
import type { DocxDocument, HeaderDef, ImageDef, Watermark } from "../index";
import { packageIssues, packageParts, partText } from "./roundtrip-integrity.helpers";

const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

const IMAGE: ImageDef = { rId: "rId100", data: PNG, mediaType: "png", fileName: "image1.png" };

interface Case {
  readonly name: string;
  readonly watermark: Watermark;
  readonly headers?: ReadonlyMap<string, HeaderDef>;
}

const header = (text: string): HeaderDef => ({
  content: { children: [Build.textParagraph(text)] }
});

const CASES: Case[] = [
  { name: "text watermark, no header", watermark: { type: "text", text: "DRAFT" } },
  {
    name: "text watermark in the default header",
    watermark: { type: "text", text: "DRAFT" },
    headers: new Map([["default", header("Header text")]])
  },
  {
    name: "image watermark in the default header",
    watermark: { type: "image", rId: IMAGE.rId!, widthPt: 300, heightPt: 120 },
    headers: new Map([["default", header("Header text")]])
  },
  {
    name: "watermark with only a first-page header",
    watermark: { type: "text", text: "DRAFT" },
    headers: new Map([["first", header("First page")]])
  }
];

function buffered(c: Case): DocxDocument {
  const d = Document.create();
  Document.addParagraph(d, "Body");
  Document.setWatermark(d, c.watermark);
  for (const [key, def] of c.headers ?? []) {
    Document.setHeader(d, key, def.content);
  }
  if (c.watermark.type === "image") {
    (d as { images?: ImageDef[] }).images = [IMAGE];
  }
  return d;
}

async function streamed(c: Case): Promise<Uint8Array> {
  const stream = new Stream.DocumentWriter({
    watermark: c.watermark,
    headers: c.headers,
    images: c.watermark.type === "image" ? [IMAGE] : undefined
  });
  stream.add(Build.textParagraph("Body"));
  return stream.finalize();
}

/** Each header part's text and watermark count, and the final section's header references. */
function headerSummary(parts: Map<string, Uint8Array>): unknown {
  const doc = partText(parts, "word/document.xml");
  const sectPr = doc.slice(doc.lastIndexOf("<w:sectPr"));
  const refs = [...sectPr.matchAll(/<w:headerReference w:type="(\w+)" r:id="([^"]+)"/g)];
  const rels = partText(parts, "word/_rels/document.xml.rels");
  return refs.map(([, type, rId]) => {
    const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)![1]!;
    const xml = partText(parts, `word/${target}`);
    const picture = /<v:imagedata r:id="([^"]+)"/.exec(xml)?.[1];
    const partRels = partText(parts, `word/_rels/${target}.rels`);
    return {
      type,
      text: [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join(""),
      watermarks: xml.match(/PowerPlusWaterMarkObject/g)?.length ?? 0,
      picture: picture ? new RegExp(`Id="${picture}"[^>]*Target="media/`).test(partRels) : null
    };
  });
}

describe("streaming writer watermark", () => {
  it.each(CASES.map(c => [c.name, c] as const))("%s", async (_name, c) => {
    const fromBuffer = await packageParts(await Io.toBuffer(buffered(c)));
    const fromStream = await packageParts(await streamed(c));
    expect(packageIssues(fromStream)).toEqual([]);
    const summary = headerSummary(fromStream);
    expect(summary).toEqual(headerSummary(fromBuffer));
    const defaults = (summary as { type: string; watermarks: number }[]).filter(
      h => h.type === "default"
    );
    expect(defaults).toHaveLength(1);
    expect(defaults[0]!.watermarks).toBe(1);
  });
});
