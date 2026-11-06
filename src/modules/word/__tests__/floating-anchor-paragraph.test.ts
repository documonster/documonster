/**
 * Where a floating image's anchor sits, through a read and a write.
 *
 * A floating image the library writes gets an empty paragraph of its own, and
 * reading one back drops that carrier. An anchor Word put among a paragraph's
 * text belongs to that paragraph: it is read as `inPrecedingParagraph` and
 * written back into it, so a round trip adds no paragraph; and a carrier that
 * also ends a section keeps its section break.
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { describe, expect, it } from "vitest";

import { utf8Decoder, utf8Encoder } from "../core/internal-utils";
import { Build, Document, Io, Layout } from "../index";
import type { FloatingImage } from "../types";

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
  0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

async function entriesOf(buffer: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(buffer).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

async function documentXml(buffer: Uint8Array): Promise<string> {
  return utf8Decoder.decode((await entriesOf(buffer)).get("word/document.xml")!);
}

/** `buffer` with its `word/document.xml` passed through `edit`. */
async function patched(buffer: Uint8Array, edit: (xml: string) => string): Promise<Uint8Array> {
  const entries = await entriesOf(buffer);
  entries.set(
    "word/document.xml",
    utf8Encoder.encode(edit(utf8Decoder.decode(entries.get("word/document.xml")!)))
  );
  const archive = zip();
  for (const [path, data] of entries) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

/** A document of one text paragraph and one floating image of its own. */
async function written(): Promise<Uint8Array> {
  const doc = Document.create();
  Document.addParagraph(doc, Build.paragraph([Build.text("text beside the image")]));
  Document.addFloatingImage(doc, PNG, "png", 914400, 914400, { wrap: { style: "square" } });
  return Io.toBuffer(doc);
}

const ANCHOR_RUN = /<w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/;

describe("a floating image anchored among a paragraph's text", () => {
  it("is read as belonging to that paragraph and written back into it", async () => {
    // Move the anchor's run into the text paragraph, as Word stores it.
    const source = await patched(await written(), xml => {
      const run = xml.match(ANCHOR_RUN)![0];
      return xml.replace(`<w:p>${run}</w:p>`, "").replace("</w:t></w:r>", `</w:t></w:r>${run}`);
    });
    expect((await documentXml(source)).match(/<w:p[ >]/g)).toHaveLength(1);

    const doc = await Io.read(source);
    expect(doc.body.map(b => b.type)).toEqual(["paragraph", "floatingImage"]);
    expect((doc.body[1] as FloatingImage).inPrecedingParagraph).toBe(true);
    expect((await documentXml(await Io.toBuffer(doc))).match(/<w:p[ >]/g)).toHaveLength(1);
  });

  it("takes no line of its own, where an image in a paragraph of its own does", async () => {
    const own = await Io.read(await written());
    const shared = await Io.read(
      await patched(await written(), xml => {
        const run = xml.match(ANCHOR_RUN)![0];
        return xml.replace(`<w:p>${run}</w:p>`, "").replace("</w:t></w:r>", `</w:t></w:r>${run}`);
      })
    );
    const paragraphs = (doc: typeof own): number =>
      Layout.compute(doc).pages[0]!.content.filter(b => b.type === "paragraph").length;
    expect(paragraphs(own)).toBe(2);
    expect(paragraphs(shared)).toBe(1);
  });

  it("keeps the section break of a paragraph holding nothing but the anchor", async () => {
    const source = await patched(await written(), xml =>
      xml.replace(
        /<w:p>(<w:r><w:drawing>)/,
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>$1'
      )
    );
    const doc = await Io.read(source);
    const breaks = doc.body.filter(
      b => b.type === "paragraph" && b.properties?.sectionProperties !== undefined
    );
    expect(breaks).toHaveLength(1);
  });
});
