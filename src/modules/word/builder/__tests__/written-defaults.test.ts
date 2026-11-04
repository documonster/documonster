import { packageParts, partText } from "@word/__tests__/roundtrip-integrity.helpers";
import { DocxError } from "@word/errors";
import { Document, Io, Layout, Streaming } from "@word/index";
import type { DocxDocument } from "@word/types";
/**
 * A document the library builds states current Word's new-document defaults —
 * Aptos 12pt, 8pt after, 1.16 lines — and is laid out with them.
 *
 * Without `w:docDefaults` (or with an empty one) Word takes the defaults from
 * the user's Normal template, which differs between Word versions, so the same
 * file paginated differently in Word and in the library.
 */
import { describe, expect, it } from "vitest";

async function stylesXml(bytes: Uint8Array): Promise<string> {
  return partText(await packageParts(bytes), "word/styles.xml");
}

function sample(): DocxDocument {
  const doc = Document.create();
  for (let i = 0; i < 80; i++) {
    Document.addParagraph(doc, `Paragraph ${i} lorem ipsum dolor sit amet `.repeat(4));
  }
  return doc;
}

describe("written docDefaults", () => {
  it("states fonts, size, language and spacing for a document with no style calls", async () => {
    const xml = await stylesXml(await Io.toBuffer(sample()));
    expect(xml).toContain("<w:rPrDefault>");
    expect(xml).toMatch(/<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/);
    expect(xml).toContain('<w:sz w:val="24"/>');
    expect(xml).toContain('<w:szCs w:val="24"/>');
    expect(xml).toContain('<w:lang w:val="en-US"/>');
    expect(xml).toMatch(
      /<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"\/>/
    );
  });

  it("lays a new document out with the defaults it states", () => {
    const doc = sample();
    const stated: DocxDocument = {
      ...doc,
      deriveEastAsianDefaults: undefined,
      docDefaults: {
        runProperties: { font: { ascii: "Aptos", hAnsi: "Aptos" }, size: 24 },
        paragraphProperties: { spacing: { after: 160, line: 278, lineRule: "auto" } }
      }
    };
    const a = Layout.getPageMap(doc);
    const b = Layout.getPageMap(stated);
    expect(a.contentPages).toEqual(b.contentPages);
    // Not the fallback a file without defaults gets.
    const bare = Layout.getPageMap({ ...doc, deriveEastAsianDefaults: undefined });
    expect(bare.pageCount).toBeLessThan(a.pageCount);
  });

  it("paginates the written document as the one in memory", async () => {
    const doc = sample();
    const back = await Io.read(await Io.toBuffer(doc));
    const a = Layout.getPageMap(doc);
    const b = Layout.getPageMap(back);
    expect(b.pageCount).toBe(a.pageCount);
    expect(b.contentPages).toEqual(a.contentPages);
  });

  it("fills only what the caller left out", async () => {
    const doc = Document.create();
    Document.setDocDefaults(doc, { runProperties: { size: 28 } });
    Document.addParagraph(doc, "x");
    const xml = await stylesXml(await Io.toBuffer(doc));
    expect(xml).toContain('<w:sz w:val="28"/>');
    expect(xml).toMatch(/<w:rFonts w:ascii="Aptos"/);
    expect(doc.docDefaults).toEqual({ runProperties: { size: 28 } });
  });

  it("keeps useDefaultStyles' values", async () => {
    const doc = Document.create();
    Document.useDefaultStyles(doc);
    const xml = await stylesXml(await Io.toBuffer(doc));
    expect(xml).toMatch(/w:after="160"/);
    expect(xml).toMatch(/w:line="259"/);
  });

  it("states them in the streaming writer too", async () => {
    const stream = Streaming.createDocxStream();
    stream.addText("x");
    const xml = await stylesXml(await stream.finalize());
    expect(xml).toContain("<w:rPrDefault>");
    expect(xml).toContain("<w:pPrDefault>");
  });

  it("leaves a read document without docDefaults as it was", async () => {
    const doc = sample();
    const read = await Io.read(await Io.toBuffer(doc));
    const bare: DocxDocument = { ...read, docDefaults: undefined };
    const xml = await stylesXml(await Io.toBuffer(bare));
    expect(xml).not.toContain("docDefaults");
  });

  it("rejects defaults in a shape it does not know instead of writing them empty", () => {
    const doc = Document.create();
    expect(() =>
      Document.setDocDefaults(doc, { run: { font: "Times New Roman" } } as never)
    ).toThrow(DocxError);
    expect(doc.docDefaults).toBeUndefined();
  });
});
