/**
 * A built-in style the content references but the document does not define
 * is written as Word defines it on first use, so Word treats a `Heading1`
 * paragraph as a heading (outline level, TOC entry) instead of as Normal.
 */

import { describe, expect, it } from "vitest";

import { Build, Document, Io, Stream } from "../index";
import type { DocxDocument } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

async function stylesXml(doc: DocxDocument): Promise<string> {
  return partText(await packageParts(await Io.toBuffer(doc)), "word/styles.xml");
}

function style(xml: string, id: string): string | undefined {
  return new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(xml)?.[0];
}

describe("referenced built-in styles", () => {
  it("defines the headings addHeading uses without useDefaultStyles", async () => {
    const doc = Document.create();
    Document.addHeading(doc, "One", 1);
    Document.addHeading(doc, "Two", 2);
    const xml = await stylesXml(doc);
    const h1 = style(xml, "Heading1");
    expect(h1).toContain('<w:name w:val="heading 1"/>');
    expect(h1).toContain('<w:outlineLvl w:val="0"/>');
    expect(style(xml, "Heading2")).toContain('<w:outlineLvl w:val="1"/>');
    expect(style(xml, "Heading3")).toBeUndefined();
    // What they are based on comes too.
    expect(style(xml, "Normal")).toContain('w:default="1"');
  });

  it("defines Hyperlink, List Paragraph, Title, Caption and TOC styles when used", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "t", { style: "Title" });
    Document.addParagraph(doc, "c", { style: "Caption" });
    Document.addParagraph(doc, "l", { style: "ListParagraph" });
    Document.addParagraph(doc, "toc", { style: "TOC2" });
    Document.addParagraph(doc, "q", { style: "FootnoteText" });
    const xml = await stylesXml(doc);
    expect(style(xml, "Title")).toContain('<w:name w:val="Title"/>');
    expect(style(xml, "Caption")).toContain('<w:name w:val="caption"/>');
    expect(style(xml, "ListParagraph")).toContain('<w:name w:val="List Paragraph"/>');
    expect(style(xml, "TOC2")).toContain('<w:name w:val="toc 2"/>');
    expect(style(xml, "FootnoteText")).toContain('<w:name w:val="footnote text"/>');
  });

  it("never replaces a defined style, nor adds an id that is not built in", async () => {
    const doc = Document.create();
    Document.addStyle(doc, {
      type: "paragraph",
      styleId: "Heading1",
      name: "heading 1",
      runProperties: { color: "FF0000" }
    });
    Document.addHeading(doc, "One", 1);
    Document.addParagraph(doc, "x", { style: "MyOwn" });
    const xml = await stylesXml(doc);
    expect(style(xml, "Heading1")).toContain("FF0000");
    expect(style(xml, "Heading1")).not.toContain("outlineLvl");
    expect(style(xml, "MyOwn")).toBeUndefined();
  });

  it("does not define a second default paragraph style or a localized built-in twice", async () => {
    const doc = await Io.read(
      await buildDocx({
        body: '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
      })
    );
    const localized: DocxDocument = {
      ...doc,
      styles: [
        { type: "paragraph", styleId: "a", name: "Normal", isDefault: true },
        { type: "paragraph", styleId: "1", name: "heading 1", basedOn: "a" }
      ]
    };
    const xml = await stylesXml(localized);
    expect(style(xml, "Heading1")).toBeUndefined();
    expect(xml.match(/w:default="1"/g)?.length).toBe(1);
  });

  it("defines them in a streamed document too", async () => {
    const stream = new Stream.DocumentWriter();
    stream.add(Build.heading("One", 2));
    const xml = partText(await packageParts(await stream.finalize()), "word/styles.xml");
    expect(style(xml, "Heading2")).toContain('<w:outlineLvl w:val="1"/>');
  });

  it("is stable across a read → write round trip", async () => {
    const doc = Document.create();
    Document.addHeading(doc, "One", 1);
    const once = await stylesXml(doc);
    const twice = await stylesXml(await Io.read(await Io.toBuffer(doc)));
    expect(twice).toBe(once);
  });
});
