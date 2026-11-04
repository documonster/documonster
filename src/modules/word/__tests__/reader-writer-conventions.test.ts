/**
 * What the writer writes is what the reader reads back: a whole field's
 * instruction (the model keeps it trimmed), a style's outline level (inside
 * its `w:pPr`), and the on/off settings (`ST_OnOff`, honouring an explicit
 * `false`, `0` or `off`).
 */

import { describe, expect, it } from "vitest";

import { Build, Document, Io, Transform } from "../index";
import type { DocumentSettings, DocxDocument } from "../types";
import { buildDocx, packageParts, partText } from "./roundtrip-integrity.helpers";

async function part(doc: DocxDocument, path: string): Promise<string> {
  return partText(await packageParts(await Io.toBuffer(doc)), path);
}

describe("whole-field instructions", () => {
  it("are written as the reader reads them back, so an update is stable", async () => {
    const doc = Document.create();
    Document.addTableOfContents(doc);
    Document.addHeading(doc, "One", 1);
    Document.addHeading(doc, "Two", 1);
    const once = await Io.toBuffer(Transform.updateFields(doc));
    const twice = await Io.toBuffer(Transform.updateFields(await Io.read(once)));
    const xml = (b: Uint8Array): Promise<string> =>
      packageParts(b).then(p => partText(p, "word/document.xml"));
    expect(await xml(twice)).toBe(await xml(once));
    expect(await xml(once)).toContain(">PAGEREF _Toc");
  });

  it("keep range markers inside the instruction at their place", async () => {
    const doc = await Io.read(
      await buildDocx({
        body:
          '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          '<w:r><w:instrText xml:space="preserve">  REF </w:instrText></w:r>' +
          '<w:bookmarkStart w:id="1" w:name="m"/>' +
          '<w:r><w:instrText xml:space="preserve">x \\h </w:instrText></w:r>' +
          '<w:bookmarkEnd w:id="1"/>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>v</w:t></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
      })
    );
    const once = await Io.toBuffer(doc);
    const twice = await Io.toBuffer(await Io.read(once));
    const xml = partText(await packageParts(once), "word/document.xml");
    expect(partText(await packageParts(twice), "word/document.xml")).toBe(xml);
    expect(xml).toMatch(/>REF <\/w:instrText>[\s\S]*bookmarkStart[\s\S]*>x \\h<\/w:instrText>/);
  });

  it("built with padding are written trimmed", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, Build.paragraph([Build.formTextField({ name: "f" })]));
    expect(await part(doc, "word/document.xml")).toContain(">FORMTEXT</w:instrText>");
  });
});

describe("StyleDef.outlineLevel", () => {
  it("is written inside the style's w:pPr and read back from there", async () => {
    const doc = Document.create();
    Document.addStyle(doc, {
      type: "paragraph",
      styleId: "Chap",
      name: "Chap",
      outlineLevel: 2,
      paragraphProperties: { keepNext: true }
    });
    const xml = await part(doc, "word/styles.xml");
    const chap = /<w:style\b[^>]*w:styleId="Chap"[\s\S]*?<\/w:style>/.exec(xml)![0];
    expect(chap).toMatch(/<w:pPr>[\s\S]*<w:outlineLvl w:val="2"\/>[\s\S]*<\/w:pPr>/);
    expect(chap.replace(/<w:pPr>[\s\S]*<\/w:pPr>/, "")).not.toContain("outlineLvl");
    const back = await Io.read(await Io.toBuffer(doc));
    const style = back.styles!.find(s => s.styleId === "Chap")!;
    expect(style.paragraphProperties?.outlineLevel).toBe(2);
    expect(style.paragraphProperties?.keepNext).toBe(true);
    expect(await part(back, "word/styles.xml")).toBe(xml);
  });

  it("is read from the w:style position older files have it in", async () => {
    const doc = await Io.read(
      await buildDocx({
        body: "<w:p/>",
        styles:
          '<w:style w:type="paragraph" w:styleId="Chap"><w:name w:val="Chap"/>' +
          '<w:outlineLvl w:val="3"/></w:style>'
      })
    );
    const style = doc.styles!.find(s => s.styleId === "Chap")!;
    expect(style.paragraphProperties?.outlineLevel).toBe(3);
  });
});

describe("on/off settings", () => {
  const flags: [string, (s: DocumentSettings) => unknown][] = [
    ["evenAndOddHeaders", s => s.evenAndOddHeaders],
    ["trackRevisions", s => s.trackRevisions],
    ["mirrorMargins", s => s.mirrorMargins],
    ["gutterAtTop", s => s.gutterAtTop],
    ["displayBackgroundShape", s => s.displayBackgroundShape],
    ["updateFields", s => s.updateFieldsOnOpen],
    ["doNotTrackMoves", s => s.doNotTrackMoves],
    ["doNotTrackFormatting", s => s.doNotTrackFormatting],
    ["doNotDemoteNonCombiningChars", s => s.doNotDemoteAsianTextFirstLine],
    ["saveSubsetFonts", s => s.saveSubsetFonts],
    ["noPunctuationKerning", s => s.noPunctuationKerning],
    ["bordersDoNotSurroundHeader", s => s.bordersDoNotSurroundHeader],
    ["bordersDoNotSurroundFooter", s => s.bordersDoNotSurroundFooter],
    ["autoHyphenation", s => s.autoHyphenation],
    ["doNotHyphenateCaps", s => s.hyphenation?.doNotHyphenateCaps]
  ];
  const read = async (xml: string): Promise<DocumentSettings> =>
    (await Io.read(await buildDocx({ body: "<w:p/>", settings: xml }))).settings ?? {};
  const withHyphenation = (name: string, xml: string): string =>
    name === "doNotHyphenateCaps" ? `<w:autoHyphenation/>${xml}` : xml;

  it.each(flags)("%s honours ST_OnOff", async (name, get) => {
    for (const val of ["false", "0", "off"]) {
      expect(
        get(await read(withHyphenation(name, `<w:${name} w:val="${val}"/>`))),
        `${name} ${val}`
      ).toBeFalsy();
    }
    for (const attr of ["", ' w:val="true"', ' w:val="1"', ' w:val="on"']) {
      expect(get(await read(withHyphenation(name, `<w:${name}${attr}/>`))), `${name}${attr}`).toBe(
        true
      );
    }
  });

  it("reads documentProtection's enforcement and formatting as ST_OnOff", async () => {
    for (const val of ["1", "true", "on"]) {
      const s = await read(
        `<w:documentProtection w:edit="readOnly" w:enforcement="${val}" w:formatting="${val}"/>`
      );
      expect(s.documentProtection?.enforcement).toBe(true);
      expect(s.documentProtection?.formatting).toBe(true);
    }
    const off = await read(
      '<w:documentProtection w:edit="readOnly" w:enforcement="off" w:formatting="false"/>'
    );
    expect(off.documentProtection?.enforcement).toBe(false);
    expect(off.documentProtection?.formatting).toBeUndefined();
  });

  it("reads writeProtection's recommended as ST_OnOff", async () => {
    expect(
      (await read('<w:writeProtection w:recommended="on"/>')).writeProtection?.recommended
    ).toBe(true);
    expect(
      (await read('<w:writeProtection w:recommended="off"/>')).writeProtection?.recommended
    ).toBeUndefined();
  });
});
