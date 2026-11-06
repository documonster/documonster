/**
 * `Editor#setMetadata` takes a `PdfMetadataUpdate`: JSON merge patch (RFC 7396) semantics —
 * absent or `undefined` keeps, a value sets, `null` removes — in /Info and XMP alike.
 */
import { parseXml } from "@xml/dom";
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError } from "../index";
import type { PdfMetadataUpdate } from "../index";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
const fromLatin1 = (text: string): Uint8Array =>
  new Uint8Array([...text].map(ch => ch.charCodeAt(0) & 0xff));

const XMP =
  '<?xpacket begin="\u00ef\u00bb\u00bf" id="W5M0MpCehiHzreSzNTczkc9d"?>' +
  '<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">' +
  '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/" xmp:CreatorTool="Writer">' +
  '<dc:title><rdf:Alt><rdf:li xml:lang="x-default">Old</rdf:li><rdf:li xml:lang="de">Alt</rdf:li></rdf:Alt></dc:title>' +
  "<dc:creator><rdf:Seq><rdf:li>Ann</rdf:li></rdf:Seq></dc:creator>" +
  "<pdf:Producer>Tool 1</pdf:Producer><pdf:Keywords>k</pdf:Keywords>" +
  "<xmp:ModifyDate>2020-01-01T00:00:00Z</xmp:ModifyDate>" +
  "</rdf:Description>" +
  // Adobe-style: the same property can sit in a second top-level description too.
  '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title><rdf:Alt><rdf:li xml:lang="x-default">Old</rdf:li></rdf:Alt></dc:title></rdf:Description>' +
  "</rdf:RDF></x:xmpmeta>" +
  '<?xpacket end="w"?>';

function source(): Uint8Array {
  const bodies = [
    "<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
    `<< /Type /Metadata /Subtype /XML /Length ${XMP.length} >>\nstream\n${XMP}\nendstream`,
    "<< /Title (Old) /Author (Ann) /Creator (Writer) /Producer (Tool 1) /Keywords (k) /ModDate (D:20200101000000Z) >>"
  ];
  let out = "%PDF-1.7\n";
  const offsets: number[] = [];
  bodies.forEach((body, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${bodies.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  out += `trailer\n<< /Size ${bodies.length + 1} /Root 1 0 R /Info 5 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return fromLatin1(out);
}

async function xmpOf(bytes: Uint8Array): Promise<string> {
  const xml = (await Pdf.extract(bytes)).metadata.xmpXml ?? "";
  parseXml(xml.replace(/^\uFEFF/, ""), { xmlns: true, processingInstructions: true });
  return xml;
}

/** The /Info dictionary text of the last revision. */
function lastInfo(bytes: Uint8Array): string {
  const raw = latin1(bytes);
  const ref = [...raw.matchAll(/\/Info (\d+) 0 R/g)].at(-1)?.[1];
  const at = raw.lastIndexOf(`\n${ref} 0 obj`);
  return raw.slice(at, raw.indexOf("endobj", at));
}

for (const incremental of [false, true]) {
  const mode = incremental ? "incremental" : "full";

  describe(`setMetadata with null removes (${mode})`, () => {
    it("removes the /Info entry and every XMP form of the property", async () => {
      const update: PdfMetadataUpdate = { title: null, creator: null, author: null };
      const out = await Pdf.Editor.read(source()).setMetadata(update).toBuffer({ incremental });
      const info = lastInfo(out);
      expect(info).not.toMatch(/\/Title\b/);
      expect(info).not.toMatch(/\/Author\b/);
      expect(info).not.toMatch(/\/Creator\b/);
      // Untouched entries are kept.
      expect(info).toMatch(/\/Producer\b/);
      expect(info).toMatch(/\/Keywords\b/);

      const xml = await xmpOf(out);
      // Element form in both descriptions, the language alternative with all its languages.
      expect(xml).not.toContain("dc:title");
      expect(xml).not.toContain(">Alt<");
      // A Seq container.
      expect(xml).not.toContain("dc:creator");
      // Attribute form.
      expect(xml).not.toContain("CreatorTool");
      expect(xml).toContain("<pdf:Producer>Tool 1</pdf:Producer>");
      expect(xml).toMatch(/<xmp:MetadataDate>[^<]+<\/xmp:MetadataDate>/);

      const meta = (await Pdf.extract(out)).metadata;
      expect(meta.title).toBe("");
      expect(meta.author).toBe("");
      expect(meta.creator).toBe("");
      expect(meta.producer).toBe("Tool 1");
      expect(meta.keywords).toBe("k");
    });

    it("removes a date, and sets and removes in one call", async () => {
      const out = await Pdf.Editor.read(source())
        .setMetadata({ modDate: null, subject: "S", producer: null })
        .toBuffer({ incremental });
      const info = lastInfo(out);
      expect(info).not.toMatch(/\/ModDate\b/);
      expect(info).not.toMatch(/\/Producer\b/);
      expect(info).toMatch(/\/Subject\b/);
      const xml = await xmpOf(out);
      expect(xml).not.toContain("xmp:ModifyDate");
      expect(xml).not.toContain("pdf:Producer");
      expect(xml).toContain("S</rdf:li>");
      expect((await Pdf.extract(out)).metadata.modDate).toBeUndefined();
    });
  });
}

describe("merge-patch semantics across calls", () => {
  it("a later value replaces an earlier null, a later null an earlier value; undefined keeps", async () => {
    const editor = Pdf.Editor.read(source());
    editor.setMetadata({ title: null, author: "Bob" });
    editor.setMetadata({ title: "Back", author: null, subject: undefined });
    const meta = (await Pdf.extract(await editor.toBuffer())).metadata;
    expect(meta.title).toBe("Back");
    expect(meta.author).toBe("");
    expect(meta.creator).toBe("Writer");
  });

  it("removing what is absent changes nothing else", async () => {
    const out = await Pdf.Editor.read(source()).setMetadata({ subject: null }).toBuffer();
    const meta = (await Pdf.extract(out)).metadata;
    expect(meta.title).toBe("Old");
    expect(meta.subject).toBe("");
  });

  it("still checks every non-null value", () => {
    const editor = Pdf.Editor.read(source());
    expect(() => editor.setMetadata({ title: 42 as never })).toThrow(PdfArgumentError);
    expect(() => editor.setMetadata({ modDate: new Date("x") })).toThrow(PdfArgumentError);
    expect(() => editor.setMetadata({ title: "a\u0001" })).toThrow(PdfArgumentError);
  });
});
