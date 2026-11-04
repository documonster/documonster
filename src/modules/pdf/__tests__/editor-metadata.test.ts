/**
 * The document information dictionary and XMP packet through an editor write: a full rewrite
 * keeps every entry it does not deliberately change, and `Editor#setMetadata` changes the ones
 * it is given — in /Info and, when the file has one, in its XMP packet — on both write paths.
 */
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError } from "../index";
import { buildRawPdf } from "./raw-pdf.helpers";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

// The file is assembled as latin-1 text, one character per byte, so the packet's U+FEFF is
// spelled as its three UTF-8 bytes; written as "\uFEFF" it became the byte 0xFF, which is not
// UTF-8, and the editor rightly leaves a packet it cannot decode as it found it.
const XMP = [
  '<?xpacket begin="\u00ef\u00bb\u00bf" id="W5M0MpCehiHzreSzNTczkc9d"?>',
  '<x:xmpmeta xmlns:x="adobe:ns:meta/">',
  '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">',
  '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"',
  '  xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/"',
  '  xmp:CreatorTool="Orig Writer">',
  '  <dc:title><rdf:Alt><rdf:li xml:lang="x-default">Old title</rdf:li></rdf:Alt></dc:title>',
  "  <pdf:Keywords>alpha, beta</pdf:Keywords>",
  "  <pdf:Producer>OrigTool 1.0</pdf:Producer>",
  "  <xmp:ModifyDate>2020-02-03T04:05:06Z</xmp:ModifyDate>",
  "</rdf:Description>",
  "</rdf:RDF>",
  "</x:xmpmeta>",
  '<?xpacket end="w"?>'
].join("\n");

/** A one-page file whose /Info has keywords, a mod date, a producer and a custom entry. */
function withInfo(options: { xmp?: boolean } = {}): Uint8Array {
  const bodies = [
    `<< /Type /Catalog /Pages 2 0 R${options.xmp ? " /Metadata 6 0 R" : ""} >>`,
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
    "<< /Length 0 >>\nstream\n\nendstream",
    "<< /Title (Old title) /Author (Ann) /Keywords (alpha, beta) /Creator (Orig Writer) " +
      "/Producer (OrigTool 1.0) /CreationDate (D:20200102030405Z) " +
      "/ModDate (D:20200203040506Z) /Company (ACME) /Trapped /False >>"
  ];
  if (options.xmp) {
    bodies.push(
      `<< /Type /Metadata /Subtype /XML /Length ${XMP.length} >>\nstream\n${XMP}\nendstream`
    );
  }
  const bytes = latin1(buildRawPdf(bodies)).replace("/Root 1 0 R >>", "/Root 1 0 R /Info 5 0 R >>");
  return new Uint8Array([...bytes].map(ch => ch.charCodeAt(0) & 0xff));
}

function utf8(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

describe("a full rewrite keeps the information dictionary", () => {
  it("keeps keywords, the mod date, the producer and custom entries", async () => {
    const editor = Pdf.Editor.read(withInfo());
    editor.getPage(0).drawText("x", { x: 10, y: 10 });
    const out = await editor.toBuffer();
    const meta = (await Pdf.extract(out)).metadata;
    // A rewrite rebuilt /Info from title, author, subject, creator and creation date only.
    expect(meta.keywords).toBe("alpha, beta");
    expect(meta.modDate?.toISOString()).toBe("2020-02-03T04:05:06.000Z");
    expect(meta.producer).toBe("OrigTool 1.0");
    expect(meta.custom).toEqual({ Company: "ACME", Trapped: "False" });
    expect(meta.title).toBe("Old title");
    expect(meta.creationDate?.toISOString()).toBe("2020-01-02T03:04:05.000Z");
  });

  it("keeps them on an encrypted rewrite too", async () => {
    const b = new Pdf.Builder();
    b.setMetadata({ title: "Secret", author: "Ann" });
    b.addPage();
    b.setEncryption({ ownerPassword: "o", userPassword: "u" });
    const editor = Pdf.Editor.read(await b.toBuffer(), { password: "u" });
    editor.getPage(0).drawText("x", { x: 10, y: 10 });
    const meta = (await Pdf.extract(await editor.toBuffer(), { password: "u" })).metadata;
    expect(meta.title).toBe("Secret");
    expect(meta.author).toBe("Ann");
  });

  for (const incremental of [false, true]) {
    it(`sets metadata on an encrypted file (${incremental ? "incremental" : "full"})`, async () => {
      const b = new Pdf.Builder();
      b.setMetadata({ title: "Secret", author: "Ann" });
      b.addPage();
      b.setPdfACompliance();
      const pdfA = await b.toBuffer();
      const enc = new Pdf.Builder();
      enc.setMetadata({ title: "Secret", author: "Ann" });
      enc.addPage();
      enc.setEncryption({ ownerPassword: "o", userPassword: "u" });
      for (const [source, password] of [
        [await enc.toBuffer(), "u"],
        [pdfA, undefined]
      ] as const) {
        const editor = Pdf.Editor.read(source, password ? { password } : undefined);
        editor.setMetadata({ title: "Changed" });
        const out = await editor.toBuffer({ incremental });
        const meta = (await Pdf.extract(out, password ? { password } : undefined)).metadata;
        expect(meta.title).toBe("Changed");
        expect(meta.author).toBe("Ann");
        if (meta.xmpXml !== null) {
          expect(meta.xmpXml).toContain(">Changed</");
        }
      }
    });
  }
});

describe("Editor#setMetadata", () => {
  for (const incremental of [false, true]) {
    const mode = incremental ? "incremental" : "full";

    it(`changes what it is given and keeps the rest (${mode})`, async () => {
      const editor = Pdf.Editor.read(withInfo());
      editor.setMetadata({
        title: "New title — ✓",
        modDate: new Date("2024-05-06T07:08:09Z")
      });
      const out = await editor.toBuffer({ incremental });
      if (incremental) {
        expect(latin1(out).startsWith(latin1(withInfo()))).toBe(true);
      }
      const meta = (await Pdf.extract(out)).metadata;
      expect(meta.title).toBe("New title — ✓");
      expect(meta.modDate?.toISOString()).toBe("2024-05-06T07:08:09.000Z");
      expect(meta.author).toBe("Ann");
      expect(meta.keywords).toBe("alpha, beta");
      expect(meta.custom.Company).toBe("ACME");
    });

    it(`keeps the XMP packet consistent with /Info (${mode})`, async () => {
      const editor = Pdf.Editor.read(withInfo({ xmp: true }));
      editor.setMetadata({
        title: "A & B",
        author: "Bob",
        creator: "New Writer",
        modDate: new Date("2024-05-06T07:08:09Z")
      });
      const out = await editor.toBuffer({ incremental });
      const xmp = (await Pdf.extract(out)).metadata.xmpXml ?? "";
      expect(xmp).toContain('<rdf:li xml:lang="x-default">A &amp; B</rdf:li>');
      expect(xmp).not.toContain("Old title");
      // dc:creator was absent: added, with its namespace.
      expect(xmp).toMatch(/<dc:creator><rdf:Seq><rdf:li>Bob<\/rdf:li><\/rdf:Seq><\/dc:creator>/);
      // The attribute form is updated in place.
      expect(xmp).toContain('xmp:CreatorTool="New Writer"');
      expect(xmp).toContain("<xmp:ModifyDate>2024-05-06T07:08:09Z</xmp:ModifyDate>");
      // What was not set is left as it was.
      expect(xmp).toContain("<pdf:Keywords>alpha, beta</pdf:Keywords>");
      expect(xmp).toContain("<pdf:Producer>OrigTool 1.0</pdf:Producer>");
      expect(utf8(new TextEncoder().encode(xmp))).toContain("<?xpacket end=");
    });
  }

  it("writes the change even with no other edit", async () => {
    const source = withInfo();
    const out = await Pdf.Editor.read(source)
      .setMetadata({ subject: "S" })
      .toBuffer({ incremental: true });
    expect(out.length).toBeGreaterThan(source.length);
    expect((await Pdf.extract(out)).metadata.subject).toBe("S");
  });

  it("adds an information dictionary to a file without one", async () => {
    const b = new Pdf.Builder();
    b.addPage();
    const text = latin1(await b.toBuffer());
    const bare = text.replace(/\/Info \d+ 0 R/, "");
    const bytes = new Uint8Array([...bare].map(ch => ch.charCodeAt(0) & 0xff));
    for (const incremental of [false, true]) {
      const out = await Pdf.Editor.read(bytes)
        .setMetadata({ title: "T" })
        .toBuffer({ incremental });
      expect((await Pdf.extract(out)).metadata.title).toBe("T");
    }
  });

  it("checks what it is given, and keeps a copy", async () => {
    const editor = Pdf.Editor.read(withInfo());
    expect(() => editor.setMetadata(undefined as never)).toThrow(PdfArgumentError);
    expect(() => editor.setMetadata({ title: 42 as never })).toThrow(PdfArgumentError);
    expect(() => editor.setMetadata({ modDate: new Date("x") })).toThrow(PdfArgumentError);
    const metadata = { title: "One" };
    editor.setMetadata(metadata);
    (metadata as { title: unknown }).title = 42;
    expect((await Pdf.extract(await editor.toBuffer())).metadata.title).toBe("One");
  });

  it("is frozen into a write when the write is called", async () => {
    const editor = Pdf.Editor.read(withInfo());
    editor.setMetadata({ title: "First" });
    const pending = editor.toBuffer();
    editor.setMetadata({ title: "Second" });
    expect((await Pdf.extract(await pending)).metadata.title).toBe("First");
    expect((await Pdf.extract(await editor.toBuffer())).metadata.title).toBe("Second");
  });

  it("reaches every part of splitPages", async () => {
    const editor = Pdf.Editor.read(withInfo()).setMetadata({ title: "Split" });
    const [part] = await editor.splitPages();
    expect((await Pdf.extract(part)).metadata.title).toBe("Split");
  });
});
