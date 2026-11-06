/**
 * Round 3: XMP editing with the real parser, Info without invented entries, page geometry
 * through a rewrite and a page copy, and the remaining boundary checks.
 */
import { parseXml } from "@xml/dom";
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError } from "../index";

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
const fromLatin1 = (text: string): Uint8Array =>
  new Uint8Array([...text].map(ch => ch.charCodeAt(0) & 0xff));

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/** Objects 1..n, `streams[i]` replacing object i+1's body with a stream of those bytes. */
function rawPdf(
  bodies: ReadonlyArray<string | { dict: string; data: Uint8Array }>,
  trailer = "/Root 1 0 R /Info 5 0 R"
): Uint8Array {
  const parts: Uint8Array[] = [fromLatin1("%PDF-1.7\n")];
  let length = parts[0].length;
  const offsets: number[] = [];
  bodies.forEach((body, i) => {
    offsets.push(length);
    const chunk =
      typeof body === "string"
        ? fromLatin1(`${i + 1} 0 obj\n${body}\nendobj\n`)
        : concat([
            fromLatin1(`${i + 1} 0 obj\n<< ${body.dict} /Length ${body.data.length} >>\nstream\n`),
            body.data,
            fromLatin1("\nendstream\nendobj\n")
          ]);
    parts.push(chunk);
    length += chunk.length;
  });
  let tail = `xref\n0 ${bodies.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    tail += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  tail += `trailer\n<< /Size ${bodies.length + 1} ${trailer} >>\nstartxref\n${length}\n%%EOF\n`;
  parts.push(fromLatin1(tail));
  return concat(parts);
}

const RDF = 'xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"';

function packet(descriptions: string): string {
  return `<?xpacket begin="\uFEFF" id="W5M0MpCehiHzreSzNTczkc9d"?>\n<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF ${RDF}>${descriptions}</rdf:RDF></x:xmpmeta>\n<?xpacket end="w"?>`;
}

function withXmp(
  xmp: string | Uint8Array,
  options: { info?: string; filter?: string } = {}
): Uint8Array {
  const data = typeof xmp === "string" ? new TextEncoder().encode(xmp) : xmp;
  return rawPdf([
    "<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
    { dict: `/Type /Metadata /Subtype /XML /SourcePage 3 0 R ${options.filter ?? ""}`, data },
    options.info ?? "<< /Producer (Other) >>"
  ]);
}

/** The XMP packet of `bytes` — parsed with namespaces, so an unbound prefix fails. */
async function xmpOf(bytes: Uint8Array): Promise<string> {
  const xml = (await Pdf.extract(bytes)).metadata.xmpXml ?? "";
  parseXml(xml.replace(/^\uFEFF/, ""), { xmlns: true, processingInstructions: true });
  return xml;
}

const MODES = [false, true] as const;
const modeName = (incremental: boolean): string => (incremental ? "incremental" : "full");

describe("XMP is edited as XML", () => {
  for (const incremental of MODES) {
    it(`binds every prefix it writes, across successive saves (${modeName(incremental)})`, async () => {
      // R3-B1: a packet that does not bind dc; the second save of the first's output.
      const first = await Pdf.Editor.read(withXmp(packet('<rdf:Description rdf:about=""/>')))
        .setMetadata({ title: "T1", author: "A" })
        .toBuffer({ incremental });
      await xmpOf(first);
      const second = await Pdf.Editor.read(first)
        .setMetadata({ title: "T2" })
        .toBuffer({ incremental });
      const xml = await xmpOf(second);
      expect(xml).toContain(">T2<");
      expect(xml).not.toContain(">T1<");
      expect((await Pdf.extract(second)).metadata.author).toBe("A");
    });

    it(`adds to the description where the namespace is bound (${modeName(incremental)})`, async () => {
      // Adobe style: one description per schema; xmp is bound only on the second.
      const source = withXmp(
        packet(
          '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:format>application/pdf</dc:format></rdf:Description>' +
            '<rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/"><xmp:Rating>3</xmp:Rating></rdf:Description>'
        )
      );
      const out = await Pdf.Editor.read(source)
        .setMetadata({ creator: "Tool", modDate: new Date("2024-01-02T03:04:05Z") })
        .toBuffer({ incremental });
      const xml = await xmpOf(out);
      expect(xml).toMatch(
        /xmlns:xmp="http:\/\/ns.adobe.com\/xap\/1.0\/"><xmp:Rating>3<\/xmp:Rating><xmp:CreatorTool>Tool<\/xmp:CreatorTool>/
      );
      expect(xml).toContain("<dc:format>application/pdf</dc:format>");
    });

    it(`changes an attribute-form property in place, any quote, any content (${modeName(incremental)})`, async () => {
      // R3-B2: the property was added a second time, or written as invalid XML.
      for (const old of [`"Bob's Tool"`, `'Old'`]) {
        const source = withXmp(
          packet(
            `<rdf:Description rdf:about="" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmp:CreatorTool=${old}/>`
          )
        );
        const out = await Pdf.Editor.read(source)
          .setMetadata({ creator: 'O\'Reilly "Tool"' })
          .toBuffer({ incremental });
        const xml = await xmpOf(out);
        expect(xml.match(/CreatorTool/g)).toHaveLength(1);
        expect((await Pdf.extract(out)).metadata.creator).toBe('O\'Reilly "Tool"');
      }
    });

    it(`never edits inside a nested description or a comment (${modeName(incremental)})`, async () => {
      // R3-B3: the title went into an xmpMM:History event.
      const source = withXmp(
        packet(
          '<!-- <dc:title>not this</dc:title> --><rdf:Description rdf:about="" xmlns:xmpMM="http://ns.adobe.com/xap/1.0/mm/" xmlns:stEvt="http://ns.adobe.com/xap/1.0/sType/ResourceEvent#" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
            '<xmpMM:History><rdf:Seq><rdf:li><rdf:Description stEvt:action="created"><dc:title>event</dc:title></rdf:Description></rdf:li></rdf:Seq></xmpMM:History>' +
            "</rdf:Description>"
        )
      );
      const out = await Pdf.Editor.read(source)
        .setMetadata({ title: "T" })
        .toBuffer({ incremental });
      const xml = await xmpOf(out);
      expect(xml).toContain("<!-- <dc:title>not this</dc:title> -->");
      expect(xml).toContain("<dc:title>event</dc:title>");
      expect(xml).toMatch(/<\/xmpMM:History><dc:title><rdf:Alt><rdf:li xml:lang="x-default">T</);
    });

    it(`keeps the other languages and writes each author (${modeName(incremental)})`, async () => {
      // R3-C2: an rdf:Alt lost its other languages; an rdf:Seq became one entry.
      const source = withXmp(
        packet(
          '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
            '<dc:title><rdf:Alt><rdf:li xml:lang="de-DE">Titel</rdf:li><rdf:li xml:lang="x-default">Title</rdf:li></rdf:Alt></dc:title>' +
            "<dc:creator><rdf:Seq><rdf:li>X</rdf:li><rdf:li>Y</rdf:li></rdf:Seq></dc:creator>" +
            "</rdf:Description>"
        )
      );
      const out = await Pdf.Editor.read(source)
        .setMetadata({ title: "New", author: "Ann; Bob" })
        .toBuffer({ incremental });
      const xml = await xmpOf(out);
      expect(xml).toContain('<rdf:li xml:lang="de-DE">Titel</rdf:li>');
      expect(xml).toContain('<rdf:li xml:lang="x-default">New</rdf:li>');
      expect(xml).toContain("<rdf:Seq><rdf:li>Ann</rdf:li><rdf:li>Bob</rdf:li></rdf:Seq>");
    });

    it(`updates xmp:MetadataDate and pdf:Producer (${modeName(incremental)})`, async () => {
      const source = withXmp(packet('<rdf:Description rdf:about=""/>'));
      const out = await Pdf.Editor.read(source)
        .setMetadata({ producer: "Mine 2.0", modDate: new Date("2024-01-02T03:04:05Z") })
        .toBuffer({ incremental });
      const xml = await xmpOf(out);
      expect(xml).toMatch(/<xmp:MetadataDate[^>]*>\d{4}-\d\d-\d\dT[\d:]+Z<\/xmp:MetadataDate>/);
      expect(xml).toMatch(/<pdf:Producer[^>]*>Mine 2.0<\/pdf:Producer>/);
      expect((await Pdf.extract(out)).metadata.producer).toBe("Mine 2.0");
    });

    it(`reads and writes back a UTF-16 packet (${modeName(incremental)})`, async () => {
      const text = packet(
        '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:format>application/pdf</dc:format></rdf:Description>'
      );
      const utf16 = new Uint8Array(2 + text.length * 2);
      utf16.set([0xfe, 0xff]);
      for (let i = 0; i < text.length; i++) {
        utf16[2 + i * 2] = text.charCodeAt(i) >> 8;
        utf16[3 + i * 2] = text.charCodeAt(i) & 0xff;
      }
      const out = await Pdf.Editor.read(withXmp(utf16))
        .setMetadata({ title: "Ünïcode" })
        .toBuffer({ incremental });
      const raw = latin1(out);
      // Written back as UTF-16BE with its BOM: the title's letters are byte pairs.
      expect(raw).toContain("\u00fe\u00ff");
      expect(raw).toContain("\u0000\u00dc\u0000n");
      const stream = raw.slice(raw.lastIndexOf("/Type /Metadata"));
      expect(stream).not.toContain("/Filter");
    });

    it(`keeps a packet it cannot decode untouched, and still updates /Info (${modeName(incremental)})`, async () => {
      // R3-B8: a stream whose filter fails was written back without its filter.
      const garbage = fromLatin1("xx\u0001\u00ff not deflate");
      const source = withXmp(garbage, { filter: "/Filter /FlateDecode" });
      const out = await Pdf.Editor.read(source)
        .setMetadata({ title: "T" })
        .toBuffer({ incremental });
      const raw = latin1(out);
      const metadataDicts = raw.match(/<<[^<>]*\/Type \/Metadata[^<>]*>>/g) ?? [];
      expect(metadataDicts.length).toBeGreaterThan(0);
      for (const dict of metadataDicts) {
        expect(dict).toContain("/Filter /FlateDecode");
      }
      expect((await Pdf.extract(out)).metadata.title).toBe("T");
      // A packet that is XML but not XMP is kept too.
      const notXmp = await Pdf.Editor.read(withXmp("<a>b</a>"))
        .setMetadata({ title: "T" })
        .toBuffer({ incremental });
      expect((await Pdf.extract(notXmp)).metadata.xmpXml).toBe("<a>b</a>");
    });
  }

  it("writes the packet uncompressed in an incremental update", async () => {
    // R3-B4: buildIncremental deflated any stream over 256 bytes.
    const b = new Pdf.Builder();
    b.addPage();
    b.setPdfACompliance();
    const out = await Pdf.Editor.read(await b.toBuffer())
      .setMetadata({ title: "T" })
      .toBuffer({ incremental: true });
    const dicts = latin1(out).match(/<<[^<>]*\/Type \/Metadata[^<>]*>>/g) ?? [];
    expect(dicts.length).toBeGreaterThanOrEqual(2);
    for (const dict of dicts) {
      expect(dict).not.toContain("/Filter");
    }
  });

  it("renumbers references in the metadata stream dictionary in a full rewrite", async () => {
    // R3-B7: /SourcePage kept the original object number.
    const out = await Pdf.Editor.read(withXmp(packet('<rdf:Description rdf:about=""/>')))
      .setMetadata({ title: "T" })
      .toBuffer();
    const raw = latin1(out);
    const ref = /\/SourcePage (\d+) 0 R/.exec(raw)?.[1];
    expect(ref).toBeDefined();
    expect(raw).toMatch(new RegExp(`(^|\\n)${ref} 0 obj\\s*<<[^]*?/Type /Page\\b`));
  });

  it("refuses text XML cannot carry, in both builder and editor", () => {
    // R3-C2: U+0001 was written into the packet.
    expect(() => new Pdf.Builder().setMetadata({ title: "a\u0001b" })).toThrow(PdfArgumentError);
    expect(() =>
      Pdf.Editor.read(withXmp(packet("<rdf:Description/>"))).setMetadata({ author: "\uffff" })
    ).toThrow(PdfArgumentError);
  });
});

describe("the information dictionary gets no invented entries", () => {
  for (const incremental of MODES) {
    it(`writes no /Creator or /Producer it was not given (${modeName(incremental)})`, async () => {
      // R3-B5: `/Creator (documonster)` was added, which the XMP did not say.
      const source = rawPdf([
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>",
        "<< >>",
        "<< /Title (Old) >>"
      ]);
      const editor = Pdf.Editor.read(source);
      editor.getPage(0).drawText("x", { x: 10, y: 10 });
      const out = await editor.setMetadata({ subject: "S" }).toBuffer({ incremental });
      const meta = (await Pdf.extract(out)).metadata;
      expect(meta.creator).toBe("");
      expect(meta.producer).toBe("");
      expect(meta.title).toBe("Old");
      expect(meta.subject).toBe("S");
    });
  }

  it("lets the builder name its producer", async () => {
    const b = new Pdf.Builder();
    b.setMetadata({ producer: "Mine" });
    b.addPage();
    expect((await Pdf.extract(await b.toBuffer())).metadata.producer).toBe("Mine");
  });
});

describe("Pdf.extract reads XMP as XML", () => {
  it("decodes character references, the attribute form and prefers x-default", async () => {
    // R3-C1.
    const xmp = packet(
      '<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmp:CreatorTool="Word">' +
        '<dc:title><rdf:Alt><rdf:li xml:lang="de">Titel</rdf:li><rdf:li xml:lang="x-default">O&#39;Brien &#x2014; R&amp;D</rdf:li></rdf:Alt></dc:title>' +
        "</rdf:Description>"
    );
    const meta = (await Pdf.extract(withXmp(xmp, { info: "<< >>" }))).metadata;
    expect(meta.title).toBe("O'Brien — R&D");
    expect(meta.creator).toBe("Word");
  });
});

describe("page geometry survives a rewrite and a page copy", () => {
  function geometry(): Uint8Array {
    return rawPdf(
      [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 /Rotate 90 /CropBox [120 120 280 280] /MediaBox [100 100 300 300] >>",
        "<< /Type /Page /Parent 2 0 R >>",
        "<< >>",
        "<< >>"
      ],
      "/Root 1 0 R"
    );
  }

  it("keeps the MediaBox origin and the inherited /Rotate and /CropBox (R3-B6)", async () => {
    const editor = Pdf.Editor.read(geometry());
    editor.getPage(0).drawText("x", { x: 150, y: 150 });
    const own = latin1(await editor.toBuffer());
    expect(own).toMatch(/\/MediaBox \[100 100 300 300\]/);
    expect(own).toMatch(/\/CropBox \[120 120 280 280\]/);
    expect(own).toMatch(/\/Rotate 90/);

    const host = new Pdf.Builder();
    host.addPage();
    const copier = Pdf.Editor.read(await host.toBuffer());
    copier.copyPagesFrom(geometry());
    const copied = await copier.toBuffer();
    const pages = (await Pdf.extract(copied)).pages;
    expect(pages[1].rotation).toBe(90);
    expect(latin1(copied)).toMatch(/\/MediaBox \[100 100 300 300\]/);
  });
});

describe("copyPagesFrom keeps annotations and form fields (R3-C3)", () => {
  async function withField(name: string, link = false): Promise<Uint8Array> {
    const b = new Pdf.Builder();
    const page = b.addPage();
    page.addFormField({ type: "text", name, rect: [10, 10, 100, 30], value: "v" });
    page.addAnnotation({ type: "Text", rect: [10, 40, 30, 60], contents: "note" });
    if (link) {
      page.drawText("x", { x: 10, y: 100 });
    }
    return b.toBuffer();
  }

  it("copies the page's annotations and adds its fields to the form", async () => {
    const editor = Pdf.Editor.read(await withField("host"));
    editor.copyPagesFrom(await withField("guest"));
    const out = await editor.toBuffer();
    const result = await Pdf.extract(out);
    expect(result.formFields.map(f => f.name).sort()).toEqual(["guest", "host"]);
    expect(result.pages[1].annotations.map(a => a.contents)).toContain("note");
  });

  for (const incremental of MODES) {
    it(`makes fields added by the editor real fields (${modeName(incremental)})`, async () => {
      // Found fixing R3-C3: a field added to an existing page was a widget in no form, and a
      // field or annotation on an added page was dropped; an update dropped both.
      const b = new Pdf.Builder();
      b.addPage();
      const editor = Pdf.Editor.read(await b.toBuffer());
      editor.getPage(0).addFormField({ type: "text", name: "added", rect: [0, 0, 50, 20] });
      editor.getPage(0).addAnnotation({ type: "Text", rect: [0, 30, 20, 50], contents: "hi" });
      editor.addPage().addFormField({ type: "text", name: "onNew", rect: [0, 0, 50, 20] });
      const result = await Pdf.extract(await editor.toBuffer({ incremental }));
      expect(result.formFields.map(f => f.name).sort()).toEqual(["added", "onNew"]);
      expect(result.pages[0].annotations.map(a => a.contents)).toContain("hi");
    });
  }

  it("refuses a copied field whose name the document already has", async () => {
    // At the call: the copied page's field names are known then (editor-call-time-checks).
    const editor = Pdf.Editor.read(await withField("same"));
    const source = await withField("same");
    expect(() => editor.copyPagesFrom(source)).toThrow(PdfArgumentError);
  });
});

describe("a full rewrite copies streams with their filters", () => {
  it("keeps /DCTDecode on a JPEG image", async () => {
    // Found fixing R3-B8: the clone decoded every stream and dropped its filters, and a
    // DCT stream "decodes" to itself — so every JPEG lost /DCTDecode on any rewrite.
    const sof = [0xff, 0xc0, 0, 17, 8, 0, 2, 0, 3, 3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0];
    const jpeg = new Uint8Array([0xff, 0xd8, ...sof, 0xff, 0xd9]);
    const b = new Pdf.Builder();
    b.addPage().drawImage({ data: jpeg, format: "jpeg", x: 0, y: 0, width: 3, height: 2 });
    const editor = Pdf.Editor.read(await b.toBuffer());
    editor.getPage(0).drawText("x", { x: 1, y: 1 });
    const out = await editor.toBuffer();
    const image = /<<[^<>]*\/Subtype \/Image[^<>]*>>/.exec(latin1(out))?.[0] ?? "";
    expect(image).toContain("/Filter /DCTDecode");
    expect((await Pdf.extract(out)).pages[0].images[0].format).toBe("jpeg");
  });
});

describe("remaining boundary checks", () => {
  it("embedFonts refuses a non-object with PdfArgumentError (R3-C4)", () => {
    expect(() => new Pdf.Builder().embedFonts(42 as never)).toThrow(PdfArgumentError);
    expect(() => new Pdf.Builder().embedFonts(null as never)).toThrow(PdfArgumentError);
  });

  it("drawText checks bold and italic (R3-C5)", () => {
    const page = new Pdf.Builder().addPage();
    expect(() => page.drawText("a", { x: 1, y: 1, bold: "yes" as never })).toThrow(
      PdfArgumentError
    );
    expect(() => page.drawText("a", { x: 1, y: 1, italic: 1 as never })).toThrow(PdfArgumentError);
  });

  it("setMetadata reports the same errors in builder and editor (R3-S1)", () => {
    const editor = Pdf.Editor.read(withXmp(packet("<rdf:Description/>")));
    for (const bad of [undefined, null, [], 42]) {
      let builderError: PdfArgumentError | undefined;
      let editorError: PdfArgumentError | undefined;
      try {
        new Pdf.Builder().setMetadata(bad as never);
      } catch (err) {
        builderError = err as PdfArgumentError;
      }
      try {
        editor.setMetadata(bad as never);
      } catch (err) {
        editorError = err as PdfArgumentError;
      }
      expect(builderError).toBeInstanceOf(PdfArgumentError);
      // The same text after the call each names: `Pdf.Builder#setMetadata` / `Pdf.Editor#…`.
      expect(editorError?.message).toBe(
        builderError?.message.replace("Pdf.Builder#", "Pdf.Editor#")
      );
      expect(editorError?.field).toBe(builderError?.field);
    }
  });
});
