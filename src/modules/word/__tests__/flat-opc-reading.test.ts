import { Archive } from "@archive/index";
/**
 * Reading Flat OPC: a package element is one of the package namespace, whatever its prefix; a part's
 * XML is taken as the file holds it — comments, processing instructions and all — and declares the
 * namespaces it uses from around it; white space between the package's elements is only white space;
 * and what is not base64 is this module's error.
 */
import { describe, expect, it } from "vitest";

import { Convert, Document, DocxParseError } from "../index";

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const text = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes);

describe("Convert.flatOpcToParts", () => {
  it("reads a package written over several lines", () => {
    const parts = Convert.flatOpcToParts(
      `<pkg:package xmlns:pkg="${PKG}">\n  <pkg:part pkg:name="/a.xml">\n    <pkg:xmlData>\n      <a/>\n    </pkg:xmlData>\n  </pkg:part>\n</pkg:package>`
    );
    expect(text(parts.get("a.xml"))).toContain("<a/>");
  });

  it("knows the package by its namespace, not its prefix", () => {
    const parts = Convert.flatOpcToParts(
      `<p:package xmlns:p="${PKG}"><p:part p:name="/a.xml"><p:xmlData><a/></p:xmlData></p:part></p:package>`
    );
    expect(text(parts.get("a.xml"))).toContain("<a/>");
    expect(() => Convert.flatOpcToParts(`<pkg:package xmlns:pkg="urn:wrong"/>`)).toThrow(
      DocxParseError
    );
  });

  it("declares on a part the namespaces it takes from around it", async () => {
    const flat = `<pkg:package xmlns:pkg="${PKG}" xmlns:w="${W}"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document><w:body><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const xml = text(Convert.flatOpcToParts(flat).get("word/document.xml"));
    expect(xml).toContain(`<w:document xmlns:w="${W}">`);
    // The package namespace is not declared on a part that does not use it.
    expect(xml).not.toContain(PKG);
    const doc = await Document.read(new TextEncoder().encode(flat));
    expect(JSON.stringify(doc.body)).toContain('"ok"');
  });

  it("keeps a part's comments, processing instructions, CDATA and references", () => {
    const parts = Convert.flatOpcToParts(
      `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/a.xml"><pkg:xmlData><a><!--keep--><?keep data?><![CDATA[x<y]]>&amp;</a></pkg:xmlData></pkg:part></pkg:package>`
    );
    expect(text(parts.get("a.xml"))).toContain(
      "<a><!--keep--><?keep data?><![CDATA[x<y]]>&amp;</a>"
    );
  });

  it("decodes base64 split over lines, and refuses what is not base64 as its own error", () => {
    const flat = (data: string): string =>
      `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/b.bin"><pkg:binaryData>${data}</pkg:binaryData></pkg:part></pkg:package>`;
    expect(text(Convert.flatOpcToParts(flat("aGVs\n  bG8=")).get("b.bin"))).toBe("hello");
    for (const bad of ["!", "aGVsbG8", "aGVs=bG8"]) {
      expect(() => Convert.flatOpcToParts(flat(bad)), bad).toThrow(DocxParseError);
    }
  });

  const part = (data: string, outer = ""): string =>
    `<pkg:package xmlns:pkg="${PKG}"${outer}><pkg:part pkg:name="/a.xml">${data}</pkg:part></pkg:package>`;

  it("declares inherited namespaces after a root name that is not ASCII", () => {
    const xml = text(
      Convert.flatOpcToParts(
        part("<pkg:xmlData><根><q:x/></根></pkg:xmlData>", ' xmlns:q="urn:q"')
      ).get("a.xml")
    );
    expect(xml).toContain('<根 xmlns:q="urn:q"><q:x/></根>');
  });

  it("requires exactly one root element in xmlData", () => {
    for (const data of ["<a/><b/>", "", " ", "x<a/>", "<a/>x", "<![CDATA[x]]><a/>"]) {
      expect(
        () => Convert.flatOpcToParts(part(`<pkg:xmlData>${data}</pkg:xmlData>`)),
        data
      ).toThrow(DocxParseError);
    }
    expect(() => Convert.flatOpcToParts(part("<pkg:xmlData/>"))).toThrow(DocxParseError);
    expect(
      text(
        Convert.flatOpcToParts(part("<pkg:xmlData><!--c--> <a/> <?p?></pkg:xmlData>")).get("a.xml")
      )
    ).toContain("<a/>");
  });

  it("refuses xmlData under a DOCTYPE, whose declarations the part would lose", () => {
    const doc = `<!DOCTYPE pkg:package [<!ATTLIST a v CDATA "VALUE">]>${part("<pkg:xmlData><a/></pkg:xmlData>")}`;
    expect(() => Convert.flatOpcToParts(doc)).toThrow(/DOCTYPE/);
    const entity = `<!DOCTYPE pkg:package [<!ENTITY e "x">]>${part("<pkg:xmlData><a/></pkg:xmlData>")}`;
    expect(() => Convert.flatOpcToParts(entity)).toThrow(/DOCTYPE/);
  });

  it("reads xmlData under a DOCTYPE that declares nothing a part would lose", () => {
    for (const doctype of [
      "<!DOCTYPE pkg:package>",
      '<!DOCTYPE pkg:package SYSTEM "package.dtd">',
      '<!DOCTYPE pkg:package [<!-- <!ATTLIST a v CDATA "x"> --><?pi <!ENTITY?>]>'
    ]) {
      const parts = Convert.flatOpcToParts(`${doctype}${part("<pkg:xmlData><a/></pkg:xmlData>")}`);
      expect(new TextDecoder().decode(parts.get("a.xml"))).toMatch(/<a\/>$/);
    }
  });

  it("refuses a second data container and elements inside binaryData", () => {
    expect(() =>
      Convert.flatOpcToParts(
        part("<pkg:xmlData><a/></pkg:xmlData><pkg:xmlData><BAD/></pkg:xmlData>")
      )
    ).toThrow(DocxParseError);
    expect(() =>
      Convert.flatOpcToParts(
        part("<pkg:binaryData>aGVs</pkg:binaryData><pkg:xmlData><a/></pkg:xmlData>")
      )
    ).toThrow(DocxParseError);
    expect(() =>
      Convert.flatOpcToParts(part("<pkg:binaryData>aGVs<x>bG8=</x></pkg:binaryData>"))
    ).toThrow(DocxParseError);
  });
});

describe("Convert.isFlatOpc", () => {
  it("is false for a ZIP, whatever text it stores", async () => {
    const zip = Archive.createZip({ level: 0 });
    zip.add("note.txt", new TextEncoder().encode(`<pkg:package xmlns:pkg="${PKG}"/>`));
    const bytes = await zip.toBuffer();
    expect(Convert.isFlatOpc(bytes)).toBe(false);
  });

  it("finds the root past a long prologue, and checks its namespace", () => {
    const body = `<pkg:package xmlns:pkg="${PKG}"/>`;
    const long = `\uFEFF<?xml version="1.0"?>\n<!--${"x".repeat(2000)}--><?pi a?><!DOCTYPE p [<!ENTITY e ">">]>${body}`;
    expect(Convert.isFlatOpc(long)).toBe(true);
    expect(Convert.isFlatOpc(new TextEncoder().encode(long))).toBe(true);
    expect(Convert.isFlatOpc(`<p:package xmlns:p="${PKG}"/>`)).toBe(true);
    expect(Convert.isFlatOpc(`<root><!-- ${PKG} --></root>`)).toBe(false);
    expect(Convert.isFlatOpc(`<pkg:package xmlns:pkg="urn:x"/>`)).toBe(false);
    expect(Convert.isFlatOpc(`<!-- <pkg:package xmlns:pkg="${PKG}"> --><a/>`)).toBe(false);
  });
});

describe("Convert.partsToFlatOpc", () => {
  const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
  const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

  it("reads [Content_Types].xml by namespace, not by prefix", () => {
    const entries = new Map([
      [
        "[Content_Types].xml",
        enc(
          `  <?xml version="1.0"?><ct:Types xmlns:ct="${CT}"><ct:Override PartName="/x.dat" ContentType="text/x-mine"/></ct:Types>`
        )
      ],
      ["x.dat", enc("aa")]
    ]);
    expect(Convert.partsToFlatOpc(entries)).toContain('pkg:contentType="text/x-mine"');
  });

  it("infers a type for an extension named like an Object member", () => {
    const entries = new Map([
      ["a.constructor", enc("aa")],
      ["b.__proto__", enc("bb")]
    ]);
    const flat = Convert.partsToFlatOpc(entries);
    expect(flat.match(/application\/octet-stream/g)).toHaveLength(2);
  });
});
