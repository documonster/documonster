/**
 * DOCX Module - Flat OPC Tests
 */

import { XmlParseError } from "@xml/errors";
import { describe, it, expect } from "vitest";

import { DocxLimitExceededError, DocxParseError } from "../errors";
import { Document, Convert, Io, Query } from "../index";

describe("Flat OPC", () => {
  it("reports text that is not XML as a DocxParseError, the XML error its cause", () => {
    let error: unknown;
    try {
      Convert.flatOpcToParts("zzz");
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(DocxParseError);
    expect((error as Error).cause).toBeInstanceOf(XmlParseError);
  });

  describe("isFlatOpc", () => {
    it("returns true for valid Flat OPC XML", () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>`;
      expect(Convert.isFlatOpc(xml)).toBe(true);
    });

    it("returns true for Uint8Array with Flat OPC content", () => {
      const xml = `<?xml version="1.0"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>`;
      const bytes = new TextEncoder().encode(xml);
      expect(Convert.isFlatOpc(bytes)).toBe(true);
    });

    it("returns false for ZIP file (starts with PK)", () => {
      const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]);
      expect(Convert.isFlatOpc(zip)).toBe(false);
    });

    it("returns false for plain text", () => {
      expect(Convert.isFlatOpc("Hello world, this is plain text.")).toBe(false);
    });
  });

  describe("toFlatOpc + parseFlatOpc round-trip", () => {
    it("round-trips entries correctly", () => {
      const entries = new Map<string, Uint8Array>();
      const encoder = new TextEncoder();

      entries.set(
        "[Content_Types].xml",
        encoder.encode(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
        )
      );
      entries.set(
        "word/document.xml",
        encoder.encode(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>`
        )
      );

      const flatOpc = Convert.partsToFlatOpc(entries);
      expect(typeof flatOpc).toBe("string");
      expect(flatOpc).toContain("pkg:package");

      const parsed = Convert.flatOpcToParts(flatOpc);
      expect(parsed.has("word/document.xml")).toBe(true);

      const docXml = new TextDecoder().decode(parsed.get("word/document.xml")!);
      expect(docXml).toContain("Hello");
    });
  });

  describe("Io.toFlatOpc", () => {
    it("converts a full document to Flat OPC format", async () => {
      const doc = Document.create();
      Document.addParagraph(doc, "Test content");
      const built = doc;

      const flatOpc = await Io.toFlatOpc(built);
      expect(typeof flatOpc).toBe("string");
      expect(flatOpc).toContain("pkg:package");
      expect(flatOpc).toContain("word/document.xml");
    });
  });
});

describe("Io.read of a Flat OPC document", () => {
  it("reads what Io.toFlatOpc writes, as text bytes", async () => {
    const doc = Document.create();
    Document.addHeading(doc, "Flat", 1);
    Document.addParagraph(doc, "body text");
    const flat = await Io.toFlatOpc(doc);
    const read = await Io.read(new TextEncoder().encode(flat));
    expect(Query.getText(read)).toContain("body text");
    // …and the same document as the package bytes give.
    const fromZip = await Io.read(await Io.toBuffer(doc));
    expect(read.body).toEqual(fromZip.body);
  });

  it("applies the security policy's limits to the parts", async () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    const bytes = new TextEncoder().encode(await Io.toFlatOpc(doc));
    await expect(Io.read(bytes, { securityPolicy: { maxPartCount: 2 } })).rejects.toThrow(
      DocxLimitExceededError
    );
  });
});
