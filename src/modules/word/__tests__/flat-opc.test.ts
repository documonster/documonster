/**
 * DOCX Module - Flat OPC Tests
 */

import { describe, it, expect } from "vitest";

import { Document, Convert, Io } from "../index";

describe("Flat OPC", () => {
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

      const flatOpc = Convert.toFlatOpc(entries);
      expect(typeof flatOpc).toBe("string");
      expect(flatOpc).toContain("pkg:package");

      const parsed = Convert.parseFlatOpc(flatOpc);
      expect(parsed.has("word/document.xml")).toBe(true);

      const docXml = new TextDecoder().decode(parsed.get("word/document.xml")!);
      expect(docXml).toContain("Hello");
    });
  });

  describe("toFlatOpcFromDoc", () => {
    it("converts a full document to Flat OPC format", async () => {
      const doc = Document.create();
      Document.addParagraph(doc, "Test content");
      const built = Document.build(doc);

      const flatOpc = await Io.toFlatOpcFromDoc(built);
      expect(typeof flatOpc).toBe("string");
      expect(flatOpc).toContain("pkg:package");
      expect(flatOpc).toContain("word/document.xml");
    });
  });
});
