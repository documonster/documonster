/**
 * DOCX Module - Font Embedding Tests
 */

import { describe, it, expect } from "vitest";

import { embedFont, embedFontFamily, addEmbeddedFonts, Document } from "../index";
import type { DocxDocument } from "../types";

// A minimal fake TTF font (starts with 0x00 0x01 0x00 0x00 magic)
function createFakeTtfData(size = 100): Uint8Array {
  const data = new Uint8Array(size);
  // TTF magic number
  data[0] = 0x00;
  data[1] = 0x01;
  data[2] = 0x00;
  data[3] = 0x00;
  // Fill some dummy data
  for (let i = 4; i < size; i++) {
    data[i] = i % 256;
  }
  return data;
}

describe("Font embedding", () => {
  describe("embedFont with obfuscation", () => {
    it("produces ODTTF output", () => {
      const result = embedFont({
        name: "TestFont",
        data: createFakeTtfData(),
        style: "regular",
        obfuscate: true
      });

      expect(result.fontDef.name).toBe("TestFont");
      expect(result.embeddedFont.fileName).toMatch(/\.odttf$/);
      expect(result.embeddedFont.fontKey).toBeDefined();
      expect(result.embeddedFont.data.length).toBeGreaterThan(0);
      // Obfuscated data should differ from original at the beginning
      expect(result.embeddedFont.data[10]).not.toBe(createFakeTtfData()[10]);
    });
  });

  describe("embedFont without obfuscation", () => {
    it("produces unobfuscated font output", () => {
      const result = embedFont({
        name: "PlainFont",
        data: createFakeTtfData(),
        style: "bold",
        obfuscate: false
      });

      expect(result.fontDef.name).toBe("PlainFont");
      expect(result.embeddedFont.fileName).toContain("PlainFont");
      expect(result.embeddedFont.fileName).toContain("bold");
      expect(result.embeddedFont.fontKey).toBeUndefined();
      // Unobfuscated should be identical to input
      const input = createFakeTtfData();
      expect(result.embeddedFont.data).toEqual(input);
    });
  });

  describe("embedFontFamily with multiple variants", () => {
    it("embeds regular and bold variants", () => {
      const results = embedFontFamily("FamilyFont", {
        regular: createFakeTtfData(80),
        bold: createFakeTtfData(90)
      });

      expect(results).toHaveLength(2);
      expect(results[0]!.fontDef.name).toBe("FamilyFont");
      expect(results[1]!.fontDef.name).toBe("FamilyFont");
      // One should be regular embed, other bold
      const hasRegular = results.some(r => r.fontDef.embedRegular !== undefined);
      const hasBold = results.some(r => r.fontDef.embedBold !== undefined);
      expect(hasRegular).toBe(true);
      expect(hasBold).toBe(true);
    });
  });

  describe("addEmbeddedFonts", () => {
    it("adds fonts to document model", () => {
      const doc = Document.build(Document.create());
      const result = embedFont({
        name: "NewFont",
        data: createFakeTtfData(),
        style: "regular"
      });

      const updated = addEmbeddedFonts(doc, [result]);
      expect(updated.fonts).toBeDefined();
      expect(updated.fonts!.some(f => f.name === "NewFont")).toBe(true);
      expect(updated.embeddedFonts).toBeDefined();
      expect(updated.embeddedFonts!.length).toBe(1);
    });

    it("merges with existing fonts", () => {
      const doc: DocxDocument = {
        body: [],
        fonts: [{ name: "ExistingFont", family: "roman", pitch: "variable" }],
        sectionProperties: {}
      } as any;

      const result = embedFont({
        name: "NewFont",
        data: createFakeTtfData(),
        style: "italic"
      });

      const updated = addEmbeddedFonts(doc, [result]);
      expect(updated.fonts!.length).toBe(2);
    });
  });
});
