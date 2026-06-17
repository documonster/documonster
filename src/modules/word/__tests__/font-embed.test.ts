/**
 * DOCX Module - Font Embedding Tests
 */

import { describe, it, expect } from "vitest";

import { deobfuscateFont } from "../font/font-obfuscation";
import { Document, Font } from "../index";
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
      const original = createFakeTtfData();
      const result = Font.embed({
        name: "TestFont",
        data: original,
        style: "regular",
        obfuscate: true
      });

      expect(result.fontDef.name).toBe("TestFont");
      expect(result.embeddedFont.fileName).toMatch(/\.odttf$/);
      expect(result.embeddedFont.fontKey).toBeDefined();
      expect(result.embeddedFont.data.length).toBe(original.length);

      // The previous assertion (`obfuscated[10] !== original[10]`) was
      // ~0.4% flaky: ODTTF XORs the first 32 bytes with a 16-byte
      // GUID-derived key, and whenever the GUID byte at position 10
      // happened to be 0x00 the byte-10 assertion would fail.
      //
      // Verify the actual contract instead: ODTTF obfuscation is a
      // reversible transform whose inverse is `deobfuscateFont` keyed
      // by the returned `fontKey`. Round-tripping must yield the
      // original bytes exactly. This:
      //   - Catches algorithmic mistakes (wrong byte range, wrong
      //     GUID reorder, wrong key reuse) that a single-byte or
      //     range-inequality assertion would miss.
      //   - Cannot be flaky: the assertion does not depend on any
      //     specific GUID byte being non-zero.
      const fontKey = result.embeddedFont.fontKey!;
      const roundTrip = deobfuscateFont(result.embeddedFont.data, fontKey);
      expect(roundTrip).toEqual(original);

      // Spec: only the first 32 bytes are obfuscated. Bytes beyond
      // that must already match the input without applying the key.
      expect(result.embeddedFont.data.subarray(32)).toEqual(original.subarray(32));
    });
  });

  describe("embedFont without obfuscation", () => {
    it("produces unobfuscated font output", () => {
      const result = Font.embed({
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
      const results = Font.embedFamily("FamilyFont", {
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
      const result = Font.embed({
        name: "NewFont",
        data: createFakeTtfData(),
        style: "regular"
      });

      const updated = Font.addEmbedded(doc, [result]);
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

      const result = Font.embed({
        name: "NewFont",
        data: createFakeTtfData(),
        style: "italic"
      });

      const updated = Font.addEmbedded(doc, [result]);
      expect(updated.fonts!.length).toBe(2);
    });
  });
});
