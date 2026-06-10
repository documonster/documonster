/**
 * DOCX Module - Glossary / Building Blocks Tests
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  createBuildingBlock,
  createGlossaryDocument,
  findBuildingBlock,
  listBuildingBlocks,
  getAutoTextEntries,
  getQuickParts,
  paragraph,
  text,
  textParagraph,
  toBuffer,
  readDocx,
  readDocxPart
} from "../index";
import type { DocxDocument } from "../types";

const utf8 = new TextDecoder();

describe("Glossary / Building Blocks", () => {
  describe("createBuildingBlock", () => {
    it("creates a building block with name, gallery, and content", () => {
      const block = createBuildingBlock("Greeting", "autoText", [textParagraph("Dear Sir,")]);
      expect(block.name).toBe("Greeting");
      expect(block.gallery).toBe("autoText");
      expect(block.content).toHaveLength(1);
      expect(block.category).toBe("General");
      expect(block.guid).toBeDefined();
    });

    it("supports custom category and description", () => {
      const block = createBuildingBlock("Logo", "quickParts", [textParagraph("img placeholder")], {
        category: "Branding",
        description: "Company logo block"
      });
      expect(block.category).toBe("Branding");
      expect(block.description).toBe("Company logo block");
    });
  });

  describe("createGlossaryDocument", () => {
    it("creates a glossary with blocks", () => {
      const b1 = createBuildingBlock("A", "autoText", [textParagraph("a")]);
      const b2 = createBuildingBlock("B", "quickParts", [textParagraph("b")]);
      const glossary = createGlossaryDocument([b1, b2]);
      expect(glossary.blocks).toHaveLength(2);
    });
  });

  describe("findBuildingBlock", () => {
    it("finds by name", () => {
      const b1 = createBuildingBlock("Greeting", "autoText", [textParagraph("hi")]);
      const b2 = createBuildingBlock("Footer", "footers", [textParagraph("foot")]);
      const glossary = createGlossaryDocument([b1, b2]);

      const found = findBuildingBlock(glossary, "Greeting");
      expect(found).toBeDefined();
      expect(found!.name).toBe("Greeting");
    });

    it("finds by name and gallery", () => {
      const b1 = createBuildingBlock("Block", "autoText", [textParagraph("at")]);
      const b2 = createBuildingBlock("Block", "quickParts", [textParagraph("qp")]);
      const glossary = createGlossaryDocument([b1, b2]);

      const found = findBuildingBlock(glossary, "Block", "quickParts");
      expect(found).toBeDefined();
      expect(found!.gallery).toBe("quickParts");
    });

    it("returns undefined when not found", () => {
      const glossary = createGlossaryDocument([]);
      expect(findBuildingBlock(glossary, "nonexistent")).toBeUndefined();
    });
  });

  describe("listBuildingBlocks", () => {
    it("filters by gallery", () => {
      const b1 = createBuildingBlock("A", "autoText", [textParagraph("a")]);
      const b2 = createBuildingBlock("B", "autoText", [textParagraph("b")]);
      const b3 = createBuildingBlock("C", "quickParts", [textParagraph("c")]);
      const glossary = createGlossaryDocument([b1, b2, b3]);

      const autoTexts = listBuildingBlocks(glossary, "autoText");
      expect(autoTexts).toHaveLength(2);
    });
  });

  describe("getAutoTextEntries", () => {
    it("returns only autoText entries", () => {
      const b1 = createBuildingBlock("A", "autoText", [textParagraph("a")]);
      const b2 = createBuildingBlock("B", "quickParts", [textParagraph("b")]);
      const glossary = createGlossaryDocument([b1, b2]);

      const entries = getAutoTextEntries(glossary);
      expect(entries).toHaveLength(1);
      expect(entries[0]!.gallery).toBe("autoText");
    });
  });

  describe("getQuickParts", () => {
    it("returns only quickParts entries", () => {
      const b1 = createBuildingBlock("A", "autoText", [textParagraph("a")]);
      const b2 = createBuildingBlock("B", "quickParts", [textParagraph("b")]);
      const b3 = createBuildingBlock("C", "quickParts", [textParagraph("c")]);
      const glossary = createGlossaryDocument([b1, b2, b3]);

      const parts = getQuickParts(glossary);
      expect(parts).toHaveLength(2);
    });
  });

  describe("serialisation via doc.glossary", () => {
    function docWithGlossary(): DocxDocument {
      const d = Document.create();
      Document.useDefaultStyles(d);
      Document.addParagraph(d, "Body.");
      const glossary = createGlossaryDocument([
        createBuildingBlock("Greeting", "autoText", [textParagraph("Dear Sir/Madam,")], {
          category: "Letters",
          description: "Opening"
        }),
        createBuildingBlock("Disclaimer", "quickParts", [paragraph([text("Provided as-is.")])])
      ]);
      return { ...Document.build(d), glossary };
    }

    it("emits word/glossary/document.xml with relationship and content type", async () => {
      const buf = await toBuffer(docWithGlossary());

      const glossaryXml = await readDocxPart(buf, "word/glossary/document.xml");
      expect(glossaryXml).toBeDefined();
      const xml = utf8.decode(glossaryXml!);
      expect(xml).toContain("<w:glossaryDocument");
      expect(xml).toContain('<w:gallery w:val="autoTxt"/>');
      expect(xml).toContain('<w:gallery w:val="quickParts"/>');
      expect(xml).toContain("Dear Sir/Madam,");
      expect(xml).toContain("Provided as-is.");

      const rels = utf8.decode((await readDocxPart(buf, "word/_rels/document.xml.rels"))!);
      expect(rels).toContain("glossaryDocument");
      expect(rels).toContain("glossary/document.xml");

      const ct = utf8.decode((await readDocxPart(buf, "[Content_Types].xml"))!);
      expect(ct).toContain("/word/glossary/document.xml");
      expect(ct).toContain("document.glossary+xml");
    });

    it("round-trips: read preserves the glossary and re-write re-emits it", async () => {
      const buf = await toBuffer(docWithGlossary());
      const reread = await readDocx(buf);
      expect(reread.glossary).toBeDefined();
      expect(reread.glossary!.rawXml).toContain("<w:glossaryDocument");
      // No phantom opaque part for the (consumed) glossary.
      expect((reread.opaqueParts ?? []).some(p => p.path.startsWith("word/glossary/"))).toBe(false);

      // Re-write must still contain the part + relationship (not dropped).
      const buf2 = await toBuffer(reread);
      const glossaryXml2 = await readDocxPart(buf2, "word/glossary/document.xml");
      expect(glossaryXml2).toBeDefined();
      const rels2 = utf8.decode((await readDocxPart(buf2, "word/_rels/document.xml.rels"))!);
      expect(rels2).toContain("glossaryDocument");
    });

    it("omits the glossary part when there are no blocks", async () => {
      const d = Document.create();
      Document.useDefaultStyles(d);
      Document.addParagraph(d, "Body.");
      const buf = await toBuffer({ ...Document.build(d), glossary: createGlossaryDocument([]) });
      expect(await readDocxPart(buf, "word/glossary/document.xml")).toBeUndefined();
    });
  });
});
