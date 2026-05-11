/**
 * DOCX Module - Glossary / Building Blocks Tests
 */

import { describe, it, expect } from "vitest";

import {
  createBuildingBlock,
  createGlossaryDocument,
  findBuildingBlock,
  listBuildingBlocks,
  getAutoTextEntries,
  getQuickParts,
  textParagraph
} from "../index";

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
});
