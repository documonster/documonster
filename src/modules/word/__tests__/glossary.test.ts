/**
 * DOCX Module - Glossary / Building Blocks Tests
 */

import { describe, it, expect } from "vitest";

import { Document, Build, Glossary, Io } from "../index";
import type { DocxDocument } from "../types";

const utf8 = new TextDecoder();

describe("Glossary / Building Blocks", () => {
  describe("createBuildingBlock", () => {
    it("creates a building block with name, gallery, and content", () => {
      const block = Glossary.createBlock("Greeting", "autoText", [
        Build.textParagraph("Dear Sir,")
      ]);
      expect(block.name).toBe("Greeting");
      expect(block.gallery).toBe("autoText");
      expect(block.content).toHaveLength(1);
      expect(block.category).toBe("General");
      expect(block.guid).toBeDefined();
    });

    it("supports custom category and description", () => {
      const block = Glossary.createBlock(
        "Logo",
        "quickParts",
        [Build.textParagraph("img placeholder")],
        {
          category: "Branding",
          description: "Company logo block"
        }
      );
      expect(block.category).toBe("Branding");
      expect(block.description).toBe("Company logo block");
    });
  });

  describe("createGlossaryDocument", () => {
    it("creates a glossary with blocks", () => {
      const b1 = Glossary.createBlock("A", "autoText", [Build.textParagraph("a")]);
      const b2 = Glossary.createBlock("B", "quickParts", [Build.textParagraph("b")]);
      const glossary = Glossary.createDocument([b1, b2]);
      expect(glossary.blocks).toHaveLength(2);
    });
  });

  describe("findBuildingBlock", () => {
    it("finds by name", () => {
      const b1 = Glossary.createBlock("Greeting", "autoText", [Build.textParagraph("hi")]);
      const b2 = Glossary.createBlock("Footer", "footers", [Build.textParagraph("foot")]);
      const glossary = Glossary.createDocument([b1, b2]);

      const found = Glossary.findBlock(glossary, "Greeting");
      expect(found).toBeDefined();
      expect(found!.name).toBe("Greeting");
    });

    it("finds by name and gallery", () => {
      const b1 = Glossary.createBlock("Block", "autoText", [Build.textParagraph("at")]);
      const b2 = Glossary.createBlock("Block", "quickParts", [Build.textParagraph("qp")]);
      const glossary = Glossary.createDocument([b1, b2]);

      const found = Glossary.findBlock(glossary, "Block", "quickParts");
      expect(found).toBeDefined();
      expect(found!.gallery).toBe("quickParts");
    });

    it("returns undefined when not found", () => {
      const glossary = Glossary.createDocument([]);
      expect(Glossary.findBlock(glossary, "nonexistent")).toBeUndefined();
    });
  });

  describe("listBuildingBlocks", () => {
    it("filters by gallery", () => {
      const b1 = Glossary.createBlock("A", "autoText", [Build.textParagraph("a")]);
      const b2 = Glossary.createBlock("B", "autoText", [Build.textParagraph("b")]);
      const b3 = Glossary.createBlock("C", "quickParts", [Build.textParagraph("c")]);
      const glossary = Glossary.createDocument([b1, b2, b3]);

      const autoTexts = Glossary.listBlocks(glossary, "autoText");
      expect(autoTexts).toHaveLength(2);
    });
  });

  describe("getAutoTextEntries", () => {
    it("returns only autoText entries", () => {
      const b1 = Glossary.createBlock("A", "autoText", [Build.textParagraph("a")]);
      const b2 = Glossary.createBlock("B", "quickParts", [Build.textParagraph("b")]);
      const glossary = Glossary.createDocument([b1, b2]);

      const entries = Glossary.autoTextEntries(glossary);
      expect(entries).toHaveLength(1);
      expect(entries[0]!.gallery).toBe("autoText");
    });
  });

  describe("getQuickParts", () => {
    it("returns only quickParts entries", () => {
      const b1 = Glossary.createBlock("A", "autoText", [Build.textParagraph("a")]);
      const b2 = Glossary.createBlock("B", "quickParts", [Build.textParagraph("b")]);
      const b3 = Glossary.createBlock("C", "quickParts", [Build.textParagraph("c")]);
      const glossary = Glossary.createDocument([b1, b2, b3]);

      const parts = Glossary.quickParts(glossary);
      expect(parts).toHaveLength(2);
    });
  });

  describe("serialisation via doc.glossary", () => {
    function docWithGlossary(): DocxDocument {
      const d = Document.create();
      Document.useDefaultStyles(d);
      Document.addParagraph(d, "Body.");
      const glossary = Glossary.createDocument([
        Glossary.createBlock("Greeting", "autoText", [Build.textParagraph("Dear Sir/Madam,")], {
          category: "Letters",
          description: "Opening"
        }),
        Glossary.createBlock("Disclaimer", "quickParts", [
          Build.paragraph([Build.text("Provided as-is.")])
        ])
      ]);
      return { ...Document.build(d), glossary };
    }

    it("emits word/glossary/document.xml with relationship and content type", async () => {
      const buf = await Io.toBuffer(docWithGlossary());

      const glossaryXml = await Io.readDocxPart(buf, "word/glossary/document.xml");
      expect(glossaryXml).toBeDefined();
      const xml = utf8.decode(glossaryXml!);
      expect(xml).toContain("<w:glossaryDocument");
      expect(xml).toContain('<w:gallery w:val="autoTxt"/>');
      expect(xml).toContain('<w:gallery w:val="custQuickParts"/>');
      expect(xml).toContain("Dear Sir/Madam,");
      expect(xml).toContain("Provided as-is.");

      const rels = utf8.decode((await Io.readDocxPart(buf, "word/_rels/document.xml.rels"))!);
      expect(rels).toContain("glossaryDocument");
      expect(rels).toContain("glossary/document.xml");

      const ct = utf8.decode((await Io.readDocxPart(buf, "[Content_Types].xml"))!);
      expect(ct).toContain("/word/glossary/document.xml");
      expect(ct).toContain("document.glossary+xml");
    });

    it("emits the self-contained glossary sub-document (companion parts + rels)", async () => {
      // Word treats word/glossary/document.xml as its own sub-document and
      // discards the whole glossary if its styles/settings/webSettings/
      // fontTable companions (referenced from its own .rels) are missing.
      const buf = await Io.toBuffer(docWithGlossary());
      for (const p of [
        "word/glossary/styles.xml",
        "word/glossary/settings.xml",
        "word/glossary/webSettings.xml",
        "word/glossary/fontTable.xml",
        "word/glossary/_rels/document.xml.rels"
      ]) {
        expect(await Io.readDocxPart(buf, p), p).toBeDefined();
      }
      const grels = utf8.decode(
        (await Io.readDocxPart(buf, "word/glossary/_rels/document.xml.rels"))!
      );
      expect(grels).toContain("styles.xml");
      expect(grels).toContain("settings.xml");
      expect(grels).toContain("webSettings.xml");
      expect(grels).toContain("fontTable.xml");

      const ct = utf8.decode((await Io.readDocxPart(buf, "[Content_Types].xml"))!);
      expect(ct).toContain("/word/glossary/styles.xml");
      expect(ct).toContain("/word/glossary/settings.xml");
      expect(ct).toContain("/word/glossary/webSettings.xml");
      expect(ct).toContain("/word/glossary/fontTable.xml");
    });

    it("emits w:docPartPr children in the schema-required order", async () => {
      // CT_DocPartPr order (ECMA-376 §17.12.1): name → style → category →
      // types → behaviors → description → guid. Word rejects the package if
      // these are out of order (e.g. description after guid).
      const buf = await Io.toBuffer(docWithGlossary());
      const xml = utf8.decode((await Io.readDocxPart(buf, "word/glossary/document.xml"))!);
      // Inspect the FormalGreeting docPart, which has both description + guid.
      const pr = /<w:docPartPr>(.*?)<\/w:docPartPr>/s.exec(xml)![1];
      // Strip the nested <w:category> so only direct children remain.
      const flat = pr.replace(/<w:category>.*?<\/w:category>/s, "<w:category/>");
      const order = [
        ...flat.matchAll(/<w:(name|style|category|types|behaviors|description|guid)\b/g)
      ].map(m => m[1]);
      const schema = ["name", "style", "category", "types", "behaviors", "description", "guid"];
      // The emitted order must be a subsequence of the schema order.
      let i = 0;
      for (const tag of order) {
        const at = schema.indexOf(tag, i);
        expect(at).toBeGreaterThanOrEqual(0);
        i = at;
      }
      // And specifically: description must precede guid.
      expect(order.indexOf("description")).toBeLessThan(order.indexOf("guid"));
    });

    it("only emits ST_DocPartGallery values from the official enum", async () => {
      // Word silently discards the whole glossary if any <w:gallery w:val>
      // is not an exact ST_DocPartGallery enum member (ECMA-376 §17.18.23) —
      // e.g. "quickParts" is NOT valid (Quick Parts map to "custQuickParts").
      const LEGAL_GALLERIES = new Set([
        "placeholder",
        "any",
        "default",
        "docParts",
        "coverPg",
        "eq",
        "ftrs",
        "hdrs",
        "pgNum",
        "tbls",
        "watermarks",
        "autoTxt",
        "txtBox",
        "pgNumT",
        "pgNumB",
        "pgNumMargins",
        "tblOfContents",
        "bib",
        "custQuickParts",
        "custCoverPg",
        "custEq",
        "custFtrs",
        "custHdrs",
        "custPgNum",
        "custTbls",
        "custWatermarks",
        "custAutoTxt",
        "custTxtBox",
        "custPgNumT",
        "custPgNumB",
        "custPgNumMargins",
        "custTblOfContents",
        "custBib",
        "custom1",
        "custom2",
        "custom3",
        "custom4",
        "custom5"
      ]);
      const allGalleries = [
        "autoText",
        "quickParts",
        "coverPages",
        "tableOfContents",
        "headers",
        "footers",
        "pageNumbers",
        "tables",
        "textBoxes",
        "watermarks",
        "equations",
        "bibliographies",
        "custom1",
        "custom2",
        "custom3",
        "custom4",
        "custom5"
      ] as const;
      const glossary = Glossary.createDocument(
        allGalleries.map((g, i) => Glossary.createBlock(`B${i}`, g, [Build.textParagraph("x")]))
      );
      const d = Document.create();
      Document.useDefaultStyles(d);
      Document.addParagraph(d, "Body.");
      const buf = await Io.toBuffer({ ...Document.build(d), glossary });
      const xml = utf8.decode((await Io.readDocxPart(buf, "word/glossary/document.xml"))!);
      const vals = [...xml.matchAll(/<w:gallery w:val="([^"]+)"/g)].map(m => m[1]);
      expect(vals.length).toBe(allGalleries.length);
      for (const v of vals) {
        expect(LEGAL_GALLERIES.has(v), `illegal ST_DocPartGallery value: ${v}`).toBe(true);
      }
    });

    it("round-trips: read preserves the glossary and re-write re-emits it", async () => {
      const buf = await Io.toBuffer(docWithGlossary());
      const reread = await Io.read(buf);
      expect(reread.glossary).toBeDefined();
      expect(reread.glossary!.rawXml).toContain("<w:glossaryDocument");
      // The glossary document.xml is consumed into doc.glossary (not left as an
      // opaque part); its companion parts round-trip via opaqueParts.
      expect((reread.opaqueParts ?? []).some(p => p.path === "word/glossary/document.xml")).toBe(
        false
      );

      // Re-write must still contain the part + relationship (not dropped).
      const buf2 = await Io.toBuffer(reread);
      const glossaryXml2 = await Io.readDocxPart(buf2, "word/glossary/document.xml");
      expect(glossaryXml2).toBeDefined();
      const rels2 = utf8.decode((await Io.readDocxPart(buf2, "word/_rels/document.xml.rels"))!);
      expect(rels2).toContain("glossaryDocument");
      // Companion parts survive the round-trip too.
      expect(await Io.readDocxPart(buf2, "word/glossary/styles.xml")).toBeDefined();
    });

    it("omits the glossary part when there are no blocks", async () => {
      const d = Document.create();
      Document.useDefaultStyles(d);
      Document.addParagraph(d, "Body.");
      const buf = await Io.toBuffer({
        ...Document.build(d),
        glossary: Glossary.createDocument([])
      });
      expect(await Io.readDocxPart(buf, "word/glossary/document.xml")).toBeUndefined();
    });
  });
});
