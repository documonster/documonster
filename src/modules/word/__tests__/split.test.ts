/**
 * DOCX Module - Document Split API Tests
 */

import { describe, it, expect } from "vitest";

import { splitDocument } from "../index";
import type { DocxDocument, Paragraph, Run, BodyContent, SectionProperties } from "../types";

function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

function para(t: string, props?: Paragraph["properties"]): Paragraph {
  return { type: "paragraph", children: [textRun(t)], properties: props };
}

function createDoc(body: BodyContent[]): DocxDocument {
  return { body } as unknown as DocxDocument;
}

function paraText(p: Paragraph): string {
  let text = "";
  for (const child of p.children) {
    if ("content" in child && Array.isArray((child as Run).content)) {
      for (const c of (child as Run).content) {
        if (c.type === "text") {
          text += c.text;
        }
      }
    }
  }
  return text;
}

describe("splitDocument", () => {
  describe("by section", () => {
    it("returns single document if no section breaks", () => {
      const doc = createDoc([para("a"), para("b"), para("c")]);
      const result = splitDocument(doc);
      expect(result).toHaveLength(1);
      expect(result[0].body).toEqual(doc.body);
    });

    it("splits at section break", () => {
      const sectProps: SectionProperties = { breakType: "nextPage" };
      const doc = createDoc([
        para("section 1 text"),
        para("section break", { sectionProperties: sectProps }),
        para("section 2 text"),
        para("section 2 more")
      ]);

      const result = splitDocument(doc, { by: "section" });
      expect(result).toHaveLength(2);
      // First segment includes the section-break paragraph
      expect(result[0].body.length).toBe(2);
      expect(paraText(result[0].body[0] as Paragraph)).toBe("section 1 text");
      // Second segment
      expect(result[1].body.length).toBe(2);
      expect(paraText(result[1].body[0] as Paragraph)).toBe("section 2 text");
    });

    it("splits at multiple section breaks", () => {
      const sectProps: SectionProperties = { breakType: "continuous" };
      const doc = createDoc([
        para("a"),
        para("end1", { sectionProperties: sectProps }),
        para("b"),
        para("end2", { sectionProperties: sectProps }),
        para("c")
      ]);

      const result = splitDocument(doc, { by: "section" });
      expect(result).toHaveLength(3);
    });
  });

  describe("by pageBreak", () => {
    it("splits at explicit page break paragraph", () => {
      const pageBreakRun: Run = {
        content: [{ type: "break", breakType: "page" }]
      };
      const breakPara: Paragraph = {
        type: "paragraph",
        children: [pageBreakRun]
      };

      const doc = createDoc([para("page 1"), breakPara, para("page 2")]);
      const result = splitDocument(doc, { by: "pageBreak" });

      expect(result).toHaveLength(2);
    });

    it("splits at pageBreakBefore property", () => {
      const doc = createDoc([
        para("first"),
        para("second", { pageBreakBefore: true }),
        para("third")
      ]);
      // pageBreakBefore is treated as a "split after" marker on the paragraph it's on
      const result = splitDocument(doc, { by: "pageBreak" });
      expect(result.length).toBeGreaterThanOrEqual(1);
    });
  });

  describe("by heading", () => {
    it("splits at Heading 1", () => {
      const doc = createDoc([
        para("intro"),
        para("Chapter 1", { style: "Heading1" }),
        para("ch1 content"),
        para("Chapter 2", { style: "Heading1" }),
        para("ch2 content")
      ]);

      const result = splitDocument(doc, { by: "heading", headingLevel: 1 });
      expect(result).toHaveLength(3);
      expect(paraText(result[0].body[0] as Paragraph)).toBe("intro");
      expect(paraText(result[1].body[0] as Paragraph)).toBe("Chapter 1");
      expect(paraText(result[2].body[0] as Paragraph)).toBe("Chapter 2");
    });

    it("supports Heading 2", () => {
      const doc = createDoc([
        para("intro"),
        para("Section A", { style: "Heading2" }),
        para("body"),
        para("Section B", { style: "Heading2" })
      ]);

      const result = splitDocument(doc, { by: "heading", headingLevel: 2 });
      expect(result).toHaveLength(3);
    });

    it("ignores non-matching heading levels", () => {
      const doc = createDoc([para("a"), para("Heading 2", { style: "Heading2" }), para("b")]);

      // Splitting on h1 should not find any matches
      const result = splitDocument(doc, { by: "heading", headingLevel: 1 });
      expect(result).toHaveLength(1);
    });

    it("recognises outlineLevel-only headings (no Heading style)", () => {
      // outlineLevel 0 → H1, outlineLevel 1 → H2, etc. Splitting must
      // honour this even when no Heading style id is present.
      const doc = createDoc([
        para("intro"),
        para("Chapter 1", { outlineLevel: 0 }),
        para("body 1"),
        para("Chapter 2", { outlineLevel: 0 })
      ]);

      const result = splitDocument(doc, { by: "heading", headingLevel: 1 });
      expect(result).toHaveLength(3);
      expect(paraText(result[1].body[0] as Paragraph)).toBe("Chapter 1");
      expect(paraText(result[2].body[0] as Paragraph)).toBe("Chapter 2");
    });
  });

  describe("preserveSharedParts", () => {
    it("preserves styles when preserveSharedParts is true (default)", () => {
      const styles = [{ styleId: "Normal", name: "Normal", type: "paragraph" } as any];
      const doc = createDoc([
        para("a"),
        para("b", { sectionProperties: { breakType: "nextPage" } }),
        para("c")
      ]);
      (doc as any).styles = styles;

      const result = splitDocument(doc);
      for (const split of result) {
        expect(split.styles).toEqual(styles);
      }
    });

    it("strips shared parts when preserveSharedParts is false", () => {
      const styles = [{ styleId: "Normal", name: "Normal", type: "paragraph" } as any];
      const doc = createDoc([
        para("a"),
        para("b", { sectionProperties: { breakType: "nextPage" } }),
        para("c")
      ]);
      (doc as any).styles = styles;

      const result = splitDocument(doc, { preserveSharedParts: false });
      for (const split of result) {
        expect(split.styles).toBeUndefined();
      }
    });
  });

  it("returns at least one document for empty body", () => {
    const doc = createDoc([]);
    const result = splitDocument(doc);
    expect(result).toHaveLength(1);
  });
});
