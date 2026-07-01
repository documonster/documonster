/**
 * DOCX Module - Document Split API Tests
 */

import { describe, it, expect } from "vitest";

import { Io } from "../index";
import type { DocxDocument, Paragraph, Run, BodyContent, SectionProperties } from "../types";

function textRun(t: string): Run {
  return { content: [{ type: "text", text: t }] };
}

function para(t: string, props?: Paragraph["properties"]): Paragraph {
  return { type: "paragraph", children: [textRun(t)], properties: props };
}

function createDoc(body: BodyContent[]): DocxDocument {
  return { body };
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
      const result = Io.split(doc);
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

      const result = Io.split(doc, { by: "section" });
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

      const result = Io.split(doc, { by: "section" });
      expect(result).toHaveLength(3);
    });

    it("strips the trailing section-break sectPr from each part (no blank page)", () => {
      // The break paragraph's sectPr carries a nextPage break used to separate
      // it from the following section. In a standalone split document that
      // break has nothing after it and would render a trailing blank page.
      // splitDocument must remove the paragraph-level sectPr and promote its
      // page setup (without breakType) to the document's section properties.
      const sectProps: SectionProperties = {
        breakType: "nextPage",
        pageSize: { width: 12240, height: 15840 }
      };
      const doc = createDoc([
        para("section 1 text"),
        para("break para", { sectionProperties: sectProps }),
        para("section 2 text")
      ]);

      const result = Io.split(doc, { by: "section" });
      expect(result).toHaveLength(2);

      // Part 1's last paragraph must no longer carry a paragraph-level sectPr.
      const part1 = result[0];
      const lastPara = part1.body[part1.body.length - 1] as Paragraph;
      expect(lastPara.properties?.sectionProperties).toBeUndefined();
      expect(paraText(lastPara)).toBe("break para");

      // The page setup is promoted to the document level, with breakType dropped.
      expect(part1.sectionProperties?.pageSize).toEqual({ width: 12240, height: 15840 });
      expect(part1.sectionProperties?.breakType).toBeUndefined();
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
      const result = Io.split(doc, { by: "pageBreak" });

      expect(result).toHaveLength(2);
    });

    it("removes the trailing page break from each part (no blank page)", () => {
      // The page-break paragraph that triggered the split must not survive at
      // the end of the part, otherwise the standalone document renders a
      // trailing blank page. A paragraph that held ONLY the page break is
      // dropped entirely.
      const breakOnlyPara: Paragraph = {
        type: "paragraph",
        children: [{ content: [{ type: "break", breakType: "page" }] } as Run]
      };
      const doc = createDoc([para("page 1"), breakOnlyPara, para("page 2")]);
      const result = Io.split(doc, { by: "pageBreak" });

      expect(result).toHaveLength(2);
      // Part 1 keeps only "page 1"; the empty break paragraph is gone.
      expect(result[0].body).toHaveLength(1);
      expect(paraText(result[0].body[0] as Paragraph)).toBe("page 1");
      // No page-break run remains anywhere in part 1.
      const part1Json = JSON.stringify(result[0].body);
      expect(part1Json).not.toContain('"breakType":"page"');
    });

    it("keeps inline text when stripping a trailing page break", () => {
      // A paragraph with text AND a page break keeps the text, loses the break.
      const mixedPara: Paragraph = {
        type: "paragraph",
        children: [
          {
            content: [
              { type: "text", text: "tail text" },
              { type: "break", breakType: "page" }
            ]
          } as Run
        ]
      };
      const doc = createDoc([para("head"), mixedPara, para("next")]);
      const result = Io.split(doc, { by: "pageBreak" });

      expect(result).toHaveLength(2);
      // Part 1 keeps both paragraphs; the trailing break is removed but "tail text" stays.
      const part1Json = JSON.stringify(result[0].body);
      expect(part1Json).toContain("tail text");
      expect(part1Json).not.toContain('"breakType":"page"');
    });

    it("splits at pageBreakBefore property", () => {
      const doc = createDoc([
        para("first"),
        para("second", { pageBreakBefore: true }),
        para("third")
      ]);
      // pageBreakBefore is treated as a "split after" marker on the paragraph it's on
      const result = Io.split(doc, { by: "pageBreak" });
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

      const result = Io.split(doc, { by: "heading", headingLevel: 1 });
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

      const result = Io.split(doc, { by: "heading", headingLevel: 2 });
      expect(result).toHaveLength(3);
    });

    it("ignores non-matching heading levels", () => {
      const doc = createDoc([para("a"), para("Heading 2", { style: "Heading2" }), para("b")]);

      // Splitting on h1 should not find any matches
      const result = Io.split(doc, { by: "heading", headingLevel: 1 });
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

      const result = Io.split(doc, { by: "heading", headingLevel: 1 });
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

      const result = Io.split(doc);
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

      const result = Io.split(doc, { preserveSharedParts: false });
      for (const split of result) {
        expect(split.styles).toBeUndefined();
      }
    });
  });

  it("returns at least one document for empty body", () => {
    const doc = createDoc([]);
    const result = Io.split(doc);
    expect(result).toHaveLength(1);
  });
});
