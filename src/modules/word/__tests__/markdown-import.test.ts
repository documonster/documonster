/**
 * DOCX Module - Markdown Import Tests
 */

import { describe, it, expect } from "vitest";

import { markdownToDocx, markdownToDocxBody } from "../convert/markdown/markdown-import";
import type { Hyperlink, Paragraph, Run, Table } from "../types";

describe("markdownToDocxBody", () => {
  describe("headings", () => {
    it("should convert ATX headings (# to ######)", () => {
      const body = markdownToDocxBody(
        "# H1\n\n## H2\n\n### H3\n\n#### H4\n\n##### H5\n\n###### H6"
      );
      const paragraphs = body.filter(b => b.type === "paragraph") as Paragraph[];

      expect(paragraphs).toHaveLength(6);
      expect(paragraphs[0].properties?.style).toBe("Heading1");
      expect(paragraphs[1].properties?.style).toBe("Heading2");
      expect(paragraphs[2].properties?.style).toBe("Heading3");
      expect(paragraphs[3].properties?.style).toBe("Heading4");
      expect(paragraphs[4].properties?.style).toBe("Heading5");
      expect(paragraphs[5].properties?.style).toBe("Heading6");
    });

    it("should convert setext headings", () => {
      const body = markdownToDocxBody("Title\n===\n\nSubtitle\n---");
      const paragraphs = body.filter(b => b.type === "paragraph") as Paragraph[];

      expect(paragraphs).toHaveLength(2);
      expect(paragraphs[0].properties?.style).toBe("Heading1");
      expect(paragraphs[1].properties?.style).toBe("Heading2");
    });

    it("should handle inline formatting in headings", () => {
      const body = markdownToDocxBody("# Hello **bold** world");
      const para = body[0] as Paragraph;
      expect(para.children).toHaveLength(3);
      const boldRun = para.children[1] as Run;
      expect(boldRun.properties?.bold).toBe(true);
    });
  });

  describe("paragraphs and inline formatting", () => {
    it("should convert plain paragraphs", () => {
      const body = markdownToDocxBody("Hello world");
      expect(body).toHaveLength(1);
      const para = body[0] as Paragraph;
      expect(para.type).toBe("paragraph");
      const run = para.children[0] as Run;
      expect(run.content[0]).toEqual({ type: "text", text: "Hello world" });
    });

    it("should handle bold text", () => {
      const body = markdownToDocxBody("Hello **bold** world");
      const para = body[0] as Paragraph;
      expect(para.children).toHaveLength(3);
      const boldRun = para.children[1] as Run;
      expect(boldRun.properties?.bold).toBe(true);
      expect(boldRun.content[0]).toEqual({ type: "text", text: "bold" });
    });

    it("should handle italic text", () => {
      const body = markdownToDocxBody("Hello *italic* world");
      const para = body[0] as Paragraph;
      const italicRun = para.children[1] as Run;
      expect(italicRun.properties?.italic).toBe(true);
    });

    it("should handle bold italic text", () => {
      const body = markdownToDocxBody("Hello ***bold italic*** world");
      const para = body[0] as Paragraph;
      // Bold wraps italic
      const boldRun = para.children[1] as Run;
      expect(boldRun.properties?.bold).toBe(true);
      expect(boldRun.properties?.italic).toBe(true);
    });

    it("should handle strikethrough text", () => {
      const body = markdownToDocxBody("Hello ~~deleted~~ world");
      const para = body[0] as Paragraph;
      const strikeRun = para.children[1] as Run;
      expect(strikeRun.properties?.strike).toBe(true);
    });

    it("should handle inline code", () => {
      const body = markdownToDocxBody("Use `console.log()` to debug");
      const para = body[0] as Paragraph;
      const codeRun = para.children[1] as Run;
      expect(codeRun.properties?.font).toBe("Courier New");
      expect(codeRun.content[0]).toEqual({ type: "text", text: "console.log()" });
    });

    it("should handle links", () => {
      const body = markdownToDocxBody("Visit [Google](https://google.com)");
      const para = body[0] as Paragraph;
      const link = para.children[1] as Hyperlink;
      expect(link.type).toBe("hyperlink");
      expect(link.url).toBe("https://google.com");
      expect(link.children[0].content[0]).toEqual({ type: "text", text: "Google" });
    });

    it("should handle links with title", () => {
      const body = markdownToDocxBody('Visit [Link](https://example.com "Example") here');
      const para = body[0] as Paragraph;
      const link = para.children.find(c => "type" in c && c.type === "hyperlink") as Hyperlink;
      expect(link).toBeDefined();
      expect(link.url).toBe("https://example.com");
      expect(link.tooltip).toBe("Example");
    });

    it("should handle hard line breaks (two trailing spaces)", () => {
      const body = markdownToDocxBody("Line one  \nLine two");
      const para = body[0] as Paragraph;
      // Should contain: "Line one", break, "Line two"
      const hasBreak = para.children.some(
        child => "content" in child && (child as Run).content.some(c => c.type === "break")
      );
      expect(hasBreak).toBe(true);
    });

    it("should handle escaped characters", () => {
      const body = markdownToDocxBody("Use \\*asterisks\\* literally");
      const para = body[0] as Paragraph;
      const textContent = para.children
        .filter((c): c is Run => "content" in c)
        .flatMap(r => r.content)
        .filter(c => c.type === "text")
        .map(c => (c as { type: "text"; text: string }).text)
        .join("");
      expect(textContent).toBe("Use *asterisks* literally");
    });
  });

  describe("code blocks", () => {
    it("should convert fenced code blocks", () => {
      const body = markdownToDocxBody("```javascript\nconst x = 1;\nconsole.log(x);\n```");
      expect(body).toHaveLength(1);
      const para = body[0] as Paragraph;
      expect(para.properties?.style).toBe("CodeBlock");
      // Should contain code text with code font
      const runs = para.children.filter((c): c is Run => "content" in c);
      expect(runs.length).toBeGreaterThan(0);
    });

    it("should handle code blocks with tilde fence", () => {
      const body = markdownToDocxBody("~~~\ncode here\n~~~");
      const para = body[0] as Paragraph;
      expect(para.properties?.style).toBe("CodeBlock");
    });
  });

  describe("lists", () => {
    it("should convert unordered lists", () => {
      const body = markdownToDocxBody("- Item 1\n- Item 2\n- Item 3");
      const paragraphs = body.filter(b => b.type === "paragraph") as Paragraph[];
      expect(paragraphs).toHaveLength(3);
      // Each should have numbering reference
      for (const para of paragraphs) {
        expect(para.properties?.numbering).toBeDefined();
        expect(para.properties?.numbering?.level).toBe(0);
      }
    });

    it("should convert ordered lists", () => {
      const body = markdownToDocxBody("1. First\n2. Second\n3. Third");
      const paragraphs = body.filter(b => b.type === "paragraph") as Paragraph[];
      expect(paragraphs).toHaveLength(3);
      for (const para of paragraphs) {
        expect(para.properties?.numbering).toBeDefined();
      }
    });

    it("respects a non-default ordered list start", () => {
      // markdownToDocx exposes the numbering instances on the document; the
      // list's first paragraph's numId must reference an instance whose
      // override starts at 3.
      const doc = markdownToDocx("3. third\n4. fourth");
      const paragraphs = doc.body.filter(b => b.type === "paragraph") as Paragraph[];
      expect(paragraphs).toHaveLength(2);
      const numId = paragraphs[0].properties?.numbering?.numId;
      expect(numId).toBeDefined();
      const inst = doc.numberingInstances?.find(n => n.numId === numId);
      expect(inst).toBeDefined();
      const override = inst?.overrides?.find(o => o.level === 0);
      expect(override?.startOverride).toBe(3);
    });

    it("should handle task lists", () => {
      const body = markdownToDocxBody("- [x] Done\n- [ ] Not done");
      const paragraphs = body.filter(b => b.type === "paragraph") as Paragraph[];
      expect(paragraphs).toHaveLength(2);
      // First item should have checkbox prefix
      const firstRun = paragraphs[0].children[0] as Run;
      const text = firstRun.content[0] as { type: "text"; text: string };
      expect(text.text).toContain("☑");
    });
  });

  describe("blockquotes", () => {
    it("should convert blockquotes", () => {
      const body = markdownToDocxBody("> This is a quote");
      const para = body[0] as Paragraph;
      expect(para.properties?.style).toBe("Quote");
      expect(para.properties?.indent?.left).toBe(720);
    });

    it("should handle multi-line blockquotes", () => {
      const body = markdownToDocxBody("> Line 1\n> Line 2");
      expect(body).toHaveLength(1);
      const para = body[0] as Paragraph;
      expect(para.properties?.style).toBe("Quote");
    });
  });

  describe("thematic breaks", () => {
    it("should convert horizontal rules", () => {
      const body = markdownToDocxBody("---");
      const para = body[0] as Paragraph;
      expect(para.properties?.thematicBreak).toBe(true);
    });

    it("should handle *** and ___", () => {
      const body1 = markdownToDocxBody("***");
      const body2 = markdownToDocxBody("___");
      expect((body1[0] as Paragraph).properties?.thematicBreak).toBe(true);
      expect((body2[0] as Paragraph).properties?.thematicBreak).toBe(true);
    });
  });

  describe("tables", () => {
    it("should convert GFM tables", () => {
      const md = "| Name | Age |\n| --- | --- |\n| Alice | 30 |\n| Bob | 25 |";
      const body = markdownToDocxBody(md);
      expect(body).toHaveLength(1);
      const table = body[0] as Table;
      expect(table.type).toBe("table");
      expect(table.rows).toHaveLength(3); // header + 2 data rows
    });

    it("should handle table alignment", () => {
      const md = "| Left | Center | Right |\n| :--- | :---: | ---: |\n| a | b | c |";
      const body = markdownToDocxBody(md);
      const table = body[0] as Table;
      // First data row cells should have alignment
      const dataRow = table.rows[1];
      const cellParas = dataRow.cells.map(c => c.content[0] as Paragraph);
      expect(cellParas[0].properties?.alignment).toBe("start");
      expect(cellParas[1].properties?.alignment).toBe("center");
      expect(cellParas[2].properties?.alignment).toBe("end");
    });

    it("should bold header cells", () => {
      const md = "| H1 | H2 |\n| --- | --- |\n| a | b |";
      const body = markdownToDocxBody(md);
      const table = body[0] as Table;
      const headerRow = table.rows[0];
      const headerPara = headerRow.cells[0].content[0] as Paragraph;
      const headerRun = headerPara.children[0] as Run;
      expect(headerRun.properties?.bold).toBe(true);
    });
  });

  describe("images", () => {
    it("should handle images as placeholder text (sync mode)", () => {
      const body = markdownToDocxBody("![Alt text](https://example.com/img.png)");
      const para = body[0] as Paragraph;
      const run = para.children[0] as Run;
      const text = run.content[0] as { type: "text"; text: string };
      expect(text.text).toContain("Alt text");
    });
  });

  describe("markdownToDocx (full document)", () => {
    it("should produce a valid DocxDocument with styles and numbering", () => {
      const doc = markdownToDocx("# Title\n\n- Item 1\n- Item 2\n\nParagraph");
      expect(doc.body).toHaveLength(4); // heading + 2 list items + paragraph
      expect(doc.styles).toBeDefined();
      expect(doc.styles!.length).toBeGreaterThan(0);
      expect(doc.abstractNumberings).toBeDefined();
      expect(doc.numberingInstances).toBeDefined();
    });

    it("should include heading styles", () => {
      const doc = markdownToDocx("# Test");
      const headingStyle = doc.styles?.find(s => s.styleId === "Heading1");
      expect(headingStyle).toBeDefined();
    });
  });

  describe("complex documents", () => {
    it("should handle a mixed markdown document", () => {
      const md = `# My Document

This is a paragraph with **bold** and *italic* text.

## Section 1

- Bullet one
- Bullet two

### Subsection

| Column A | Column B |
| -------- | -------- |
| Value 1  | Value 2  |

> A blockquote here

\`\`\`typescript
const x = 42;
\`\`\`

---

End of document.`;

      const body = markdownToDocxBody(md);
      expect(body.length).toBeGreaterThan(5);

      // Check various block types are present
      const types = body.map(b => b.type);
      expect(types).toContain("paragraph");
      expect(types).toContain("table");
    });

    it("should handle autolinks", () => {
      const body = markdownToDocxBody("Visit <https://example.com> for more");
      const para = body[0] as Paragraph;
      const link = para.children.find(c => "type" in c && c.type === "hyperlink") as Hyperlink;
      expect(link).toBeDefined();
      expect(link.url).toBe("https://example.com");
    });
  });
});
