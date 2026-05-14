/**
 * DOCX Module - HTML Import Tests
 *
 * Tests for htmlToDocxBody conversion.
 */

import { describe, it, expect } from "vitest";

import { htmlToDocxBody } from "../convert/html/html-import";
import type { BodyContent, Paragraph, Run, Table } from "../types";

// Helper to extract paragraph text
function paraText(block: BodyContent): string {
  if (block.type !== "paragraph") {
    return "";
  }
  let t = "";
  for (const child of (block as Paragraph).children) {
    if ("content" in child && Array.isArray((child as Run).content)) {
      for (const c of (child as Run).content) {
        if (c.type === "text") {
          t += c.text;
        }
      }
    }
  }
  return t;
}

describe("htmlToDocxBody", () => {
  describe("basic elements", () => {
    it("converts plain text", () => {
      const blocks = htmlToDocxBody("Hello World");
      expect(blocks.length).toBeGreaterThan(0);
      expect(paraText(blocks[0])).toBe("Hello World");
    });

    it("converts <p> elements", () => {
      const blocks = htmlToDocxBody("<p>First</p><p>Second</p>");
      expect(blocks.length).toBe(2);
      expect(paraText(blocks[0])).toBe("First");
      expect(paraText(blocks[1])).toBe("Second");
    });

    it("converts headings (h1-h6)", () => {
      const blocks = htmlToDocxBody("<h1>Title</h1><h2>Subtitle</h2>");
      expect(blocks.length).toBe(2);
      expect(paraText(blocks[0])).toBe("Title");
      expect(paraText(blocks[1])).toBe("Subtitle");
      // Check heading style applied
      const h1 = blocks[0] as Paragraph;
      expect(h1.properties?.style).toMatch(/[Hh]eading1/);
    });

    it("converts <br> to line break", () => {
      const blocks = htmlToDocxBody("<p>Line1<br>Line2</p>");
      expect(blocks.length).toBe(1);
      const para = blocks[0] as Paragraph;
      // Should have a break content somewhere in the runs
      let hasBreak = false;
      for (const child of para.children) {
        if ("content" in child) {
          for (const c of (child as Run).content) {
            if (c.type === "break") {
              hasBreak = true;
            }
          }
        }
      }
      expect(hasBreak).toBe(true);
    });

    it("converts <hr> to thematic break", () => {
      const blocks = htmlToDocxBody("<hr>");
      expect(blocks.length).toBeGreaterThan(0);
      // HR creates a paragraph with bottom border (thematic break)
      const para = blocks[0] as Paragraph;
      expect(para.properties?.borders?.bottom || para.properties?.thematicBreak).toBeTruthy();
    });
  });

  describe("inline formatting", () => {
    it("converts <strong>/<b> to bold", () => {
      const blocks = htmlToDocxBody("<p><strong>bold</strong></p>");
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.bold).toBe(true);
    });

    it("converts <em>/<i> to italic", () => {
      const blocks = htmlToDocxBody("<p><em>italic</em></p>");
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.italic).toBe(true);
    });

    it("converts <u> to underline", () => {
      const blocks = htmlToDocxBody("<p><u>underlined</u></p>");
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.underline).toBeTruthy();
    });

    it("converts <s>/<del> to strikethrough", () => {
      const blocks = htmlToDocxBody("<p><s>struck</s></p>");
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.strike).toBe(true);
    });

    it("converts <sub> to subscript", () => {
      const blocks = htmlToDocxBody("<p>H<sub>2</sub>O</p>");
      const para = blocks[0] as Paragraph;
      let hasSubscript = false;
      for (const child of para.children) {
        if ("content" in child && (child as Run).properties?.vertAlign === "subscript") {
          hasSubscript = true;
        }
      }
      expect(hasSubscript).toBe(true);
    });

    it("converts <sup> to superscript", () => {
      const blocks = htmlToDocxBody("<p>x<sup>2</sup></p>");
      const para = blocks[0] as Paragraph;
      let hasSuperscript = false;
      for (const child of para.children) {
        if ("content" in child && (child as Run).properties?.vertAlign === "superscript") {
          hasSuperscript = true;
        }
      }
      expect(hasSuperscript).toBe(true);
    });

    it("converts nested formatting (bold + italic)", () => {
      const blocks = htmlToDocxBody("<p><strong><em>bold italic</em></strong></p>");
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.bold).toBe(true);
      expect(run.properties?.italic).toBe(true);
    });
  });

  describe("links", () => {
    it("converts <a> to hyperlink", () => {
      const blocks = htmlToDocxBody('<p><a href="https://example.com">click</a></p>');
      const para = blocks[0] as Paragraph;
      let hasHyperlink = false;
      for (const child of para.children) {
        if ("type" in child && (child as any).type === "hyperlink") {
          hasHyperlink = true;
        }
      }
      expect(hasHyperlink).toBe(true);
    });

    it("strips javascript: URLs from <a href> in normal closed-link form", () => {
      const blocks = htmlToDocxBody('<p><a href="javascript:alert(1)">click</a></p>');
      const para = blocks[0] as Paragraph;
      for (const child of para.children) {
        if ("type" in child && (child as any).type === "hyperlink") {
          // sanitizeUrl strips dangerous schemes — url should NOT contain
          // "javascript:".
          expect((child as any).url).not.toMatch(/javascript:/i);
        }
      }
    });

    it("strips javascript: URLs from <a href> when the tag is never closed (EOF fallback)", () => {
      // Previously the EOF path used the raw href attribute, bypassing
      // sanitizeUrl. Confirm the unsafe scheme is dropped on this path too.
      const blocks = htmlToDocxBody('<p><a href="javascript:alert(1)">click');
      const para = blocks[0] as Paragraph;
      for (const child of para.children) {
        if ("type" in child && (child as any).type === "hyperlink") {
          expect((child as any).url).not.toMatch(/javascript:/i);
        }
      }
    });
  });

  describe("lists", () => {
    it("converts unordered list", () => {
      const blocks = htmlToDocxBody("<ul><li>Item 1</li><li>Item 2</li></ul>");
      expect(blocks.length).toBe(2);
      const para = blocks[0] as Paragraph;
      expect(para.properties?.numbering).toBeTruthy();
    });

    it("converts ordered list", () => {
      const blocks = htmlToDocxBody("<ol><li>First</li><li>Second</li></ol>");
      expect(blocks.length).toBe(2);
      const para = blocks[0] as Paragraph;
      expect(para.properties?.numbering).toBeTruthy();
    });

    it("converts nested lists", () => {
      const blocks = htmlToDocxBody("<ul><li>A<ul><li>A1</li></ul></li></ul>");
      expect(blocks.length).toBeGreaterThanOrEqual(2);
    });
  });

  describe("tables", () => {
    it("converts basic table", () => {
      const html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>";
      const blocks = htmlToDocxBody(html);
      expect(blocks.length).toBe(1);
      expect(blocks[0].type).toBe("table");
      const table = blocks[0] as Table;
      expect(table.rows.length).toBe(2);
      expect(table.rows[0].cells.length).toBe(2);
    });

    it("converts table with header row", () => {
      const html =
        "<table><thead><tr><th>H1</th><th>H2</th></tr></thead><tbody><tr><td>A</td><td>B</td></tr></tbody></table>";
      const blocks = htmlToDocxBody(html);
      expect(blocks[0].type).toBe("table");
      const table = blocks[0] as Table;
      expect(table.rows.length).toBe(2);
    });

    it("handles colspan", () => {
      const html = '<table><tr><td colspan="2">Wide</td></tr><tr><td>A</td><td>B</td></tr></table>';
      const blocks = htmlToDocxBody(html);
      const table = blocks[0] as Table;
      // First row should have a cell with gridSpan
      const firstCell = table.rows[0].cells[0];
      expect(firstCell.properties?.gridSpan).toBe(2);
    });
  });

  describe("styles", () => {
    it("converts inline style text-align", () => {
      const blocks = htmlToDocxBody('<p style="text-align: center">Centered</p>');
      const para = blocks[0] as Paragraph;
      expect(para.properties?.alignment).toBe("center");
    });

    it("converts inline style color", () => {
      const blocks = htmlToDocxBody('<p><span style="color: #ff0000">Red</span></p>');
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect((run.properties?.color as string)?.toLowerCase()).toBe("ff0000");
    });

    it("converts inline style font-size", () => {
      const blocks = htmlToDocxBody('<p><span style="font-size: 24pt">Big</span></p>');
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      // 24pt = 48 half-points
      expect(run.properties?.size).toBe(48);
    });
  });

  describe("edge cases", () => {
    it("handles empty string", () => {
      const blocks = htmlToDocxBody("");
      expect(blocks).toEqual([]);
    });

    it("handles whitespace-only", () => {
      const blocks = htmlToDocxBody("   \n  ");
      // Either empty or single paragraph with whitespace
      expect(blocks.length).toBeLessThanOrEqual(1);
    });

    it("handles HTML entities", () => {
      const blocks = htmlToDocxBody("<p>&lt;tag&gt; &amp; &quot;quotes&quot;</p>");
      expect(paraText(blocks[0])).toBe('<tag> & "quotes"');
    });

    it("handles pre/code blocks", () => {
      const blocks = htmlToDocxBody("<pre>  formatted\n  code</pre>");
      expect(blocks.length).toBeGreaterThan(0);
    });

    it("handles blockquote", () => {
      const blocks = htmlToDocxBody("<blockquote>Quoted text</blockquote>");
      expect(blocks.length).toBeGreaterThan(0);
      expect(paraText(blocks[0])).toContain("Quoted text");
    });

    it("strips <!doctype>, comments, and <head> contents from the body", () => {
      // Regression: tokenize() used to emit `!doctype html>` as text and
      // parseBlocks rendered <title>/<meta> as runs.
      const html = `<!doctype html>
        <html>
          <head>
            <title>Page title</title>
            <meta charset="utf-8"/>
            <link rel="stylesheet" href="x.css"/>
            <style>.x { color: red }</style>
          </head>
          <body>
            <p>Body</p>
          </body>
        </html>`;
      const blocks = htmlToDocxBody(html);
      const allText = blocks.map(paraText).join("\n");
      expect(allText).toContain("Body");
      // None of the head-only content should leak into the body
      expect(allText).not.toContain("Page title");
      expect(allText).not.toContain("doctype");
      expect(allText).not.toContain(".x");
      expect(allText).not.toContain("stylesheet");
    });

    it("strips HTML comments", () => {
      const html = `<p>before</p><!-- hidden --><p>after</p>`;
      const blocks = htmlToDocxBody(html);
      const allText = blocks.map(paraText).join("|");
      expect(allText).not.toContain("hidden");
      expect(allText).toContain("before");
      expect(allText).toContain("after");
    });

    it("data: image URI does not become a non-canonical r:embed reference", () => {
      // Previously the importer set `rId: "data:image/png;base64,..."`,
      // which Word rejects. Make sure the placeholder image either
      // surfaces an empty rId (skipped at render time) or is omitted.
      const dataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC";
      const blocks = htmlToDocxBody(`<p><img src="${dataUri}" alt="dot"/></p>`);
      // Walk every InlineImageContent and assert no rId carries a data URI
      const imageRIds: string[] = [];
      for (const block of blocks) {
        if (block.type !== "paragraph") {
          continue;
        }
        for (const child of block.children) {
          if ("content" in child) {
            for (const c of child.content) {
              if (c.type === "image") {
                imageRIds.push(c.rId);
              }
            }
          }
        }
      }
      // Either no image content, or rId is empty placeholder (never the data URI)
      for (const rId of imageRIds) {
        expect(rId.startsWith("data:")).toBe(false);
      }
    });
  });

  describe("options", () => {
    it("respects defaultFont option", () => {
      const blocks = htmlToDocxBody("<p>Text</p>", { defaultFont: "Arial" });
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      // Font should be set on the run as a FontSpec record (not bare string)
      // so we check the relevant slots rather than relying on string `toContain`.
      const font = run.properties?.font;
      expect(font).toBeDefined();
      if (typeof font === "object" && font !== null) {
        expect(font.ascii).toBe("Arial");
        expect(font.hAnsi).toBe("Arial");
      } else {
        expect(font).toBe("Arial");
      }
    });

    it("respects defaultFontSize option", () => {
      const blocks = htmlToDocxBody("<p>Text</p>", { defaultFontSize: 28 });
      const para = blocks[0] as Paragraph;
      const run = para.children[0] as Run;
      expect(run.properties?.size).toBe(28);
    });
  });
});
