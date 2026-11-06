/**
 * A line break or a tab inside a run's text (`Build.text("a\nb")`) is one: the
 * DOCX writer writes `<w:br/>` for a line break and keeps a tab in `<w:t>`,
 * which Word sets as a tab, and the layout and the converters read them the
 * same way. The layout used to draw them as characters, which no face has a
 * glyph for — a PDF showed `a�b`.
 */
import { describe, expect, it } from "vitest";

import { WordHtml } from "../html";
import { Build, Document } from "../index";
import type { DocxDocument, Paragraph } from "../index";
import { WordMarkdown } from "../markdown";

function docOf(text: string): DocxDocument {
  const doc = Document.create();
  Document.addParagraph(doc, Build.paragraph([Build.text(text)]));
  return doc;
}

describe("a line break or a tab inside run text", () => {
  it("is written as <w:br/>, with the tab kept in the text", async () => {
    const back = await Document.read(await Document.toBuffer(docOf("a\nb\tc\r\nd")));
    const content = (back.body[0] as Paragraph).children.flatMap(c =>
      "content" in c ? c.content.map(x => (x.type === "text" ? x.text : x.type)) : []
    );
    expect(content).toEqual(["a", "break", "b\tc", "break", "d"]);
  });

  it("converts to HTML as a <br> and a tab", () => {
    const { html } = WordHtml.fromDocx(docOf("a\nb\tc"));
    expect(html).toMatch(/a\s*<br>\s*b\s*<span class="docx-tab"[^>]*><\/span>\s*c/);
  });

  it("converts to Markdown as a hard line break", () => {
    expect(WordMarkdown.fromDocx(docOf("a\nb"))).toBe("a  \nb\n");
  });
});
