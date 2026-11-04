/**
 * A table that names no style behaves exactly as if it named the document's
 * default table style (`w:default="1"`) — its paragraph and run properties
 * included, not only margins and shading. Word for Mac 16.114 sets the text of
 * such a table bold with the style's spacing.
 */
import { renderToHtml } from "@word/convert/html/html-renderer";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { LayoutTable } from "@word/layout/layout-model";
import type { DocxDocument, Paragraph, StyleDef, Table } from "@word/types";
import { describe, expect, it } from "vitest";

const PLAIN: StyleDef = {
  type: "table",
  styleId: "Plain",
  name: "Plain",
  isDefault: true,
  paragraphProperties: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } },
  runProperties: { bold: true }
};

function cellPara(text: string): Paragraph {
  return { type: "paragraph", children: [{ content: [{ type: "text", text }] }] };
}

function doc(style?: string): DocxDocument {
  const table: Table = {
    type: "table",
    ...(style ? { properties: { style } } : {}),
    columnWidths: [4000],
    rows: [
      { cells: [{ content: [cellPara("first"), cellPara("second")] }] },
      { cells: [{ content: [cellPara("third")] }] }
    ]
  };
  return {
    docDefaults: {
      paragraphProperties: { spacing: { after: 160, line: 259, lineRule: "auto" } },
      runProperties: { font: { ascii: "Calibri", hAnsi: "Calibri" }, size: 22 }
    },
    styles: [{ type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true }, PLAIN],
    body: [table]
  };
}

function laidTable(d: DocxDocument): LayoutTable {
  return layoutDocumentFull(d).pages[0]!.content.find((c): c is LayoutTable => c.type === "table")!;
}

describe("an unstyled table takes the default table style", () => {
  it("lays its text out as the default table style sets it", () => {
    const unstyled = laidTable(doc());
    const runs = unstyled.cells.flatMap(c =>
      c.content.flatMap(p => (p.type === "paragraph" ? p.lines.flatMap(l => l.runs) : []))
    );
    expect(runs.length).toBeGreaterThan(0);
    expect(runs.every(r => "bold" in r && r.bold === true)).toBe(true);
    // Spacing after 0 from the style, not the document default's 8pt.
    const [first, second] = unstyled.cells[0]!.content;
    expect(second!.rect.y - (first!.rect.y + first!.rect.height)).toBeCloseTo(0, 6);
  });

  it("is laid out exactly as the same table naming the style", () => {
    expect(laidTable(doc())).toEqual(laidTable(doc("Plain")));
  });

  it("exports its text bold to HTML and Markdown, as the named style does", () => {
    const html = renderToHtml(doc()).html;
    expect(html).toMatch(/<strong>\s*first\s*<\/strong>/);
    expect(html).toBe(renderToHtml(doc("Plain")).html);
    const md = renderToMarkdown(doc());
    expect(md).toContain("**first**");
    expect(md).toBe(renderToMarkdown(doc("Plain")));
  });
});
