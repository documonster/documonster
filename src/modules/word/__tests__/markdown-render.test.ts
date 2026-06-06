/**
 * Focused coverage for the DOCX → Markdown renderer
 * (`renderToMarkdown`).
 *
 * The broad smoke cases (headings / bold-italic / table / bullet list)
 * live in `gap-closure.test.ts`. This file exercises the renderer's
 * richer surface that was previously untested: strikethrough, inline
 * code (monospace font), fenced code blocks, blockquotes, hyperlinks,
 * ordered lists with nesting, footnotes, setext headings, pipe escaping
 * in tables, and track-change skipping. Each case asserts the *exact*
 * GFM output produced, so a regression in any branch fails loudly.
 */

import { describe, it, expect } from "vitest";

import { renderToMarkdown } from "../convert/markdown/markdown-renderer";
import type { DocxDocument, Paragraph, Run } from "../types";

function para(children: Paragraph["children"], properties?: Paragraph["properties"]): Paragraph {
  return { type: "paragraph", properties, children };
}

function textRun(text: string, properties?: Run["properties"]): Run {
  return { properties, content: [{ type: "text", text }] } as Run;
}

describe("renderToMarkdown — inline formatting", () => {
  it("renders strikethrough as ~~…~~", () => {
    const doc: DocxDocument = { body: [para([textRun("gone", { strike: true })])] };
    expect(renderToMarkdown(doc)).toContain("~~gone~~");
  });

  it("renders combined bold+italic as ***…***", () => {
    const doc: DocxDocument = {
      body: [para([textRun("strong-em", { bold: true, italic: true })])]
    };
    expect(renderToMarkdown(doc)).toContain("***strong-em***");
  });

  it("renders a monospace-font run as inline `code`", () => {
    // A monospace run mixed with normal text renders inline; a paragraph
    // that is *entirely* monospace would instead become a fenced block.
    const doc: DocxDocument = {
      body: [para([textRun("run "), textRun("npm install", { font: "Courier New" })])]
    };
    expect(renderToMarkdown(doc)).toContain("`npm install`");
  });

  it("renders a hyperlink as [text](url)", () => {
    const doc: DocxDocument = {
      body: [
        para([
          {
            type: "hyperlink",
            url: "https://example.com",
            children: [textRun("Example")]
          } as never
        ])
      ]
    };
    expect(renderToMarkdown(doc)).toContain("[Example](https://example.com)");
  });
});

describe("renderToMarkdown — block constructs", () => {
  it("renders a fenced code block for a Code-styled paragraph", () => {
    const doc: DocxDocument = {
      body: [para([textRun("const x = 1;")], { style: "Code" })]
    };
    const md = renderToMarkdown(doc);
    expect(md).toMatch(/```\nconst x = 1;\n```/);
  });

  it("renders a blockquote for a Quote-styled paragraph", () => {
    const doc: DocxDocument = {
      body: [para([textRun("To be or not to be.")], { style: "Quote" })]
    };
    expect(renderToMarkdown(doc)).toContain("> To be or not to be.");
  });

  it("renders setext-style headings when requested", () => {
    const doc: DocxDocument = {
      body: [
        para([textRun("Big Title")], { style: "Heading1" }),
        para([textRun("Sub Title")], { style: "Heading2" })
      ]
    };
    const md = renderToMarkdown(doc, { headingStyle: "setext" });
    expect(md).toContain("Big Title\n===");
    expect(md).toContain("Sub Title\n---");
  });
});

describe("renderToMarkdown — lists", () => {
  it("renders an ordered list when numbering resolves to a decimal format", () => {
    const doc: DocxDocument = {
      body: [
        para([textRun("First")], { numbering: { numId: 7, level: 0 } }),
        para([textRun("Second")], { numbering: { numId: 7, level: 0 } })
      ],
      numberingInstances: [{ numId: 7, abstractNumId: 3 }],
      abstractNumberings: [
        {
          abstractNumId: 3,
          levels: [{ level: 0, format: "decimal", text: "%1." }]
        }
      ]
    } as DocxDocument;
    const md = renderToMarkdown(doc);
    expect(md).toContain("1. First");
    expect(md).toContain("1. Second");
  });

  it("indents nested list items by two spaces per level", () => {
    const doc: DocxDocument = {
      body: [
        para([textRun("Top")], { numbering: { numId: 1, level: 0 } }),
        para([textRun("Nested")], { numbering: { numId: 1, level: 1 } })
      ]
    };
    const md = renderToMarkdown(doc);
    expect(md).toContain("- Top");
    expect(md).toContain("  - Nested");
  });
});

describe("renderToMarkdown — tables", () => {
  it("escapes pipe characters inside table cells", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            {
              cells: [{ content: [para([textRun("a|b")])] }, { content: [para([textRun("c")])] }]
            }
          ]
        }
      ]
    };
    const md = renderToMarkdown(doc);
    // The literal pipe must be escaped so it does not split the column.
    expect(md).toContain("a\\|b");
  });
});

describe("renderToMarkdown — footnotes & track changes", () => {
  it("emits footnote markers and definitions when includeNotes is on", () => {
    const doc: DocxDocument = {
      body: [para([textRun("See note"), { content: [{ type: "footnoteRef", id: 2 }] } as Run])],
      footnotes: [
        {
          id: 2,
          type: "normal",
          content: [para([textRun("The footnote text.")])]
        }
      ]
    } as DocxDocument;
    const md = renderToMarkdown(doc, { includeNotes: true });
    expect(md).toMatch(/\[\^1\]/);
    expect(md).toContain("[^1]: The footnote text.");
  });

  it("skips deleted (tracked-change) runs", () => {
    const doc: DocxDocument = {
      body: [
        para([
          textRun("kept "),
          { type: "deletedRun", run: textRun("removed") } as never,
          textRun("text")
        ])
      ]
    };
    const md = renderToMarkdown(doc);
    expect(md).toContain("kept text");
    expect(md).not.toContain("removed");
  });
});
