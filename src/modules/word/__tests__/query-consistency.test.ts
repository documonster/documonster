/**
 * The public query APIs and converters share one reading of the styles:
 * headings follow the shared heading rule, run formatting is resolved, and
 * conversion internals do not leak into returned objects.
 */

import { docxToSemantic } from "@word/convert/docx-to-semantic";
import { renderToHtml } from "@word/convert/html/html-renderer";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { Query } from "@word/index";
import type { StyleIndex } from "@word/index";
import { getHeadings } from "@word/query/search";
import { splitDocument } from "@word/query/split";
import type {
  BodyContent,
  DocxDocument,
  Paragraph,
  ParagraphProperties,
  RunProperties,
  StyleDef
} from "@word/types";
import { describe, expect, it } from "vitest";

function para(text: string, properties?: ParagraphProperties, rPr?: RunProperties): Paragraph {
  return {
    type: "paragraph",
    properties,
    children: [{ properties: rPr, content: [{ type: "text", text }] }]
  };
}

function firstText(p: Paragraph): string {
  const first = p.children[0];
  const content = first && "content" in first ? first.content[0] : undefined;
  return content?.type === "text" ? content.text : "";
}

function doc(
  body: BodyContent[],
  styles: StyleDef[] = [],
  docDefaults?: DocxDocument["docDefaults"]
): DocxDocument {
  return { body, styles, docDefaults } as DocxDocument;
}

function style(styleId: string, extra: Partial<StyleDef> = {}): StyleDef {
  return { type: "paragraph", styleId, name: styleId, ...extra };
}

const STYLES: StyleDef[] = [
  style("Heading1", { name: "heading 1" }),
  style("Heading2", { name: "heading 2" }),
  style("Title", { name: "Title" }),
  style("Chapter", { name: "Chapter", paragraphProperties: { outlineLevel: 0 } }),
  style("ChapterLike", { name: "Chapter Like", basedOn: "Chapter" }),
  style("BodyBase", { name: "Body Base", paragraphProperties: { outlineLevel: 9 } }),
  style("NotAHeading", { name: "Not A Heading", basedOn: "BodyBase" }),
  // A heading style whose own chain marks it body text.
  style("Heading2Body", { name: "heading 2", basedOn: "BodyBase" })
];

describe("getHeadings and splitDocument use the shared heading rule", () => {
  const d = doc(
    [
      para("Title text", { style: "Title" }),
      para("Inherited", { style: "ChapterLike" }),
      para("Body", { style: "NotAHeading" }),
      para("Body heading", { style: "Heading2Body" }),
      para("Direct wins", { style: "Heading2", outlineLevel: 0 }),
      para("Plain H2", { style: "Heading2" }),
      para("Direct body", { style: "Heading1", outlineLevel: 9 })
    ],
    STYLES
  );

  it("getHeadings follows basedOn and outline level 9, and omits Title", () => {
    expect(getHeadings(d).map(h => [h.text, h.level])).toEqual([
      ["Inherited", 1],
      ["Direct wins", 1],
      ["Plain H2", 2]
    ]);
  });

  it("agrees with Query.resolveHeadingLevel-based converters", () => {
    // Markdown is produced by the shared rule; every heading getHeadings
    // reports is a Markdown heading at the same level (Title aside).
    const md = renderToMarkdown(d);
    for (const h of getHeadings(d)) {
      expect(md).toContain(`${"#".repeat(h.level)} ${h.text}`);
    }
  });

  it("splitDocument splits at the same level-1 headings", () => {
    const parts = splitDocument(d, { by: "heading", headingLevel: 1 });
    const firsts = parts.map(p => firstText(p.body[0] as Paragraph));
    expect(firsts).toEqual(["Title text", "Inherited", "Direct wins"]);
  });

  it("splitDocument no longer treats a bare style id like H1 or Title1 as a heading", () => {
    const d2 = doc([para("a"), para("b", { style: "H1" }), para("c", { style: "Title1" })]);
    expect(splitDocument(d2, { by: "heading" })).toHaveLength(1);
  });
});

describe("Query publishes the style index its resolvers accept", () => {
  it("indexStyles builds a reusable StyleIndex", () => {
    const d = doc([], [style("Base", { runProperties: { bold: true } })]);
    const index: StyleIndex = Query.indexStyles(d);
    const p = para("x", { style: "Base" });
    expect(Query.resolveStyle(d, p, undefined, index)).toEqual(Query.resolveStyle(d, p));
  });
});

describe("Markdown code-block detection reads resolved run fonts", () => {
  it("treats a paragraph whose style sets a monospace font as code", () => {
    const d = doc(
      [para("let x = 1;", { style: "Source" })],
      [style("Source", { name: "Source", runProperties: { font: "Consolas" } })]
    );
    expect(renderToMarkdown(d)).toContain("```\nlet x = 1;\n```");
  });

  it("treats runs with a monospace character style as code", () => {
    const d = doc(
      [para("echo hi", undefined, { style: "Mono" })],
      [{ type: "character", styleId: "Mono", name: "Mono", runProperties: { font: "Courier New" } }]
    );
    expect(renderToMarkdown(d)).toContain("```\necho hi\n```");
  });
});

describe("docxToSemantic context", () => {
  it("returns only the public ConversionContext members", () => {
    const { context } = docxToSemantic(doc([para("x")]));
    expect(Object.keys(context).sort()).toEqual([
      "addWarning",
      "assets",
      "registerAsset",
      "warnings"
    ]);
  });

  it("still collects warnings and assets made during conversion", () => {
    const { context, document } = docxToSemantic(doc([para("x")]));
    expect(context.assets).toBe(document.assets);
  });
});

describe("HTML writes document-default run formatting once", () => {
  const defaults = { runProperties: { font: "Georgia", size: 22 } };
  const html = (d: DocxDocument) =>
    renderToHtml(d, { fullDocument: false, includeStyles: false }).html;

  it("puts the defaults on the container and not on plain runs", () => {
    const out = html(doc([para("a"), para("b"), para("c", undefined, { size: 28 })], [], defaults));
    expect(out).toContain(
      `<div class="docx-document" style="font-family:'Georgia';font-size:11pt">`
    );
    expect(out.match(/font-family:'Georgia'/g)).toHaveLength(1);
    expect(out).toMatch(/<p>\s*a\s*<\/p>\s*<p>\s*b\s*<\/p>/);
    // A size that differs from the default is still written.
    expect(out).toMatch(/<span style="font-size:14pt">\s*c\s*<\/span>/);
  });

  it("keeps the size where the container's does not reach the run unchanged", () => {
    const d = doc(
      [
        para("Head", { style: "Heading1" }),
        {
          type: "table",
          rows: [{ cells: [{ content: [para("cell")] }] }]
        } as unknown as BodyContent
      ],
      [style("Heading1", { name: "heading 1" })],
      defaults
    );
    const out = html(d);
    // A heading is sized by the user agent; a quirks-mode table does not inherit.
    expect(out).toMatch(/<h1[^>]*>\s*<span style="font-size:11pt">\s*Head/);
    expect(out).toMatch(/<td[^>]*>[\s\S]*<span style="font-size:11pt">\s*cell/);
    expect(out).not.toMatch(/<span style="[^"]*font-family:'Georgia'/);
  });

  it("writes no container style for a document without defaults", () => {
    expect(html(doc([para("a")]))).toContain(`<div class="docx-document">`);
  });
});
