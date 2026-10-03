/**
 * The page layout (Word → PDF) shows the same final view as the converters:
 * revisions accepted, hidden text omitted, field results and symbols drawn,
 * and headings detected by the shared `resolveHeadingLevel` rule.
 */

import { updateTableOfContents } from "@word/advanced/field-engine";
import { renderToMarkdown } from "@word/convert/markdown/markdown-renderer";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type {
  DocxDocument,
  Paragraph,
  ParagraphChild,
  RevisionInfo,
  Run,
  RunProperties,
  StyleDef
} from "@word/types";
import { describe, expect, it } from "vitest";

const rev: RevisionInfo = { id: 1, author: "t" } as RevisionInfo;

function run(text: string, properties?: RunProperties): Run {
  return { content: [{ type: "text", text }], ...(properties ? { properties } : {}) } as Run;
}

function para(children: ParagraphChild[], style?: string): Paragraph {
  return { type: "paragraph", children, ...(style ? { properties: { style } } : {}) };
}

function doc(body: Paragraph[], styles: StyleDef[] = []): DocxDocument {
  return { body, styles } as DocxDocument;
}

/** Every drawn run, page by page, in reading order. */
function laidRuns(d: DocxDocument) {
  const out: { text: string; fontSize: number; font: string }[] = [];
  for (const page of layoutDocumentFull(d).pages) {
    for (const block of page.content) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const line of block.lines) {
        for (const r of line.runs) {
          if (r.type !== "image") {
            out.push({ text: r.text, fontSize: r.fontSize, font: r.font });
          }
        }
      }
    }
  }
  return out;
}

function laidText(d: DocxDocument): string {
  return laidRuns(d)
    .map(r => r.text)
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

function markdownText(d: DocxDocument): string {
  return renderToMarkdown(d).replace(/\s+/g, " ").trim();
}

describe("layout final view", () => {
  it("draws inserted and moved-to text, not deleted or moved-from", () => {
    const d = doc([
      para([
        run("A "),
        { type: "insertedRun", revision: rev, run: run("B ") },
        { type: "deletedRun", revision: rev, run: run("C ") },
        { type: "movedToRun", revision: rev, run: run("D ") },
        { type: "movedFromRun", revision: rev, run: run("E ") },
        {
          type: "hyperlink",
          url: "https://example.com",
          children: [
            { type: "insertedRun", revision: rev, run: run("F") } as unknown as Run,
            { type: "deletedRun", revision: rev, run: run("G") } as unknown as Run
          ]
        }
      ])
    ]);
    expect(laidText(d)).toBe("A B D F");
  });

  it("matches the Markdown converter's text for the same paragraph", () => {
    const d = doc([
      para([
        run("one "),
        { type: "insertedRun", revision: rev, run: run("two ") },
        { type: "deletedRun", revision: rev, run: run("gone ") },
        { type: "movedToRun", revision: rev, run: run("three") }
      ])
    ]);
    expect(laidText(d)).toBe(markdownText(d));
  });

  it("omits hidden text, set directly or through a character style", () => {
    const styles = [
      { type: "character", styleId: "Secret", name: "Secret", runProperties: { vanish: true } }
    ] as StyleDef[];
    const d = doc(
      [
        para([run("shown "), run("direct ", { vanish: true }), run("styled ", { style: "Secret" })])
      ],
      styles
    );
    expect(laidText(d)).toBe("shown");
  });

  it("draws a field's cached result", () => {
    const d = doc([
      para([
        run("Total: "),
        { content: [{ type: "field", instruction: " =1+1 ", cachedValue: "2" }] } as Run
      ])
    ]);
    expect(laidText(d)).toBe("Total: 2");
    expect(laidText(d)).toBe(markdownText(d));
  });

  it("draws a symbol character", () => {
    const d = doc([
      para([
        run("x"),
        { content: [{ type: "symbol", font: "Arial", char: "2192" }] } as Run,
        run("y")
      ])
    ]);
    expect(laidText(d)).toBe("x\u2192y");
  });

  it("scales a custom style based on Heading1 like Heading1", () => {
    const styles = [
      { type: "paragraph", styleId: "Heading1", name: "heading 1" },
      { type: "paragraph", styleId: "MyHead", name: "My Head", basedOn: "Heading1" }
    ] as StyleDef[];
    const lineHeight = (style: string) => {
      const page = layoutDocumentFull(doc([para([run("Head")], style)], styles)).pages[0];
      const block = page.content[0];
      return block.type === "paragraph" ? block.lines[0].height : NaN;
    };
    const body = lineHeight("Normal");
    expect(lineHeight("Heading1")).toBeGreaterThan(body * 1.5);
    expect(lineHeight("MyHead")).toBeCloseTo(lineHeight("Heading1"), 5);
  });

  it("lists the same headings in a TOC as resolveHeadingLevel finds", () => {
    const styles = [
      { type: "paragraph", styleId: "Heading1", name: "heading 1" },
      { type: "paragraph", styleId: "MyHead", name: "My Head", basedOn: "Heading1" },
      { type: "paragraph", styleId: "Title", name: "Title" }
    ] as StyleDef[];
    const d = {
      body: [
        { type: "tableOfContents" },
        para([run("Doc title")], "Title"),
        para([run("Custom")], "MyHead"),
        para([run("Body")])
      ],
      styles
    } as unknown as DocxDocument;
    const toc = updateTableOfContents(d).body[0];
    expect(toc.type).toBe("tableOfContents");
    const entries = (toc.type === "tableOfContents" ? (toc.cachedParagraphs ?? []) : []).map(p =>
      p.children
        .flatMap(c => ("content" in c ? c.content : []))
        .map(c => (c.type === "text" ? c.text : ""))
        .join("")
    );
    expect(entries.some(e => e.includes("Custom"))).toBe(true);
    // Title has no outline level, so Word's TOC omits it.
    expect(entries.some(e => e.includes("Doc title"))).toBe(false);
    expect(entries.some(e => e.includes("Body"))).toBe(false);
  });

  it("shows the current page number for a PAGE field in a header", () => {
    const pageField = {
      content: [{ type: "field", instruction: " PAGE ", cachedValue: "1" }]
    } as Run;
    const d = {
      body: [
        para([run("one")]),
        { type: "paragraph", properties: { pageBreakBefore: true }, children: [run("two")] }
      ],
      sectionProperties: { headers: [{ type: "default", rId: "rH" }] },
      headers: new Map([["rH", { content: { children: [para([run("p"), pageField])] } }]]),
      styles: []
    } as unknown as DocxDocument;
    const headerTexts = layoutDocumentFull(d).pages.map(page =>
      (page.header ?? [])
        .flatMap(b => (b.type === "paragraph" ? b.lines.flatMap(l => l.runs) : []))
        .map(r => (r.type === "image" ? "" : r.text))
        .join("")
    );
    expect(headerTexts).toEqual(["p1", "p2"]);
  });

  describe("page-number fields in headers", () => {
    const field = (instruction: string, cachedValue = "?") =>
      ({ content: [{ type: "field", instruction, cachedValue }] }) as Run;
    const pageBreak = (text: string) => ({
      type: "paragraph",
      properties: { pageBreakBefore: true },
      children: [run(text)]
    });
    const headerTexts = (doc: DocxDocument): string[] =>
      layoutDocumentFull(doc).pages.map(page =>
        (page.header ?? [])
          .flatMap(b => (b.type === "paragraph" ? b.lines.flatMap(l => l.runs) : []))
          .map(r => (r.type === "image" ? "" : r.text))
          .join("")
      );
    const withHeader = (children: Run[], body: unknown[], extra: object = {}): DocxDocument =>
      ({
        body,
        sectionProperties: { headers: [{ type: "default", rId: "rH" }], ...extra },
        headers: new Map([["rH", { content: { children: [para(children)] } }]]),
        styles: []
      }) as unknown as DocxDocument;

    it("PAGE and NUMPAGES agree — never 'Page 3 of 1'", () => {
      const doc = withHeader(
        [
          run("Page "),
          field(" PAGE   \\* MERGEFORMAT ", "1"),
          run(" of "),
          field(" NUMPAGES ", "1")
        ],
        [para([run("one")]), pageBreak("two"), pageBreak("three")]
      );
      expect(headerTexts(doc)).toEqual(["Page 1 of 3", "Page 2 of 3", "Page 3 of 3"]);
    });

    it("honours numeric format switches and the section's page format", () => {
      const doc = withHeader(
        [field(" PAGE \\* ROMAN "), run("/"), field(" NUMPAGES \\* alphabetic ")],
        [para([run("one")]), pageBreak("two")]
      );
      expect(headerTexts(doc)).toEqual(["I/b", "II/b"]);
      const lower = withHeader([field(" PAGE ")], [para([run("one")]), pageBreak("two")], {
        pageNumbering: { start: 4, format: "lowerRoman" }
      });
      expect(headerTexts(lower)).toEqual(["iv", "v"]);
    });

    it("keeps the cached result for switches it cannot compute", () => {
      const doc = withHeader(
        [field(' PAGE \\# "00" ', "cached")],
        [para([run("one")]), pageBreak("two")]
      );
      expect(headerTexts(doc)).toEqual(["cached", "cached"]);
    });

    it("does not move body content when totals are filled in", () => {
      const body = [para([run("one")]), pageBreak("two")];
      const plain = layoutDocumentFull(withHeader([run("x")], body));
      const totals = layoutDocumentFull(withHeader([field(" NUMPAGES ")], body));
      expect(totals.pages.map(p => p.content)).toEqual(plain.pages.map(p => p.content));
    });
  });
});
