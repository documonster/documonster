import { extractAll } from "@archive/unzip/extract";
/**
 * Regressions in the Markdown → DOCX style port.
 *
 * The formatting moved out of individual paragraphs and into named styles. These
 * tests pin the consequences of that move which went wrong the first time.
 */
import { describe, it, expect } from "vitest";

import { markdownToDocx, markdownToDocxBody } from "../convert/markdown/markdown-import";
import { toBuffer } from "../document-io";
import { layoutDocumentFull } from "../layout/layout-full";
import type { Table } from "../types";

/** Read one part out of a packaged .docx as text. */
async function part(bytes: Uint8Array, path: string): Promise<string> {
  const files = await extractAll(bytes);
  const entry = files.get(path);
  expect(entry).toBeDefined();
  return new TextDecoder().decode(entry!.data);
}

describe("markdownToDocxBody", () => {
  it("returns the styles and defaults the body only references by name", async () => {
    // The body's paragraphs name `Quote`, `CodeBlock` and `ListParagraph`; every
    // visual property lives in those definitions. Without them a caller splicing
    // the body into a host document got unstyled text.
    const result = await markdownToDocxBody("> quoted\n\n```ts\ncode\n```\n\n- item");
    const referenced = new Set(
      result.body
        .filter(b => b.type === "paragraph")
        .map(b => b.properties?.style)
        .filter((s): s is string => s !== undefined)
    );
    expect(referenced.size).toBeGreaterThan(0);
    const provided = new Set(result.styles.map(s => s.styleId));
    for (const name of referenced) {
      expect(provided.has(name)).toBe(true);
    }
    expect(result.docDefaults.runProperties?.size).toBeGreaterThan(0);
  });
});

describe("blockquote", () => {
  const quoteStyle = async () =>
    (await markdownToDocx("> quoted")).styles?.find(s => s.styleId === "Quote")
      ?.paragraphProperties;

  it("does not overwrite the style of a block that carries its own", async () => {
    // Everything inside a quote was forced to `Quote`, which stripped a fenced code block of its
    // frame, background and leading, and flattened headings.
    const doc = await markdownToDocx("> ```ts\n> code\n> ```\n\n> # heading\n\n> plain");
    const styles = doc.body.filter(b => b.type === "paragraph").map(b => b.properties?.style);
    expect(styles).toEqual(["CodeBlock", "Heading1", "Quote"]);
  });

  it("is never a table, however tempting the padding is", async () => {
    // It has been a single-cell table twice, because a cell is the only OOXML construct with real
    // padding. A cell *is* a table, and a reader is entitled to treat it as one: with the PDF tagged
    // it reports the callout as a table and draws cell furniture round it when clicked, and untagged
    // its ruling lines feed the reader's own table detection. Either way the callout stops being one.
    const doc = await markdownToDocx("| A | B |\n|---|---|\n| 1 | 2 |\n\n> quoted");
    expect(doc.body.map(b => b.type)).toEqual(["table", "paragraph"]);
  });

  it("takes its bar from the theme, and never italicises the text", async () => {
    // `markdown.css` sets no colours on a block quote — the *webview host* does, via
    // `--vscode-textBlockQuote-border`. Reading only the extension's sheet turned the bar grey once.
    const props = await quoteStyle();
    expect(props?.borders?.left?.color).toBe("79B6DF");
    const quote = (await markdownToDocx("> quoted")).styles?.find(s => s.styleId === "Quote");
    expect(quote?.runProperties?.italic).toBeUndefined();
    expect(quote?.runProperties?.color).toBeUndefined();
  });

  it("carries no background at all", async () => {
    // The one change that actually fixed it. A band of fill the width of the text column, below a
    // table whose row rules span that same width, is geometrically another row of it — Preview
    // selected the callout as a cell of the table above, and did so whether the callout was built as
    // a paragraph or as a table. Removing the fill is what stopped it, and it is also what made the
    // padding below possible.
    const props = await quoteStyle();
    expect(props?.shading).toBeUndefined();
    // And no edge but the bar: an "invisible" border painted in the page or fill colour is still a
    // real ruling line in the PDF, which is the other thing a table detector latches onto.
    expect(Object.keys(props?.borders ?? {})).toEqual(["left"]);
  });

  it("carries comfortable padding, which only a quote with no fill can have", async () => {
    // `w:pBdr`'s `w:space` is the gap between the bar and the text, and it is *white*. On a tinted
    // panel that made it a visible break between the bar and the panel it was meant to be the edge
    // of, so it had to be zero and there was no left padding to be had. With no fill, white is the
    // page and the same gap is simply `padding-left`.
    const props = await quoteStyle();
    const px = (v: number) => Math.round(v * (11 / 14) * 20);
    // No outside margin. The indent carries only bar + inner padding,
    // because the bar is drawn at `indent - space - size`.
    expect(props?.indent?.left).toBe(px(5 + 14));
    expect(props?.indent?.right).toBe(px(18));
    expect(props?.borders?.left?.space).toBe(Math.round(14 * (11 / 14)));
    // A small external gap keeps the callout clear of the content above it;
    // vertical breathing room inside the text remains in the larger leading.
    expect(props?.spacing?.before).toBeGreaterThan(0);
    expect(props?.spacing?.line).toBeGreaterThan(240);
    expect(props?.spacing?.after).toBeGreaterThan(0);
  });

  it("puts the bar on the measure edge and spends all horizontal space inside", async () => {
    // The hard separation from a continued table is the page boundary after
    // the callout. Styling no longer needs a defensive outside inset.
    const props = await quoteStyle();
    const pt = (twips: number) => twips / 20;
    const barAndPadding =
      (props!.borders!.left!.size ?? 0) / 8 + (props!.borders!.left!.space ?? 0);
    // Rounding px independently into twips / eighth-points leaves under a
    // tenth of a point, visually the measure edge.
    expect(Math.abs(pt(props!.indent!.left!) - barAndPadding)).toBeLessThan(0.1);
    expect(pt(props!.indent!.right!)).toBeGreaterThan(10);
  });

  it("draws the bar at the width the sheet asks for, and inside the measure", async () => {
    const props = await quoteStyle();
    // `border-left: 5px` in eighths of a point — not `5 * 8`, which is five *points* and made the bar
    // 27% too thick.
    expect(props?.borders?.left?.size).toBe(Math.round(5 * (11 / 14) * 8));
    // Bar plus padding must fit inside the indent, or the bar is drawn off the page.
    const barPt = (props!.borders!.left!.size ?? 0) / 8 + (props!.borders!.left!.space ?? 0);
    expect(barPt).toBeLessThanOrEqual(props!.indent!.left! / 20);
  });

  it("sets a quote a step below body text", async () => {
    const doc = await markdownToDocx("> quoted");
    expect(doc.styles?.find(s => s.styleId === "Quote")?.runProperties?.size).toBeLessThan(
      doc.docDefaults!.runProperties!.size!
    );
  });

  it("gives a code block a background but no frame", async () => {
    const code = (await markdownToDocx("```ts\nx\n```")).styles?.find(
      s => s.styleId === "CodeBlock"
    );
    expect(code?.paragraphProperties?.shading?.fill).toBe("F1F1F1");
    expect(code?.paragraphProperties?.borders).toBeUndefined();
  });
});

describe("ruling lines", () => {
  /**
   * An invisible border is still a real line in the PDF, and a PDF reader's table detector looks for
   * exactly that. Declaring these edges in the page colour — to suppress Word's non-printing
   * on-screen gridlines — drew a vertical rule down the column boundary; Preview built a grid from
   * it, extended the grid past the last horizontal rule, and split the block *below* the table into
   * two pieces at that boundary. A callout was cut in half and reported as two cells of a row it had
   * nothing to do with.
   *
   * Word's gridlines never print. These tests pin the trade in the direction that survived.
   */
  it("draws no vertical rule, but closes the table after its last row", async () => {
    const doc = await markdownToDocx("| A | B |\n|---|---|\n| 1 | 2 |");
    const borders = (doc.body[0] as Table).properties?.borders;
    expect(borders?.insideV?.style).toBe("none");
    for (const edge of ["top", "left", "right"] as const) {
      expect(borders?.[edge]?.style, edge).toBe("none");
    }
    // Row separators and the closing edge are the only rules. Without the latter Preview takes the
    // next full-width rule on the page as the table's bottom, so an unrelated callout between them
    // becomes a phantom final row — two cells when a vertical divider existed, one when it did not.
    expect(borders?.insideH?.style).toBe("single");
    expect(borders?.bottom?.style).toBe("single");
    expect(borders?.bottom?.color).toBe(borders?.insideH?.color);
    expect(borders?.bottom?.size).toBe(borders?.insideH?.size);
  });

  it("puts that closing rule on every cell of the last row", async () => {
    // Asserted after layout rather than only in the style: border conflict resolution is what reaches
    // the PDF, and that is where an apparently correct declaration can disappear.
    const doc = await markdownToDocx("| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |");
    const table = layoutDocumentFull(doc).pages[0]!.content.find(block => block.type === "table")!;
    const last = Math.max(...table.cells.map(cell => cell.row));
    const cells = table.cells.filter(cell => cell.row === last);
    expect(cells).toHaveLength(2);
    for (const cell of cells) {
      expect(cell.borders?.bottom?.width).toBeGreaterThan(0);
      expect(cell.borders?.bottom?.color).toBe("D1D1D1");
    }
  });

  it("leaves a table with no vertical rule at all", async () => {
    // Measured on the laid-out page rather than the declarations, because a table border reaches the
    // page through cell-level resolution and that is where the earlier version leaked: every cell
    // came out with `right: FFFFFF`, a real line down the column boundary that a reader's table
    // detector then built a grid from.
    const doc = await markdownToDocx("| A | B |\n|---|---|\n| 1 | 2 |\n\n> quoted");
    const vertical: string[] = [];
    for (const block of layoutDocumentFull(doc).pages[0]!.content) {
      if (block.type !== "table") {
        continue;
      }
      for (const cell of block.cells ?? []) {
        for (const edge of ["left", "right"] as const) {
          const border = cell.borders?.[edge];
          if (border) {
            vertical.push(`${edge}:${border.color}`);
          }
        }
      }
    }
    expect(vertical).toEqual([]);
  });
});

describe("what follows a table", () => {
  const TABLE = "| A | B |\n|---|---|\n| 1 | 2 |\n";

  /** `spacing.before` on the block at `index`, in twips. */
  const before = (doc: Awaited<ReturnType<typeof markdownToDocx>>, index: number) => {
    const block = doc.body[index];
    return block?.type === "paragraph" ? (block.properties?.spacing?.before ?? 0) : undefined;
  };

  it("gives a plain paragraph the margin the table cannot carry", async () => {
    // This stylesheet's rhythm is bottom margins only, and OOXML has no space-after on an inline
    // table — so the chain breaks and the next block sits flush against the last rule.
    const doc = await markdownToDocx(`${TABLE}\nAfter.`);
    expect(before(doc, 1)).toBeGreaterThan(0);
  });

  it("leaves a block that does not follow a table alone", async () => {
    const doc = await markdownToDocx("One.\n\nTwo.");
    expect(before(doc, 1)).toBe(0);
  });

  it("does not reach past a table to the block after that", async () => {
    const doc = await markdownToDocx(`${TABLE}\nFirst.\n\nSecond.`);
    expect(before(doc, 1)).toBeGreaterThan(0);
    expect(before(doc, 2)).toBe(0);
  });

  it("puts a paragraph between two adjacent tables", async () => {
    // ECMA-376 §17.13.5.34: a `<w:tbl>` must be followed by a paragraph before the next may begin,
    // or Word collapses the pair into one malformed table — which is exactly the "it took the
    // layout of the table above" failure.
    const doc = await markdownToDocx(`${TABLE}\n${TABLE}`);
    expect(doc.body.map(b => b.type)).toEqual(["table", "paragraph", "table"]);
    const xml = await part(await toBuffer(doc), "word/document.xml");
    const firstEnd = xml.indexOf("</w:tbl>");
    const secondStart = xml.indexOf("<w:tbl>", firstEnd);
    expect(secondStart).toBeGreaterThan(firstEnd);
    expect(xml.slice(firstEnd + "</w:tbl>".length, secondStart)).toContain("<w:p");
  });

  it("makes that separator cost a hairline, not a blank line", async () => {
    // The writer synthesises a bare `<w:p/>` when it has to, which is correct and expensive: with no
    // properties it inherits the body's 22px line and 16px bottom margin, so a spec formality costs
    // 30 points of blank page. Emitted here instead it can be sized — an exact 1pt line, and the
    // table's own bottom margin as the gap that belongs there anyway.
    const doc = await markdownToDocx(`${TABLE}\n${TABLE}`);
    const separator = layoutDocumentFull(doc).pages[0]!.content[1]!;
    expect(separator.type).toBe("paragraph");
    // The 16px margin, plus a point of line. A bare `<w:p/>` here measured 26.
    expect(separator.rect.height).toBeLessThan(16);
    expect(separator.rect.height).toBeGreaterThan(12);
  });
});

describe("list leading", () => {
  it("leads a list item like the paragraphs around it", async () => {
    // `ListParagraph` declares only `spacing.after`. Style resolution replaced
    // the whole `w:spacing` value instead of merging its attributes, so the
    // document default's `line` was dropped and every list item fell back to
    // single spacing — 13.2pt against the 17.27pt of the prose beside it, a 24%
    // difference that made every list look cramped.
    const doc = await markdownToDocx("Body text here.\n\n- first item\n- second item");
    const page = layoutDocumentFull(doc).pages[0];
    const paragraphs = page.content.filter(b => b.type === "paragraph");
    const leadingOf = (index: number): number => {
      const lines = paragraphs[index].lines;
      expect(lines.length).toBeGreaterThan(0);
      return lines[0].height;
    };
    expect(leadingOf(1)).toBeCloseTo(leadingOf(0), 5);
    // 11pt × 1.2 natural × 314/240 — the `--markdown-line-height: 22px` default.
    expect(leadingOf(1)).toBeCloseTo(17.27, 2);
  });
});

describe("fenced code blocks", () => {
  /** Courier's advance is 600/1000 em, so a column is 0.6 × the size. */
  const COLUMN_RATIO = 0.6;
  /** Letter content width less the `CodeBlock` style's 16px of padding a side. */
  const MEASURE_PT = (12240 - 2 * 1440) / 20 - (2 * Math.round(16 * (11 / 14) * 20)) / 20;

  /** The run size, in half-points, of the single code block in `markdown`. */
  async function codeSize(markdown: string, options?: Parameters<typeof markdownToDocx>[1]) {
    const doc = await markdownToDocx(markdown, options);
    const block = doc.body.find(b => b.type === "paragraph" && b.properties?.style === "CodeBlock");
    expect(block).toBeDefined();
    if (block?.type !== "paragraph") {
      throw new Error("unreachable");
    }
    const sizes = new Set(
      block.children.flatMap(child =>
        "content" in child && child.content.some(c => c.type === "text" && c.text.length > 0)
          ? [child.properties?.size]
          : []
      )
    );
    // One size for the whole block: a listing set at two sizes is not a listing.
    expect(sizes.size).toBe(1);
    return [...sizes][0]!;
  }

  const fence = (lines: readonly string[]) => "```\n" + lines.join("\n") + "\n```\n";

  it("leaves a block that already fits at the body size", async () => {
    expect(await codeSize(fence(["short line", "another"]))).toBe(22);
  });

  it("sets a block small enough that its longest line does not wrap", async () => {
    // `pre` in the preview scrolls; a page cannot, and wrapping one line destroys
    // the column alignment of the whole block — which for a directory tree or an
    // aligned comment column is most of what it was for.
    const columns = 108;
    const size = await codeSize(fence(["x".repeat(columns), "y"]));
    expect(size).toBeLessThan(22);
    expect((size / 2) * COLUMN_RATIO * columns).toBeLessThanOrEqual(MEASURE_PT);
    // And no smaller than it has to be: one more half-point would overflow.
    expect(((size + 1) / 2) * COLUMN_RATIO * columns).toBeGreaterThan(MEASURE_PT);
  });

  it("stops shrinking at the floor and lets the rest wrap", async () => {
    // 60% of the code size. Without a floor one 400-character line would render
    // the other twenty lines of its block unreadable.
    expect(await codeSize(fence(["x".repeat(4000)]))).toBe(Math.floor(22 * 0.6));
  });

  it("honours codeBlockFit: wrap", async () => {
    expect(await codeSize(fence(["x".repeat(108)]), { codeBlockFit: "wrap" })).toBe(22);
  });

  it("fits to the measure the caller declares", async () => {
    // 90 columns needs shrinking on Letter but not to the floor, so the three
    // measures give three different answers.
    const lines = fence(["x".repeat(90)]);
    const narrow = await codeSize(lines, { contentWidth: 4680 });
    const wide = await codeSize(lines, { contentWidth: 20000 });
    const letter = await codeSize(lines);
    expect(letter).toBeLessThan(22);
    expect(narrow).toBeLessThan(letter);
    expect(wide).toBe(22);
  });

  it("sizes each block for itself", async () => {
    const doc = await markdownToDocx(
      fence(["x".repeat(108)]) + "\ntext between\n\n" + fence(["short"])
    );
    const sizes = doc.body
      .filter(b => b.type === "paragraph" && b.properties?.style === "CodeBlock")
      .map(b =>
        b.type === "paragraph" ? b.children.find(c => "content" in c)?.properties?.size : undefined
      );
    expect(sizes).toHaveLength(2);
    expect(sizes[0]).toBeLessThan(22);
    expect(sizes[1]).toBe(22);
  });
});

describe("declared content width", () => {
  it("sizes table columns against it", async () => {
    const table = "| a | b |\n| - | - |\n| one | two |";
    const doc = await markdownToDocx(table, { contentWidth: 5000 });
    const t = doc.body.find(b => b.type === "table");
    expect(t?.type).toBe("table");
    if (t?.type !== "table") {
      throw new Error("unreachable");
    }
    expect(t.columnWidths?.reduce((a, b) => a + b, 0)).toBe(5000);
  });
});
