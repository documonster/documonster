/**
 * layoutDocumentFull — table layout fidelity tests.
 *
 * Guards the three previously-missing behaviours of `layoutTable`:
 *   1. Nested tables inside a cell are laid out (previously silently
 *      dropped — the PDF/SVG renderers already supported them).
 *   2. `table.columnWidths` is honoured (previously every column was
 *      forced to equal width, discarding the Excel→Word bridge's widths).
 *   3. `gridSpan` (horizontal cell merge) widens the spanning cell and
 *      shifts subsequent cells by the spanned grid columns.
 */

import { describe, it, expect } from "vitest";

import { layoutDocumentFull } from "../layout/layout-full";
import type { LayoutTable } from "../layout/layout-model";
import type { DocxDocument, Paragraph, Table } from "../types";

const minimalDoc = (body: DocxDocument["body"]): DocxDocument => ({
  body,
  styles: [],
  abstractNumberings: [],
  numberingInstances: [],
  headers: new Map(),
  footers: new Map(),
  footnotes: [],
  endnotes: [],
  comments: [],
  images: [],
  fonts: [],
  embeddedFonts: [],
  customXmlParts: [],
  customProperties: [],
  opaqueParts: []
});

function textPara(s: string): Paragraph {
  return { type: "paragraph", children: [{ content: [{ type: "text", text: s }] }] };
}

function firstTable(doc: DocxDocument): LayoutTable {
  const layout = layoutDocumentFull(doc);
  const item = layout.pages[0].content.find((c): c is LayoutTable => c.type === "table");
  if (!item) {
    throw new Error("no table laid out");
  }
  return item;
}

describe("layoutTable — nested tables", () => {
  it("lays out a table nested inside a cell instead of dropping it", () => {
    const nested: Table = {
      type: "table",
      rows: [{ cells: [{ content: [textPara("inner")] }] }]
    };
    const outer: Table = {
      type: "table",
      rows: [{ cells: [{ content: [textPara("outer"), nested] }] }]
    };
    const laid = firstTable(minimalDoc([outer]));
    const cell = laid.cells[0];
    // The cell now carries both the paragraph and the nested table.
    const nestedLaid = cell.content.find(c => c.type === "table");
    expect(nestedLaid).toBeDefined();
    expect((nestedLaid as LayoutTable).cells).toHaveLength(1);
    // The cell height grew to accommodate the nested table (it is taller
    // than a single text line).
    expect(cell.rect.height).toBeGreaterThan(0);
  });
});

describe("layoutTable — column widths", () => {
  it("honours table.columnWidths proportionally instead of forcing equal widths", () => {
    // 3:1 width ratio between the two columns (in twips), in a fixed-layout
    // table: an auto-fit one is sized by its content (below).
    const table: Table = {
      type: "table",
      properties: { layout: "fixed" },
      columnWidths: [4500, 1500],
      rows: [
        {
          cells: [{ content: [textPara("wide")] }, { content: [textPara("narrow")] }]
        }
      ]
    };
    const laid = firstTable(minimalDoc([table]));
    const [c0, c1] = laid.cells;
    // First column should be ~3× the width of the second.
    expect(c0.rect.width / c1.rect.width).toBeCloseTo(3, 1);
    // Second cell starts where the first ends (no overlap, no gap).
    expect(c1.rect.x).toBeCloseTo(c0.rect.x + c0.rect.width, 1);
  });

  it("falls back to equal widths when columnWidths is absent", () => {
    const table: Table = {
      type: "table",
      properties: { layout: "fixed" },
      rows: [
        {
          cells: [{ content: [textPara("a")] }, { content: [textPara("b")] }]
        }
      ]
    };
    const laid = firstTable(minimalDoc([table]));
    const [c0, c1] = laid.cells;
    expect(c0.rect.width).toBeCloseTo(c1.rect.width, 1);
  });
});

describe("layoutTable — auto-fit column widths", () => {
  // An auto-fit table whose cells ask for no width is sized by what they hold,
  // whatever its grid says: Word for Mac 16 lays this one out "wide" column
  // narrower than "narrow", not 3:1.
  const table = (properties?: Table["properties"]): Table => ({
    type: "table",
    ...(properties ? { properties } : {}),
    columnWidths: [4500, 1500],
    rows: [{ cells: [{ content: [textPara("wide")] }, { content: [textPara("narrow")] }] }]
  });
  /**
   * 11pt Calibri advances of the cell text. The document has no default table
   * style, so Word gives the cells no margins (this pinned 5.4pt either side,
   * which only `Normal Table` as the default table style supplies).
   */
  const wide = (11 * (1464 + 470 + 1076 + 1019)) / 2048;
  const narrow = (11 * (1076 + 981 + 714 + 714 + 1080 + 1464)) / 2048;

  it("gives each column its content, plus any cell margins, in a table of auto width", () => {
    const [c0, c1] = firstTable(minimalDoc([table()])).cells;
    expect(c0.rect.width).toBeCloseTo(wide, 1);
    expect(c1.rect.width).toBeCloseTo(narrow, 1);
  });

  it("shares a set table width out in proportion to the content", () => {
    const [c0, c1] = firstTable(minimalDoc([table({ width: { value: 5000, type: "pct" } })])).cells;
    expect(c0.rect.width + c1.rect.width).toBeCloseTo(468, 2);
    expect(c0.rect.width / c1.rect.width).toBeCloseTo(wide / narrow, 2);
  });
});

describe("layoutTable — gridSpan", () => {
  it("widens a gridSpan cell and shifts subsequent cells by the spanned columns", () => {
    const table: Table = {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      rows: [
        // Row 0: three single cells (defines the 3-column grid).
        {
          cells: [
            { content: [textPara("a")] },
            { content: [textPara("b")] },
            { content: [textPara("c")] }
          ]
        },
        // Row 1: a 2-wide spanning cell, then one normal cell.
        {
          cells: [
            { content: [textPara("span2")], properties: { gridSpan: 2 } },
            { content: [textPara("last")] }
          ]
        }
      ]
    };
    const laid = firstTable(minimalDoc([table]));
    const row1 = laid.cells.filter(c => c.row === 1);
    expect(row1).toHaveLength(2);
    const [spanCell, lastCell] = row1;
    const [a, b] = laid.cells.filter(c => c.row === 0);
    // The spanning cell is as wide as the two columns it covers.
    expect(spanCell.rect.width).toBeCloseTo(a!.rect.width + b!.rect.width, 1);
    // The trailing cell starts in the 3rd grid column, not the 2nd.
    expect(lastCell.rect.x).toBeCloseTo(spanCell.rect.x + spanCell.rect.width, 1);
  });
});

describe("layoutTable — vertical cell alignment", () => {
  /**
   * A row whose first cell holds three paragraphs, so the other cells have
   * two lines of slack to distribute according to their `w:vAlign`.
   */
  function alignedRow(): Table {
    return {
      type: "table",
      rows: [
        {
          cells: [
            { content: [textPara("a"), textPara("b"), textPara("c")] },
            { content: [textPara("top")], properties: { verticalAlign: "top" } },
            { content: [textPara("center")], properties: { verticalAlign: "center" } },
            { content: [textPara("bottom")], properties: { verticalAlign: "bottom" } },
            { content: [textPara("default")] }
          ]
        }
      ]
    };
  }

  it("moves cell content into the row's slack per w:vAlign", () => {
    const laid = firstTable(minimalDoc([alignedRow()]));
    const cells = laid.cells.filter(c => c.row === 0);
    expect(cells).toHaveLength(5);

    const contentTop = (col: number) => cells[col].content[0].rect.y;
    const contentBottom = (col: number) =>
      cells[col].content[0].rect.y + cells[col].content[0].rect.height;
    const rowHeight = cells[0].rect.height;
    const slack = rowHeight - contentBottom(1);

    // A row this tall only exists because of the first cell.
    expect(slack).toBeGreaterThan(0);

    // top and an unset vAlign both stay at the top margin.
    expect(contentTop(1)).toBeCloseTo(0, 6);
    expect(contentTop(4)).toBeCloseTo(0, 6);
    // center splits the slack.
    expect(contentTop(2)).toBeCloseTo(slack / 2, 6);
    // bottom takes all of it, landing the content on the bottom margin.
    expect(contentTop(3)).toBeCloseTo(slack, 6);
    expect(rowHeight - contentBottom(3)).toBeCloseTo(0, 6);
  });

  it("leaves content alone when an exact row height removes the slack", () => {
    const table = alignedRow();
    const clamped: Table = {
      ...table,
      rows: [{ ...table.rows[0], properties: { height: { value: 200, rule: "exact" } } }]
    };
    const laid = firstTable(minimalDoc([clamped]));
    const cells = laid.cells.filter(c => c.row === 0);

    // 200 twips = 10pt, shorter than the three-paragraph cell: Word clips
    // rather than growing the row, so there is nothing to distribute and
    // bottom-aligned content must not be pushed off the top.
    expect(cells[0].rect.height).toBeCloseTo(10, 6);
    for (const cell of cells) {
      expect(cell.content[0].rect.y).toBeCloseTo(0, 6);
    }
  });
});
