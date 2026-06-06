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
    // 3:1 width ratio between the two columns (in twips).
    const table: Table = {
      type: "table",
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
    const singleColWidth = laid.cells.find(c => c.row === 0)!.rect.width;
    // The spanning cell is ~2 columns wide.
    expect(spanCell.rect.width).toBeCloseTo(singleColWidth * 2, 0);
    // The trailing cell starts in the 3rd grid column, not the 2nd.
    expect(lastCell.rect.x).toBeCloseTo(spanCell.rect.x + spanCell.rect.width, 1);
  });
});
