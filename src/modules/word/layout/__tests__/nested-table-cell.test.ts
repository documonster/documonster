import { Build, Document } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { LayoutTable } from "@word/layout/layout-model";
import type { Paragraph, Table } from "@word/types";
/**
 * The paragraph that closes a cell after a nested table.
 *
 * A cell has to end with a paragraph, so a nested table at the end of a cell
 * is followed by an empty one. Word gives that paragraph no height: in final5
 * M2-noimg (Word for Mac 16.114) row 1 of the first table is 97.5pt tall, where
 * laying the empty paragraph out as a line made it 121.4pt, and every heading
 * after such a table sat 47.7pt (two rows of one line each) lower than Word's.
 * The same holds after a fixed-layout table, fitting the cell or not (final7
 * R0–Rg; final6 Qae's extra 15pt is its last nested row wrapping to two lines).
 *
 * The nested table itself keeps its own width (final7 R*), see
 * `nestedTableWidthTwips`.
 */
import { describe, expect, it } from "vitest";

function rowHeight(cellContent: (Paragraph | Table)[]): number {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addTable(doc, Build.table([Build.row([Build.cell(cellContent)])]));
  const table = layoutDocumentFull(doc).pages[0]!.content.find(
    (c): c is LayoutTable => c.type === "table"
  )!;
  return table.cells[0]!.rect.height;
}

const nested = (): Table =>
  Build.table([Build.row([Build.cell("a"), Build.cell("b")])], { borders: Build.gridBorders() });

/** A fixed-layout nested table with two columns of `twips` each. */
const fixedNested = (twips: number): Table =>
  Build.table(
    [Build.row([Build.cell("a"), Build.cell("b")])],
    { borders: Build.gridBorders(), layout: "fixed" },
    [twips, twips]
  );

describe("cell-end paragraph after a nested table", () => {
  it("takes no height", () => {
    const withEnd = rowHeight([Build.textParagraph("text"), nested(), Build.paragraph([])]);
    const without = rowHeight([Build.textParagraph("text"), nested()]);
    expect(withEnd).toBeCloseTo(without, 6);
  });

  it("still takes its height when it holds text", () => {
    const withText = rowHeight([
      Build.textParagraph("text"),
      nested(),
      Build.textParagraph("more")
    ]);
    const without = rowHeight([Build.textParagraph("text"), nested()]);
    expect(withText).toBeGreaterThan(without + 10);
  });

  it("still takes its height when it is not the last paragraph", () => {
    const twoEmpty = rowHeight([
      Build.textParagraph("text"),
      nested(),
      Build.paragraph([]),
      Build.paragraph([])
    ]);
    const without = rowHeight([Build.textParagraph("text"), nested()]);
    expect(twoEmpty).toBeGreaterThan(without + 10);
  });

  it("takes no height after a fixed-layout table that fits the cell", () => {
    const withEnd = rowHeight([
      Build.textParagraph("text"),
      fixedNested(2000),
      Build.paragraph([])
    ]);
    const without = rowHeight([Build.textParagraph("text"), fixedNested(2000)]);
    expect(withEnd).toBeCloseTo(without, 6);
  });

  it("takes no height after a fixed-layout table wider than the cell", () => {
    const withEnd = rowHeight([
      Build.textParagraph("text"),
      fixedNested(6000),
      Build.paragraph([])
    ]);
    const without = rowHeight([Build.textParagraph("text"), fixedNested(6000)]);
    expect(withEnd).toBeCloseTo(without, 6);
  });
});

/** Column widths of the table nested in the first cell of a two-column outer table. */
function nestedColumns(inner: Table, outer: { fixed: boolean }): number[] {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  const cell = (content: (Paragraph | Table)[]) =>
    Build.cell(content, outer.fixed ? { width: { value: 4680, type: "dxa" } } : undefined);
  Document.addTable(
    doc,
    Build.table(
      [
        Build.row([
          cell([Build.textParagraph("top"), inner, Build.paragraph([])]),
          cell([Build.textParagraph("x")])
        ])
      ],
      { borders: Build.gridBorders(), ...(outer.fixed ? { layout: "fixed" as const } : {}) },
      [4680, 4680]
    )
  );
  const table = layoutDocumentFull(doc).pages[0]!.content.find(
    (c): c is LayoutTable => c.type === "table"
  )!;
  const nested = table.cells[0]!.content.find((c): c is LayoutTable => c.type === "table")!;
  return nested.cells.slice(0, 2).map(c => Math.round(c.rect.width * 100) / 100);
}

/** A nested 1×2 table of one-word cells over a `twips`-per-column grid. */
const innerTable = (twips: number, fixed: boolean, tblW?: number): Table =>
  Build.table(
    [Build.row([Build.cell("a"), Build.cell("b")])],
    {
      borders: Build.gridBorders(),
      ...(fixed ? { layout: "fixed" as const } : {}),
      ...(tblW ? { width: { value: tblW, type: "dxa" as const } } : {})
    },
    [twips, twips]
  );

// Widths Word for Mac 16.114 measured (final7 R*), in a 4464-twip cell.
describe("a nested table's width", () => {
  it("is a fixed-layout table's grid, not the cell's width", () => {
    expect(nestedColumns(innerTable(2232, true), { fixed: true })).toEqual([111.6, 111.6]);
    expect(nestedColumns(innerTable(2000, true), { fixed: true })).toEqual([100, 100]);
  });

  it("overflows the cell when the grid is wider", () => {
    expect(nestedColumns(innerTable(2340, true), { fixed: true })).toEqual([117, 117]);
    expect(nestedColumns(innerTable(3000, true), { fixed: true })).toEqual([150, 150]);
  });

  it("is the declared table width over the grid", () => {
    expect(nestedColumns(innerTable(2000, true, 3000), { fixed: true })).toEqual([75, 75]);
    expect(nestedColumns(innerTable(3000, true, 3000), { fixed: true })).toEqual([75, 75]);
    expect(nestedColumns(innerTable(1500, false, 3000), { fixed: true })).toEqual([75, 75]);
  });

  // In an auto-fit outer table Word fits an auto-fit nested one to its content
  // within its declared width instead: the R3c-nest probe pins that.

  it("is an auto-fit table's grid inside a fixed-layout table", () => {
    expect(nestedColumns(innerTable(1000, false), { fixed: true })).toEqual([50, 50]);
    expect(nestedColumns(innerTable(2000, false), { fixed: true })).toEqual([100, 100]);
  });
});
