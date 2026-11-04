import { Build, Document } from "@word/index";
import { layoutDocumentFull } from "@word/layout/layout-full";
import type { LayoutParagraph, LayoutTable } from "@word/layout/layout-model";
import type { DocxDocument, TableRowProperties } from "@word/types";
/**
 * A table row breaks across pages, as Word breaks it, unless it says it
 * cannot (`w:cantSplit`). Word for Mac 16.114 on final5 M2-nonest breaks the
 * first row of the fourth table across pages 4 and 5; moving rows only whole
 * put everything after it a page late.
 */
import { describe, expect, it } from "vitest";

/** A document whose page has `room` pt left above a two-row table. */
function docWith(
  rows: { text: string; props?: TableRowProperties }[],
  room = 100,
  widowControl = true
): DocxDocument {
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  // Letter with 1in margins: 648pt of text area.
  // Built by hand: a `room` above the text area makes `before` negative, which the layout takes
  // and `Build.paragraph` refuses (`w:before` is unsigned).
  Document.addParagraph(doc, {
    type: "paragraph",
    properties: { spacing: { before: Math.round((648 - room - 30) * 20), after: 0 } },
    children: [Build.text("lead")]
  });
  Document.addTable(
    doc,
    Build.table(
      rows.map(r =>
        Build.row(
          [
            Build.cell([
              Build.paragraph(
                [Build.text(r.text)],
                widowControl ? undefined : { widowControl: false }
              )
            ])
          ],
          r.props
        )
      ),
      { borders: Build.gridBorders() }
    )
  );
  return doc;
}

const long = "word ".repeat(300);

function tables(doc: DocxDocument): LayoutTable[][] {
  return layoutDocumentFull(doc).pages.map(p =>
    p.content.filter((c): c is LayoutTable => c.type === "table")
  );
}

function lineCount(table: LayoutTable, row: number): number {
  return table.cells
    .filter(c => c.row === row)
    .flatMap(c => c.content)
    .reduce((n, b) => n + (b.type === "paragraph" ? (b as LayoutParagraph).lines.length : 0), 0);
}

describe("table rows across pages", () => {
  it("breaks a row that does not fit, line by line", () => {
    const [first, second] = tables(docWith([{ text: "short" }, { text: long }]));
    const head = first![0]!;
    const tail = second![0]!;
    // Row 1 starts on the first page and continues on the second.
    expect(lineCount(head, 1)).toBeGreaterThan(1);
    expect(lineCount(tail, 1)).toBeGreaterThan(1);
    const total = tables(docWith([{ text: "short" }, { text: long }], 2000)).flat()[0]!;
    expect(lineCount(head, 1) + lineCount(tail, 1)).toBe(lineCount(total, 1));
  });

  it("moves a w:cantSplit row whole", () => {
    const [first, second] = tables(
      docWith([{ text: "short" }, { text: long, props: { cantSplit: true } }])
    );
    expect(first![0]!.cells.some(c => c.row === 1)).toBe(false);
    expect(second![0]!.cells.some(c => c.row === 1)).toBe(true);
  });

  it("repeats header rows above the continued piece", () => {
    const [, second] = tables(
      docWith([{ text: "header", props: { tableHeader: true } }, { text: long }])
    );
    const tail = second![0]!;
    expect(tail.cells.filter(c => c.row === 0)).toHaveLength(1);
    const header = tail.cells.find(c => c.row === 0)!;
    const piece = tail.cells.find(c => c.row === 1)!;
    expect(piece.rect.y).toBeCloseTo(header.rect.y + header.rect.height, 6);
  });

  it("keeps a cell paragraph's widow/orphan rule when it breaks", () => {
    // Three lines with room for two: widow control cannot leave one behind and
    // cannot carry one alone, so the row moves whole.
    const three = "word ".repeat(45);
    const lines = lineCount(tables(docWith([{ text: three }], 2000)).flat()[0]!, 0);
    expect(lines).toBe(3);
    const pages = tables(docWith([{ text: "short" }, { text: three }], 45));
    expect(pages[0]![0]!.cells.some(c => c.row === 1)).toBe(false);
    // Without widow control the same row breaks after the line that fits.
    const loose = tables(docWith([{ text: "short" }, { text: three }], 45, false));
    expect(loose[0]![0]!.cells.some(c => c.row === 1)).toBe(true);
  });
});
