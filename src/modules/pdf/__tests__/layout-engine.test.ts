/**
 * Focused tests for layout engine pagination helpers and page placement.
 */
import { FontManager } from "@pdf/font/font-manager";
import { layoutSheet, paginateRows } from "@pdf/render/layout-engine";
import { borderStyleToLineWidth } from "@pdf/render/style-converter";
import type {
  LayoutCell,
  LayoutPage,
  PdfCellData,
  PdfColumnData,
  PdfRowData,
  PdfSheetData,
  ResolvedPdfOptions,
  PdfPageSetupData
} from "@pdf/types";
import { PdfCellType } from "@pdf/types";
import { describe, expect, it } from "vitest";

describe("layout-engine pagination", () => {
  it("should repeat header rows on subsequent pages", () => {
    const pages = paginateRows([10, 10, 10, 10], 25, 1, new Set());
    expect(pages).toEqual([
      [0, 1],
      [0, 2],
      [0, 3]
    ]);
  });

  it("should avoid emitting repeat-row-only pages when headers cannot fit with body rows", () => {
    const pages = paginateRows([30, 30, 10], 35, 2, new Set());
    expect(pages).toEqual([[0], [1], [2]]);
  });

  it("should honor manual row breaks", () => {
    const pages = paginateRows([10, 10, 10, 10], 100, 0, new Set([2]));
    expect(pages).toEqual([
      [0, 1],
      [2, 3]
    ]);
  });

  // A manual break used to stay "active" after the repeated title rows were
  // re-added, so the loop flushed title-only pages forever and the export died
  // with a heap OOM. Each of these terminates and yields finite pages.
  it("should terminate when repeat rows meet a manual break", () => {
    expect(paginateRows([10, 10, 10, 10], 100, 1, new Set([2]))).toEqual([
      [0, 1],
      [0, 2, 3]
    ]);
  });

  it("should terminate with multiple breaks and a repeated prefix", () => {
    expect(paginateRows([10, 10, 10, 10], 100, 1, new Set([1, 3]))).toEqual([
      [0],
      [0, 1, 2],
      [0, 3]
    ]);
  });

  it("should terminate when the repeat count spans the break", () => {
    expect(paginateRows([10, 10, 10, 10], 100, 2, new Set([1]))).toEqual([[0], [0, 1, 2, 3]]);
  });
});

// =============================================================================
// Page placement — "Center on page" (Excel printOptions)
// =============================================================================

const MARGIN = 72;
const PAGE_WIDTH = 595.28;
const PAGE_HEIGHT = 841.89;

function buildSheet(
  colCount: number,
  rowCount: number,
  colWidth: number,
  pageSetup?: PdfPageSetupData
): PdfSheetData {
  const columns = new Map<number, PdfColumnData>();
  for (let c = 1; c <= colCount; c++) {
    columns.set(c, { width: colWidth });
  }
  const rows = new Map<number, PdfRowData>();
  for (let r = 1; r <= rowCount; r++) {
    const cells = new Map<number, PdfCellData>();
    for (let c = 1; c <= colCount; c++) {
      cells.set(c, {
        type: PdfCellType.String,
        value: `R${r}C${c}`,
        text: `R${r}C${c}`,
        col: c
      });
    }
    rows.set(r, { cells });
  }
  return {
    name: "Sheet1",
    bounds: { top: 1, left: 1, bottom: rowCount, right: colCount },
    columns,
    rows,
    pageSetup
  };
}

function buildOptions(overrides: Partial<ResolvedPdfOptions> = {}): ResolvedPdfOptions {
  return {
    pageSize: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
    orientation: "portrait",
    margins: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
    ignorePrintArea: false,
    fitToPage: true,
    scale: 1,
    fitToWidth: 0,
    fitToHeight: 0,
    showGridLines: false,
    gridLineColor: { r: 0.8, g: 0.8, b: 0.8 },
    showRowColHeaders: false,
    horizontalCentered: false,
    verticalCentered: false,
    pageOrder: "downThenOver",
    blackAndWhite: false,
    draft: false,
    errors: "displayed",
    cellComments: "none",
    repeatRows: false,
    repeatCols: false,
    defaultFontFamily: "Helvetica",
    defaultFontSize: 11,
    showSheetNames: false,
    showPageNumbers: false,
    includeHeadersFooters: false,
    headerMargin: 21.6,
    footerMargin: 21.6,
    sourceFileName: "",
    sourceFilePath: "",
    headerFooterDate: new Date(0),
    title: "",
    author: "",
    subject: "",
    creator: "documonster",
    ...overrides
  };
}

describe("layout-engine page placement", () => {
  it("should left-align a narrow sheet by default (issue #203)", async () => {
    const pages = await layoutSheet(buildSheet(2, 5, 12), buildOptions(), new FontManager());

    expect(pages).toHaveLength(1);
    expect(pages[0].columnOffsets[0]).toBeCloseTo(MARGIN, 5);
  });

  it("should center a narrow sheet horizontally when horizontalCentered is set", async () => {
    const sheet = buildSheet(2, 5, 12);
    const pages = await layoutSheet(
      sheet,
      buildOptions({ horizontalCentered: true }),
      new FontManager()
    );

    const contentWidth = PAGE_WIDTH - 2 * MARGIN;
    const tableWidth = pages[0].columnWidths.reduce((s, w) => s + w, 0);
    expect(tableWidth).toBeLessThan(contentWidth);
    expect(pages[0].columnOffsets[0]).toBeCloseTo(MARGIN + (contentWidth - tableWidth) / 2, 5);
  });

  it("should keep all horizontal page groups left-aligned by default", async () => {
    // 30 columns at width 12 overflow the content area, so the last column
    // group is narrower than the page — it must still start at the margin.
    const pages = await layoutSheet(
      buildSheet(30, 3, 12),
      buildOptions({ fitToPage: false }),
      new FontManager()
    );

    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages) {
      expect(page.columnOffsets[0]).toBeCloseTo(MARGIN, 5);
    }
  });

  it("should top-align rows by default", async () => {
    const pages = await layoutSheet(buildSheet(2, 5, 12), buildOptions(), new FontManager());

    expect(pages[0].rowYPositions[0]).toBeCloseTo(PAGE_HEIGHT - MARGIN, 5);
  });

  it("should center rows vertically when verticalCentered is set", async () => {
    const sheet = buildSheet(2, 5, 12);
    const pages = await layoutSheet(
      sheet,
      buildOptions({ verticalCentered: true }),
      new FontManager()
    );

    const availableHeight = PAGE_HEIGHT - 2 * MARGIN;
    const tableHeight = pages[0].rowHeights.reduce((s, h) => s + h, 0);
    expect(pages[0].rowYPositions[0]).toBeCloseTo(
      PAGE_HEIGHT - MARGIN - (availableHeight - tableHeight) / 2,
      5
    );
  });

  it("should derive centering from the sheet pageSetup when no option is given", async () => {
    // resolveOptions() (pdf-exporter) performs the pageSetup fallback; this
    // asserts the layout engine honours whatever the resolved flag says.
    const sheet = buildSheet(2, 5, 12, { horizontalCentered: true });
    const left = await layoutSheet(sheet, buildOptions(), new FontManager());
    const centered = await layoutSheet(
      sheet,
      buildOptions({ horizontalCentered: true }),
      new FontManager()
    );

    expect(left[0].columnOffsets[0]).toBeCloseTo(MARGIN, 5);
    expect(centered[0].columnOffsets[0]).toBeGreaterThan(MARGIN);
  });
});

// =============================================================================
// Merged regions split by a page or column-group boundary (issue #226)
// =============================================================================

/**
 * A grid whose every cell carries a `thin` border, with `medium` on the outer
 * ring of the region under test. The two widths are what let an assertion say
 * *which* border a piece drew: the region's own outline, or the ordinary grid.
 */
function borderedSheet(opts: {
  rowCount: number;
  colCount: number;
  merges?: string[];
  /** Region whose outer ring gets `medium` borders instead of `thin`. */
  ring?: { top: number; left: number; bottom: number; right: number };
  /** Cells given a fill, keyed as "row:col". */
  fills?: string[];
  hiddenRows?: number[];
  hiddenCols?: number[];
  pageSetup?: PdfPageSetupData;
  /** Text per cell, keyed as "row:col". Cells outside the map are empty. */
  text?: Record<string, string>;
}): PdfSheetData {
  const thin = { style: "thin", color: { argb: "FF000000" } };
  const medium = { style: "medium", color: { argb: "FF000000" } };
  const ring = opts.ring;

  const columns = new Map<number, PdfColumnData>();
  for (let c = 1; c <= opts.colCount; c++) {
    columns.set(c, { width: 12, hidden: opts.hiddenCols?.includes(c) || undefined });
  }

  const rows = new Map<number, PdfRowData>();
  for (let r = 1; r <= opts.rowCount; r++) {
    const cells = new Map<number, PdfCellData>();
    for (let c = 1; c <= opts.colCount; c++) {
      const inRing =
        ring !== undefined &&
        r >= ring.top &&
        r <= ring.bottom &&
        c >= ring.left &&
        c <= ring.right;
      const text = opts.text?.[`${r}:${c}`] ?? "";
      cells.set(c, {
        type: PdfCellType.String,
        value: text,
        text,
        col: c,
        style: {
          border: {
            top: inRing && r === ring.top ? medium : thin,
            bottom: inRing && r === ring.bottom ? medium : thin,
            left: inRing && c === ring.left ? medium : thin,
            right: inRing && c === ring.right ? medium : thin
          },
          fill: opts.fills?.includes(`${r}:${c}`)
            ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } }
            : undefined
        }
      });
    }
    rows.set(r, { cells, hidden: opts.hiddenRows?.includes(r) });
  }

  return {
    name: "Sheet1",
    bounds: { top: 1, left: 1, bottom: opts.rowCount, right: opts.colCount },
    columns,
    rows,
    merges: opts.merges,
    pageSetup: opts.pageSetup
  };
}

/** The laid-out cells whose left edge sits at the given column offset. */
function cellsInColumn(page: LayoutPage, columnIndex: number): LayoutCell[] {
  const x = page.columnOffsets[columnIndex];
  return page.cells.filter(cell => Math.abs(cell.rect.x - x) < 0.01);
}

// Taken from the converter rather than written as numbers: these assertions are
// about which of the two borders a piece drew, not about the points it maps to.
const MEDIUM_WIDTH = borderStyleToLineWidth("medium");
const THIN_WIDTH = borderStyleToLineWidth("thin");

describe("layout-engine merged regions across page boundaries", () => {
  /** A single-column merge four times taller than one page fits. */
  const tallMerge = () =>
    borderedSheet({
      rowCount: 200,
      colCount: 4,
      merges: ["A1:A200"],
      ring: { top: 1, left: 1, bottom: 200, right: 4 },
      fills: ["1:1"],
      text: { "1:1": "MERGED" }
    });

  it("should lay out a merged region on every page it spans (issue #226)", async () => {
    // The region used to be laid out only on the page holding its master, which
    // left the pages after it with no cell in that column at all — no fill, no
    // outline, nothing. Excel prints the region on each page it reaches.
    const pages = await layoutSheet(tallMerge(), buildOptions(), new FontManager());
    expect(pages.length).toBeGreaterThan(2);

    for (const page of pages) {
      const pieces = cellsInColumn(page, 0);
      expect(pieces).toHaveLength(1);
      // The piece covers the page's whole row band, not one row of it.
      const bandHeight = page.rowHeights.reduce((sum, h) => sum + h, 0);
      expect(pieces[0].rect.height).toBeCloseTo(bandHeight, 5);
      expect(pieces[0].rowSpan).toBe(page.sheetRows.length);
      expect(pieces[0].fillColor).not.toBeNull();
    }
  });

  it("should draw the merged value once, on the piece holding the master", async () => {
    const pages = await layoutSheet(tallMerge(), buildOptions(), new FontManager());

    const withText = pages.filter(page => cellsInColumn(page, 0)[0].text === "MERGED");
    expect(withText).toHaveLength(1);
    expect(withText[0]).toBe(pages[0]);
  });

  it("should close a merged region's outline only where the region really ends", async () => {
    // A page break cuts the region's interior. Closing the cut would print a
    // line Excel does not, and — because the boundary cell was read regardless
    // of where the rectangle had been clamped to — used to print it at the
    // region's full `medium` weight in the middle of the table.
    const pages = await layoutSheet(tallMerge(), buildOptions(), new FontManager());

    const tops = pages.map(page => cellsInColumn(page, 0)[0].borders.top);
    const bottoms = pages.map(page => cellsInColumn(page, 0)[0].borders.bottom);

    expect(tops[0]?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(tops.slice(1)).toEqual(tops.slice(1).map(() => null));
    expect(bottoms.at(-1)?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(bottoms.slice(0, -1)).toEqual(bottoms.slice(0, -1).map(() => null));

    // The left edge is never cut, so every piece carries it — at the region's
    // own weight, while the right edge is only the surrounding grid's.
    for (const page of pages) {
      const piece = cellsInColumn(page, 0)[0];
      expect(piece.borders.left?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
      expect(piece.borders.right?.width).toBeCloseTo(THIN_WIDTH, 5);
    }
  });

  it("should continue a merged region across a column-group boundary", async () => {
    // 40 columns at width 12 need several column groups; the merge spans all of
    // them, so each group holds one piece of it.
    const sheet = borderedSheet({
      rowCount: 3,
      colCount: 40,
      merges: ["A2:AN2"],
      ring: { top: 2, left: 1, bottom: 2, right: 40 },
      text: { "2:1": "WIDE" }
    });
    const pages = await layoutSheet(sheet, buildOptions({ fitToPage: false }), new FontManager());
    expect(pages.length).toBeGreaterThan(2);

    const pieces = pages.map(page => {
      const row = page.sheetRows.indexOf(2);
      const piece = page.cells.find(
        cell => Math.abs(cell.rect.y + cell.rect.height - page.rowYPositions[row]) < 0.01
      );
      expect(piece).toBeDefined();
      // One piece per group, spanning the whole group and starting at its left.
      expect(piece!.colSpan).toBe(page.sheetCols.length);
      expect(piece!.rect.x).toBeCloseTo(page.columnOffsets[0], 5);
      expect(piece!.rect.width).toBeCloseTo(
        page.columnWidths.reduce((sum, w) => sum + w, 0),
        5
      );
      return piece!;
    });

    expect(pieces[0].borders.left?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(pieces.slice(1).map(p => p.borders.left)).toEqual(pieces.slice(1).map(() => null));
    expect(pieces.at(-1)!.borders.right?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(pieces.slice(0, -1).map(p => p.borders.right)).toEqual(
      pieces.slice(0, -1).map(() => null)
    );
    expect(pieces.filter(p => p.text === "WIDE")).toHaveLength(1);

    // The region's own top and bottom are never cut, so every piece keeps them —
    // and at the region's weight rather than the surrounding grid's.
    for (const piece of pieces) {
      expect(piece.borders.top?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
      expect(piece.borders.bottom?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    }
  });

  it("should keep a merged region whole across a hidden row inside it", async () => {
    // Hidden rows leave a gap in the worksheet numbering but not in the visible
    // sequence, so the region is still one piece — testing worksheet adjacency
    // instead would cut it in two at the gap.
    const sheet = borderedSheet({
      rowCount: 8,
      colCount: 3,
      merges: ["A2:A7"],
      ring: { top: 2, left: 1, bottom: 7, right: 1 },
      hiddenRows: [4],
      text: { "2:1": "M" }
    });
    const pages = await layoutSheet(sheet, buildOptions(), new FontManager());

    expect(pages).toHaveLength(1);
    const pieces = cellsInColumn(pages[0], 0);
    // Rows 1, 2–7 (one piece, five visible rows), 8.
    expect(pieces).toHaveLength(3);
    const merged = pieces.find(cell => cell.text === "M")!;
    expect(merged.rowSpan).toBe(5);
    expect(merged.borders.top?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(merged.borders.bottom?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
  });

  // Issue #231: a merge in a collapsed outline group has its master's row
  // hidden while the rest of it is displayed. No piece starts at the master,
  // so keying the value and the outline on the master's row dropped both.
  it("should draw a merged region's value and outline when its master row is hidden", async () => {
    const sheet = borderedSheet({
      rowCount: 6,
      colCount: 2,
      merges: ["A2:A4"],
      ring: { top: 2, left: 1, bottom: 4, right: 1 },
      hiddenRows: [2, 3],
      text: { "2:1": "M" }
    });
    const pages = await layoutSheet(sheet, buildOptions(), new FontManager());

    const merged = cellsInColumn(pages[0], 0).find(cell => cell.text === "M");
    expect(merged).toBeDefined();
    expect(merged!.rowSpan).toBe(1);
    expect(merged!.borders.top?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(merged!.borders.bottom?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
  });

  it("should draw a merged region's value and outline when its master column is hidden", async () => {
    const sheet = borderedSheet({
      rowCount: 2,
      colCount: 4,
      merges: ["A1:C1"],
      ring: { top: 1, left: 1, bottom: 1, right: 3 },
      hiddenCols: [1],
      text: { "1:1": "M" }
    });
    const pages = await layoutSheet(sheet, buildOptions(), new FontManager());

    const merged = pages[0].cells.find(cell => cell.text === "M");
    expect(merged).toBeDefined();
    expect(merged!.colSpan).toBe(2);
    expect(merged!.rect.x).toBeCloseTo(pages[0].columnOffsets[0], 5);
    expect(merged!.borders.left?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(merged!.borders.right?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
  });

  it("should not lay out a merged region whose every row is hidden", async () => {
    const sheet = borderedSheet({
      rowCount: 5,
      colCount: 2,
      merges: ["A2:A3"],
      hiddenRows: [2, 3],
      text: { "2:1": "M" }
    });
    const pages = await layoutSheet(sheet, buildOptions(), new FontManager());
    expect(pages[0].cells.some(cell => cell.text === "M")).toBe(false);
  });

  it("should tile a merged region that also appears in a repeated title band", async () => {
    // The band re-emits the region's first row at the top of every page, right
    // above body rows that are far from it in the sheet. The band's copy and the
    // body's are separate pieces: growing the first to the page's full height
    // would have it overlap the second.
    const sheet = borderedSheet({
      rowCount: 120,
      colCount: 3,
      merges: ["A1:A120"],
      ring: { top: 1, left: 1, bottom: 120, right: 1 },
      pageSetup: { printTitlesRow: "1:1" },
      text: { "1:1": "T" }
    });
    const pages = await layoutSheet(
      sheet,
      buildOptions({ repeatRows: { first: 1, last: 1 } }),
      new FontManager()
    );
    expect(pages.length).toBeGreaterThan(1);

    for (const page of pages.slice(1)) {
      const pieces = cellsInColumn(page, 0).sort((a, b) => b.rect.y - a.rect.y);
      expect(pieces).toHaveLength(2);
      // Contiguous and non-overlapping: the band's piece ends exactly where the
      // body's begins.
      expect(pieces[0].rect.y).toBeCloseTo(pieces[1].rect.y + pieces[1].rect.height, 5);
      expect(pieces[0].rect.height).toBeCloseTo(page.rowHeights[0], 5);
    }
  });

  it("should not grow a merged region past a title band emitted out of order", async () => {
    // A print-title band outside the print area is emitted *ahead* of it, so the
    // visible rows run 25, 26, 10, 11, … — not ascending. A membership test
    // against the region's near edge alone reads row 10 as still inside a region
    // ending at 26, and the piece then covers the whole page on top of the body
    // cells already laid out there.
    const sheet = borderedSheet({
      rowCount: 30,
      colCount: 3,
      merges: ["A25:A26"],
      ring: { top: 25, left: 1, bottom: 26, right: 1 },
      pageSetup: { printArea: "A10:C20" },
      text: { "25:1": "BAND" }
    });
    const pages = await layoutSheet(
      sheet,
      buildOptions({ repeatRows: { first: 25, last: 26 } }),
      new FontManager()
    );

    expect(pages).toHaveLength(1);
    expect(pages[0].sheetRows.slice(0, 3)).toEqual([25, 26, 10]);

    const band = cellsInColumn(pages[0], 0).find(cell => cell.text === "BAND")!;
    expect(band.rowSpan).toBe(2);
    expect(band.rect.height).toBeCloseTo(pages[0].rowHeights[0] + pages[0].rowHeights[1], 5);
    // The region is whole here, so both cut edges are its own.
    expect(band.borders.top?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
    expect(band.borders.bottom?.width).toBeCloseTo(MEDIUM_WIDTH, 5);
  });

  it("should never overlap two laid-out rectangles", async () => {
    // A piece that reached past the tracks it was anchored over would land on
    // top of the next one, which no assertion about a single cell can see. Both
    // arrangements below produced one before: the title band splitting a region
    // in the middle of a page, and the band emitted out of visible order.
    const cases = [
      {
        sheet: borderedSheet({
          rowCount: 120,
          colCount: 6,
          merges: ["A1:A120", "B1:B60", "B61:B120", "C5:D9"],
          ring: { top: 1, left: 1, bottom: 120, right: 6 },
          pageSetup: { printTitlesRow: "1:1" }
        }),
        band: { first: 1, last: 1 }
      },
      {
        sheet: borderedSheet({
          rowCount: 30,
          colCount: 3,
          merges: ["A25:A26", "B10:B14"],
          ring: { top: 25, left: 1, bottom: 26, right: 1 },
          pageSetup: { printArea: "A10:C20" }
        }),
        band: { first: 25, last: 26 }
      }
    ];

    for (const { sheet, band } of cases) {
      const pages = await layoutSheet(sheet, buildOptions({ repeatRows: band }), new FontManager());

      for (const page of pages) {
        for (let i = 0; i < page.cells.length; i++) {
          for (let j = i + 1; j < page.cells.length; j++) {
            const a = page.cells[i].rect;
            const b = page.cells[j].rect;
            const overlaps =
              a.x < b.x + b.width - 0.01 &&
              b.x < a.x + a.width - 0.01 &&
              a.y < b.y + b.height - 0.01 &&
              b.y < a.y + a.height - 0.01;
            expect(overlaps).toBe(false);
          }
        }
      }
    }
  });
});
