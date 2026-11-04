/**
 * Shared layout constants for the layout / pagination / SVG render modules.
 *
 * These defaults are used when a `SectionProperties` does not specify a
 * page size or margin. Keeping them in a single file avoids drift between
 * `layout.ts`, `layout-full.ts`, and `render-page.ts`.
 *
 * The values match Microsoft Word's defaults for US Letter paper:
 *   - Page size: 8.5 in × 11 in (12240 × 15840 twips)
 *   - Margins:   1 in on all sides (1440 twips)
 */

import { finalViewRuns } from "@word/query/final-view";
import type { StyleResolver } from "@word/query/style-resolver";
import type {
  Paragraph,
  Run,
  RunProperties,
  Table,
  TableCellProperties,
  TableProperties,
  TableWidth
} from "@word/types";

/** Default page width in twips (US Letter, 8.5 in). */
export const DEFAULT_PAGE_WIDTH_TWIPS = 12240;

/** Default page height in twips (US Letter, 11 in). */
export const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;

/** Default page margin in twips (1 in on each side). */
export const DEFAULT_PAGE_MARGIN_TWIPS = 1440;

/**
 * Default body font size in points, for a run whose size no style in its chain
 * declares.
 *
 * Keeping this in one place matters because the paginator, the positioner and
 * the SVG renderer all have to agree on the height of a size-less run: when
 * they disagreed, the page each item was assigned to no longer matched the
 * space it actually occupied.
 *
 * **This is a fallback, not Word's answer.** A file with no `w:docDefaults`
 * gets its defaults in Word from the user's Normal template, which differs
 * between Word versions (Word for Mac 16.114 uses Aptos 12pt, 8pt after,
 * 1.16 lines), so no fallback can match Word for such a file. It applies
 * only to such files: a document the library builds states current Word's
 * new-document defaults and is laid out and written with them
 * (`withStatedDefaults`).
 */
export const DEFAULT_FONT_SIZE_PT = 11;

/** Default body font size in half-points (the unit `w:sz` uses). */
export const DEFAULT_FONT_SIZE_HALF_PT = DEFAULT_FONT_SIZE_PT * 2;

/**
 * Natural (single-spaced) line height as a multiple of the font size.
 *
 * Word derives a line's natural height from the font metrics of the largest run
 * on that line, not from a document-wide constant. 1.2 is the usual
 * approximation for the Latin faces we ship metrics for.
 */
export const LINE_HEIGHT_FACTOR = 1.2;

/**
 * Rendered size of a superscript/subscript, as a fraction of its source size.
 *
 * Kept beside {@link SCRIPT_BASELINE_SHIFT_FACTOR}: the two describe one effect,
 * and measurement, line extents and drawing all have to apply both.
 */
export const SCRIPT_FONT_SIZE_RATIO = 0.65;

/** Baseline rise/drop for a superscript/subscript, relative to its rendered size. */
export const SCRIPT_BASELINE_SHIFT_FACTOR = 0.33;

// =============================================================================
// Table cell margins
// =============================================================================

/**
 * The four inner margins of a table cell, in twips.
 *
 * Word calls these cell margins (`w:tblCellMar` on the table, `w:tcMar` per
 * cell); CSS would call them padding. They set both how far the content is inset
 * from the cell's borders and how tall the row has to be, so the paginator and
 * the positioner have to derive them identically — hence one resolver here
 * rather than a hardcoded guess in each.
 */
export interface CellMarginsTwips {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
}

/**
 * Word's cell margins when neither the cell nor the table — its own
 * properties or its table style's (the document's default table style, for a
 * table naming none) — sets any: none at all. Word's own default table style, `Normal Table`, is what
 * carries the familiar 0.08 inch left and right; a document without one
 * (`w:default="1"`) gets flush cells (Word for Mac 16.114).
 */
export const DEFAULT_CELL_MARGINS_TWIPS: CellMarginsTwips = {
  top: 0,
  right: 0,
  bottom: 0,
  left: 0
};

/**
 * Resolve one `w:tblCellMar` / `w:tcMar` entry to twips.
 *
 * `type: "nil"` is an explicit zero. Cell margins are always authored in twips
 * (`dxa`), so any other unit is ignored rather than misread as a twip count.
 */
function marginTwips(width: TableWidth | undefined, fallback: number): number | undefined {
  if (!width) {
    return undefined;
  }
  if (width.type === "nil") {
    return 0;
  }
  return width.type === "dxa" ? width.value : fallback;
}

/**
 * Resolve a cell's effective inner margins in twips: the cell's own `w:tcMar`
 * wins, otherwise the table's `w:tblCellMar`, otherwise
 * {@link DEFAULT_CELL_MARGINS_TWIPS}. `tableProps` are the table's effective
 * properties (`StyleResolver.tableProperties`), its table style's included.
 */
export function resolveCellMarginsTwips(
  tableProps: TableProperties | undefined,
  cellProps: TableCellProperties | undefined
): CellMarginsTwips {
  const table = tableProps?.cellMargins;
  const cell = cellProps?.margins;
  const pick = (
    own: TableWidth | undefined,
    inherited: TableWidth | undefined,
    fallback: number
  ): number => marginTwips(own, fallback) ?? marginTwips(inherited, fallback) ?? fallback;

  // `start`/`end` are the logical (bidi-aware) spellings of left/right; in the
  // left-to-right layout these engines implement they mean the same thing.
  return {
    top: pick(cell?.top, table?.top, DEFAULT_CELL_MARGINS_TWIPS.top),
    left: pick(
      cell?.left ?? cell?.start,
      table?.left ?? table?.start,
      DEFAULT_CELL_MARGINS_TWIPS.left
    ),
    bottom: pick(cell?.bottom, table?.bottom, DEFAULT_CELL_MARGINS_TWIPS.bottom),
    right: pick(
      cell?.right ?? cell?.end,
      table?.right ?? table?.end,
      DEFAULT_CELL_MARGINS_TWIPS.right
    )
  };
}

// =============================================================================
// Heading heuristics
// =============================================================================

/**
 * How much larger than body text a heading of this level is drawn.
 *
 * Applied **only** when the style supplies no concrete `w:sz` — a document with a
 * styles table gets its declared sizes and this heuristic stays out of the way.
 *
 * Shared because both layout passes must agree on it. The positioned pass applied
 * it and the pagination pass did not, so a run of `Heading 1` paragraphs was
 * estimated at body height: 40 of them reported one page and paginated to two, and
 * `pageCount` is what `NUMPAGES`, `PAGE`, TOC entries and `PAGEREF` are resolved
 * from — so the page numbers written into the document were a quarter of the truth.
 */
export function getHeadingFontScale(level: number): number {
  switch (level) {
    case 1:
      return 2.0;
    case 2:
      return 1.5;
    case 3:
      return 1.17;
    case 4:
      return 1.0;
    case 5:
      return 0.83;
    case 6:
      return 0.67;
    default:
      return 1.0;
  }
}

/**
 * The heading scale in force for a paragraph, given its resolved style run props.
 *
 * The single expression both passes use, so neither can forget the `w:sz` opt-out.
 */
export function resolveHeadingScale(
  styles: StyleResolver,
  para: Paragraph,
  styleRunSize: number | null | undefined
): number {
  if (styleRunSize != null) {
    return 1;
  }
  const heading = styles.heading(para);
  return heading ? getHeadingFontScale(heading.level) : 1;
}

/**
 * Slack allowed when asking whether a block still fits on the page.
 *
 * Heights accumulate by repeated addition, so a page that should be filled exactly
 * comes up a few ulps short: fourteen 43.2pt paragraphs sum to 604.80000000000007,
 * leaving 43.199999999999932 for a fifteenth that measures 43.2 — and the block was
 * pushed to the next page over 7e-14pt. That is invisible on the page and visible in
 * `pageCount`, which `NUMPAGES`, `PAGE`, TOC entries and `PAGEREF` are resolved from.
 *
 * A hundredth of a point is far below anything a document can express (a twip, the
 * finest OOXML unit, is 0.05pt) and far above the accumulated error.
 */
export const FIT_EPSILON_PT = 0.01;

// =============================================================================
// Run property inheritance
// =============================================================================

/**
 * Merge a run's own properties over the ones it inherits.
 *
 * `w:rFonts` is merged **per slot**, not replaced whole. A spread — `{...inherited,
 * ...own}` — throws away the inherited `w:ascii` the moment a run names only
 * `w:eastAsia`, which is how a Chinese run overriding just its East Asian face is
 * written. Word resolves each attribute of `w:rFonts` independently, so the Latin
 * face from the style chain has to survive.
 *
 * The pagination pass papered over this by threading the inherited Latin name
 * through a separate parameter; the positioned pass had nothing and fell back to
 * Calibri, measuring the Latin half of every such run against the wrong face —
 * narrower than the truth, so the two passes disagreed about page count in the
 * direction that writes a wrong `NUMPAGES` into the document.
 */
export function mergeRunProperties(
  inherited: RunProperties | undefined,
  own: RunProperties | undefined
): RunProperties | undefined {
  if (!inherited) {
    return own;
  }
  if (!own) {
    return inherited;
  }
  const merged: RunProperties = { ...inherited, ...own };
  if (inherited.font !== undefined && own.font !== undefined) {
    const a = typeof inherited.font === "string" ? { ascii: inherited.font } : inherited.font;
    const b = typeof own.font === "string" ? { ascii: own.font } : own.font;
    return { ...merged, font: { ...a, ...b } };
  }
  return merged;
}

/**
 * A run's effective formatting for layout: document defaults → paragraph style
 * (`inherited`) → character style chain → direct properties. Both layout passes
 * use this, so they agree on every run's size, face and visibility
 * (`isHiddenRun` of the result).
 *
 * A run without a character style takes the cheap merge; with one, the full
 * `resolveRunStyle` precedence (including toggle properties) applies.
 */
export function layoutRunProperties(
  styles: StyleResolver,
  run: Run,
  inherited: RunProperties | undefined
): RunProperties | undefined {
  return run.properties?.style
    ? styles.run(run, inherited).runProperties
    : mergeRunProperties(inherited, run.properties);
}

// =============================================================================
// Table geometry
// =============================================================================

/**
 * Column widths in twips, scaled to the measure the table is laid into.
 *
 * `w:tblGrid` is advisory: Word fits the grid to the available width, shrinking a
 * table whose columns overflow and expanding one that under-fills. The pagination
 * pass used the declared twips as-is, so a table arriving from the Excel or HTML
 * bridge with a grid twice the page width was estimated with columns twice as wide
 * as they are drawn — half the wrapped lines, and a page count short of the truth.
 *
 * Returns equal columns when no usable grid is declared.
 */
/**
 * How many grid columns a table spans: its `w:tblGrid`, or the widest row
 * counted in grid columns (`w:gridBefore`, each cell's `w:gridSpan`,
 * `w:gridAfter`) when that is wider. A row's cell count is not it — a row of a
 * three-column span and one cell covers four columns.
 */
export function gridColumnCount(table: Table): number {
  let count = table.columnWidths?.length ?? 0;
  for (const row of table.rows) {
    let cols = (row.properties?.gridBefore ?? 0) + (row.properties?.gridAfter ?? 0);
    for (const cell of row.cells) {
      cols += Math.max(1, cell.properties?.gridSpan ?? 1);
    }
    count = Math.max(count, cols);
  }
  return count;
}

export function resolveColumnWidthsTwips(
  table: Table,
  numCols: number,
  availableWidthTwips: number
): number[] {
  if (numCols <= 0) {
    return [];
  }
  const declared = table.columnWidths;
  if (declared && declared.length >= numCols) {
    const widths = declared.slice(0, numCols);
    const total = widths.reduce((a, b) => a + b, 0);
    if (total > 0) {
      const scale = availableWidthTwips / total;
      return widths.map(w => w * scale);
    }
  }
  return new Array<number>(numCols).fill(availableWidthTwips / numCols);
}

/**
 * The width, in twips, a table nested in a cell sets out at, or `undefined`
 * when the cell decides it (an auto-fit table in an auto-fit table, sized by
 * its content).
 *
 * A nested table is not stretched or shrunk to the cell holding it. Its
 * preferred width (`w:tblW`, ECMA-376 17.4.64) wins where one is declared — a
 * percentage of the cell's content width; otherwise a fixed-layout table
 * (`w:tblLayout`, 17.4.53) keeps its grid, and so does an auto-fit one inside
 * a fixed-layout table. The grid is scaled to the declared width. Word for Mac
 * 16.114 lays out every case this way, overflowing the cell where the width is
 * wider (final7 R0–R3c, Re, Rg: 2×2232 → 111.6pt, 2×2340 → 117pt, tblW 3000
 * over a 2×2000 or 2×3000 grid → 75pt, 2×1000 auto-fit → 50pt, 2×3000 → 150pt).
 */
export function nestedTableWidthTwips(
  properties: TableProperties | undefined,
  columnWidths: readonly number[] | undefined,
  cellContentWidthTwips: number,
  outerFixed: boolean
): number | undefined {
  const width = properties?.width;
  if (width?.type === "dxa" && width.value > 0) {
    return width.value;
  }
  if (width?.type === "pct" && width.value > 0) {
    return (cellContentWidthTwips * width.value) / 5000;
  }
  const grid = columnWidths?.reduce((a, b) => a + b, 0) ?? 0;
  if (grid > 0 && (properties?.layout === "fixed" || outerFixed)) {
    return grid;
  }
  return undefined;
}

/**
 * The shortest a table row can be, in points.
 *
 * One line of body text plus the row's own cell margins. The two passes had
 * different answers — the pagination pass used a line height plus the margins while
 * the positioned pass hardcoded `11 × 1.5 = 16.5pt` and ignored the margins — so a
 * compact table of single-line cells was estimated 3.3pt per row short. Forty rows
 * of that is a page.
 */
export function minimumRowHeightPt(
  tableProperties: TableProperties | undefined,
  defaultFontSizePt: number
): number {
  const margins = resolveCellMarginsTwips(tableProperties, undefined);
  return defaultFontSizePt * LINE_HEIGHT_FACTOR + (margins.top + margins.bottom) / 20;
}

/**
 * Whether block `index` of a cell is the empty paragraph that closes it after a
 * nested table. A cell must end with a paragraph, so one follows every nested
 * table at the end of a cell; Word for Mac 16.114 gives it no height whatever
 * the nested table's layout and whether or not it fits the cell (final5
 * M2-noimg: a 97.5pt row, 121.4pt with the paragraph laid out as a line; final7
 * R0–R3c, Re, Rg, Qae, Qag: fixed and auto-fit, fitting and overflowing).
 */
export function isCellEndAfterNestedTable(
  content: readonly (Paragraph | Table)[],
  index: number
): boolean {
  const block = content[index];
  return (
    index === content.length - 1 &&
    index > 0 &&
    content[index - 1]!.type === "table" &&
    block?.type === "paragraph" &&
    [...finalViewRuns(block.children)].every(run => run.content.length === 0)
  );
}

/**
 * Rotated cell text (`w:textDirection` `tbRl`, `btLr`, …), as Word for Mac 16 sets
 * it. An auto-fit rotated column is as wide as its paragraphs are high, each laid
 * out on one line ({@link ROTATED_LINE_WIDTH_PT} wide), plus
 * {@link ROTATED_PARAGRAPH_EXTRA_PT} per paragraph and
 * {@link ROTATED_COLUMN_EXTRA_PT} once (pages2 P21: one to five 11pt paragraphs
 * are 28.9, 74.8 and 120.7pt of text room — 22.95pt each, 22.49 laid out — a 20pt
 * line 40.9pt, a single-spaced line without spacing 20.0pt). Fitted to Word's
 * measurements; the reason for the residue is not known.
 */
export const ROTATED_LINE_WIDTH_PT = 1e6;
export const ROTATED_PARAGRAPH_EXTRA_PT = 0.46;
export const ROTATED_COLUMN_EXTRA_PT = 5.95;

/** Whether a cell's text runs top to bottom or bottom to top. */
export function isRotatedTextDirection(direction: string | undefined): boolean {
  return (
    direction === "tbRl" || direction === "btLr" || direction === "tbRlV" || direction === "tbLrV"
  );
}
