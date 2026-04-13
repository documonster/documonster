/**
 * Layout engine for PDF generation.
 *
 * Takes a PdfSheetData and produces LayoutPage objects that describe exactly
 * where each cell, border, and piece of text should be drawn on each PDF page.
 *
 * This module is fully independent of the Excel module — it works with
 * the PDF module's own data model (PdfSheetData, PdfCellData, etc.).
 *
 * Key responsibilities:
 * - Convert column widths (character units) to PDF points
 * - Convert row heights (points already, but may need scaling)
 * - Handle merged cells spanning multiple rows/columns
 * - Paginate content across multiple pages
 * - Handle fitToPage scaling
 * - Handle repeated header rows
 * - Skip hidden rows and columns
 */

import type {
  PdfSheetData,
  PdfCellData,
  PdfCellStyle,
  PdfRichTextRunData,
  PdfSheetImage,
  PdfAlignmentData,
  PdfCellTypeValue,
  ResolvedPdfOptions,
  LayoutPage,
  LayoutCell,
  LayoutBorder,
  LayoutRichTextRun
} from "../types";
import { PdfCellType } from "../types";
import type { FontManager } from "../font/font-manager";
import { resolvePdfFontName } from "../font/font-manager";
import {
  extractFontProperties,
  excelFillToPdfColor,
  excelBordersToPdf,
  excelHAlignToPdf,
  excelVAlignToPdf,
  borderStyleToLineWidth
} from "./style-converter";
import { wrapTextLines } from "./page-renderer";
import {
  CELL_PADDING_H,
  CELL_PADDING_V,
  LINE_HEIGHT_FACTOR,
  INDENT_WIDTH,
  MAX_DIGIT_WIDTH_PX,
  EXCEL_COLUMN_PADDING_PX,
  PX_TO_PT
} from "./constants";
import { yieldToEventLoop } from "@utils/utils.base";

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_COLUMN_WIDTH = 8.43;
const DEFAULT_ROW_HEIGHT = 15;
const MIN_COLUMN_WIDTH = 3;

// =============================================================================
// Type-based Default Alignment
// =============================================================================

/**
 * Resolve horizontal alignment, using Excel's type-based defaults when
 * no explicit alignment is set (or when alignment is "general"):
 * - Numbers/Dates → right
 * - Booleans/Errors → center
 * - Text/RichText/Hyperlink → left
 * - Formulas → based on result type
 */
function resolveHorizontalAlign(
  alignment: Partial<PdfAlignmentData> | undefined,
  cellType: PdfCellTypeValue | undefined,
  formulaResult?: unknown
): "left" | "center" | "right" {
  // If explicitly set (and not "general"), use the explicit alignment
  if (alignment?.horizontal && alignment.horizontal !== "general") {
    return excelHAlignToPdf(alignment);
  }

  // Use type-based default
  if (cellType !== undefined) {
    switch (cellType) {
      case PdfCellType.Number:
      case PdfCellType.Date:
        return "right";
      case PdfCellType.Boolean:
      case PdfCellType.Error:
        return "center";
      case PdfCellType.Formula:
        if (typeof formulaResult === "number" || formulaResult instanceof Date) {
          return "right";
        }
        if (typeof formulaResult === "boolean") {
          return "center";
        }
        return "left";
      default:
        return "left";
    }
  }

  return "left";
}

// =============================================================================
// Layout Engine
// =============================================================================

/**
 * Compute the layout for a sheet across one or more PDF pages.
 * Yields to the event loop between each output page.
 */
export async function layoutSheet(
  sheet: PdfSheetData,
  options: ResolvedPdfOptions,
  fontManager: FontManager
): Promise<LayoutPage[]> {
  const ctx = prepareLayout(sheet, options, fontManager);
  if (!ctx) {
    return [createEmptyPage(sheet, options)];
  }

  const layoutPages: LayoutPage[] = [];
  const totalOutputPages = ctx.rowPages.length * ctx.colGroups.length;

  for (const rowPage of ctx.rowPages) {
    for (const colGroup of ctx.colGroups) {
      layoutPages.push(
        buildPageLayout(ctx, rowPage, colGroup, layoutPages.length, sheet, options, fontManager)
      );
      if (layoutPages.length < totalOutputPages) {
        await yieldToEventLoop();
      }
    }
  }

  if (layoutPages.length > 0 && sheet.images) {
    assignImagesToPages(sheet.images, layoutPages, ctx.scaleFactor);
  }

  return layoutPages;
}

// =============================================================================
// Internal — Shared Layout Pipeline
// =============================================================================

/** Pre-computed layout context for the layout pipeline. */
interface LayoutContext {
  pageWidth: number;
  pageHeight: number;
  contentWidth: number;
  headerHeight: number;
  scaleFactor: number;
  scaledColumnWidths: number[];
  rowHeights: number[];
  visibleRows: number[];
  visibleCols: number[];
  mergeMap: Map<string, MergeInfo>;
  rowPages: number[][];
  colGroups: number[][];
  margins: { top: number; right: number; bottom: number; left: number };
}

/**
 * Steps 1–5: compute columns, scale, rows, merges, pagination.
 * Returns null if the sheet has no visible columns (→ caller should emit an empty page).
 */
function prepareLayout(
  sheet: PdfSheetData,
  options: ResolvedPdfOptions,
  fontManager: FontManager
): LayoutContext | null {
  const { margins } = options;

  let pageWidth = options.pageSize.width;
  let pageHeight = options.pageSize.height;
  if (options.orientation === "landscape") {
    [pageWidth, pageHeight] = [pageHeight, pageWidth];
  }

  const contentWidth = pageWidth - margins.left - margins.right;
  const contentHeight = pageHeight - margins.top - margins.bottom;
  const headerHeight = options.showSheetNames ? 20 : 0;
  const footerHeight = options.showPageNumbers ? 20 : 0;
  const availableHeight = contentHeight - headerHeight - footerHeight;

  const printRange = getPrintRange(sheet);

  // --- Step 1: Visible columns and widths ---
  const { columnWidths, visibleCols } = computeColumnWidths(sheet, printRange);
  if (visibleCols.length === 0) {
    return null;
  }

  // --- Step 2: Scale ---
  const totalTableWidth = columnWidths.reduce((sum, w) => sum + w, 0);
  let scaleFactor = options.scale;
  if (options.fitToPage && totalTableWidth > 0) {
    const fitScale = contentWidth / totalTableWidth;
    if (fitScale < 1) {
      scaleFactor *= fitScale;
    }
  }
  const scaledColumnWidths = columnWidths.map(w => w * scaleFactor);

  // --- Step 3: Visible rows and heights ---
  const { rowHeights, visibleRows } = computeRowHeights(
    sheet,
    scaleFactor,
    printRange,
    fontManager,
    options
  );

  // --- Step 4: Merge map ---
  const mergeMap = buildMergeMap(sheet);

  // --- Step 5: Paginate ---
  const repeatRowCount = typeof options.repeatRows === "number" ? options.repeatRows : 0;
  const rowBreakSet = buildRowBreakSet(sheet, visibleRows);
  const rowPages = paginateRows(rowHeights, availableHeight, repeatRowCount, rowBreakSet);
  const colGroups = paginateColumns(scaledColumnWidths, contentWidth, sheet, visibleCols);

  return {
    pageWidth,
    pageHeight,
    contentWidth,
    headerHeight,
    scaleFactor,
    scaledColumnWidths,
    rowHeights,
    visibleRows,
    visibleCols,
    mergeMap,
    rowPages,
    colGroups,
    margins
  };
}

/**
 * Build the LayoutPage for a single rowPage × colGroup combination.
 */
function buildPageLayout(
  ctx: LayoutContext,
  rowPage: number[],
  colGroup: number[],
  currentPageCount: number,
  sheet: PdfSheetData,
  options: ResolvedPdfOptions,
  fontManager: FontManager
): LayoutPage {
  const {
    scaledColumnWidths,
    rowHeights,
    visibleRows,
    visibleCols,
    mergeMap,
    pageWidth,
    pageHeight,
    contentWidth,
    headerHeight,
    scaleFactor,
    margins
  } = ctx;

  const cells: LayoutCell[] = [];

  // Compute column offsets for this column group
  const groupColWidths = colGroup.map(ci => scaledColumnWidths[ci]);
  const groupTotalWidth = groupColWidths.reduce((s, w) => s + w, 0);
  const groupColOffsets: number[] = [];
  let gx = margins.left;
  if (groupTotalWidth < contentWidth) {
    gx = margins.left + (contentWidth - groupTotalWidth) / 2;
  }
  for (const w of groupColWidths) {
    groupColOffsets.push(gx);
    gx += w;
  }

  // Row Y positions
  const rowYPositions: number[] = [];
  const pageRowHeights: number[] = [];
  let currentY = pageHeight - margins.top - headerHeight;
  for (const rowIdx of rowPage) {
    const rowH = rowHeights[rowIdx] ?? DEFAULT_ROW_HEIGHT * scaleFactor;
    rowYPositions.push(currentY);
    pageRowHeights.push(rowH);
    currentY -= rowH;
  }

  // Build cells for this row page × column group
  const cellGrid = new Map<string, LayoutCell>();

  for (let ri = 0; ri < rowPage.length; ri++) {
    const visibleRowIdx = rowPage[ri];
    const wsRowNumber = visibleRows[visibleRowIdx];

    for (let gci = 0; gci < colGroup.length; gci++) {
      const ci = colGroup[gci];
      const wsColNumber = visibleCols[ci];

      const mergeKey = `${wsRowNumber}:${wsColNumber}`;
      const mergeInfo = mergeMap.get(mergeKey);
      if (mergeInfo && !mergeInfo.isMaster) {
        continue;
      }

      const row = sheet.rows.get(wsRowNumber);
      const cell = row?.cells.get(wsColNumber);

      let colSpan = 1;
      let rowSpan = 1;
      if (mergeInfo && mergeInfo.isMaster) {
        const mergeEndCol = wsColNumber + mergeInfo.colSpan - 1;
        colSpan = 0;
        for (let s = gci; s < colGroup.length; s++) {
          if (visibleCols[colGroup[s]] <= mergeEndCol) {
            colSpan++;
          } else {
            break;
          }
        }
        const mergeEndRow = wsRowNumber + mergeInfo.rowSpan - 1;
        rowSpan = 0;
        for (let s = visibleRowIdx; s < visibleRows.length; s++) {
          if (visibleRows[s] <= mergeEndRow) {
            rowSpan++;
          } else {
            break;
          }
        }
        colSpan = Math.max(colSpan, 1);
        rowSpan = Math.max(rowSpan, 1);
      }

      const cellX = groupColOffsets[gci];
      const cellY = rowYPositions[ri];
      let cellWidth = 0;
      for (let s = 0; s < colSpan && gci + s < groupColWidths.length; s++) {
        cellWidth += groupColWidths[gci + s];
      }
      let cellHeight = 0;
      for (let s = 0; s < rowSpan && ri + s < pageRowHeights.length; s++) {
        cellHeight += pageRowHeights[ri + s];
      }
      const rectY = cellY - cellHeight;

      cells.push(
        buildLayoutCell(
          cell,
          cellX,
          rectY,
          cellWidth,
          cellHeight,
          colSpan,
          rowSpan,
          options,
          fontManager,
          scaleFactor
        )
      );

      const layoutCell = cells[cells.length - 1];

      // Propagate merged cell borders from boundary cells
      if (mergeInfo?.isMaster) {
        propagateMergeBorders(layoutCell, mergeInfo, wsRowNumber, wsColNumber, sheet);
      }

      cellGrid.set(`${ri}:${gci}`, layoutCell);
    }
  }

  // Resolve shared borders: on each shared edge between adjacent cells, keep
  // only the winning border for drawing but preserve insets for both cells.
  resolveSharedBorders(cellGrid, rowPage.length, colGroup.length);

  // Compute text overflow widths for non-wrapped cells
  computeTextOverflows(
    cellGrid,
    rowPage,
    colGroup,
    visibleRows,
    visibleCols,
    groupColWidths,
    mergeMap,
    fontManager
  );

  return {
    pageNumber: currentPageCount + 1,
    options,
    cells,
    width: pageWidth,
    height: pageHeight,
    sheetName: sheet.name,
    sheetCols: colGroup.map(ci => visibleCols[ci]),
    columnOffsets: groupColOffsets,
    columnWidths: groupColWidths,
    sheetRows: rowPage.map(ri => visibleRows[ri]),
    rowYPositions,
    rowHeights: pageRowHeights,
    images: [],
    scaleFactor
  };
}

function createEmptyPage(sheet: PdfSheetData, options: ResolvedPdfOptions): LayoutPage {
  let pageWidth = options.pageSize.width;
  let pageHeight = options.pageSize.height;
  if (options.orientation === "landscape") {
    [pageWidth, pageHeight] = [pageHeight, pageWidth];
  }

  return {
    pageNumber: 1,
    options,
    cells: [],
    width: pageWidth,
    height: pageHeight,
    sheetName: sheet.name,
    sheetCols: [],
    columnOffsets: [],
    columnWidths: [],
    sheetRows: [],
    rowYPositions: [],
    rowHeights: [],
    images: [],
    scaleFactor: 1
  };
}

// =============================================================================
// Range Parsing (standalone — no @excel dependency)
// =============================================================================

interface CellRef {
  /** 0-indexed column */
  c: number;
  /** 0-indexed row */
  r: number;
}

interface RangeRef {
  s: CellRef;
  e: CellRef;
}

/**
 * Parse a cell reference like "A1" into 0-indexed { c, r }.
 */
function parseCellRef(ref: string): CellRef {
  const upper = ref.replace(/\$/g, "").toUpperCase();
  let col = 0;
  let i = 0;
  while (i < upper.length && upper.charCodeAt(i) >= 65 && upper.charCodeAt(i) <= 90) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
    i++;
  }
  const row = parseInt(upper.substring(i), 10);
  return { c: col - 1, r: row - 1 };
}

/**
 * Parse a range string like "A1:B2" into 0-indexed start/end.
 */
function parseRangeRef(range: string): RangeRef {
  const idx = range.indexOf(":");
  if (idx === -1) {
    const cell = parseCellRef(range);
    return { s: cell, e: { ...cell } };
  }
  return {
    s: parseCellRef(range.slice(0, idx)),
    e: parseCellRef(range.slice(idx + 1))
  };
}

// =============================================================================
// Print Range
// =============================================================================

interface PrintRange {
  startRow: number; // 1-based
  endRow: number;
  startCol: number;
  endCol: number;
}

/**
 * Get the print area range from the sheet's pageSetup.
 * Returns null if no print area is set.
 */
function getPrintRange(sheet: PdfSheetData): PrintRange | null {
  const printArea = sheet.pageSetup?.printArea;
  if (!printArea || typeof printArea !== "string") {
    return null;
  }
  // printArea may be multi-range separated by "&&" (e.g. "A1:B2&&D1:E2").
  // Use the first range for PDF export.
  const firstRange = printArea.split("&&")[0].trim();
  if (!firstRange) {
    return null;
  }
  try {
    const range = parseRangeRef(firstRange);
    return {
      startRow: range.s.r + 1,
      endRow: range.e.r + 1,
      startCol: range.s.c + 1,
      endCol: range.e.c + 1
    };
  } catch {
    return null;
  }
}

// =============================================================================
// Column Width Computation
// =============================================================================

function computeColumnWidths(
  sheet: PdfSheetData,
  printRange: PrintRange | null
): {
  columnWidths: number[];
  visibleCols: number[];
} {
  const bounds = sheet.bounds;
  const hasData = bounds.top > 0 && bounds.left > 0;

  if (!hasData) {
    return { columnWidths: [], visibleCols: [] };
  }

  const startCol = printRange?.startCol ?? bounds.left;
  const endCol = printRange?.endCol ?? bounds.right;
  const columnWidths: number[] = [];
  const visibleCols: number[] = [];

  for (let c = startCol; c <= endCol; c++) {
    const col = sheet.columns.get(c);
    if (col?.hidden) {
      continue;
    }
    const excelWidth = col?.width ?? DEFAULT_COLUMN_WIDTH;
    const pixelWidth = excelWidth * MAX_DIGIT_WIDTH_PX + EXCEL_COLUMN_PADDING_PX;
    const pointWidth = Math.max(pixelWidth * PX_TO_PT, MIN_COLUMN_WIDTH);
    columnWidths.push(pointWidth);
    visibleCols.push(c);
  }

  return { columnWidths, visibleCols };
}

// =============================================================================
// Row Height Computation
// =============================================================================

function computeRowHeights(
  sheet: PdfSheetData,
  scaleFactor: number,
  printRange: PrintRange | null,
  fontManager: FontManager,
  options: ResolvedPdfOptions
): { rowHeights: number[]; visibleRows: number[] } {
  const bounds = sheet.bounds;
  if (bounds.top <= 0) {
    return { rowHeights: [], visibleRows: [] };
  }

  const startRow = printRange?.startRow ?? bounds.top;
  const endRow = printRange?.endRow ?? bounds.bottom;
  const rowHeights: number[] = [];
  const visibleRows: number[] = [];

  for (let r = startRow; r <= endRow; r++) {
    const row = sheet.rows.get(r);
    if (row?.hidden) {
      continue;
    }

    let height: number;
    if (row?.height && row.customHeight) {
      // Custom height explicitly set by user — use as-is
      height = row.height;
    } else if (row?.height) {
      // Excel auto-calculated height — trust it as-is
      height = row.height;
    } else {
      // No height info: auto-size based on cell content
      height = DEFAULT_ROW_HEIGHT;
      if (row) {
        for (const cell of row.cells.values()) {
          const fontSize = getCellFontSize(cell);
          const wrapLineCount = countWrapLines(
            cell,
            fontSize,
            scaleFactor,
            sheet,
            fontManager,
            options
          );
          const lineHeight = fontSize * LINE_HEIGHT_FACTOR;
          // Account for border width: half of each border extends inward
          const borderTop = cell.style?.border?.top?.style
            ? borderStyleToLineWidth(cell.style.border.top.style) / 2
            : 0;
          const borderBottom = cell.style?.border?.bottom?.style
            ? borderStyleToLineWidth(cell.style.border.bottom.style) / 2
            : 0;
          const neededHeight =
            fontSize +
            (wrapLineCount - 1) * lineHeight +
            (CELL_PADDING_V + borderTop + borderBottom) * 2;
          if (neededHeight > height) {
            height = neededHeight;
          }
        }
      }
    }

    rowHeights.push(height * scaleFactor);
    visibleRows.push(r);
  }

  return { rowHeights, visibleRows };
}

/**
 * Get the largest font size for a cell, checking rich text runs.
 */
function getCellFontSize(cell: PdfCellData): number {
  let fontSize = cell.style?.font?.size ?? 11;

  if (cell.type === PdfCellType.RichText) {
    const value = cell.value;
    if (value && typeof value === "object" && "richText" in value) {
      const runs = (value as { richText: PdfRichTextRunData[] }).richText;
      for (const run of runs) {
        const runSize = run.font?.size ?? fontSize;
        if (runSize > fontSize) {
          fontSize = runSize;
        }
      }
    }
  }

  return fontSize;
}

/**
 * Count the wrap-line count for a cell, using actual font measurements
 * so row heights match the page renderer exactly.
 */
function countWrapLines(
  cell: PdfCellData,
  fontSize: number,
  scaleFactor: number,
  sheet: PdfSheetData,
  fontManager: FontManager,
  options: ResolvedPdfOptions
): number {
  const text = typeof cell.text === "string" ? cell.text : String(cell.text ?? "");
  const lineCount = Math.max(1, (text.match(/\n/g) ?? []).length + 1);

  if (!cell.style?.alignment?.wrapText || text.length === 0) {
    return lineCount;
  }

  const col = sheet.columns.get(cell.col);
  const colWidth = col?.width ?? DEFAULT_COLUMN_WIDTH;
  const scaledColPts =
    (colWidth * MAX_DIGIT_WIDTH_PX + EXCEL_COLUMN_PADDING_PX) * PX_TO_PT * scaleFactor;
  const indent = cell.style.alignment.indent ?? 0;
  const borderLeft = cell.style?.border?.left?.style
    ? borderStyleToLineWidth(cell.style.border.left.style) / 2
    : 0;
  const borderRight = cell.style?.border?.right?.style
    ? borderStyleToLineWidth(cell.style.border.right.style) / 2
    : 0;
  const padding =
    CELL_PADDING_H + borderLeft + (CELL_PADDING_H + borderRight) + indent * INDENT_WIDTH;
  const effectiveWidth = Math.max(scaledColPts - padding, 1);

  const scaledFontSize = fontSize * scaleFactor;
  const fontProps = extractFontProperties(
    cell.style.font,
    options.defaultFontFamily,
    options.defaultFontSize
  );
  const pdfFontName = resolvePdfFontName(fontProps.fontFamily, fontProps.bold, fontProps.italic);
  const resourceName = fontManager.hasEmbeddedFont()
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(pdfFontName);
  const measure = (s: string) => fontManager.measureText(s, resourceName, scaledFontSize);
  const wrappedLines = wrapTextLines(text, measure, effectiveWidth);

  return Math.max(lineCount, wrappedLines.length);
}

// =============================================================================
// Row Breaks
// =============================================================================

/**
 * Build a set of visible-row indices where manual page breaks occur.
 */
function buildRowBreakSet(sheet: PdfSheetData, visibleRows: number[]): Set<number> {
  const breaks = new Set<number>();
  const rowBreaks = sheet.rowBreaks ?? [];
  if (rowBreaks.length === 0) {
    return breaks;
  }
  // Map row numbers to visible-row indices
  const rowToIndex = new Map<number, number>();
  for (let i = 0; i < visibleRows.length; i++) {
    rowToIndex.set(visibleRows[i], i);
  }
  for (const brk of rowBreaks) {
    const idx = rowToIndex.get(brk);
    if (idx !== undefined) {
      // Break AFTER this row, so the next row starts a new page
      breaks.add(idx + 1);
    }
  }
  return breaks;
}

// =============================================================================
// Merge Map
// =============================================================================

interface MergeInfo {
  isMaster: boolean;
  rowSpan: number;
  colSpan: number;
}

/**
 * Build a map of all merged cell regions.
 * Key: "row:col" (1-based), Value: merge info
 */
function buildMergeMap(sheet: PdfSheetData): Map<string, MergeInfo> {
  const map = new Map<string, MergeInfo>();

  const merges = sheet.merges;
  if (!merges || merges.length === 0) {
    return map;
  }

  for (const rangeStr of merges) {
    const range = parseRangeRef(rangeStr);
    const top = range.s.r + 1;
    const left = range.s.c + 1;
    const bottom = range.e.r + 1;
    const right = range.e.c + 1;

    const rowSpan = bottom - top + 1;
    const colSpan = right - left + 1;

    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        map.set(`${r}:${c}`, {
          isMaster: r === top && c === left,
          rowSpan,
          colSpan
        });
      }
    }
  }

  return map;
}

// =============================================================================
// Pagination
// =============================================================================

export function paginateRows(
  rowHeights: number[],
  availableHeight: number,
  repeatRowCount: number,
  rowBreaks: Set<number>
): number[][] {
  if (rowHeights.length === 0) {
    return [[]];
  }

  const pages: number[][] = [];
  let currentPage: number[] = [];
  let currentPageHeight = 0;
  let isFirstPage = true;
  let repeatedPrefixCount = 0;

  const addRepeatRows = () => {
    repeatedPrefixCount = 0;
    for (let h = 0; h < repeatRowCount && h < rowHeights.length; h++) {
      if (currentPageHeight + rowHeights[h] > availableHeight && currentPage.length > 0) {
        break;
      }
      currentPage.push(h);
      currentPageHeight += rowHeights[h];
      repeatedPrefixCount++;
    }
  };

  for (let i = 0; i < rowHeights.length; i++) {
    const rowHeight = rowHeights[i];
    const pageAvailable = availableHeight;
    let skipRepeatedRow = false;

    while (true) {
      // Force page break at row break positions, or when content overflows
      const forceBreak = rowBreaks.has(i) && currentPage.length > 0;
      if ((forceBreak || currentPageHeight + rowHeight > pageAvailable) && currentPage.length > 0) {
        const pageHasOnlyRepeatRows =
          !forceBreak &&
          !isFirstPage &&
          currentPage.length > 0 &&
          currentPage.length === repeatedPrefixCount;

        if (pageHasOnlyRepeatRows) {
          currentPage = [];
          currentPageHeight = 0;
          repeatedPrefixCount = 0;
          continue;
        }

        pages.push(currentPage);
        currentPage = [];
        currentPageHeight = 0;
        repeatedPrefixCount = 0;
        isFirstPage = false;
        addRepeatRows();
        continue;
      }

      if (!isFirstPage && i < repeatRowCount && currentPage.includes(i)) {
        skipRepeatedRow = true;
        break;
      }

      currentPage.push(i);
      currentPageHeight += rowHeight;
      break;
    }

    if (skipRepeatedRow) {
      continue;
    }
  }

  if (currentPage.length > 0) {
    pages.push(currentPage);
  }

  return pages.length > 0 ? pages : [[]];
}

/**
 * Split columns into groups for horizontal pagination.
 */
function paginateColumns(
  columnWidths: number[],
  contentWidth: number,
  sheet: PdfSheetData,
  visibleCols: number[]
): number[][] {
  if (columnWidths.length === 0) {
    return [[]];
  }

  // Build col break set (indices into visibleCols)
  const colBreaks = new Set<number>();
  const wsColBreaks = sheet.colBreaks ?? [];
  if (wsColBreaks.length > 0) {
    const colToIndex = new Map<number, number>();
    for (let i = 0; i < visibleCols.length; i++) {
      colToIndex.set(visibleCols[i], i);
    }
    for (const brk of wsColBreaks) {
      const idx = colToIndex.get(brk);
      if (idx !== undefined) {
        colBreaks.add(idx + 1);
      }
    }
  }

  const groups: number[][] = [];
  let currentGroup: number[] = [];
  let currentWidth = 0;

  for (let i = 0; i < columnWidths.length; i++) {
    const colWidth = columnWidths[i];

    const forceBreak = colBreaks.has(i) && currentGroup.length > 0;
    if ((forceBreak || currentWidth + colWidth > contentWidth + 0.01) && currentGroup.length > 0) {
      groups.push(currentGroup);
      currentGroup = [];
      currentWidth = 0;
    }

    currentGroup.push(i);
    currentWidth += colWidth;
  }

  if (currentGroup.length > 0) {
    groups.push(currentGroup);
  }

  return groups.length > 0 ? groups : [Array.from({ length: columnWidths.length }, (_, i) => i)];
}

// =============================================================================
// Cell Layout
// =============================================================================

function buildLayoutCell(
  cell: PdfCellData | undefined,
  x: number,
  y: number,
  width: number,
  height: number,
  colSpan: number,
  rowSpan: number,
  options: ResolvedPdfOptions,
  fontManager: FontManager,
  scaleFactor: number
): LayoutCell {
  const text = cell?.text ?? "";
  const style: Partial<PdfCellStyle> = cell?.style ?? {};

  const fontProps = extractFontProperties(
    style.font,
    options.defaultFontFamily,
    options.defaultFontSize
  );

  // Scale font size proportionally when fitToPage shrinks the layout
  const scaledFontSize = fontProps.fontSize * scaleFactor;

  // Register font and track text for subsetting
  if (fontManager.hasEmbeddedFont()) {
    fontManager.trackText(text);
  } else {
    const pdfFontName = resolvePdfFontName(fontProps.fontFamily, fontProps.bold, fontProps.italic);
    fontManager.ensureFont(pdfFontName);
  }

  // Rich text runs
  const richText = buildRichTextRuns(cell, options, fontManager, scaleFactor);

  const borders = excelBordersToPdf(style.border);

  return {
    text,
    rect: { x, y, width, height },
    fontFamily: fontProps.fontFamily,
    fontSize: scaledFontSize,
    bold: fontProps.bold,
    italic: fontProps.italic,
    strike: fontProps.strike,
    underline: fontProps.underline,
    textColor: fontProps.textColor,
    fillColor: excelFillToPdfColor(style.fill),
    horizontalAlign: resolveHorizontalAlign(style.alignment, cell?.type, cell?.result),
    verticalAlign: excelVAlignToPdf(style.alignment),
    wrapText: style.alignment?.wrapText ?? false,
    borders,
    borderInsets: {
      top: (borders.top?.width ?? 0) / 2,
      right: (borders.right?.width ?? 0) / 2,
      bottom: (borders.bottom?.width ?? 0) / 2,
      left: (borders.left?.width ?? 0) / 2
    },
    colSpan,
    rowSpan,
    hyperlink: cell?.hyperlink ?? null,
    richText,
    indent: style.alignment?.indent ?? 0,
    textRotation:
      style.alignment?.textRotation === 255 ? "vertical" : (style.alignment?.textRotation ?? 0),
    textOverflowWidth: 0
  };
}

// =============================================================================
// Shared-Edge Border Resolution
// =============================================================================

/**
 * Border precedence weight.
 *
 * When two adjacent cells both declare a border on a shared edge the winning
 * border is chosen by:  1. thicker wins,  2. solid beats dashed,
 * 3. double beats single,  4. darker colour wins (tie-break).
 *
 * Returns a numeric score – higher score wins.
 */
export function borderPrecedence(b: LayoutBorder): number {
  let score = b.width * 1000; // width dominates
  if (b.dashPattern.length === 0) {
    score += 100; // solid beats dashed
  }
  if (b.isDouble) {
    score += 50; // double beats single
  }
  // Darker colour = lower sum of RGB → higher score
  const brightness = b.color.r + b.color.g + b.color.b;
  score += (3 - brightness) * 10; // max RGB sum = 3 → adds up to 30
  return score;
}

/**
 * Resolve shared borders between adjacent cells.
 *
 * For each shared edge, determine the winning border (by precedence), then:
 * - The cell that "owns" the winning border keeps it in `borders` for drawing.
 * - The losing cell has that border side set to `null` (it won't draw).
 * - Both cells' `borderInsets` are updated to reflect the winning border's
 *   half-width, so text padding accounts for the line that is actually there.
 */
export function resolveSharedBorders(
  cellGrid: Map<string, LayoutCell>,
  rowCount: number,
  colCount: number
): void {
  for (let ri = 0; ri < rowCount; ri++) {
    for (let gci = 0; gci < colCount; gci++) {
      const cell = cellGrid.get(`${ri}:${gci}`);
      if (!cell) {
        continue;
      }

      // Horizontal shared edge: this cell's right border vs right neighbour's left
      if (cell.borders.right) {
        const rightNeighbor = cellGrid.get(`${ri}:${gci + 1}`);
        if (rightNeighbor?.borders.left) {
          const myScore = borderPrecedence(cell.borders.right);
          const theirScore = borderPrecedence(rightNeighbor.borders.left);
          if (theirScore > myScore) {
            // Neighbour wins — this cell stops drawing, but its inset = winner's half-width
            cell.borderInsets.right = rightNeighbor.borders.left.width / 2;
            cell.borders.right = null;
          } else {
            // This cell wins (or tie) — neighbour stops drawing
            rightNeighbor.borderInsets.left = cell.borders.right.width / 2;
            rightNeighbor.borders.left = null;
          }
        }
      }

      // Vertical shared edge: this cell's bottom border vs below neighbour's top
      if (cell.borders.bottom) {
        const belowNeighbor = cellGrid.get(`${ri + 1}:${gci}`);
        if (belowNeighbor?.borders.top) {
          const myScore = borderPrecedence(cell.borders.bottom);
          const theirScore = borderPrecedence(belowNeighbor.borders.top);
          if (theirScore > myScore) {
            cell.borderInsets.bottom = belowNeighbor.borders.top.width / 2;
            cell.borders.bottom = null;
          } else {
            belowNeighbor.borderInsets.top = cell.borders.bottom.width / 2;
            belowNeighbor.borders.top = null;
          }
        }
      }
    }
  }
}

// =============================================================================
// Image Placement
// =============================================================================

/**
 * Assign pre-collected images to the pages that contain their top-left anchor.
 */
function assignImagesToPages(
  images: PdfSheetImage[],
  layoutPages: LayoutPage[],
  scaleFactor: number
): void {
  for (const img of images) {
    const tl = img.range.tl;
    const tlCol = (tl.nativeCol ?? tl.col ?? 0) + 1; // convert 0-indexed to 1-indexed
    const tlRow = (tl.nativeRow ?? tl.row ?? 0) + 1;

    const targetPage = layoutPages.find(
      page => page.sheetCols.includes(tlCol) && page.sheetRows.includes(tlRow)
    );
    if (!targetPage) {
      continue;
    }

    const pageColIndex = targetPage.sheetCols.indexOf(tlCol);
    const pageRowIndex = targetPage.sheetRows.indexOf(tlRow);
    const baseX = targetPage.columnOffsets[pageColIndex] ?? targetPage.options.margins.left;
    const baseY =
      targetPage.rowYPositions[pageRowIndex] ??
      targetPage.height -
        targetPage.options.margins.top -
        (targetPage.options.showSheetNames ? 20 : 0);

    // Apply sub-cell offsets (EMU: 1pt = 12700 EMU), scaled to match page layout
    const tlColOff = ((tl.nativeColOff ?? 0) / 12700 || 0) * scaleFactor;
    const tlRowOff = ((tl.nativeRowOff ?? 0) / 12700 || 0) * scaleFactor;
    const imgX = baseX + tlColOff;
    const imgY = baseY - tlRowOff;

    // Determine image size
    let imgWidth = 100;
    let imgHeight = 100;
    if (img.range.ext) {
      imgWidth = (img.range.ext.width ?? 100) * 0.75 * scaleFactor;
      imgHeight = (img.range.ext.height ?? 100) * 0.75 * scaleFactor;
    } else if (img.range.br) {
      const br = img.range.br;
      const brCol = (br.nativeCol ?? br.col ?? 0) + 1;
      const brRow = (br.nativeRow ?? br.row ?? 0) + 1;
      const brPageColIndex = targetPage.sheetCols.indexOf(brCol);
      const brPageRowIndex = targetPage.sheetRows.indexOf(brRow);
      const brBaseX =
        brPageColIndex >= 0
          ? targetPage.columnOffsets[brPageColIndex]
          : imgX + (targetPage.columnWidths[pageColIndex] ?? 100);
      const brBaseY =
        brPageRowIndex >= 0
          ? targetPage.rowYPositions[brPageRowIndex]
          : imgY - (targetPage.rowHeights[pageRowIndex] ?? 100);
      const brColOff = ((br.nativeColOff ?? 0) / 12700 || 0) * scaleFactor;
      const brRowOff = ((br.nativeRowOff ?? 0) / 12700 || 0) * scaleFactor;
      const brX = brBaseX + brColOff;
      const brY = brBaseY - brRowOff;
      imgWidth = brX - imgX;
      imgHeight = imgY - brY;
    }

    targetPage.images.push({
      data: img.data,
      format: img.format,
      rect: {
        x: imgX,
        y: imgY - imgHeight,
        width: Math.abs(imgWidth),
        height: Math.abs(imgHeight)
      }
    });
  }
}

// =============================================================================
// Merge Border Propagation
// =============================================================================

/**
 * Excel stores merged-cell borders on the boundary cells, not on the master.
 * Copy the right border from the rightmost column cell and the bottom border
 * from the bottom row cell so the layout cell renders them correctly.
 */
function propagateMergeBorders(
  layoutCell: LayoutCell,
  mergeInfo: MergeInfo,
  wsRowNumber: number,
  wsColNumber: number,
  sheet: PdfSheetData
): void {
  if (mergeInfo.colSpan > 1) {
    const rightCol = wsColNumber + mergeInfo.colSpan - 1;
    const rightCellData = sheet.rows.get(wsRowNumber)?.cells.get(rightCol);
    if (rightCellData?.style?.border?.right) {
      const converted = excelBordersToPdf({ right: rightCellData.style.border.right });
      if (converted.right) {
        layoutCell.borders.right = converted.right;
        layoutCell.borderInsets.right = converted.right.width / 2;
      }
    }
  }
  if (mergeInfo.rowSpan > 1) {
    const bottomRowNum = wsRowNumber + mergeInfo.rowSpan - 1;
    const bottomCellData = sheet.rows.get(bottomRowNum)?.cells.get(wsColNumber);
    if (bottomCellData?.style?.border?.bottom) {
      const converted = excelBordersToPdf({ bottom: bottomCellData.style.border.bottom });
      if (converted.bottom) {
        layoutCell.borders.bottom = converted.bottom;
        layoutCell.borderInsets.bottom = converted.bottom.width / 2;
      }
    }
  }
}

// =============================================================================
// Text Overflow Calculation
// =============================================================================

/**
 * In Excel, non-wrapped text overflows into adjacent empty cells.
 * Fill color alone does NOT block overflow — only text content does.
 * Computes `textOverflowWidth` for cells whose text exceeds the cell width.
 */
function computeTextOverflows(
  cellGrid: Map<string, LayoutCell>,
  rowPage: number[],
  colGroup: number[],
  visibleRows: number[],
  visibleCols: number[],
  groupColWidths: number[],
  mergeMap: Map<string, MergeInfo>,
  fontManager: FontManager
): void {
  for (let ri = 0; ri < rowPage.length; ri++) {
    for (let gci = 0; gci < colGroup.length; gci++) {
      const cell = cellGrid.get(`${ri}:${gci}`);
      if (
        !cell ||
        cell.wrapText ||
        cell.colSpan > 1 ||
        !cell.text ||
        cell.richText ||
        (typeof cell.textRotation === "number" && cell.textRotation !== 0) ||
        cell.textRotation === "vertical"
      ) {
        continue;
      }

      const resourceName = fontManager.hasEmbeddedFont()
        ? fontManager.getEmbeddedResourceName()
        : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));
      const textWidth = fontManager.measureText(cell.text, resourceName, cell.fontSize);
      const cellContentWidth =
        cell.rect.width -
        (CELL_PADDING_H + cell.borderInsets.left) -
        (CELL_PADDING_H + cell.borderInsets.right);

      if (textWidth <= cellContentWidth) {
        continue;
      }

      const overflowNeeded = textWidth - cellContentWidth;
      let overflowAvailable = 0;

      for (let j = gci + 1; j < colGroup.length; j++) {
        const visibleRowIdx = rowPage[ri];
        const wsRow = visibleRows[visibleRowIdx];
        const wsCol = visibleCols[colGroup[j]];

        if (mergeMap.has(`${wsRow}:${wsCol}`)) {
          break;
        }

        const neighborCell = cellGrid.get(`${ri}:${j}`);
        if (neighborCell?.text) {
          break;
        }

        overflowAvailable += groupColWidths[j];
        if (overflowAvailable >= overflowNeeded) {
          break;
        }
      }

      if (overflowAvailable > 0) {
        cell.textOverflowWidth = Math.min(overflowNeeded, overflowAvailable);
      }
    }
  }
}

// =============================================================================
// Rich Text
// =============================================================================

/**
 * Build rich text runs from a RichText cell.
 * Returns null for non-RichText cells.
 */
function buildRichTextRuns(
  cell: PdfCellData | undefined,
  options: ResolvedPdfOptions,
  fontManager: FontManager,
  scaleFactor: number
): LayoutRichTextRun[] | null {
  if (!cell || cell.type !== PdfCellType.RichText) {
    return null;
  }

  const value = cell.value;
  if (!value || typeof value !== "object" || !("richText" in value)) {
    return null;
  }

  const runs = (value as { richText: PdfRichTextRunData[] }).richText;
  if (runs.length === 0) {
    return null;
  }

  return runs.map(run => {
    const fontProps = extractFontProperties(
      run.font,
      options.defaultFontFamily,
      options.defaultFontSize
    );

    // Register font for this run
    if (fontManager.hasEmbeddedFont()) {
      fontManager.trackText(run.text);
    } else {
      const pdfFontName = resolvePdfFontName(
        fontProps.fontFamily,
        fontProps.bold,
        fontProps.italic
      );
      fontManager.ensureFont(pdfFontName);
    }

    return {
      text: run.text,
      fontFamily: fontProps.fontFamily,
      fontSize: fontProps.fontSize * scaleFactor,
      bold: fontProps.bold,
      italic: fontProps.italic,
      strike: fontProps.strike,
      underline: fontProps.underline,
      textColor: fontProps.textColor
    };
  });
}
