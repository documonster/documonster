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
  ResolvedPdfOptions,
  LayoutPage,
  LayoutCell,
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
  excelVAlignToPdf
} from "./style-converter";

// =============================================================================
// Constants
// =============================================================================

const EXCEL_CHAR_WIDTH_TO_POINTS = 7;
const DEFAULT_COLUMN_WIDTH = 8.43;
const DEFAULT_ROW_HEIGHT = 15;
const MIN_COLUMN_WIDTH = 5;

// =============================================================================
// Layout Engine
// =============================================================================

/**
 * Compute the layout for a sheet across one or more PDF pages.
 */
export function layoutSheet(
  sheet: PdfSheetData,
  options: ResolvedPdfOptions,
  fontManager: FontManager
): LayoutPage[] {
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

  // Determine print area bounds (if set)
  const printRange = getPrintRange(sheet);

  // --- Step 1: Visible columns and widths ---
  const { columnWidths, visibleCols } = computeColumnWidths(sheet, printRange);
  const columnCount = visibleCols.length;

  if (columnCount === 0) {
    return [emptyPage(pageWidth, pageHeight, sheet.name, options)];
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
  const { rowHeights, visibleRows } = computeRowHeights(sheet, scaleFactor, printRange);

  // --- Step 4: Merge map ---
  const mergeMap = buildMergeMap(sheet);

  // --- Step 5: Paginate vertically (rows) and horizontally (columns) ---
  const repeatRowCount = typeof options.repeatRows === "number" ? options.repeatRows : 0;
  const rowBreakSet = buildRowBreakSet(sheet, visibleRows);
  const rowPages = paginateRows(rowHeights, availableHeight, repeatRowCount, rowBreakSet);
  const colGroups = paginateColumns(scaledColumnWidths, contentWidth, sheet, visibleCols);

  // --- Step 6: Layout cells per page (row page × column page) ---
  const layoutPages: LayoutPage[] = [];

  for (const rowPage of rowPages) {
    for (const colGroup of colGroups) {
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
      for (let ri = 0; ri < rowPage.length; ri++) {
        const visibleRowIdx = rowPage[ri];
        const wsRowNumber = visibleRows[visibleRowIdx];

        for (let gci = 0; gci < colGroup.length; gci++) {
          const ci = colGroup[gci]; // index into visibleCols
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
        }
      }

      layoutPages.push({
        pageNumber: layoutPages.length + 1,
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
        images: []
      });
    }
  }

  // --- Step 7: Place images on the correct pages ---
  if (layoutPages.length > 0 && sheet.images) {
    assignImagesToPages(sheet.images, layoutPages);
  }

  return layoutPages;
}

// =============================================================================
// Helpers
// =============================================================================

function emptyPage(
  width: number,
  height: number,
  sheetName: string,
  options: ResolvedPdfOptions
): LayoutPage {
  return {
    pageNumber: 1,
    options,
    cells: [],
    width,
    height,
    sheetName,
    sheetCols: [],
    columnOffsets: [],
    columnWidths: [],
    sheetRows: [],
    rowYPositions: [],
    rowHeights: [],
    images: []
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
    const pointWidth = Math.max(excelWidth * EXCEL_CHAR_WIDTH_TO_POINTS, MIN_COLUMN_WIDTH);
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
  printRange: PrintRange | null
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
    if (row && row.hidden) {
      continue;
    }

    let height: number;
    if (row?.height) {
      // Explicit row height set by user
      height = row.height;
    } else {
      // Auto-size: scan cells in this row to find the largest needed height.
      height = DEFAULT_ROW_HEIGHT;
      if (row) {
        for (const cell of row.cells.values()) {
          let fontSize = cell.style?.font?.size ?? 11;

          // For rich text cells, find the largest font size across all runs
          const rtValue = cell.value as { richText?: Array<{ font?: { size?: number } }> } | null;
          if (rtValue?.richText) {
            for (const run of rtValue.richText) {
              const runSize = run.font?.size ?? fontSize;
              if (runSize > fontSize) {
                fontSize = runSize;
              }
            }
          }

          const lineHeight = fontSize * 1.5;

          // Count lines: explicit newlines in the text
          const text = cell.text ?? "";
          const lineCount = Math.max(1, (text.match(/\n/g) ?? []).length + 1);

          // For wrapText cells, estimate how many lines word-wrapping produces
          let wrapLineCount = lineCount;
          if (cell.style?.alignment?.wrapText && lineCount === 1 && text.length > 0) {
            const col = sheet.columns.get(cell.col);
            const colWidth = col?.width ?? DEFAULT_COLUMN_WIDTH;
            const colPts = colWidth * EXCEL_CHAR_WIDTH_TO_POINTS * scaleFactor;
            const avgCharWidth = fontSize * 0.55; // rough average char width
            const charsPerLine = Math.max(1, Math.floor(colPts / avgCharWidth));
            wrapLineCount = Math.ceil(text.length / charsPerLine);
          }

          const neededHeight = lineHeight * wrapLineCount;

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
    horizontalAlign: excelHAlignToPdf(style.alignment),
    verticalAlign: excelVAlignToPdf(style.alignment),
    wrapText: style.alignment?.wrapText ?? false,
    borders: excelBordersToPdf(style.border),
    colSpan,
    rowSpan,
    hyperlink: cell?.hyperlink ?? null,
    richText,
    indent: style.alignment?.indent ?? 0,
    textRotation: style.alignment?.textRotation ?? 0
  };
}

// =============================================================================
// Image Placement
// =============================================================================

/**
 * Assign pre-collected images to the pages that contain their top-left anchor.
 */
function assignImagesToPages(images: PdfSheetImage[], layoutPages: LayoutPage[]): void {
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

    // Apply sub-cell offsets (EMU: 1pt = 12700 EMU)
    const tlColOff = (tl.nativeColOff ?? 0) / 12700 || 0;
    const tlRowOff = (tl.nativeRowOff ?? 0) / 12700 || 0;
    const imgX = baseX + tlColOff;
    const imgY = baseY - tlRowOff;

    // Determine image size
    let imgWidth = 100;
    let imgHeight = 100;
    if (img.range.ext) {
      imgWidth = (img.range.ext.width ?? 100) * 0.75;
      imgHeight = (img.range.ext.height ?? 100) * 0.75;
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
      const brColOff = (br.nativeColOff ?? 0) / 12700 || 0;
      const brRowOff = (br.nativeRowOff ?? 0) / 12700 || 0;
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

  const rtValue = cell.value as { richText?: PdfRichTextRunData[] };
  if (!rtValue?.richText || rtValue.richText.length === 0) {
    return null;
  }

  return rtValue.richText.map(run => {
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
