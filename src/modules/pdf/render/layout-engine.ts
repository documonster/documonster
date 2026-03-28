/**
 * Layout engine for Excel-to-PDF conversion.
 *
 * Takes a worksheet and produces LayoutPage objects that describe exactly
 * where each cell, border, and piece of text should be drawn on each PDF page.
 *
 * Key responsibilities:
 * - Convert Excel column widths (character units) to PDF points
 * - Convert Excel row heights (points already, but may need scaling)
 * - Handle merged cells spanning multiple rows/columns
 * - Paginate content across multiple pages
 * - Handle fitToPage scaling
 * - Handle repeated header rows
 * - Skip hidden rows and columns
 */

import type { Worksheet } from "@excel/worksheet";
import type { Cell } from "@excel/cell";
import type { Style, CellErrorValue, CellRichTextValue } from "@excel/types";
import { ValueType } from "@excel/enums";
import { decodeRange } from "@excel/utils/address";
import { formatCellValue } from "@excel/utils/cell-format";
import { base64ToUint8Array } from "@utils/utils.base";
import type { ResolvedPdfOptions, LayoutPage, LayoutCell, LayoutRichTextRun } from "../types";
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
 * Compute the layout for a worksheet across one or more PDF pages.
 */
export function layoutWorksheet(
  worksheet: Worksheet,
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
  const printRange = getPrintRange(worksheet);

  // --- Step 1: Visible columns and widths ---
  const { columnWidths, visibleCols } = computeColumnWidths(worksheet, printRange);
  const columnCount = visibleCols.length;

  if (columnCount === 0) {
    return [emptyPage(pageWidth, pageHeight, worksheet.name, options)];
  }

  // --- Step 2: Scale ---
  let totalTableWidth = columnWidths.reduce((sum, w) => sum + w, 0);
  let scaleFactor = options.scale;
  if (options.fitToPage && totalTableWidth > 0) {
    const fitScale = contentWidth / totalTableWidth;
    if (fitScale < 1) {
      scaleFactor *= fitScale;
    }
  }

  const scaledColumnWidths = columnWidths.map(w => w * scaleFactor);
  totalTableWidth = scaledColumnWidths.reduce((sum, w) => sum + w, 0);

  // Column x-offsets
  const columnOffsets: number[] = [];
  let xOffset = margins.left;
  if (totalTableWidth < contentWidth) {
    xOffset = margins.left + (contentWidth - totalTableWidth) / 2;
  }
  for (let i = 0; i < scaledColumnWidths.length; i++) {
    columnOffsets.push(xOffset);
    xOffset += scaledColumnWidths[i];
  }

  // --- Step 3: Visible rows and heights ---
  const { rowHeights, visibleRows } = computeRowHeights(worksheet, scaleFactor, printRange);

  // --- Step 4: Merge map ---
  const mergeMap = buildMergeMap(worksheet);

  // --- Step 5: Paginate vertically (rows) and horizontally (columns) ---
  const repeatRowCount = typeof options.repeatRows === "number" ? options.repeatRows : 0;
  const rowBreakSet = buildRowBreakSet(worksheet, visibleRows);
  const rowPages = paginateRows(rowHeights, availableHeight, repeatRowCount, rowBreakSet);
  const colGroups = paginateColumns(scaledColumnWidths, contentWidth, worksheet, visibleCols);

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

          const row = worksheet.findRow(wsRowNumber);
          const cell = row ? row.findCell(wsColNumber) : undefined;

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
              ci,
              visibleRowIdx,
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
        sheetName: worksheet.name,
        worksheetCols: colGroup.map(ci => visibleCols[ci]),
        columnOffsets: groupColOffsets,
        columnWidths: groupColWidths,
        worksheetRows: rowPage.map(ri => visibleRows[ri]),
        rowYPositions,
        rowHeights: pageRowHeights,
        images: []
      });
    }
  }

  // --- Step 7: Collect images and place them on the correct pages ---
  if (layoutPages.length > 0) {
    assignImagesToPages(worksheet, layoutPages);
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
    worksheetCols: [],
    columnOffsets: [],
    columnWidths: [],
    worksheetRows: [],
    rowYPositions: [],
    rowHeights: [],
    images: []
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
 * Get the print area range from the worksheet's pageSetup.
 * Returns null if no print area is set.
 */
function getPrintRange(worksheet: Worksheet): PrintRange | null {
  const printArea = (worksheet as any).pageSetup?.printArea;
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
    const range = decodeRange(firstRange);
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
  worksheet: Worksheet,
  printRange: PrintRange | null
): {
  columnWidths: number[];
  visibleCols: number[];
} {
  const dimensions = worksheet.dimensions;
  const hasData = dimensions && dimensions.model.top > 0 && dimensions.model.left > 0;

  if (!hasData) {
    return { columnWidths: [], visibleCols: [] };
  }

  const startCol = printRange?.startCol ?? dimensions.model.left;
  const endCol = printRange?.endCol ?? dimensions.model.right;
  const columnWidths: number[] = [];
  const visibleCols: number[] = [];

  for (let c = startCol; c <= endCol; c++) {
    const col = worksheet.getColumn(c);
    if (col.hidden) {
      continue;
    }
    const excelWidth = col.width ?? DEFAULT_COLUMN_WIDTH;
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
  worksheet: Worksheet,
  scaleFactor: number,
  printRange: PrintRange | null
): { rowHeights: number[]; visibleRows: number[] } {
  const dimensions = worksheet.dimensions;
  if (!dimensions || dimensions.model.top <= 0) {
    return { rowHeights: [], visibleRows: [] };
  }

  const startRow = printRange?.startRow ?? dimensions.model.top;
  const endRow = printRange?.endRow ?? dimensions.model.bottom;
  const rowHeights: number[] = [];
  const visibleRows: number[] = [];

  for (let r = startRow; r <= endRow; r++) {
    const row = worksheet.findRow(r);
    if (row && row.hidden) {
      continue;
    }

    let height: number;
    if (row?.height) {
      // Explicit row height set by user
      height = row.height;
    } else {
      // Auto-size: scan cells in this row to find the largest needed height.
      // Account for font size and multi-line content (explicit newlines or wrapText).
      height = DEFAULT_ROW_HEIGHT;
      if (row) {
        row.eachCell({ includeEmpty: false }, cell => {
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
          if (cell.alignment?.wrapText && lineCount === 1 && text.length > 0) {
            const colWidth = worksheet.getColumn(cell.col).width ?? DEFAULT_COLUMN_WIDTH;
            const colPts = colWidth * EXCEL_CHAR_WIDTH_TO_POINTS * scaleFactor;
            const avgCharWidth = fontSize * 0.55; // rough average char width
            const charsPerLine = Math.max(1, Math.floor(colPts / avgCharWidth));
            wrapLineCount = Math.ceil(text.length / charsPerLine);
          }

          const neededHeight = lineHeight * wrapLineCount;

          if (neededHeight > height) {
            height = neededHeight;
          }
        });
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
function buildRowBreakSet(worksheet: Worksheet, visibleRows: number[]): Set<number> {
  const breaks = new Set<number>();
  const rowBreaks: Array<{ id: number }> = (worksheet as any).rowBreaks ?? [];
  if (rowBreaks.length === 0) {
    return breaks;
  }
  // Map worksheet row numbers to visible-row indices
  const rowToIndex = new Map<number, number>();
  for (let i = 0; i < visibleRows.length; i++) {
    rowToIndex.set(visibleRows[i], i);
  }
  for (const brk of rowBreaks) {
    const idx = rowToIndex.get(brk.id);
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
 * Build a map of all merged cell regions in the worksheet.
 * Uses the worksheet model's public mergeCells property.
 * Key: "row:col" (1-based), Value: merge info
 */
function buildMergeMap(worksheet: Worksheet): Map<string, MergeInfo> {
  const map = new Map<string, MergeInfo>();

  if (!worksheet.hasMerges) {
    return map;
  }

  const mergeCells = worksheet.model.mergeCells;
  if (!mergeCells) {
    return map;
  }

  for (const rangeStr of mergeCells) {
    const range = decodeRange(rangeStr);
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
          // If repeated header rows consume too much space to allow any body row,
          // fall back to placing body rows without repeated headers rather than
          // emitting a header-only page or overflowing the page.
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
 * Each group is an array of column indices (into the visibleCols/scaledColumnWidths arrays).
 * If the total width fits in contentWidth and there are no colBreaks, returns a single group.
 */
function paginateColumns(
  columnWidths: number[],
  contentWidth: number,
  worksheet: Worksheet,
  visibleCols: number[]
): number[][] {
  if (columnWidths.length === 0) {
    return [[]];
  }

  // Build col break set (indices into visibleCols)
  const colBreaks = new Set<number>();
  const wsColBreaks: Array<{ id: number }> = (worksheet as any).colBreaks ?? [];
  if (wsColBreaks.length > 0) {
    const colToIndex = new Map<number, number>();
    for (let i = 0; i < visibleCols.length; i++) {
      colToIndex.set(visibleCols[i], i);
    }
    for (const brk of wsColBreaks) {
      const idx = colToIndex.get(brk.id);
      if (idx !== undefined) {
        // Break AFTER this column, so the next column starts a new group
        colBreaks.add(idx + 1);
      }
    }
  }

  const groups: number[][] = [];
  let currentGroup: number[] = [];
  let currentWidth = 0;

  for (let i = 0; i < columnWidths.length; i++) {
    const colWidth = columnWidths[i];

    // Force break at column break positions or when exceeding page width
    // Use a small epsilon (0.01pt) to avoid floating-point precision issues
    // when fitToPage scales columns to exactly match content width
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
  cell: Cell | undefined,
  _colIndex: number,
  _rowIndex: number,
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
  const text = getCellText(cell);
  const style: Partial<Style> = cell?.style ?? {};

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

  // Rich text runs (buildRichTextRuns handles font tracking internally)
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
// Image Collection
// =============================================================================

/**
 * Collect images from a worksheet and assign them to the page that contains
 * their top-left anchor.
 */
function assignImagesToPages(worksheet: Worksheet, layoutPages: LayoutPage[]): void {
  // Access worksheet images via getImages()
  const wsImages = (worksheet as any).getImages?.();
  if (!wsImages || !Array.isArray(wsImages)) {
    return;
  }

  // Access the workbook for image data
  const workbook = (worksheet as any).workbook;
  if (!workbook) {
    return;
  }

  for (const wsImage of wsImages) {
    if (!wsImage.range?.tl) {
      continue;
    }

    const imageId = wsImage.imageId;
    const mediaItem = workbook.getImage?.(Number(imageId));
    if (!mediaItem) {
      continue;
    }

    // Get image data
    let data: Uint8Array | undefined;
    if (mediaItem.buffer instanceof Uint8Array) {
      data = mediaItem.buffer;
    } else if (mediaItem.base64) {
      data = base64ToUint8Array(mediaItem.base64);
    }
    if (!data || data.length === 0) {
      continue;
    }

    const format = mediaItem.extension as string;
    if (format !== "jpeg" && format !== "png") {
      continue; // Only JPEG and PNG are supported
    }

    // Calculate position from anchor
    const tl = wsImage.range.tl;
    const tlCol = (tl.nativeCol ?? tl.col ?? 0) + 1; // convert 0-indexed to 1-indexed
    const tlRow = (tl.nativeRow ?? tl.row ?? 0) + 1;

    const targetPage = layoutPages.find(
      page => page.worksheetCols.includes(tlCol) && page.worksheetRows.includes(tlRow)
    );
    if (!targetPage) {
      continue;
    }

    const pageColIndex = targetPage.worksheetCols.indexOf(tlCol);
    const pageRowIndex = targetPage.worksheetRows.indexOf(tlRow);
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
    const imgY = baseY - tlRowOff; // PDF y-axis is bottom-up, offset moves down

    // Determine image size
    let imgWidth = 100;
    let imgHeight = 100;
    if (wsImage.range.ext) {
      // ext.width and ext.height are in pixels; convert to points (1px ≈ 0.75pt)
      imgWidth = (wsImage.range.ext.width ?? 100) * 0.75;
      imgHeight = (wsImage.range.ext.height ?? 100) * 0.75;
    } else if (wsImage.range.br) {
      // Calculate from bottom-right anchor
      const br = wsImage.range.br;
      const brCol = (br.nativeCol ?? br.col ?? 0) + 1;
      const brRow = (br.nativeRow ?? br.row ?? 0) + 1;
      const brPageColIndex = targetPage.worksheetCols.indexOf(brCol);
      const brPageRowIndex = targetPage.worksheetRows.indexOf(brRow);
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

    // PDF coordinates: y is bottom of image
    targetPage.images.push({
      data,
      format,
      rect: {
        x: imgX,
        y: imgY - imgHeight,
        width: Math.abs(imgWidth),
        height: Math.abs(imgHeight)
      }
    });
  }
}

/**
 * Extract display text from a cell, applying numFmt formatting.
 */
function getCellText(cell: Cell | undefined): string {
  if (!cell) {
    return "";
  }

  switch (cell.type) {
    case ValueType.Null:
    case ValueType.Merge:
      return "";
    case ValueType.RichText: {
      // RichText cell.value is a CellRichTextValue object; cell.text joins the run texts
      return cell.text;
    }
    case ValueType.Hyperlink:
      return cell.text;
    case ValueType.Error: {
      // Error cells have value = { error: "#N/A" } etc.
      const errValue = cell.value as CellErrorValue | null;
      return errValue?.error ?? cell.text;
    }
    case ValueType.Formula: {
      const result = cell.result;
      if (result !== undefined && result !== null) {
        if (typeof result === "object" && "error" in result) {
          return (result as CellErrorValue).error;
        }
        return formatCellValueSafe(result, cell.style?.numFmt);
      }
      return cell.text;
    }
    default: {
      const value = cell.value;
      if (value === null || value === undefined) {
        return "";
      }
      return formatCellValueSafe(value, cell.style?.numFmt);
    }
  }
}

/**
 * Safely format a cell value using its numFmt.
 * Falls back to toString if formatting fails or no format is specified.
 */
function formatCellValueSafe(
  value: unknown,
  numFmt: string | { formatCode: string } | undefined
): string {
  const fmt = typeof numFmt === "string" ? numFmt : numFmt?.formatCode;
  if (fmt && (typeof value === "number" || value instanceof Date || typeof value === "boolean")) {
    try {
      return formatCellValue(value, fmt);
    } catch {
      // Fall through to default
    }
  }
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }
  return String(value);
}

// =============================================================================
// Rich Text
// =============================================================================

/**
 * Build rich text runs from a RichText cell.
 * Returns null for non-RichText cells.
 */
function buildRichTextRuns(
  cell: Cell | undefined,
  options: ResolvedPdfOptions,
  fontManager: FontManager,
  scaleFactor: number
): LayoutRichTextRun[] | null {
  if (!cell || cell.type !== ValueType.RichText) {
    return null;
  }

  const rtValue = cell.value as CellRichTextValue;
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
