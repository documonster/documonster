/**
 * Layout engine: converts Excel Worksheet data into positioned PDF operations.
 *
 * Handles:
 * - Column widths and row heights → PDF points
 * - Cell text rendering with font, alignment, color
 * - Cell borders (thin, medium, thick, dashed, etc.)
 * - Cell fills (solid pattern fills)
 * - Merged cells
 * - Text wrapping
 * - Image placement
 * - Pagination (splitting across multiple pages)
 */

import { mapFontFamily, getPdfFontName, measureTextWidth, wrapText } from "@pdf/font-metrics";
import type { PdfOp, PdfPageDef, PdfColor, PdfImageData } from "@pdf/pdf-writer";
import type { Worksheet } from "@excel/worksheet";
import type { Cell } from "@excel/cell";
import type { Font, Borders, BorderStyle, Fill, Alignment, Color, NumFmt } from "@excel/types";
import { ValueType } from "@excel/enums";
import type { WorkbookMedia } from "@excel/workbook.browser";
import { formatCellValue } from "@excel/utils/cell-format";

// =============================================================================
// Constants
// =============================================================================

/**
 * Points per Excel column width unit.
 * Excel column width is in "character widths" of the default font.
 * For Calibri 11pt at 96 DPI, maxDigitWidth ≈ 7 pixels.
 * pixel_width ≈ width * 7 ; point_width = pixel_width * 72/96 = width * 5.25
 */
const COL_WIDTH_FACTOR = 5.25;
/** Points per Excel row height unit (row height is in points already) */
const ROW_HEIGHT_FACTOR = 1;
/** Default column width in Excel units */
const DEFAULT_COL_WIDTH = 9;
/** Default row height in points */
const DEFAULT_ROW_HEIGHT = 15;
/** Cell padding in points (matches Excel's ~3px default cell margin) */
const CELL_PADDING = 3;
/** Points per indent level (one character width of default font) */
const INDENT_WIDTH = 7;
/** Default font size in points */
const DEFAULT_FONT_SIZE = 11;
/** Default page width (A4 portrait) */
const A4_WIDTH = 595.28;
/** Default page height (A4 portrait) */
const A4_HEIGHT = 841.89;
/** Letter page width */
const LETTER_WIDTH = 612;
/** Letter page height */
const LETTER_HEIGHT = 792;

// =============================================================================
// Options
// =============================================================================

export interface PdfLayoutOptions {
  /** Page width in points (default: A4 = 595.28) */
  pageWidth?: number;
  /** Page height in points (default: A4 = 841.89) */
  pageHeight?: number;
  /** Page margins in points */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Whether to show grid lines (default: true) */
  gridLines?: boolean;
  /** Scale factor (default: 1.0) */
  scale?: number;
  /** Orientation override ('portrait' | 'landscape') */
  orientation?: "portrait" | "landscape";
  /** Whether to fit all columns on one page width (default: false) */
  fitToWidth?: boolean;
  /** Custom column widths override (1-based index to width in Excel units) */
  columnWidths?: Record<number, number>;
  /** Specific worksheet indices to include (0-based). If omitted, all worksheets. */
  worksheets?: number[];
}

// =============================================================================
// Color conversion
// =============================================================================

/**
 * Default Office theme color palette (indices 0–11).
 * These match the standard Office 2013+ / Office 365 theme.
 * Stored as [R, G, B] in 0–255 range.
 */
const DEFAULT_THEME_COLORS: [number, number, number][] = [
  [255, 255, 255], // 0  lt1 / bg1 (white)
  [0, 0, 0], //       1  dk1 / tx1 (black)
  [231, 230, 230], // 2  lt2 / bg2
  [68, 84, 106], //   3  dk2 / tx2
  [68, 114, 196], //  4  accent1
  [237, 125, 49], //  5  accent2
  [165, 165, 165], // 6  accent3
  [255, 192, 0], //   7  accent4
  [91, 155, 213], //  8  accent5
  [112, 173, 71], //  9  accent6
  [5, 99, 193], //    10 hlink
  [149, 79, 114] //   11 folHlink
];

function argbToColor(argb?: string): PdfColor | undefined {
  if (!argb) {
    return undefined;
  }
  // ARGB format: "FF000000" or "000000"
  let hex = argb;
  if (hex.length === 8) {
    hex = hex.substring(2); // Strip alpha
  }
  if (hex.length !== 6) {
    return undefined;
  }
  const r = Number.parseInt(hex.substring(0, 2), 16) / 255;
  const g = Number.parseInt(hex.substring(2, 4), 16) / 255;
  const b = Number.parseInt(hex.substring(4, 6), 16) / 255;
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return undefined;
  }
  return { r, g, b };
}

/**
 * Apply an Excel tint value to a base color component (0–1 range).
 *
 *  - tint > 0 → lighten (mix towards white)
 *  - tint < 0 → darken (mix towards black)
 */
function applyTint(value: number, tint: number): number {
  if (tint > 0) {
    return value + (1 - value) * tint;
  }
  if (tint < 0) {
    return value * (1 + tint);
  }
  return value;
}

function themeToColor(theme: number, tint?: number): PdfColor | undefined {
  const rgb = DEFAULT_THEME_COLORS[theme];
  if (!rgb) {
    return undefined;
  }
  let r = rgb[0] / 255;
  let g = rgb[1] / 255;
  let b = rgb[2] / 255;

  if (tint !== undefined && tint !== 0) {
    r = applyTint(r, tint);
    g = applyTint(g, tint);
    b = applyTint(b, tint);
  }

  return {
    r: Math.max(0, Math.min(1, r)),
    g: Math.max(0, Math.min(1, g)),
    b: Math.max(0, Math.min(1, b))
  };
}

function excelColorToRgb(color?: Partial<Color>): PdfColor | undefined {
  if (!color) {
    return undefined;
  }
  if (color.argb) {
    return argbToColor(color.argb);
  }
  if (color.theme !== undefined) {
    return themeToColor(color.theme, (color as any).tint);
  }
  return undefined;
}

const BLACK: PdfColor = { r: 0, g: 0, b: 0 };
const WHITE: PdfColor = { r: 1, g: 1, b: 1 };
const GRID_COLOR: PdfColor = { r: 0.8, g: 0.8, b: 0.8 };

// =============================================================================
// Border helpers
// =============================================================================

function borderLineWidth(style?: BorderStyle): number {
  if (!style) {
    return 0;
  }
  switch (style) {
    case "hair":
      return 0.25;
    case "thin":
      return 0.75;
    case "medium":
      return 1.5;
    case "thick":
      return 2.5;
    case "double":
      return 0.75;
    case "dashed":
      return 0.75;
    case "mediumDashed":
      return 1.5;
    case "dotted":
      return 0.5;
    case "dashDot":
    case "dashDotDot":
    case "slantDashDot":
      return 0.75;
    case "mediumDashDot":
    case "mediumDashDotDot":
      return 1.5;
    default:
      return 0.75;
  }
}

function borderDash(style?: BorderStyle): number[] | undefined {
  switch (style) {
    case "dashed":
    case "mediumDashed":
      return [4, 2];
    case "dotted":
      return [1, 1];
    case "dashDot":
    case "mediumDashDot":
      return [4, 2, 1, 2];
    case "dashDotDot":
    case "mediumDashDotDot":
      return [4, 2, 1, 2, 1, 2];
    default:
      return undefined;
  }
}

// =============================================================================
// Fill helpers
// =============================================================================

function getFillColor(fill?: Fill): PdfColor | undefined {
  if (!fill) {
    return undefined;
  }
  if (fill.type === "pattern") {
    if (fill.pattern === "solid" && fill.fgColor) {
      return excelColorToRgb(fill.fgColor);
    }
    if (fill.pattern === "none") {
      return undefined;
    }
    // Other patterns: use fgColor if available
    if (fill.fgColor) {
      return excelColorToRgb(fill.fgColor);
    }
  }
  // Gradient fills: use the first stop color
  if (fill.type === "gradient" && "stops" in fill && fill.stops?.length > 0) {
    return excelColorToRgb(fill.stops[0].color);
  }
  return undefined;
}

// =============================================================================
// Text extraction
// =============================================================================

/**
 * Resolve the numFmt string from a cell's numFmt property.
 */
function resolveNumFmt(numFmt: string | NumFmt | undefined): string {
  if (!numFmt) {
    return "General";
  }
  return typeof numFmt === "string" ? numFmt : (numFmt.formatCode ?? "General");
}

/**
 * Format a primitive value using the cell's number format.
 */
function formatPrimitive(value: Date | number | boolean | string, fmt: string): string {
  return formatCellValue(value, fmt);
}

function getCellText(cell: Cell): string {
  if (!cell || cell.type === ValueType.Null || cell.type === ValueType.Merge) {
    return "";
  }

  const val = cell.value;
  if (val === null || val === undefined) {
    return "";
  }

  const fmt = resolveNumFmt(cell.numFmt);

  // Rich text
  if (typeof val === "object" && "richText" in val && Array.isArray(val.richText)) {
    return val.richText.map(rt => rt.text).join("");
  }

  // Hyperlink
  if (typeof val === "object" && "hyperlink" in val) {
    return (val as any).text ?? (val as any).hyperlink ?? "";
  }

  // Formula
  if (typeof val === "object" && "formula" in val) {
    const result = (val as any).result;
    if (result !== undefined && result !== null) {
      if (
        result instanceof Date ||
        typeof result === "number" ||
        typeof result === "boolean" ||
        typeof result === "string"
      ) {
        return formatPrimitive(result, fmt);
      }
      return String(result);
    }
    return "";
  }

  // Shared formula
  if (typeof val === "object" && "sharedFormula" in val) {
    const result = (val as any).result;
    if (result !== undefined && result !== null) {
      if (
        result instanceof Date ||
        typeof result === "number" ||
        typeof result === "boolean" ||
        typeof result === "string"
      ) {
        return formatPrimitive(result, fmt);
      }
      return String(result);
    }
    return "";
  }

  // Error
  if (typeof val === "object" && "error" in val) {
    return (val as any).error;
  }

  // Checkbox
  if (typeof val === "object" && "checkbox" in val) {
    return (val as any).checkbox ? "\u2611" : "\u2610";
  }

  // Primitive values: apply number format
  if (
    val instanceof Date ||
    typeof val === "number" ||
    typeof val === "boolean" ||
    typeof val === "string"
  ) {
    return formatPrimitive(val, fmt);
  }

  if (typeof val === "object") {
    return "";
  }
  return String(val);
}

// =============================================================================
// Merge map builder
// =============================================================================

interface MergeInfo {
  /** The master cell address (e.g., "A1") */
  master: string;
  /** Number of columns this merge spans */
  colSpan: number;
  /** Number of rows this merge spans */
  rowSpan: number;
  /** Starting row (1-based) */
  startRow: number;
  /** Starting col (1-based) */
  startCol: number;
}

// =============================================================================
// Shared context/rect types used by rendering helpers
// =============================================================================

/** Bounding rectangle of a PDF cell in points */
interface CellRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Context bundle passed to page-generation helpers */
interface PageGenCtx {
  ws: Worksheet;
  startRow: number;
  endRow: number;
  colCount: number;
  colWidths: number[];
  rowHeights: number[];
  margins: { top: number; right: number; bottom: number; left: number };
  pageHeight: number;
  scale: number;
  mergeMap: Map<string, MergeInfo>;
  showGrid: boolean;
  wsImages: any[];
}

/** Context for text rendering helpers */
interface TextDrawCtx {
  pdfFontName: string;
  family: string;
  bold: boolean;
  italic: boolean;
  fontSize: number;
  textColor: PdfColor;
  hAlign: string;
  vAlign: string;
  padding: number;
  font?: Partial<Font>;
}

function buildMergeMap(ws: Worksheet): Map<string, MergeInfo> {
  const mergeMap = new Map<string, MergeInfo>();
  const model = ws.model;

  if (model.mergeCells) {
    for (const mergeRange of model.mergeCells) {
      // Parse range like "A1:C3"
      const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(mergeRange);
      if (!match) {
        continue;
      }
      const startCol = letterToCol(match[1]);
      const startRow = Number.parseInt(match[2], 10);
      const endCol = letterToCol(match[3]);
      const endRow = Number.parseInt(match[4], 10);

      const colSpan = endCol - startCol + 1;
      const rowSpan = endRow - startRow + 1;
      const master = match[1].toUpperCase() + match[2];

      const info: MergeInfo = { master, colSpan, rowSpan, startRow, startCol };

      // Mark all cells in the merge range
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const addr = colToLetter(c) + r;
          mergeMap.set(addr, info);
        }
      }
    }
  }

  return mergeMap;
}

function letterToCol(letter: string): number {
  let col = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + ((upper.codePointAt(i) ?? 64) - 64);
  }
  return col;
}

function colToLetter(col: number): string {
  let result = "";
  let c = col;
  while (c > 0) {
    const rem = (c - 1) % 26;
    result = String.fromCodePoint(65 + rem) + result;
    c = Math.floor((c - 1) / 26);
  }
  return result;
}

// =============================================================================
// Row height auto-expansion
// =============================================================================

/**
 * Auto-expand row heights so that cell text (especially wrapped text) fits
 * without overflowing into adjacent rows.
 */
function autoExpandRowHeights(
  ws: Worksheet,
  rowCount: number,
  colCount: number,
  rowHeights: number[],
  colWidths: number[],
  mergeMap: Map<string, MergeInfo>,
  scale: number
): void {
  const DEFAULT_RH = DEFAULT_ROW_HEIGHT * scale;
  const padding = CELL_PADDING * scale;
  const processedMerges = new Set<string>();

  for (let r = 1; r <= rowCount; r++) {
    const row = ws.findRow(r);
    if (!row) {
      continue;
    }

    let maxRequiredHeight = rowHeights[r] ?? DEFAULT_RH;

    for (let c = 1; c <= colCount; c++) {
      const cell = row.findCell(c);
      if (!cell || cell.type === ValueType.Null || cell.type === ValueType.Merge) {
        continue;
      }

      const cellAddr = colToLetter(c) + r;
      const mergeInfo = mergeMap.get(cellAddr);

      // Only process master cells and non-merged cells
      if (mergeInfo && cellAddr !== mergeInfo.master) {
        continue;
      }
      if (mergeInfo && processedMerges.has(mergeInfo.master)) {
        continue;
      }

      const text = getCellText(cell);
      if (!text) {
        continue;
      }

      const font = cell.font;
      const alignment = cell.alignment;
      const family = mapFontFamily(font?.name);
      const bold = !!font?.bold;
      const italic = !!font?.italic;
      const fontSize = (font?.size ?? DEFAULT_FONT_SIZE) * scale;
      const shouldWrap = alignment?.wrapText ?? false;

      // Compute available width for text
      let cellWidth: number;
      if (mergeInfo) {
        cellWidth = 0;
        for (let mc = mergeInfo.startCol; mc < mergeInfo.startCol + mergeInfo.colSpan; mc++) {
          cellWidth += colWidths[mc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
        }
      } else {
        cellWidth = colWidths[c] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
      }

      const availableWidth = cellWidth - padding * 2;
      const lineHeight = fontSize * 1.2;

      let lines: string[];
      if (shouldWrap) {
        lines = wrapText(text, family, bold, italic, fontSize, availableWidth);
      } else {
        lines = text.split(/\r?\n/);
      }

      const requiredHeight = lines.length * lineHeight + padding * 2;

      if (mergeInfo) {
        processedMerges.add(mergeInfo.master);
        // For merged cells spanning multiple rows, check total height
        let totalMergeHeight = 0;
        for (let mr = mergeInfo.startRow; mr < mergeInfo.startRow + mergeInfo.rowSpan; mr++) {
          totalMergeHeight += rowHeights[mr] ?? DEFAULT_RH;
        }
        if (requiredHeight > totalMergeHeight) {
          // Add extra to the last row of the merge
          const lastRow = mergeInfo.startRow + mergeInfo.rowSpan - 1;
          rowHeights[lastRow] =
            (rowHeights[lastRow] ?? DEFAULT_RH) + (requiredHeight - totalMergeHeight);
        }
      } else if (requiredHeight > maxRequiredHeight) {
        maxRequiredHeight = requiredHeight;
      }
    }

    if (maxRequiredHeight > (rowHeights[r] ?? DEFAULT_RH)) {
      rowHeights[r] = maxRequiredHeight;
    }
  }
}

// =============================================================================
// Layout Engine
// =============================================================================

export interface LayoutResult {
  pages: PdfPageDef[];
  images: PdfImageData[];
}

export function layoutWorksheet(
  ws: Worksheet,
  media: WorkbookMedia[],
  options: PdfLayoutOptions = {}
): LayoutResult {
  const scale = options.scale ?? 1;
  let pageWidth = options.pageWidth ?? A4_WIDTH;
  let pageHeight = options.pageHeight ?? A4_HEIGHT;

  // Handle orientation
  const orientation = options.orientation ?? ws.pageSetup?.orientation ?? "portrait";
  if (orientation === "landscape") {
    [pageWidth, pageHeight] = [pageHeight, pageWidth];
  }

  const margins = {
    top: (options.margins?.top ?? 40) * scale,
    right: (options.margins?.right ?? 30) * scale,
    bottom: (options.margins?.bottom ?? 40) * scale,
    left: (options.margins?.left ?? 30) * scale
  };

  const printableWidth = pageWidth - margins.left - margins.right;
  const printableHeight = pageHeight - margins.top - margins.bottom;

  // Determine column widths
  const rowCount = ws.rowCount;
  let colCount = ws.columnCount;
  if (colCount === 0) {
    colCount = 1;
  }

  const colWidths: number[] = []; // 1-indexed, in PDF points
  for (let c = 1; c <= colCount; c++) {
    let w: number;
    if (options.columnWidths?.[c] === undefined) {
      const col = ws.getColumn(c);
      w = col.width ?? ws.properties.defaultColWidth ?? DEFAULT_COL_WIDTH;
    } else {
      w = options.columnWidths[c];
    }
    colWidths[c] = w * COL_WIDTH_FACTOR * scale;
  }

  // Determine row heights
  const rowHeights: number[] = []; // 1-indexed, in PDF points
  for (let r = 1; r <= rowCount; r++) {
    const row = ws.findRow(r);
    const h = row?.height ?? ws.properties.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
    rowHeights[r] = h * ROW_HEIGHT_FACTOR * scale;
  }

  // If fitToWidth, scale all column widths
  if (options.fitToWidth) {
    const totalWidth = colWidths.reduce((sum, w, i) => (i > 0 ? sum + w : sum), 0);
    if (totalWidth > printableWidth) {
      const fitScale = printableWidth / totalWidth;
      for (let c = 1; c <= colCount; c++) {
        colWidths[c] *= fitScale;
      }
    }
  }

  // Build merge map
  const mergeMap = buildMergeMap(ws);

  // Auto-expand row heights based on cell text content
  autoExpandRowHeights(ws, rowCount, colCount, rowHeights, colWidths, mergeMap, scale);

  // Collect images
  const pdfImages: PdfImageData[] = [];
  const wsImages = ws.getImages();

  for (const img of wsImages) {
    const imgModel = img.model;
    if (imgModel.type !== "image") {
      continue;
    }
    const mediaItem = media[Number(imgModel.imageId)];
    if (!mediaItem) {
      continue;
    }

    let imageData: Uint8Array | undefined;
    if (mediaItem.buffer) {
      imageData =
        mediaItem.buffer instanceof Uint8Array
          ? mediaItem.buffer
          : new Uint8Array(mediaItem.buffer);
    } else if (mediaItem.base64) {
      imageData = base64ToBytes(mediaItem.base64);
    }

    if (!imageData || imageData.length === 0) {
      continue;
    }

    const ext = mediaItem.extension as "jpeg" | "png" | "gif";
    const format = ext === "png" ? "png" : "jpeg";

    // Detect image dimensions from binary
    const dims = detectImageDimensions(imageData, format);

    pdfImages.push({
      key: imgModel.imageId,
      data: imageData,
      width: dims.width,
      height: dims.height,
      format
    });
  }

  // Paginate rows
  const pages: PdfPageDef[] = [];
  let startRow = 1;

  while (startRow <= rowCount) {
    const ops: PdfOp[] = [];
    let yUsed = 0;
    let endRow = startRow;

    // Determine how many rows fit on this page
    for (let r = startRow; r <= rowCount; r++) {
      const rh = rowHeights[r] ?? DEFAULT_ROW_HEIGHT * scale;
      if (yUsed + rh > printableHeight && r > startRow) {
        break;
      }
      yUsed += rh;
      endRow = r;
    }

    // Generate content for rows startRow..endRow
    const showGrid = options.gridLines === true;
    const pageCtx: PageGenCtx = {
      ws,
      startRow,
      endRow,
      colCount,
      colWidths,
      rowHeights,
      margins,
      pageHeight,
      scale,
      mergeMap,
      showGrid,
      wsImages
    };
    generatePageOps(ops, pageCtx);

    pages.push({ width: pageWidth, height: pageHeight, ops });
    startRow = endRow + 1;
  }

  // If no rows, create at least one blank page
  if (pages.length === 0) {
    pages.push({ width: pageWidth, height: pageHeight, ops: [] });
  }

  return { pages, images: pdfImages };
}

// =============================================================================
// Page content generation
// =============================================================================

function generatePageOps(ops: PdfOp[], ctx: PageGenCtx): void {
  const {
    ws,
    startRow,
    endRow,
    colCount,
    colWidths,
    rowHeights,
    margins,
    pageHeight,
    scale,
    mergeMap,
    showGrid,
    wsImages
  } = ctx;
  const DEFAULT_RH = DEFAULT_ROW_HEIGHT * scale;

  // Precompute column x positions (left edge of each column)
  const colX: number[] = [0]; // 1-indexed
  colX[1] = margins.left;
  for (let c = 2; c <= colCount + 1; c++) {
    colX[c] = colX[c - 1] + (colWidths[c - 1] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale);
  }

  // Precompute row y positions (top edge of each row) – PDF y is bottom-up
  // We compute from top of printable area downward
  let yOffset = margins.top;
  const rowYTop: number[] = []; // 1-indexed, from page top
  for (let r = startRow; r <= endRow; r++) {
    rowYTop[r] = yOffset;
    yOffset += rowHeights[r] ?? DEFAULT_RH;
  }

  // Two-pass rendering: fills/grids/borders first, then text.
  // This prevents adjacent cell fills from covering overflowed text.
  const textOps: PdfOp[] = [];

  // Track which merge cells we've already drawn
  const drawnMerges = new Set<string>();

  // Helper to compute merge dimensions and detect slave cells
  function resolveMerge(
    cellAddr: string,
    cw: number,
    rh: number,
    mergeInfo: MergeInfo | undefined
  ): { effectiveCw: number; effectiveRh: number; isMergedSlave: boolean } {
    let effectiveCw = cw;
    let effectiveRh = rh;
    let isMergedSlave = false;

    if (mergeInfo) {
      if (cellAddr !== mergeInfo.master) {
        isMergedSlave = true;
      } else if (drawnMerges.has(mergeInfo.master)) {
        isMergedSlave = true;
      } else {
        drawnMerges.add(mergeInfo.master);
        effectiveCw = 0;
        for (let mc = mergeInfo.startCol; mc < mergeInfo.startCol + mergeInfo.colSpan; mc++) {
          effectiveCw += colWidths[mc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
        }
        effectiveRh = 0;
        const mergeEndRow = Math.min(mergeInfo.startRow + mergeInfo.rowSpan - 1, endRow);
        for (let mr = mergeInfo.startRow; mr <= mergeEndRow; mr++) {
          effectiveRh += rowHeights[mr] ?? DEFAULT_RH;
        }
      }
    }

    return { effectiveCw, effectiveRh, isMergedSlave };
  }

  // ---- Pass 1: Fills and grid lines ----
  // ---- (borders collected separately to draw after all fills) ----
  const borderOps: PdfOp[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.findRow(r);
    const rh = rowHeights[r] ?? DEFAULT_RH;
    const pdfRowTop = pageHeight - rowYTop[r];

    for (let c = 1; c <= colCount; c++) {
      const cellAddr = colToLetter(c) + r;
      const cw = colWidths[c] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
      const pdfCellLeft = colX[c];
      const mergeInfo = mergeMap.get(cellAddr);
      const { effectiveCw, effectiveRh, isMergedSlave } = resolveMerge(cellAddr, cw, rh, mergeInfo);

      if (isMergedSlave) {
        continue;
      }

      const cell = row?.findCell(c);

      // White fill for merged cells to cover any internal gridlines
      if (cellAddr === mergeInfo?.master) {
        ops.push({
          kind: "rect",
          x: pdfCellLeft,
          y: pdfRowTop - effectiveRh,
          width: effectiveCw,
          height: effectiveRh,
          fillColor: WHITE
        });
      }

      // Background fill
      const fillColor = cell ? getFillColor(cell.fill) : undefined;
      if (fillColor) {
        ops.push({
          kind: "rect",
          x: pdfCellLeft,
          y: pdfRowTop - effectiveRh,
          width: effectiveCw,
          height: effectiveRh,
          fillColor
        });
      }

      // Grid lines
      if (showGrid) {
        ops.push({
          kind: "rect",
          x: pdfCellLeft,
          y: pdfRowTop - effectiveRh,
          width: effectiveCw,
          height: effectiveRh,
          strokeColor: GRID_COLOR,
          lineWidth: 0.25
        });
      }

      // Collect cell borders (drawn after all fills so no fill overwrites them)
      if (cell?.border) {
        drawBorders(borderOps, cell.border, pdfCellLeft, pdfRowTop, effectiveCw, effectiveRh);
      }
    }
  }

  // Append borders after all fills so they are not covered
  for (const op of borderOps) {
    ops.push(op);
  }

  // ---- Pass 2: Text (with overflow) ----
  // Reset merge tracking for the second pass
  drawnMerges.clear();

  for (let r = startRow; r <= endRow; r++) {
    const row = ws.findRow(r);
    const rh = rowHeights[r] ?? DEFAULT_RH;
    const pdfRowTop = pageHeight - rowYTop[r];

    for (let c = 1; c <= colCount; c++) {
      const cellAddr = colToLetter(c) + r;
      const cw = colWidths[c] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
      const pdfCellLeft = colX[c];
      const mergeInfo = mergeMap.get(cellAddr);
      const { effectiveCw, effectiveRh, isMergedSlave } = resolveMerge(cellAddr, cw, rh, mergeInfo);

      if (isMergedSlave) {
        continue;
      }

      const cell = row?.findCell(c);

      // Cell text
      if (cell && cell.type !== ValueType.Null && cell.type !== ValueType.Merge) {
        const text = getCellText(cell);
        if (text) {
          // Compute text overflow width for non-merged, non-wrapping cells
          let renderWidth = effectiveCw;
          let renderLeft = pdfCellLeft;
          if (!mergeInfo && !cell.alignment?.wrapText) {
            const font = cell.font;
            const family = mapFontFamily(font?.name);
            const bold = !!font?.bold;
            const italic = !!font?.italic;
            const fontSize = (font?.size ?? DEFAULT_FONT_SIZE) * scale;
            const padding = CELL_PADDING * scale;
            const textWidth = measureTextWidth(text, family, bold, italic, fontSize);

            if (textWidth > effectiveCw - padding * 2) {
              const hAlign = cell.alignment?.horizontal ?? "left";
              if (hAlign === "right") {
                let overflowWidth = effectiveCw;
                for (let oc = c - 1; oc >= 1; oc--) {
                  const adjCell = row?.findCell(oc);
                  const adjAddr = colToLetter(oc) + r;
                  if (mergeMap.get(adjAddr)) {
                    break;
                  }
                  if (
                    adjCell &&
                    adjCell.type !== ValueType.Null &&
                    adjCell.type !== ValueType.Merge
                  ) {
                    break;
                  }
                  overflowWidth += colWidths[oc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
                  if (overflowWidth >= textWidth + padding * 2) {
                    break;
                  }
                }
                renderLeft = pdfCellLeft - (overflowWidth - effectiveCw);
                renderWidth = overflowWidth;
              } else if (hAlign === "center" || hAlign === "centerContinuous") {
                let leftExtra = 0;
                for (let oc = c - 1; oc >= 1; oc--) {
                  const adjCell = row?.findCell(oc);
                  const adjAddr = colToLetter(oc) + r;
                  if (mergeMap.get(adjAddr)) {
                    break;
                  }
                  if (
                    adjCell &&
                    adjCell.type !== ValueType.Null &&
                    adjCell.type !== ValueType.Merge
                  ) {
                    break;
                  }
                  leftExtra += colWidths[oc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
                }
                let rightExtra = 0;
                for (let oc = c + 1; oc <= colCount; oc++) {
                  const adjCell = row?.findCell(oc);
                  const adjAddr = colToLetter(oc) + r;
                  if (mergeMap.get(adjAddr)) {
                    break;
                  }
                  if (
                    adjCell &&
                    adjCell.type !== ValueType.Null &&
                    adjCell.type !== ValueType.Merge
                  ) {
                    break;
                  }
                  rightExtra += colWidths[oc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
                }
                renderWidth = leftExtra + effectiveCw + rightExtra;
                renderLeft = pdfCellLeft - leftExtra;
              } else {
                // Left-aligned: overflow rightward
                let overflowWidth = effectiveCw;
                for (let oc = c + 1; oc <= colCount; oc++) {
                  const adjCell = row?.findCell(oc);
                  const adjAddr = colToLetter(oc) + r;
                  if (mergeMap.get(adjAddr)) {
                    break;
                  }
                  if (
                    adjCell &&
                    adjCell.type !== ValueType.Null &&
                    adjCell.type !== ValueType.Merge
                  ) {
                    break;
                  }
                  overflowWidth += colWidths[oc] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
                  if (overflowWidth >= textWidth + padding * 2) {
                    break;
                  }
                }
                renderWidth = overflowWidth;
              }
            }
          }
          drawCellText(
            textOps,
            text,
            cell,
            { left: renderLeft, top: pdfRowTop, width: renderWidth, height: effectiveRh },
            scale
          );
        }
      }
    }
  }

  // Append text ops after all fills/borders so text renders on top
  for (const op of textOps) {
    ops.push(op);
  }

  // Draw images that are positioned within this page's row range
  for (const img of wsImages) {
    const imgModel = img.model;
    if (imgModel.type !== "image" || !imgModel.range) {
      continue;
    }

    const range = imgModel.range;
    const tl = range.tl;
    if (!tl) {
      continue;
    }

    const imgRow = (tl.nativeRow ?? tl.row ?? 0) + 1;
    const imgCol = (tl.nativeCol ?? tl.col ?? 0) + 1;

    // Check if image falls within this page's rows
    if (imgRow < startRow || imgRow > endRow) {
      continue;
    }

    const emuToPt = (emu: number) => (emu / 914400) * 72;

    // Base position: top-left corner of the anchor cell
    let imgX = colX[Math.min(imgCol, colCount)] ?? margins.left;
    let imgYTopPos = rowYTop[Math.min(imgRow, endRow)] ?? margins.top;

    // Add sub-cell offset from tl anchor (nativeColOff/nativeRowOff are in EMU)
    if (tl.nativeColOff) {
      imgX += emuToPt(tl.nativeColOff) * scale;
    }
    if (tl.nativeRowOff) {
      imgYTopPos += emuToPt(tl.nativeRowOff) * scale;
    }

    let imgW: number;
    let imgH: number;

    if (range.br) {
      // Two-cell anchor: compute size from tl and br cell grid positions
      const brCol = (range.br.nativeCol ?? 0) + 1;
      const brRow = (range.br.nativeRow ?? 0) + 1;

      let brX = colX[Math.min(brCol, colCount)] ?? margins.left;
      if (range.br.nativeColOff) {
        brX += emuToPt(range.br.nativeColOff) * scale;
      }

      let brYTopPos: number;
      if (brRow >= startRow && brRow <= endRow) {
        brYTopPos = rowYTop[brRow] ?? margins.top;
      } else {
        // br is beyond current page, extrapolate from last row
        brYTopPos = rowYTop[endRow] + (rowHeights[endRow] ?? DEFAULT_RH);
        for (let rr = endRow + 1; rr < brRow; rr++) {
          brYTopPos += rowHeights[rr] ?? DEFAULT_RH;
        }
      }
      if (range.br.nativeRowOff) {
        brYTopPos += emuToPt(range.br.nativeRowOff) * scale;
      }

      imgW = Math.max(1, brX - imgX);
      imgH = Math.max(1, brYTopPos - imgYTopPos);
    } else if (range.ext && (range.ext.width || range.ext.height)) {
      // One-cell anchor: ext.width/height are in pixels (converted from EMU by ExtXform)
      imgW = (((range.ext.width ?? 100) * 72) / 96) * scale;
      imgH = (((range.ext.height ?? 75) * 72) / 96) * scale;
    } else {
      // Fallback: fit to the anchor cell dimensions
      imgW = colWidths[imgCol] ?? DEFAULT_COL_WIDTH * COL_WIDTH_FACTOR * scale;
      imgH = rowHeights[imgRow] ?? DEFAULT_RH;
    }

    const pdfImgY = pageHeight - imgYTopPos;

    ops.push({
      kind: "image",
      imageKey: imgModel.imageId,
      x: imgX,
      y: pdfImgY - imgH,
      width: imgW,
      height: imgH
    });
  }
}

// =============================================================================
// Border drawing
// =============================================================================

function drawBorders(
  ops: PdfOp[],
  borders: Partial<Borders>,
  x: number,
  yTop: number,
  w: number,
  h: number
): void {
  const DOUBLE_GAP = 1.5; // gap between the two lines in a double border

  // Top border
  if (borders.top?.style) {
    const color = excelColorToRgb(borders.top.color) ?? BLACK;
    const lw = borderLineWidth(borders.top.style);
    const dash = borderDash(borders.top.style);
    if (borders.top.style === "double") {
      ops.push(
        {
          kind: "line",
          x1: x,
          y1: yTop + DOUBLE_GAP / 2,
          x2: x + w,
          y2: yTop + DOUBLE_GAP / 2,
          color,
          lineWidth: lw,
          dash
        },
        {
          kind: "line",
          x1: x,
          y1: yTop - DOUBLE_GAP / 2,
          x2: x + w,
          y2: yTop - DOUBLE_GAP / 2,
          color,
          lineWidth: lw,
          dash
        }
      );
    } else {
      ops.push({ kind: "line", x1: x, y1: yTop, x2: x + w, y2: yTop, color, lineWidth: lw, dash });
    }
  }

  // Bottom border
  if (borders.bottom?.style) {
    const color = excelColorToRgb(borders.bottom.color) ?? BLACK;
    const lw = borderLineWidth(borders.bottom.style);
    const dash = borderDash(borders.bottom.style);
    if (borders.bottom.style === "double") {
      ops.push(
        {
          kind: "line",
          x1: x,
          y1: yTop - h + DOUBLE_GAP / 2,
          x2: x + w,
          y2: yTop - h + DOUBLE_GAP / 2,
          color,
          lineWidth: lw,
          dash
        },
        {
          kind: "line",
          x1: x,
          y1: yTop - h - DOUBLE_GAP / 2,
          x2: x + w,
          y2: yTop - h - DOUBLE_GAP / 2,
          color,
          lineWidth: lw,
          dash
        }
      );
    } else {
      ops.push({
        kind: "line",
        x1: x,
        y1: yTop - h,
        x2: x + w,
        y2: yTop - h,
        color,
        lineWidth: lw,
        dash
      });
    }
  }

  // Left border
  if (borders.left?.style) {
    const color = excelColorToRgb(borders.left.color) ?? BLACK;
    const lw = borderLineWidth(borders.left.style);
    const dash = borderDash(borders.left.style);
    if (borders.left.style === "double") {
      ops.push(
        {
          kind: "line",
          x1: x - DOUBLE_GAP / 2,
          y1: yTop,
          x2: x - DOUBLE_GAP / 2,
          y2: yTop - h,
          color,
          lineWidth: lw,
          dash
        },
        {
          kind: "line",
          x1: x + DOUBLE_GAP / 2,
          y1: yTop,
          x2: x + DOUBLE_GAP / 2,
          y2: yTop - h,
          color,
          lineWidth: lw,
          dash
        }
      );
    } else {
      ops.push({ kind: "line", x1: x, y1: yTop, x2: x, y2: yTop - h, color, lineWidth: lw, dash });
    }
  }

  // Right border
  if (borders.right?.style) {
    const color = excelColorToRgb(borders.right.color) ?? BLACK;
    const lw = borderLineWidth(borders.right.style);
    const dash = borderDash(borders.right.style);
    if (borders.right.style === "double") {
      ops.push(
        {
          kind: "line",
          x1: x + w - DOUBLE_GAP / 2,
          y1: yTop,
          x2: x + w - DOUBLE_GAP / 2,
          y2: yTop - h,
          color,
          lineWidth: lw,
          dash
        },
        {
          kind: "line",
          x1: x + w + DOUBLE_GAP / 2,
          y1: yTop,
          x2: x + w + DOUBLE_GAP / 2,
          y2: yTop - h,
          color,
          lineWidth: lw,
          dash
        }
      );
    } else {
      ops.push({
        kind: "line",
        x1: x + w,
        y1: yTop,
        x2: x + w,
        y2: yTop - h,
        color,
        lineWidth: lw,
        dash
      });
    }
  }
}

// =============================================================================
// Text drawing
// =============================================================================

function drawCellText(ops: PdfOp[], text: string, cell: Cell, rect: CellRect, scale: number): void {
  const { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight } = rect;
  const font = cell.font;
  const alignment = cell.alignment;

  // Font properties
  const family = mapFontFamily(font?.name);
  const bold = !!font?.bold;
  const italic = !!font?.italic;
  const fontSize = (font?.size ?? DEFAULT_FONT_SIZE) * scale;
  const pdfFontName = getPdfFontName(family, bold, italic);

  // Text color
  const textColor = excelColorToRgb(font?.color) ?? BLACK;

  // Horizontal alignment
  const hAlign = alignment?.horizontal ?? "left";
  // Vertical alignment
  const vAlign = alignment?.vertical ?? "bottom";
  // Wrap text
  const shouldWrap = alignment?.wrapText ?? false;
  // Text rotation
  const textRotation = alignment?.textRotation;
  // Indent level
  const indent = alignment?.indent ?? 0;

  const padding = CELL_PADDING * scale;
  const indentPx = indent * INDENT_WIDTH * scale;

  // Handle stacked vertical text (textRotation === "vertical" or 255)
  if (textRotation === "vertical") {
    drawVerticalStackedText(
      ops,
      text,
      { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight },
      { pdfFontName, family, bold, italic, fontSize, textColor, hAlign, vAlign, padding, font }
    );
    return;
  }

  // Handle rotated text (numeric rotation)
  const rotation = typeof textRotation === "number" && textRotation !== 0 ? textRotation : 0;

  if (rotation !== 0) {
    drawRotatedText(
      ops,
      text,
      { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight },
      { pdfFontName, family, bold, italic, fontSize, textColor, hAlign, vAlign, padding, font },
      rotation
    );
    return;
  }

  const availableWidth = cellWidth - padding * 2 - indentPx;
  const lineHeight = fontSize * 1.2;

  // Get text lines (split on explicit newlines even without wrapText)
  let lines: string[];
  if (shouldWrap) {
    lines = wrapText(text, family, bold, italic, fontSize, availableWidth);
  } else {
    lines = text.split(/\r?\n/);
  }

  // Clip lines to fit within cell height to prevent overflow into adjacent rows
  const maxVisibleLines = Math.max(1, Math.floor((cellHeight - padding * 2) / lineHeight));
  if (lines.length > maxVisibleLines) {
    lines = lines.slice(0, maxVisibleLines);
  }

  const totalTextHeight = lines.length * lineHeight;

  // Compute vertical starting position (PDF y from bottom)
  let yStart: number;
  switch (vAlign) {
    case "top":
      yStart = cellTop - padding - fontSize;
      break;
    case "middle":
      yStart = cellTop - (cellHeight - totalTextHeight) / 2 - fontSize;
      break;
    case "bottom":
    default:
      yStart = cellTop - cellHeight + padding + (lines.length - 1) * lineHeight;
      break;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) {
      continue;
    }
    const lineWidth = measureTextWidth(line, family, bold, italic, fontSize);

    // Compute horizontal position (accounting for indent)
    let xPos: number;
    switch (hAlign) {
      case "center":
      case "centerContinuous":
        xPos = cellLeft + (cellWidth - lineWidth) / 2;
        break;
      case "right":
        xPos = cellLeft + cellWidth - padding - indentPx - lineWidth;
        break;
      case "left":
      case "fill":
      case "justify":
      case "distributed":
      default:
        xPos = cellLeft + padding + indentPx;
        break;
    }

    const yPos = yStart - i * lineHeight;

    ops.push({
      kind: "text",
      x: xPos,
      y: yPos,
      text: line,
      fontRef: pdfFontName,
      fontSize,
      color: textColor,
      underline: !!(font?.underline && font.underline !== "none"),
      strike: !!font?.strike
    });

    // Draw underline precisely
    if (font?.underline && font.underline !== "none") {
      const ulY = yPos - fontSize * 0.15;
      ops.push({
        kind: "line",
        x1: xPos,
        y1: ulY,
        x2: xPos + lineWidth,
        y2: ulY,
        color: textColor,
        lineWidth: fontSize * 0.05
      });
    }

    // Draw strikethrough precisely
    if (font?.strike) {
      const stY = yPos + fontSize * 0.25;
      ops.push({
        kind: "line",
        x1: xPos,
        y1: stY,
        x2: xPos + lineWidth,
        y2: stY,
        color: textColor,
        lineWidth: fontSize * 0.05
      });
    }
  }

  // Handle rich text with mixed formatting
  const val = cell.value;
  if (val && typeof val === "object" && "richText" in val && Array.isArray(val.richText)) {
    drawRichText(
      ops,
      val.richText,
      { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight },
      scale,
      alignment
    );
  }
}

// =============================================================================
// Vertical stacked text rendering
// =============================================================================

function drawVerticalStackedText(
  ops: PdfOp[],
  text: string,
  rect: CellRect,
  ctx: TextDrawCtx
): void {
  const { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight } = rect;
  const { pdfFontName, fontSize, textColor, hAlign, vAlign, padding } = ctx;
  // Each character is drawn on a separate line, stacked vertically (not rotated)
  const chars = [...text];
  const charHeight = fontSize * 1.3;
  const totalHeight = chars.length * charHeight;

  // Vertical positioning
  let yStart: number;
  switch (vAlign) {
    case "top":
      yStart = cellTop - padding - fontSize;
      break;
    case "middle":
      yStart = cellTop - (cellHeight - totalHeight) / 2 - fontSize;
      break;
    case "bottom":
    default:
      yStart = cellTop - cellHeight + padding + (chars.length - 1) * charHeight;
      break;
  }

  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (!ch || ch === "\n" || ch === "\r") {
      continue;
    }
    // Center each character horizontally in the cell
    const charWidth = measureTextWidth(ch, "helvetica", false, false, fontSize);
    let xPos: number;
    switch (hAlign) {
      case "center":
      case "centerContinuous":
        xPos = cellLeft + (cellWidth - charWidth) / 2;
        break;
      case "right":
        xPos = cellLeft + cellWidth - padding - charWidth;
        break;
      default:
        xPos = cellLeft + (cellWidth - charWidth) / 2;
        break;
    }

    const yPos = yStart - i * charHeight;

    ops.push({
      kind: "text",
      x: xPos,
      y: yPos,
      text: ch,
      fontRef: pdfFontName,
      fontSize,
      color: textColor,
      underline: false,
      strike: false
    });
  }
}

// =============================================================================
// Rotated text rendering
// =============================================================================

function drawRotatedText(
  ops: PdfOp[],
  text: string,
  rect: CellRect,
  ctx: TextDrawCtx,
  rotation: number
): void {
  const { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight } = rect;
  const { pdfFontName, family, bold, italic, fontSize, textColor, hAlign, vAlign, padding, font } =
    ctx;
  let lines = text.split(/\r?\n/);
  const lineHeight = fontSize * 1.2;
  const rad = (rotation * Math.PI) / 180;
  const absSin = Math.abs(Math.sin(rad));
  const absCos = Math.abs(Math.cos(rad));

  // For 90° or -90° rotation, clip text to fit within cell dimensions.
  // Each "line" occupies lineHeight in the horizontal direction.
  // Text length runs along the cell height direction.
  if (rotation === 90 || rotation === -90) {
    const maxLines = Math.max(1, Math.floor((cellWidth - padding * 2) / lineHeight));
    if (lines.length > maxLines) {
      lines = lines.slice(0, maxLines);
    }
    // Clip long text lines to fit within available cell height
    const availableTextLength = cellHeight - padding * 2;
    for (let i = 0; i < lines.length; i++) {
      let line = lines[i];
      let tw = measureTextWidth(line, family, bold, italic, fontSize);
      if (tw > availableTextLength && line.length > 1) {
        // Truncate to fit
        while (line.length > 1 && tw > availableTextLength) {
          line = line.slice(0, -1);
          tw = measureTextWidth(line, family, bold, italic, fontSize);
        }
        lines[i] = line;
      }
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line) {
      continue;
    }

    const textWidth = measureTextWidth(line, family, bold, italic, fontSize);

    // For rotated text, compute the position within the cell
    // The text is rotated around its starting point
    let xPos: number;
    let yPos: number;

    // Horizontal center of cell is the anchor for rotated text
    const cellCenterX = cellLeft + cellWidth / 2;
    const cellCenterY = cellTop - cellHeight / 2;

    if (rotation === 90 || rotation === -90) {
      // 90°: text reads bottom-to-top; -90°: text reads top-to-bottom
      const rotatedWidth = fontSize;

      switch (hAlign) {
        case "center":
        case "centerContinuous":
          xPos = cellCenterX - rotatedWidth / 2 + i * lineHeight;
          break;
        case "right":
          xPos = cellLeft + cellWidth - padding - rotatedWidth;
          break;
        default:
          xPos = cellLeft + padding;
          break;
      }

      if (rotation === 90) {
        // Bottom-to-top: start from bottom
        switch (vAlign) {
          case "top":
            yPos = cellTop - padding - textWidth;
            break;
          case "middle":
            yPos = cellCenterY - textWidth / 2;
            break;
          default:
            yPos = cellTop - cellHeight + padding;
            break;
        }
        xPos += fontSize;
      } else {
        // Top-to-bottom: start from top
        switch (vAlign) {
          case "top":
            yPos = cellTop - padding;
            break;
          case "middle":
            yPos = cellCenterY + textWidth / 2;
            break;
          default:
            yPos = cellTop - cellHeight + padding + textWidth;
            break;
        }
      }
    } else {
      // Arbitrary angle rotation
      xPos = cellCenterX - (textWidth * absCos) / 2;
      yPos = cellCenterY - (textWidth * absSin) / 2 + i * lineHeight;
    }

    ops.push({
      kind: "text",
      x: xPos,
      y: yPos,
      text: line,
      fontRef: pdfFontName,
      fontSize,
      color: textColor,
      underline: !!(font?.underline && font.underline !== "none"),
      strike: !!font?.strike,
      rotation
    });
  }
}

// =============================================================================
// Rich text rendering
// =============================================================================

function drawRichText(
  _ops: PdfOp[],
  _richText: Array<{ text: string; font?: Partial<Font> }>,
  _rect: CellRect,
  _scale: number,
  _alignment?: Partial<Alignment>
): void {
  // For rich text, we draw each fragment inline.
  // The base drawCellText already drew the plain concatenated text.
  // We remove the last ops that were the plain text and redraw with rich formatting.
  // However, since we don't know how many ops were added, we'll just add overlapping
  // rich text. To avoid double rendering, the caller should check for rich text first.
  // Since we're called after the plain text is drawn, we skip this to avoid complexity.
  // The plain text fallback already handles rich text by concatenation.
  // This is a placeholder for future enhancement with proper inline rich text layout.
}

// =============================================================================
// Image dimension detection
// =============================================================================

function detectImageDimensions(
  data: Uint8Array,
  format: "jpeg" | "png"
): { width: number; height: number } {
  if (format === "png") {
    return detectPngDimensions(data);
  }
  return detectJpegDimensions(data);
}

function detectPngDimensions(data: Uint8Array): { width: number; height: number } {
  // PNG: width at offset 16, height at offset 20 (IHDR chunk)
  if (data.length < 24) {
    return { width: 100, height: 100 };
  }
  const width = ((data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19]) >>> 0;
  const height = ((data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23]) >>> 0;
  return { width: width || 100, height: height || 100 };
}

function detectJpegDimensions(data: Uint8Array): { width: number; height: number } {
  // JPEG: scan for SOF0 or SOF2 marker (0xFF 0xC0 or 0xFF 0xC2)
  let pos = 2; // Skip SOI marker
  while (pos < data.length - 8) {
    if (data[pos] !== 0xff) {
      pos++;
      continue;
    }
    const marker = data[pos + 1];
    if (marker === 0xc0 || marker === 0xc2) {
      const height = (data[pos + 5] << 8) | data[pos + 6];
      const width = (data[pos + 7] << 8) | data[pos + 8];
      return { width: width || 100, height: height || 100 };
    }
    // Skip to next marker
    const length = (data[pos + 2] << 8) | data[pos + 3];
    pos += 2 + length;
  }
  return { width: 100, height: 100 };
}

// =============================================================================
// Base64 decoding (minimal, no dependencies)
// =============================================================================

function base64ToBytes(base64: string): Uint8Array {
  const binaryStr = atob(base64);
  const bytes = new Uint8Array(binaryStr.length);
  for (let i = 0; i < binaryStr.length; i++) {
    bytes[i] = binaryStr.codePointAt(i) ?? 0;
  }
  return bytes;
}
