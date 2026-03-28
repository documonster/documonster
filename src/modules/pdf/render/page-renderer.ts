/**
 * Page renderer for PDF generation.
 *
 * Takes LayoutPage objects (produced by the layout engine) and renders them
 * as PDF content streams. Handles:
 * - Cell background fills
 * - Cell borders (with proper overlap handling)
 * - Text rendering with alignment, wrapping, and clipping
 * - Grid lines
 * - Page headers (sheet names) and footers (page numbers)
 */

import { PdfContentStream } from "../core/pdf-stream";
import type { FontManager } from "../font/font-manager";
import { resolvePdfFontName } from "../font/font-manager";
import type {
  LayoutPage,
  LayoutCell,
  LayoutBorder,
  LayoutRichTextRun,
  ResolvedPdfOptions,
  PdfRect
} from "../types";

// =============================================================================
// Constants
// =============================================================================

/** Internal cell padding in points */
const CELL_PADDING_H = 3;
const CELL_PADDING_V = 2;

// =============================================================================
// Page Renderer
// =============================================================================

/**
 * Result of rendering a page.
 */
export interface PageRenderResult {
  stream: PdfContentStream;
  /** Set of unique alpha values (0-1) used on this page. Empty if all opaque. */
  alphaValues: Set<number>;
}

/**
 * Render a single page to a PDF content stream.
 */
export function renderPage(
  page: LayoutPage,
  options: ResolvedPdfOptions,
  fontManager: FontManager,
  totalPages: number
): PageRenderResult {
  const stream = new PdfContentStream();
  const alphaValues = new Set<number>();

  // --- Step 1: Draw grid lines (behind everything) ---
  if (options.showGridLines) {
    drawGridLines(stream, page, options);
  }

  // --- Step 2: Draw cell backgrounds ---
  for (const cell of page.cells) {
    if (cell.fillColor) {
      drawCellFill(stream, cell, alphaValues);
    }
  }

  // --- Step 3: Draw cell borders ---
  for (const cell of page.cells) {
    drawCellBorders(stream, cell);
  }

  // --- Step 4: Draw cell text ---
  for (const cell of page.cells) {
    if (cell.text) {
      drawCellText(stream, cell, fontManager, alphaValues);
    }
  }

  // --- Step 5: Draw page header (sheet name) ---
  if (options.showSheetNames) {
    drawPageHeader(stream, page, options, fontManager);
  }

  // --- Step 6: Draw page footer (page number) ---
  if (options.showPageNumbers) {
    drawPageFooter(stream, page, options, fontManager, totalPages);
  }

  return { stream, alphaValues };
}

// =============================================================================
// Grid Lines
// =============================================================================

function drawGridLines(
  stream: PdfContentStream,
  page: LayoutPage,
  options: ResolvedPdfOptions
): void {
  if (page.columnWidths.length === 0 || page.rowHeights.length === 0) {
    return;
  }

  const color = options.gridLineColor;
  const lineWidth = 0.25;

  stream.save();
  stream.setStrokeColor(color);
  stream.setLineWidth(lineWidth);

  // Vertical grid lines
  const topY = page.rowYPositions[0];
  const lastRowIdx = page.rowYPositions.length - 1;
  const bottomY = page.rowYPositions[lastRowIdx] - page.rowHeights[lastRowIdx];

  for (let i = 0; i <= page.columnWidths.length; i++) {
    const x =
      i < page.columnWidths.length
        ? page.columnOffsets[i]
        : page.columnOffsets[i - 1] + page.columnWidths[i - 1];
    stream.moveTo(x, topY);
    stream.lineTo(x, bottomY);
  }

  // Horizontal grid lines
  const leftX = page.columnOffsets[0];
  const lastColIdx = page.columnOffsets.length - 1;
  const rightX = page.columnOffsets[lastColIdx] + page.columnWidths[lastColIdx];

  for (let i = 0; i <= page.rowYPositions.length; i++) {
    const y =
      i < page.rowYPositions.length
        ? page.rowYPositions[i]
        : page.rowYPositions[i - 1] - page.rowHeights[i - 1];
    stream.moveTo(leftX, y);
    stream.lineTo(rightX, y);
  }

  stream.stroke();
  stream.restore();
}

// =============================================================================
// Cell Fill
// =============================================================================

function drawCellFill(stream: PdfContentStream, cell: LayoutCell, alphaValues: Set<number>): void {
  if (!cell.fillColor) {
    return;
  }
  const alpha = cell.fillColor.a;
  if (alpha !== undefined && alpha < 1) {
    // Use ExtGState for transparency
    const gsName = alphaGsName(alpha);
    alphaValues.add(alpha);
    stream.save();
    stream.setGraphicsState(gsName);
    stream.fillRect(cell.rect.x, cell.rect.y, cell.rect.width, cell.rect.height, cell.fillColor);
    stream.restore();
  } else {
    stream.fillRect(cell.rect.x, cell.rect.y, cell.rect.width, cell.rect.height, cell.fillColor);
  }
}

// =============================================================================
// Cell Borders
// =============================================================================

function drawCellBorders(stream: PdfContentStream, cell: LayoutCell): void {
  const { rect, borders } = cell;
  const { x, y, width, height } = rect;

  if (borders.top) {
    drawBorderLine(stream, borders.top, x, y + height, x + width, y + height);
  }
  if (borders.bottom) {
    drawBorderLine(stream, borders.bottom, x, y, x + width, y);
  }
  if (borders.left) {
    drawBorderLine(stream, borders.left, x, y, x, y + height);
  }
  if (borders.right) {
    drawBorderLine(stream, borders.right, x + width, y, x + width, y + height);
  }
}

function drawBorderLine(
  stream: PdfContentStream,
  border: LayoutBorder,
  x1: number,
  y1: number,
  x2: number,
  y2: number
): void {
  stream.drawLine(x1, y1, x2, y2, border.color, border.width, border.dashPattern);
}

// =============================================================================
// Cell Text
// =============================================================================

function drawCellText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  alphaValues: Set<number>
): void {
  const { rect, text, fontSize, horizontalAlign, verticalAlign, wrapText } = cell;

  if (!text && !cell.richText) {
    return;
  }

  const availWidth = rect.width - CELL_PADDING_H * 2;
  const availHeight = rect.height - CELL_PADDING_V * 2;
  if (availWidth <= 0 || availHeight <= 0) {
    return;
  }

  const indentPts = cell.indent * INDENT_WIDTH;

  // Clip to cell bounds
  stream.save();
  stream.rect(rect.x, rect.y, rect.width, rect.height);
  stream.clip();
  stream.endPath();

  // Apply text color alpha if needed
  const textAlpha = cell.textColor.a;
  if (textAlpha !== undefined && textAlpha < 1) {
    alphaValues.add(textAlpha);
    stream.setGraphicsState(alphaGsName(textAlpha));
  }

  // Handle text rotation
  if (cell.textRotation === "vertical") {
    drawVerticalStackedText(stream, cell, fontManager, indentPts);
    stream.restore();
    return;
  }
  if (typeof cell.textRotation === "number" && cell.textRotation !== 0) {
    drawRotatedText(stream, cell, fontManager, indentPts);
    stream.restore();
    return;
  }

  // Handle rich text runs
  if (cell.richText && cell.richText.length > 0) {
    drawRichText(stream, cell, fontManager, indentPts);
    stream.restore();
    return;
  }

  // --- Plain text rendering ---
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  const measure = (s: string) => fontManager.measureText(s, resourceName, fontSize);
  // Leave a small buffer (1pt) for wrap width to account for font metrics rounding
  const effectiveWidth = availWidth - indentPts - 1;
  const lines = wrapText ? wrapTextLines(text, measure, effectiveWidth) : [text];

  const lineHeight = fontSize * 1.2;
  const ascent = fontManager.getFontAscent(resourceName, fontSize);
  const totalTextHeight = lines.length * lineHeight;
  const textStartY = computeTextStartY(verticalAlign, rect, totalTextHeight, ascent);

  stream.setFillColor(cell.textColor);
  stream.beginText();
  stream.setFont(resourceName, fontSize);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineY = textStartY - i * lineHeight;
    const textWidth = measure(line);
    const textX = computeTextX(horizontalAlign, rect, textWidth, indentPts);

    stream.setTextMatrix(1, 0, 0, 1, textX, lineY);
    const hexEncoded = fontManager.encodeText(line, resourceName);
    if (hexEncoded) {
      stream.showTextHex(hexEncoded);
    } else {
      stream.showText(line);
    }
  }

  stream.endText();

  drawTextDecorations(
    stream,
    cell,
    lines,
    lineHeight,
    textStartY,
    measure,
    resourceName,
    fontManager,
    indentPts
  );
  stream.restore();
}

// =============================================================================
// Rich Text Rendering
// =============================================================================

function drawRichText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  indentPts: number
): void {
  const { rect, horizontalAlign, verticalAlign, wrapText } = cell;
  const runs = cell.richText!;

  // Use the largest font size across all runs for line height calculation
  let maxFontSize = cell.fontSize;
  for (const run of runs) {
    if (run.fontSize > maxFontSize) {
      maxFontSize = run.fontSize;
    }
  }
  const primaryFontSize = maxFontSize;
  const lineHeight = primaryFontSize * 1.2;

  const isEmbedded = fontManager.hasEmbeddedFont();

  // Helper: resolve resource name for a run
  const runResource = (run: LayoutRichTextRun) =>
    isEmbedded
      ? fontManager.getEmbeddedResourceName()
      : fontManager.ensureFont(resolvePdfFontName(run.fontFamily, run.bold, run.italic));

  // --- Wrapping path ---
  if (wrapText) {
    const availWidth = rect.width - CELL_PADDING_H * 2 - indentPts - 1;
    if (availWidth <= 0) {
      return;
    }

    // Concatenate full text and wrap it
    const fullText = runs.map(r => r.text).join("");
    const primaryResource = runResource(runs[0]);
    const measure = (s: string) => fontManager.measureText(s, primaryResource, primaryFontSize);
    const lines = wrapTextLines(fullText, measure, availWidth);

    // Build a character-to-run mapping so we know which run each char belongs to
    const runForChar: number[] = [];
    for (let ri = 0; ri < runs.length; ri++) {
      for (let ci = 0; ci < runs[ri].text.length; ci++) {
        runForChar.push(ri);
      }
    }

    const primaryResourceName = runResource(runs[0]);
    const ascent = fontManager.getFontAscent(primaryResourceName, primaryFontSize);
    const totalTextHeight = lines.length * lineHeight;
    const textStartY = computeTextStartY(verticalAlign, rect, totalTextHeight, ascent);

    let charPos = 0;
    for (let li = 0; li < lines.length; li++) {
      const lineY = textStartY - li * lineHeight;
      const lineLen = lines[li].length;

      // Split the line into segments by run
      const segments: Array<{ run: LayoutRichTextRun; text: string; resourceName: string }> = [];
      const segStart = charPos;
      for (let ci = 0; ci < lineLen; ci++) {
        const globalIdx = charPos + ci;
        const ri = runForChar[globalIdx] ?? runForChar.length - 1;
        const last = segments[segments.length - 1];
        if (last && last.run === runs[ri]) {
          last.text += lines[li][ci];
        } else {
          segments.push({
            run: runs[ri],
            text: lines[li][ci],
            resourceName: runResource(runs[ri])
          });
        }
      }
      charPos += lineLen;

      // Skip whitespace chars consumed by word-wrap between lines
      while (charPos < runForChar.length && charPos < segStart + lineLen + 1) {
        const nextLineStart =
          li + 1 < lines.length ? fullText.indexOf(lines[li + 1], charPos) : fullText.length;
        if (nextLineStart > charPos) {
          charPos = nextLineStart;
          break;
        }
        break;
      }

      // Measure total line width for alignment
      let lineWidth = 0;
      for (const seg of segments) {
        lineWidth += fontManager.measureText(seg.text, seg.resourceName, seg.run.fontSize);
      }

      let textX = computeTextX(horizontalAlign, rect, lineWidth, indentPts);
      for (const seg of segments) {
        const { run, text, resourceName } = seg;
        const segWidth = fontManager.measureText(text, resourceName, run.fontSize);

        stream.setFillColor(run.textColor);
        stream.beginText();
        stream.setFont(resourceName, run.fontSize);
        stream.setTextMatrix(1, 0, 0, 1, textX, lineY);
        const hex = fontManager.encodeText(text, resourceName);
        if (hex) {
          stream.showTextHex(hex);
        } else {
          stream.showText(text);
        }
        stream.endText();

        if (run.strike) {
          const descent = fontManager.getFontDescent(resourceName, run.fontSize);
          const y = lineY + descent + run.fontSize * 0.3;
          stream.drawLine(textX, y, textX + segWidth, y, run.textColor, 0.5);
        }
        if (run.underline) {
          const descent = fontManager.getFontDescent(resourceName, run.fontSize);
          const y = lineY + descent * 0.5;
          stream.drawLine(textX, y, textX + segWidth, y, run.textColor, 0.5);
        }

        textX += segWidth;
      }
    }
    return;
  }

  // --- Single-line (no wrap) path ---
  // Measure total width of all runs
  let totalWidth = 0;
  const runMetrics: Array<{ resourceName: string; width: number }> = [];
  for (const run of runs) {
    const resourceName = runResource(run);
    const w = fontManager.measureText(run.text, resourceName, run.fontSize);
    runMetrics.push({ resourceName, width: w });
    totalWidth += w;
  }

  const primaryResourceName = runMetrics[0]?.resourceName ?? "F1";
  const ascent = fontManager.getFontAscent(primaryResourceName, primaryFontSize);
  const textStartY = computeTextStartY(verticalAlign, rect, lineHeight, ascent);
  let textX = computeTextX(horizontalAlign, rect, totalWidth, indentPts);

  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const { resourceName } = runMetrics[i];

    stream.setFillColor(run.textColor);
    stream.beginText();
    stream.setFont(resourceName, run.fontSize);
    stream.setTextMatrix(1, 0, 0, 1, textX, textStartY);

    const hexEncoded = fontManager.encodeText(run.text, resourceName);
    if (hexEncoded) {
      stream.showTextHex(hexEncoded);
    } else {
      stream.showText(run.text);
    }

    stream.endText();

    // Draw per-run decorations (strikethrough, underline)
    const runWidth = runMetrics[i].width;
    if (run.strike) {
      const descent = fontManager.getFontDescent(resourceName, run.fontSize);
      const y = textStartY + descent + run.fontSize * 0.3;
      stream.drawLine(textX, y, textX + runWidth, y, run.textColor, 0.5);
    }
    if (run.underline) {
      const descent = fontManager.getFontDescent(resourceName, run.fontSize);
      const y = textStartY + descent * 0.5;
      stream.drawLine(textX, y, textX + runWidth, y, run.textColor, 0.5);
    }

    textX += runMetrics[i].width;
  }
}

// =============================================================================
// Rotated Text
// =============================================================================

function drawRotatedText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  indentPts: number
): void {
  const { rect, text } = cell;
  let { fontSize } = cell;
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  // Convert Excel rotation to radians
  // 1-90: counterclockwise, 91-180: clockwise (value-90 degrees)
  let degrees: number;
  if (typeof cell.textRotation === "number") {
    if (cell.textRotation <= 90) {
      degrees = cell.textRotation;
    } else {
      degrees = -(cell.textRotation - 90);
    }
  } else {
    degrees = 0;
  }
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  // Scale font size down if rotated bounding box exceeds cell dimensions
  const textWidth = fontManager.measureText(text, resourceName, fontSize);
  const absSin = Math.abs(sin);
  const absCos = Math.abs(cos);
  const rotatedWidth = textWidth * absCos + fontSize * absSin;
  const rotatedHeight = textWidth * absSin + fontSize * absCos;
  const maxWidth = rect.width - CELL_PADDING_H * 2;
  const maxHeight = rect.height - CELL_PADDING_V * 2;
  if (maxWidth > 0 && maxHeight > 0 && (rotatedWidth > maxWidth || rotatedHeight > maxHeight)) {
    const fitScale = Math.min(maxWidth / rotatedWidth, maxHeight / rotatedHeight);
    if (fitScale < 1) {
      fontSize = fontSize * fitScale;
    }
  }

  // Center text at rotation point
  const indentOffset =
    cell.horizontalAlign === "left"
      ? indentPts / 2
      : cell.horizontalAlign === "right"
        ? -indentPts / 2
        : 0;
  const cx = rect.x + rect.width / 2 + indentOffset;
  const cy = rect.y + rect.height / 2;
  const finalTextWidth = fontManager.measureText(text, resourceName, fontSize);
  const ascent = fontManager.getFontAscent(resourceName, fontSize);
  // Offset to center the text around the rotation point
  const offsetX = -finalTextWidth / 2;
  const offsetY = -ascent / 2;
  const tx = cx + offsetX * cos - offsetY * sin;
  const ty = cy + offsetX * sin + offsetY * cos;

  stream.setFillColor(cell.textColor);
  stream.beginText();
  stream.setFont(resourceName, fontSize);
  stream.setTextMatrix(cos, sin, -sin, cos, tx, ty);

  const hexEncoded = fontManager.encodeText(text, resourceName);
  if (hexEncoded) {
    stream.showTextHex(hexEncoded);
  } else {
    stream.showText(text);
  }
  stream.endText();
}

/**
 * Draw vertical stacked text (each character top-to-bottom).
 */
function drawVerticalStackedText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  _indentPts: number
): void {
  const { rect, text, fontSize } = cell;
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  const charHeight = fontSize * 1.3;
  const ascent = fontManager.getFontAscent(resourceName, fontSize);
  const startX = rect.x + rect.width / 2;
  let currentY = rect.y + rect.height - CELL_PADDING_V - ascent;

  stream.setFillColor(cell.textColor);

  for (let i = 0; i < text.length; i++) {
    // Stop if next character would be below cell bottom
    if (currentY < rect.y + CELL_PADDING_V) {
      break;
    }
    const ch = text[i];
    const charWidth = fontManager.measureText(ch, resourceName, fontSize);

    stream.beginText();
    stream.setFont(resourceName, fontSize);
    stream.setTextMatrix(1, 0, 0, 1, startX - charWidth / 2, currentY);

    const hexEncoded = fontManager.encodeText(ch, resourceName);
    if (hexEncoded) {
      stream.showTextHex(hexEncoded);
    } else {
      stream.showText(ch);
    }
    stream.endText();
    currentY -= charHeight;
  }
}

// =============================================================================
// Alpha / ExtGState Helpers
// =============================================================================

/**
 * Generate a deterministic ExtGState resource name for a given alpha value.
 * Uses 4 decimal digits to avoid collisions between close alpha values.
 * E.g. alpha=0.504 → "GS5040", alpha=0.506 → "GS5060"
 */
export function alphaGsName(alpha: number): string {
  return `GS${Math.round(alpha * 10000)}`;
}

// =============================================================================
// Text Layout Helpers
// =============================================================================

/** Indent width per level in points (~3 characters at 11pt) */
const INDENT_WIDTH = 10;

export function computeTextStartY(
  verticalAlign: "top" | "middle" | "bottom",
  rect: PdfRect,
  totalTextHeight: number,
  ascent: number
): number {
  let y: number;
  switch (verticalAlign) {
    case "top":
      y = rect.y + rect.height - CELL_PADDING_V - ascent;
      break;
    case "middle":
      y = rect.y + rect.height / 2 + totalTextHeight / 2 - ascent;
      break;
    case "bottom":
    default:
      y = rect.y + CELL_PADDING_V + (totalTextHeight - ascent);
      break;
  }
  // Clamp: ensure text ascent doesn't exceed the cell top
  const maxY = rect.y + rect.height - CELL_PADDING_V - ascent;
  if (y > maxY) {
    y = maxY;
  }
  // Clamp: ensure text descent doesn't go below cell bottom
  const minY = rect.y + CELL_PADDING_V;
  if (y < minY) {
    y = minY;
  }
  return y;
}

export function computeTextX(
  align: "left" | "center" | "right",
  rect: { x: number; width: number },
  textWidth: number,
  indentPts = 0
): number {
  let x: number;
  switch (align) {
    case "center":
      x = rect.x + (rect.width - textWidth) / 2;
      break;
    case "right":
      x = rect.x + rect.width - CELL_PADDING_H - textWidth;
      break;
    default:
      x = rect.x + CELL_PADDING_H + indentPts;
      break;
  }
  // Clamp: don't start before cell left edge
  const minX = rect.x + CELL_PADDING_H;
  if (x < minX) {
    x = minX;
  }
  return x;
}

function drawTextDecorations(
  stream: PdfContentStream,
  cell: LayoutCell,
  lines: string[],
  lineHeight: number,
  textStartY: number,
  measure: (s: string) => number,
  resourceName: string,
  fontManager: FontManager,
  indentPts: number
): void {
  if (cell.strike) {
    const descent = fontManager.getFontDescent(resourceName, cell.fontSize);
    const strikeY = textStartY + descent + cell.fontSize * 0.3;
    for (let i = 0; i < lines.length; i++) {
      const lineY = strikeY - i * lineHeight;
      const lw = measure(lines[i]);
      const startX = computeTextX(cell.horizontalAlign, cell.rect, lw, indentPts);
      stream.drawLine(startX, lineY, startX + lw, lineY, cell.textColor, 0.5);
    }
  }
  if (cell.underline) {
    const descent = fontManager.getFontDescent(resourceName, cell.fontSize);
    const underlineOffset = descent * 0.5;
    for (let i = 0; i < lines.length; i++) {
      const lineY = textStartY - i * lineHeight + underlineOffset;
      const lw = measure(lines[i]);
      const startX = computeTextX(cell.horizontalAlign, cell.rect, lw, indentPts);
      stream.drawLine(startX, lineY, startX + lw, lineY, cell.textColor, 0.5);
    }
  }
}

// =============================================================================
// Text Wrapping
// =============================================================================

/**
 * Wrap text into lines that fit within the given width.
 * Uses a greedy word-wrap algorithm.
 */
export function wrapTextLines(
  text: string,
  measure: (s: string) => number,
  maxWidth: number
): string[] {
  if (!text) {
    return [""];
  }

  const paragraphs = text.split(/\r?\n/);
  const allLines: string[] = [];

  for (const paragraph of paragraphs) {
    if (!paragraph) {
      allLines.push("");
      continue;
    }

    const words = paragraph.split(/\s+/);
    let currentLine = "";

    for (const word of words) {
      if (!currentLine) {
        currentLine = word;
        continue;
      }

      const testLine = currentLine + " " + word;
      if (measure(testLine) <= maxWidth) {
        currentLine = testLine;
      } else {
        allLines.push(currentLine);
        currentLine = word;
      }
    }

    if (currentLine) {
      allLines.push(currentLine);
    }
  }

  return allLines.length > 0 ? allLines : [""];
}

// =============================================================================
// Page Header / Footer
// =============================================================================

function drawPageHeader(
  stream: PdfContentStream,
  page: LayoutPage,
  options: ResolvedPdfOptions,
  fontManager: FontManager
): void {
  const headerFontSize = 10;
  const headerText = page.sheetName;
  const resourceName = fontManager.hasEmbeddedFont()
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(options.defaultFontFamily, true, false));

  const textWidth = fontManager.measureText(headerText, resourceName, headerFontSize);
  const x = (page.width - textWidth) / 2;
  const y = page.height - options.margins.top + 5;

  stream.save();
  stream.setFillColor({ r: 0.3, g: 0.3, b: 0.3 });
  stream.beginText();
  stream.setFont(resourceName, headerFontSize);
  stream.setTextMatrix(1, 0, 0, 1, x, y);
  const hex = fontManager.encodeText(headerText, resourceName);
  if (hex) {
    stream.showTextHex(hex);
  } else {
    stream.showText(headerText);
  }
  stream.endText();
  stream.restore();
}

function drawPageFooter(
  stream: PdfContentStream,
  page: LayoutPage,
  options: ResolvedPdfOptions,
  fontManager: FontManager,
  totalPages: number
): void {
  const footerFontSize = 9;
  const footerText = `Page ${page.pageNumber} of ${totalPages}`;
  // Footer always uses Type1 (page numbers are ASCII)
  const resourceName = fontManager.ensureFont(
    resolvePdfFontName(options.defaultFontFamily, false, false)
  );

  const textWidth = fontManager.measureText(footerText, resourceName, footerFontSize);
  const x = (page.width - textWidth) / 2;
  const y = Math.max(5, options.margins.bottom - 15);

  stream.save();
  stream.setFillColor({ r: 0.5, g: 0.5, b: 0.5 });
  stream.beginText();
  stream.setFont(resourceName, footerFontSize);
  stream.setTextMatrix(1, 0, 0, 1, x, y);
  stream.showText(footerText);
  stream.endText();
  stream.restore();
}
