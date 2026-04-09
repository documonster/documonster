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
import { CELL_PADDING_H, CELL_PADDING_V, LINE_HEIGHT_FACTOR, INDENT_WIDTH } from "./constants";

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
  const sf = page.scaleFactor;
  for (const cell of page.cells) {
    if (cell.text) {
      drawCellText(stream, cell, fontManager, alphaValues, sf);
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
// Rotation Helpers
// =============================================================================

/**
 * Convert Excel textRotation to standard signed degrees.
 * Excel uses 1-90 for CCW and 91-180 for CW (where 91 = -1°, 180 = -90°).
 * Returns 0 for non-numeric values (e.g. "vertical").
 */
function excelRotationToDegrees(textRotation: number | "vertical"): number {
  if (typeof textRotation !== "number") {
    return 0;
  }
  return textRotation <= 90 ? textRotation : -(textRotation - 90);
}

// =============================================================================
// Cell Borders
// =============================================================================

/**
 * Compute the horizontal slant offset for parallelogram borders.
 * For general rotation angles (not 0°/90°), Excel renders cell borders as a
 * parallelogram whose left/right edges tilt to match the text rotation angle.
 * Returns 0 for straight borders (no rotation, 90°, -90°, or vertical stacked).
 */
function computeSlantOffset(textRotation: number | "vertical", height: number): number {
  const degrees = excelRotationToDegrees(textRotation);
  if (degrees === 0) {
    return 0;
  }
  const absDeg = Math.abs(degrees);
  if (absDeg < 0.01 || absDeg > 89.99) {
    return 0;
  }
  const radians = (absDeg * Math.PI) / 180;
  const offset = (height * Math.cos(radians)) / Math.sin(radians);
  return degrees < 0 ? -offset : offset;
}

function drawCellBorders(stream: PdfContentStream, cell: LayoutCell): void {
  const { rect, borders, textRotation } = cell;
  const { x, y, width, height } = rect;

  // Compute slant for parallelogram borders on general-angle rotated cells
  const slant = computeSlantOffset(textRotation, height);

  if (borders.top) {
    drawBorderLine(stream, borders.top, x + slant, y + height, x + width + slant, y + height, true);
  }
  if (borders.bottom) {
    drawBorderLine(stream, borders.bottom, x, y, x + width, y, true);
  }
  if (borders.left) {
    drawBorderLine(stream, borders.left, x, y, x + slant, y + height, false);
  }
  if (borders.right) {
    drawBorderLine(stream, borders.right, x + width, y, x + width + slant, y + height, false);
  }
}

function drawBorderLine(
  stream: PdfContentStream,
  border: LayoutBorder,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  isHorizontal: boolean
): void {
  if (border.isDouble) {
    // Draw two parallel thin lines with a small gap between them
    const offset = 0.4;
    const thinWidth = Math.min(border.width, 0.25);
    if (isHorizontal) {
      stream.drawLine(
        x1,
        y1 + offset,
        x2,
        y2 + offset,
        border.color,
        thinWidth,
        border.dashPattern
      );
      stream.drawLine(
        x1,
        y1 - offset,
        x2,
        y2 - offset,
        border.color,
        thinWidth,
        border.dashPattern
      );
    } else {
      stream.drawLine(
        x1 + offset,
        y1,
        x2 + offset,
        y2,
        border.color,
        thinWidth,
        border.dashPattern
      );
      stream.drawLine(
        x1 - offset,
        y1,
        x2 - offset,
        y2,
        border.color,
        thinWidth,
        border.dashPattern
      );
    }
  } else {
    stream.drawLine(x1, y1, x2, y2, border.color, border.width, border.dashPattern);
  }
}

// =============================================================================
// Cell Text
// =============================================================================

function drawCellText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  alphaValues: Set<number>,
  scaleFactor = 1
): void {
  const { rect, text, fontSize, horizontalAlign, verticalAlign, wrapText } = cell;

  if (!text && !cell.richText) {
    return;
  }

  const padH = CELL_PADDING_H * scaleFactor;
  const padV = CELL_PADDING_V * scaleFactor;
  const availWidth = rect.width - padH * 2;
  const availHeight = rect.height - padV * 2;
  if (availWidth <= 0 || availHeight <= 0) {
    return;
  }

  const indentPts = cell.indent * INDENT_WIDTH * scaleFactor;

  // Clip to cell bounds (extend for text overflow into adjacent empty cells)
  // For rotated text with slanted borders, use a parallelogram clip path
  const clipWidth = rect.width + (cell.textOverflowWidth || 0);
  stream.save();

  const slantClip = computeSlantOffset(cell.textRotation, rect.height);

  if (slantClip !== 0) {
    // Parallelogram clip: bottom-left, bottom-right, top-right (shifted), top-left (shifted)
    stream.moveTo(rect.x, rect.y);
    stream.lineTo(rect.x + clipWidth, rect.y);
    stream.lineTo(rect.x + clipWidth + slantClip, rect.y + rect.height);
    stream.lineTo(rect.x + slantClip, rect.y + rect.height);
    stream.closePath();
  } else {
    stream.rect(rect.x, rect.y, clipWidth, rect.height);
  }
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
    drawVerticalStackedText(stream, cell, fontManager, indentPts, scaleFactor);
    stream.restore();
    return;
  }
  if (typeof cell.textRotation === "number" && cell.textRotation !== 0) {
    drawRotatedText(stream, cell, fontManager, indentPts, scaleFactor);
    stream.restore();
    return;
  }

  // Handle rich text runs
  if (cell.richText && cell.richText.length > 0) {
    drawRichText(stream, cell, fontManager, indentPts, scaleFactor);
    stream.restore();
    return;
  }

  // --- Plain text rendering ---
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  const measure = (s: string) => fontManager.measureText(s, resourceName, fontSize);
  const effectiveWidth = availWidth - indentPts;
  // Always split on explicit newlines; additionally word-wrap if wrapText is set
  const lines = wrapText ? wrapTextLines(text, measure, effectiveWidth) : text.split(/\r?\n/);

  const lineHeight = fontSize * LINE_HEIGHT_FACTOR;
  const ascent = fontManager.getFontAscent(resourceName, fontSize);
  const totalTextHeight = lines.length * lineHeight;
  const textStartY = computeTextStartY(verticalAlign, rect, totalTextHeight, ascent, padV);

  stream.setFillColor(cell.textColor);
  stream.beginText();
  stream.setFont(resourceName, fontSize);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineY = textStartY - i * lineHeight;
    const textWidth = measure(line);
    const textX = computeTextX(horizontalAlign, rect, textWidth, indentPts, padH);

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
  indentPts: number,
  scaleFactor = 1
): void {
  const { rect, horizontalAlign, verticalAlign, wrapText } = cell;
  const runs = cell.richText!;
  const padH = CELL_PADDING_H * scaleFactor;
  const padV = CELL_PADDING_V * scaleFactor;

  // Use the largest font size across all runs for line height calculation
  let maxFontSize = cell.fontSize;
  for (const run of runs) {
    if (run.fontSize > maxFontSize) {
      maxFontSize = run.fontSize;
    }
  }
  const primaryFontSize = maxFontSize;
  const lineHeight = primaryFontSize * LINE_HEIGHT_FACTOR;

  const isEmbedded = fontManager.hasEmbeddedFont();

  // Helper: resolve resource name for a run
  const runResource = (run: LayoutRichTextRun) =>
    isEmbedded
      ? fontManager.getEmbeddedResourceName()
      : fontManager.ensureFont(resolvePdfFontName(run.fontFamily, run.bold, run.italic));

  // --- Wrapping path ---
  if (wrapText) {
    const availWidth = rect.width - padH * 2 - indentPts;
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
    const textStartY = computeTextStartY(verticalAlign, rect, totalTextHeight, ascent, padV);

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

      let textX = computeTextX(horizontalAlign, rect, lineWidth, indentPts, padH);
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
  const textStartY = computeTextStartY(verticalAlign, rect, lineHeight, ascent, padV);
  let textX = computeTextX(horizontalAlign, rect, totalWidth, indentPts, padH);

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
  indentPts: number,
  scaleFactor = 1
): void {
  const { rect, wrapText } = cell;
  let { fontSize } = cell;
  const padH = CELL_PADDING_H * scaleFactor;
  const padV = CELL_PADDING_V * scaleFactor;
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  // Convert Excel rotation to degrees
  const degrees = excelRotationToDegrees(cell.textRotation);

  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const absSin = Math.abs(sin);
  const absCos = Math.abs(cos);

  const maxWidth = rect.width - padH * 2;
  const maxHeight = rect.height - padV * 2;

  // Available length along the text flow direction for wrapping
  let availTextLength: number;
  if (absSin > 0.01 && absCos > 0.01) {
    availTextLength = Math.min(maxHeight / absSin, maxWidth / absCos);
  } else if (absSin > 0.01) {
    availTextLength = maxHeight / absSin;
  } else {
    availTextLength = maxWidth;
  }

  const measure = (s: string) => fontManager.measureText(s, resourceName, fontSize);

  // Split on explicit newlines first, then optionally word-wrap each paragraph
  let lines: string[];
  if (wrapText) {
    lines = wrapTextLines(cell.text, measure, Math.max(availTextLength - 1, 1));
  } else {
    lines = cell.text.split(/\r?\n/);
  }

  const lineHeight = fontSize * LINE_HEIGHT_FACTOR;
  const totalTextHeight = lines.length * lineHeight;

  // For non-wrapping text: scale font down if the rotated bounding box exceeds cell
  if (!wrapText) {
    let maxLineWidth = 0;
    for (const line of lines) {
      const w = measure(line);
      if (w > maxLineWidth) {
        maxLineWidth = w;
      }
    }
    const rotatedWidth = maxLineWidth * absCos + totalTextHeight * absSin;
    const rotatedHeight = maxLineWidth * absSin + totalTextHeight * absCos;
    if (maxWidth > 0 && maxHeight > 0 && (rotatedWidth > maxWidth || rotatedHeight > maxHeight)) {
      const fitScale = Math.min(maxWidth / rotatedWidth, maxHeight / rotatedHeight);
      if (fitScale < 1) {
        fontSize = fontSize * fitScale;
      }
    }
  }

  const scaledLineHeight = fontSize * LINE_HEIGHT_FACTOR;
  const ascent = fontManager.getFontAscent(resourceName, fontSize);

  const is90 = Math.abs(degrees - 90) < 0.01;
  const isMinus90 = Math.abs(degrees + 90) < 0.01;

  stream.setFillColor(cell.textColor);

  if (is90) {
    // Text reads bottom-to-top. Each line becomes a column drawn left-to-right.
    drawRotated90(
      stream,
      cell,
      lines,
      fontManager,
      resourceName,
      fontSize,
      scaledLineHeight,
      ascent,
      padH,
      padV
    );
  } else if (isMinus90) {
    // Text reads top-to-bottom. Each line becomes a column drawn right-to-left.
    drawRotatedMinus90(
      stream,
      cell,
      lines,
      fontManager,
      resourceName,
      fontSize,
      scaledLineHeight,
      ascent,
      padH,
      padV
    );
  } else {
    // General rotation — center multi-line text block in cell
    drawRotatedGeneral(
      stream,
      cell,
      lines,
      fontManager,
      resourceName,
      fontSize,
      scaledLineHeight,
      ascent,
      cos,
      sin,
      indentPts
    );
  }
}

/** 90° CCW: text reads bottom-to-top, lines stack left-to-right. */
function drawRotated90(
  stream: PdfContentStream,
  cell: LayoutCell,
  lines: string[],
  fontManager: FontManager,
  resourceName: string,
  fontSize: number,
  lineHeight: number,
  ascent: number,
  padH: number,
  padV: number
): void {
  const { rect, horizontalAlign, verticalAlign } = cell;
  const totalColumnsWidth = lines.length * lineHeight;

  // horizontalAlign controls X placement of line columns (same visual axis)
  let startX: number;
  if (horizontalAlign === "center") {
    startX = rect.x + rect.width / 2 - totalColumnsWidth / 2 + ascent;
  } else if (horizontalAlign === "right") {
    startX = rect.x + rect.width - padH - totalColumnsWidth + ascent;
  } else {
    // left (default)
    startX = rect.x + padH + ascent;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineWidth = fontManager.measureText(line, resourceName, fontSize);
    const colX = startX + i * lineHeight;

    // verticalAlign controls Y placement (text flows upward from ty)
    // In PDF coords: higher y = top of cell
    let ty: number;
    if (verticalAlign === "top") {
      // text at top → text end near top → ty starts at bottom so text reaches top
      ty = rect.y + rect.height - padV - lineWidth;
    } else if (verticalAlign === "middle") {
      ty = rect.y + (rect.height - lineWidth) / 2;
    } else {
      // bottom (default) → text at bottom → ty near bottom
      ty = rect.y + padV;
    }
    ty = Math.max(ty, rect.y + padV);

    stream.beginText();
    stream.setFont(resourceName, fontSize);
    stream.setTextMatrix(0, 1, -1, 0, colX, ty);
    emitText(stream, fontManager, line, resourceName);
    stream.endText();
  }
}

/** -90° (270° CW): text reads top-to-bottom, lines stack right-to-left. */
function drawRotatedMinus90(
  stream: PdfContentStream,
  cell: LayoutCell,
  lines: string[],
  fontManager: FontManager,
  resourceName: string,
  fontSize: number,
  lineHeight: number,
  ascent: number,
  padH: number,
  padV: number
): void {
  const { rect, horizontalAlign, verticalAlign } = cell;
  const totalColumnsWidth = lines.length * lineHeight;

  // horizontalAlign controls X placement: lines stack right-to-left
  let startX: number;
  if (horizontalAlign === "center") {
    startX = rect.x + rect.width / 2 + totalColumnsWidth / 2 - lineHeight + ascent;
  } else if (horizontalAlign === "right") {
    startX = rect.x + rect.width - padH - lineHeight + ascent;
  } else {
    // left (default)
    startX = rect.x + padH + totalColumnsWidth - lineHeight + ascent;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineWidth = fontManager.measureText(line, resourceName, fontSize);
    const colX = startX - i * lineHeight;

    // verticalAlign controls Y placement (text flows downward from ty)
    // In PDF coords: higher y = top of cell; text drawn downward = toward lower y
    let ty: number;
    if (verticalAlign === "top") {
      // text at top → ty near top (high PDF y)
      ty = rect.y + rect.height - padV;
    } else if (verticalAlign === "middle") {
      ty = rect.y + (rect.height + lineWidth) / 2;
    } else {
      // bottom (default) → text at bottom → ty so text ends at bottom
      ty = rect.y + padV + lineWidth;
    }
    ty = Math.min(ty, rect.y + rect.height - padV);

    stream.beginText();
    stream.setFont(resourceName, fontSize);
    stream.setTextMatrix(0, -1, 1, 0, colX, ty);
    emitText(stream, fontManager, line, resourceName);
    stream.endText();
  }
}

/** General rotation — center a multi-line text block in the cell. */
function drawRotatedGeneral(
  stream: PdfContentStream,
  cell: LayoutCell,
  lines: string[],
  fontManager: FontManager,
  resourceName: string,
  fontSize: number,
  lineHeight: number,
  ascent: number,
  cos: number,
  sin: number,
  indentPts: number
): void {
  const { rect, horizontalAlign, verticalAlign } = cell;
  const padH = CELL_PADDING_H;
  const padV = CELL_PADDING_V;

  // Compute the rotated bounding box of the text block
  let maxLineWidth = 0;
  for (const line of lines) {
    const w = fontManager.measureText(line, resourceName, fontSize);
    if (w > maxLineWidth) {
      maxLineWidth = w;
    }
  }
  const totalTextHeight = lines.length * lineHeight;
  const absSin = Math.abs(sin);
  const absCos = Math.abs(cos);
  const rotatedWidth = maxLineWidth * absCos + totalTextHeight * absSin;
  const rotatedHeight = maxLineWidth * absSin + totalTextHeight * absCos;

  // Compute slant offset to match parallelogram border shape
  const slantShift = computeSlantOffset(cell.textRotation, rect.height) / 2;

  // Determine vertical position first, then horizontal (because slant depends on Y position)
  const indentOffset =
    horizontalAlign === "left" ? indentPts / 2 : horizontalAlign === "right" ? -indentPts / 2 : 0;

  let cy: number;
  if (verticalAlign === "top") {
    cy = rect.y + rect.height - padV - rotatedHeight / 2;
  } else if (verticalAlign === "bottom") {
    cy = rect.y + padV + rotatedHeight / 2;
  } else {
    // middle (default)
    cy = rect.y + rect.height / 2;
  }

  // For slanted parallelogram, the horizontal offset depends on the vertical position
  // At bottom (y), left edge is at x; at top (y+height), left edge is at x+slantOffset
  // At cy, the horizontal shift is proportional: slantOffset * (cy - y) / height
  const verticalRatio = rect.height > 0 ? (cy - rect.y) / rect.height : 0.5;
  const slantAtCy = slantShift * 2 * verticalRatio; // slantShift*2 = full slantOffset

  let cx: number;
  if (horizontalAlign === "right") {
    cx = rect.x + rect.width - padH - rotatedWidth / 2 + indentOffset + slantAtCy;
  } else if (horizontalAlign === "left") {
    cx = rect.x + padH + rotatedWidth / 2 + indentOffset + slantAtCy;
  } else {
    // center (default for rotated)
    cx = rect.x + rect.width / 2 + indentOffset + slantAtCy;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineWidth = fontManager.measureText(line, resourceName, fontSize);
    const lineOffset = (i - (lines.length - 1) / 2) * lineHeight;
    const offsetX = -lineWidth / 2;
    const offsetY = -ascent / 2 - lineOffset;
    const tx = cx + offsetX * cos - offsetY * sin;
    const ty = cy + offsetX * sin + offsetY * cos;

    stream.beginText();
    stream.setFont(resourceName, fontSize);
    stream.setTextMatrix(cos, sin, -sin, cos, tx, ty);
    emitText(stream, fontManager, line, resourceName);
    stream.endText();
  }
}

/** Emit a text string with hex encoding if available. */
function emitText(
  stream: PdfContentStream,
  fontManager: FontManager,
  text: string,
  resourceName: string
): void {
  const hex = fontManager.encodeText(text, resourceName);
  if (hex) {
    stream.showTextHex(hex);
  } else {
    stream.showText(text);
  }
}

/**
 * Draw vertical stacked text (each character top-to-bottom).
 * Newlines (\n) start a new column to the right.
 */
function drawVerticalStackedText(
  stream: PdfContentStream,
  cell: LayoutCell,
  fontManager: FontManager,
  _indentPts: number,
  scaleFactor = 1
): void {
  const { rect, text, fontSize, horizontalAlign, verticalAlign } = cell;
  const padH = CELL_PADDING_H * scaleFactor;
  const padV = CELL_PADDING_V * scaleFactor;
  const isEmbedded = fontManager.hasEmbeddedFont();
  const resourceName = isEmbedded
    ? fontManager.getEmbeddedResourceName()
    : fontManager.ensureFont(resolvePdfFontName(cell.fontFamily, cell.bold, cell.italic));

  const charHeight = fontSize * 1.3;
  const ascent = fontManager.getFontAscent(resourceName, fontSize);

  // Split on newlines — each segment becomes a new column
  const columns = text.split(/\r?\n/);
  const columnWidth = fontSize * 1.4;
  const totalColumnsWidth = columns.length * columnWidth;

  // Horizontal alignment controls column X positioning
  let startX: number;
  if (horizontalAlign === "center") {
    startX = rect.x + rect.width / 2 - totalColumnsWidth / 2 + columnWidth / 2;
  } else if (horizontalAlign === "right") {
    startX = rect.x + rect.width - padH - totalColumnsWidth + columnWidth / 2;
  } else {
    // left (default)
    startX = rect.x + padH + columnWidth / 2;
  }

  stream.setFillColor(cell.textColor);

  for (let colIdx = 0; colIdx < columns.length; colIdx++) {
    const colText = columns[colIdx];
    const colX = startX + colIdx * columnWidth;
    const totalTextHeight = colText.length * charHeight;

    // Vertical alignment controls starting Y position (PDF y-axis: higher = top of cell)
    let currentY: number;
    if (verticalAlign === "middle") {
      currentY = rect.y + rect.height / 2 + totalTextHeight / 2 - ascent;
    } else if (verticalAlign === "bottom") {
      currentY = rect.y + padV + totalTextHeight - ascent;
    } else {
      // top (default)
      currentY = rect.y + rect.height - padV - ascent;
    }

    for (const ch of colText) {
      if (currentY < rect.y + padV) {
        break;
      }
      const charWidth = fontManager.measureText(ch, resourceName, fontSize);

      stream.beginText();
      stream.setFont(resourceName, fontSize);
      stream.setTextMatrix(1, 0, 0, 1, colX - charWidth / 2, currentY);
      emitText(stream, fontManager, ch, resourceName);
      stream.endText();
      currentY -= charHeight;
    }
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

export function computeTextStartY(
  verticalAlign: "top" | "middle" | "bottom",
  rect: PdfRect,
  totalTextHeight: number,
  ascent: number,
  padV = CELL_PADDING_V
): number {
  let y: number;
  switch (verticalAlign) {
    case "top":
      y = rect.y + rect.height - padV - ascent;
      break;
    case "middle":
      y = rect.y + rect.height / 2 + totalTextHeight / 2 - ascent;
      break;
    case "bottom":
    default:
      y = rect.y + padV + (totalTextHeight - ascent);
      break;
  }
  // Clamp: ensure text ascent doesn't exceed the cell top
  const maxY = rect.y + rect.height - padV - ascent;
  if (y > maxY) {
    y = maxY;
  }
  // Clamp: ensure text descent doesn't go below cell bottom
  const minY = rect.y + padV;
  if (y < minY) {
    y = minY;
  }
  return y;
}

export function computeTextX(
  align: "left" | "center" | "right",
  rect: { x: number; width: number },
  textWidth: number,
  indentPts = 0,
  padH = CELL_PADDING_H
): number {
  let x: number;
  switch (align) {
    case "center":
      x = rect.x + (rect.width - textWidth) / 2;
      break;
    case "right":
      x = rect.x + rect.width - padH - textWidth;
      break;
    default:
      x = rect.x + padH + indentPts;
      break;
  }
  // Clamp: don't start before cell left edge
  const minX = rect.x + padH;
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
