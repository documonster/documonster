/**
 * Text measurement engine for auto-fit column width and row height calculation.
 *
 * ## Algorithm Summary
 *
 * **Width calculation** follows a 3-tier approach:
 *
 * 1. **Calibri 11pt**: Use pre-computed bitmap pixel widths (exact match with Excel)
 * 2. **Known fonts at any size**: Use FUnit advance widths with the formula:
 *    `pixelWidth = ROUND(advanceFU / unitsPerEm * ROUND(fontSize / 72 * 96))`
 * 3. **Unknown fonts**: Fall back to excelize-style category-average factors
 *
 * **Height calculation** uses the ClosedXML-verified formula:
 *    `lineHeight = (unitsPerEm + usWinDescent) / unitsPerEm * fontSizePx`
 *
 * **Unit conversions** follow the ECMA-376 spec as verified by ClosedXML:
 * - Column width in XLSX = `TRUNC(pixelWidth / MDW * 256) / 256`
 * - MDW = max digit width in pixels (Calibri 11pt: 7)
 * - Pixel Padding (PP) = `2 * CEIL(MDW / 4) + 1`
 *
 * ## Key References
 * - ClosedXML wiki Cell Dimensions
 * - ECMA-376 §18.3.1.13 (col width)
 * - rust_xlsxwriter utility.rs (Calibri 11pt pixel table)
 * - excelize col.go (factor-based calculation)
 */
import type { Font, Alignment, NumFmt, RichText } from "@excel/types";
import { ValueType } from "@excel/enums";
import { getCellDisplayText } from "@excel/utils/cell-format";
import {
  getCalibri11PtPixelWidth,
  getFontMetrics,
  getDefaultFontMetrics,
  getCharAdvance,
  getFontWidthFactors,
  isWideCharacter,
  hasBoldMetrics,
  type FontMetrics
} from "@excel/utils/font-data";

// =============================================================================
// Constants
// =============================================================================

/** Default DPI for Excel rendering */
const DPI = 96;

/** Default font size in points */
const DEFAULT_FONT_SIZE = 11;

/** Default font name */
const DEFAULT_FONT_NAME = "Calibri";

/**
 * Calibri 11pt bitmap MDW (Max Digit Width) in pixels.
 * All Calibri digits 0-9 have advance width 1038 FU.
 * With bitmap metrics at ppem=15, the actual rendered width is 7px.
 * The outline formula gives ROUND(1038/2048*15) = 8, but bitmap overrides.
 */
const CALIBRI_11PT_MDW = 7;

/** Maximum auto-fit column width in pixels (~255 characters) */
const MAX_AUTOFIT_WIDTH_PX = 1790;

/** Maximum auto-fit column width in character units */
const MAX_COLUMN_WIDTH = 255;

/** Autofilter dropdown arrow width in pixels at 96 DPI */
const AUTOFILTER_ARROW_PX = 16;

// =============================================================================
// Pixel Width Calculation (per-character)
// =============================================================================

/**
 * Calculate the pixel width of a single character given font parameters.
 *
 * For Calibri 11pt (both regular and bold), uses the bitmap pixel table (Tier 1).
 * Bold adjustment is handled by the caller via the 1.05 multiplier.
 * For other fonts/sizes with FUnit data, uses outline formula (Tier 2).
 * Returns undefined if no data is available for this font (use Tier 3).
 */
function getCharPixelWidth(
  codePoint: number,
  fontName: string,
  fontSize: number,
  _bold: boolean,
  metrics: FontMetrics | undefined
): number | undefined {
  // Tier 1: Calibri 11pt bitmap (used for both regular and bold base measurement)
  if (fontName === "calibri" && fontSize === DEFAULT_FONT_SIZE) {
    const px = getCalibri11PtPixelWidth(codePoint);
    if (px !== undefined) {
      return px;
    }
    // Character not in bitmap table; fall through to Tier 2
  }

  // Tier 2: FUnit outline calculation
  if (metrics) {
    const advanceFU = getCharAdvance(metrics, codePoint);
    const ppem = Math.round((fontSize / 72) * DPI);
    return Math.round((advanceFU / metrics.header.unitsPerEm) * ppem);
  }

  return undefined;
}

// =============================================================================
// Text Width Measurement
// =============================================================================

/** Resolved font parameters for measurement */
interface ResolvedFont {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  vertAlign?: "superscript" | "subscript";
}

/**
 * Resolve a partial Font to concrete measurement parameters.
 */
function resolveFont(font?: Partial<Font>): ResolvedFont {
  return {
    name: (font?.name ?? DEFAULT_FONT_NAME).toLowerCase(),
    size: font?.size ?? DEFAULT_FONT_SIZE,
    bold: font?.bold ?? false,
    italic: font?.italic ?? false,
    vertAlign: font?.vertAlign
  };
}

/**
 * Measure the pixel width of a text string with a given font.
 *
 * Handles:
 * - Per-character precise measurement (Tier 1 & 2)
 * - Category-average fallback for unknown fonts (Tier 3)
 * - Multi-line text (returns width of widest line)
 * - Bold/italic modifiers
 * - Superscript/subscript scaling
 */
export function measureTextWidthPx(text: string, font?: Partial<Font>): number {
  if (!text) {
    return 0;
  }

  const resolved = resolveFont(font);
  const metrics = getFontMetrics(resolved.name, resolved.bold);

  // Split by newlines, measure each line, return max
  const lines = text.split(/\r\n|\r|\n/);
  let maxWidth = 0;

  for (const line of lines) {
    const width = measureLineWidthPx(line, resolved, metrics);
    if (width > maxWidth) {
      maxWidth = width;
    }
  }

  return maxWidth;
}

/**
 * Measure a single line of text (no newlines) in pixels.
 */
function measureLineWidthPx(
  line: string,
  resolved: ResolvedFont,
  metrics: FontMetrics | undefined
): number {
  if (!line) {
    return 0;
  }

  // Try per-character measurement (Tier 1 & 2)
  if (metrics || (resolved.name === "calibri" && resolved.size === DEFAULT_FONT_SIZE)) {
    return measureLineWithGlyphs(line, resolved, metrics);
  }

  // Tier 3: Factor-based fallback
  return measureLineWithFactors(line, resolved);
}

/**
 * Tier 1 & 2: Per-character pixel width measurement.
 */
function measureLineWithGlyphs(
  line: string,
  resolved: ResolvedFont,
  metrics: FontMetrics | undefined
): number {
  let totalWidth = 0;

  for (const char of line) {
    totalWidth += _measureCharPx(char.codePointAt(0)!, resolved, metrics);
  }

  // Superscript/subscript renders at ~60% size
  if (resolved.vertAlign) {
    totalWidth = Math.ceil(totalWidth * 0.6);
  }

  return totalWidth;
}

/**
 * Tier 3: Factor-based width measurement for unknown fonts.
 * Uses excelize-style category averages (lowercase, uppercase, wide).
 */
function measureLineWithFactors(line: string, resolved: ResolvedFont): number {
  const factors = getFontWidthFactors(resolved.name);
  const lowerFactor = factors?.[0] ?? 1.0;
  const upperFactor = factors?.[1] ?? 1.3;
  const wideFactor = factors?.[2] ?? 1.0;

  let lowerUnits = 0;
  let upperUnits = 0;
  let wideUnits = 0;

  for (const char of line) {
    const cp = char.codePointAt(0)!;
    if (isWideCharacter(cp)) {
      wideUnits += 2;
    } else if (char >= "A" && char <= "Z") {
      upperUnits++;
    } else {
      lowerUnits++;
    }
  }

  // Width in "character units" (where 1 unit = average char at 11pt)
  const charWidth =
    (lowerUnits * lowerFactor + upperUnits * upperFactor + wideUnits * wideFactor) *
    (resolved.size / DEFAULT_FONT_SIZE);

  // Apply bold/italic
  let width = charWidth;
  if (resolved.bold) {
    width *= 1.05;
  }
  if (resolved.italic) {
    width *= 1.02;
  }
  if (resolved.vertAlign) {
    width *= 0.6;
  }

  // Convert from character units to pixels using the font's MDW
  const mdw = getMaxDigitWidth({ name: resolved.name, size: resolved.size });
  return Math.ceil(width * mdw);
}

// =============================================================================
// Rich Text Width Measurement
// =============================================================================

/**
 * Measure the pixel width of rich text (mixed fonts).
 * Each run may have a different font; width is summed per run.
 * Line breaks reset the accumulator.
 */
export function measureRichTextWidthPx(richText: RichText[], defaultFont?: Partial<Font>): number {
  let maxLineWidth = 0;
  let currentLineWidth = 0;

  for (const run of richText) {
    const font = run.font ? { ...defaultFont, ...run.font } : defaultFont;

    // Handle newlines within a run
    const parts = run.text.split(/\r\n|\r|\n/);
    for (let i = 0; i < parts.length; i++) {
      if (i > 0) {
        // New line: save current line width and reset
        if (currentLineWidth > maxLineWidth) {
          maxLineWidth = currentLineWidth;
        }
        currentLineWidth = 0;
      }
      if (parts[i]) {
        currentLineWidth += measureTextWidthPx(parts[i], font);
      }
    }
  }

  if (currentLineWidth > maxLineWidth) {
    maxLineWidth = currentLineWidth;
  }

  return maxLineWidth;
}

// =============================================================================
// Unit Conversions (ClosedXML-verified formulas)
// =============================================================================

/**
 * Calculate the Max Digit Width (MDW) in pixels for a given font.
 *
 * Special case: Calibri 11pt returns 7 (bitmap metrics override).
 * For other fonts: MDW = ROUND(maxDigitAdvanceFU / unitsPerEm * ppem)
 */
export function getMaxDigitWidth(font?: Partial<Font>): number {
  const name = (font?.name ?? DEFAULT_FONT_NAME).toLowerCase();
  const size = font?.size ?? DEFAULT_FONT_SIZE;

  // Calibri 11pt: bitmap MDW = 7
  if (name === "calibri" && size === DEFAULT_FONT_SIZE) {
    return CALIBRI_11PT_MDW;
  }

  const metrics = getFontMetrics(name);
  if (metrics) {
    const ppem = Math.round((size / 72) * DPI);
    return Math.round((metrics.header.maxDigitAdvance / metrics.header.unitsPerEm) * ppem);
  }

  // Fallback: scale from Calibri proportionally
  return Math.max(1, Math.round(CALIBRI_11PT_MDW * (size / DEFAULT_FONT_SIZE)));
}

/**
 * Calculate Pixel Padding (PP) from MDW.
 * Formula: PP = 2 * CEIL(MDW / 4) + 1
 */
export function getPixelPadding(mdw: number): number {
  return 2 * Math.ceil(mdw / 4) + 1;
}

/**
 * Convert pixel width to Excel column character width (stored in XLSX).
 *
 * Formula: TRUNC(pixels / MDW * 256) / 256
 * This gives the column width in MDW-based character units with 1/256 precision.
 */
export function pixelToCharWidth(pixels: number, mdw: number): number {
  if (mdw <= 0) {
    return 0;
  }
  return Math.trunc((pixels / mdw) * 256) / 256;
}

/**
 * Convert Excel character width to pixel width.
 *
 * The formula differs for widths below 1 character:
 * - width < 1: pixels = ROUND(width * (MDW + PP))
 * - width >= 1: pixels = ROUND(width * MDW) + PP
 */
export function charWidthToPixel(width: number, mdw: number): number {
  if (mdw <= 0) {
    return 0;
  }
  const pp = getPixelPadding(mdw);
  if (width === 0) {
    return 0;
  }
  if (width < 1) {
    return Math.round(width * (mdw + pp));
  }
  return Math.round(width * mdw) + pp;
}

/**
 * Convert pixel height to points.
 * 1 point = 1/72 inch, 1 pixel = 1/DPI inch
 * points = pixels * 72 / DPI
 */
export function pixelToPoints(pixels: number): number {
  return (pixels * 72) / DPI;
}

/**
 * Convert points to pixels.
 * pixels = points * DPI / 72
 */
export function pointsToPixel(points: number): number {
  return (points * DPI) / 72;
}

// =============================================================================
// Auto-Fit Column Width
// =============================================================================

/**
 * Calculate the auto-fit column width in character units for a cell's text.
 *
 * This is the main entry point for column auto-fit calculation.
 *
 * @param textWidthPx - The pixel width of the cell's text content
 * @param mdw - Max digit width in pixels for the workbook's default font
 * @param hasAutoFilter - Whether the column is part of an auto-filter
 * @returns Column width in Excel character units
 */
export function calculateAutoFitWidth(
  textWidthPx: number,
  mdw: number,
  hasAutoFilter?: boolean
): number {
  if (textWidthPx <= 0) {
    return 0;
  }

  // ClosedXML padding formula:
  // oneSidePadding = CEIL(textWidth * 0.03 + mdw / 4)
  // totalWidth = textWidth + 2 * oneSidePadding + 1 (gridline)
  const oneSidePadding = Math.ceil(textWidthPx * 0.03 + mdw / 4);
  let totalPx = textWidthPx + 2 * oneSidePadding + 1;

  // Add autofilter dropdown space
  if (hasAutoFilter) {
    totalPx += AUTOFILTER_ARROW_PX;
  }

  // Clamp to maximum
  if (totalPx > MAX_AUTOFIT_WIDTH_PX) {
    totalPx = MAX_AUTOFIT_WIDTH_PX;
  }

  // Convert to character units
  const charWidth = pixelToCharWidth(totalPx, mdw);
  return Math.min(charWidth, MAX_COLUMN_WIDTH);
}

// =============================================================================
// Auto-Fit Row Height
// =============================================================================

/**
 * Calculate the line height in pixels for a font.
 *
 * Uses the ClosedXML-verified formula:
 *   lineHeight = (unitsPerEm + usWinDescent) / unitsPerEm * fontSizePx
 *
 * This matches Excel's actual row height calculation.
 */
export function getLineHeightPx(font?: Partial<Font>): number {
  const name = (font?.name ?? DEFAULT_FONT_NAME).toLowerCase();
  const size = font?.size ?? DEFAULT_FONT_SIZE;
  const fontSizePx = Math.round((size / 72) * DPI);

  const metrics = getFontMetrics(name);
  if (metrics) {
    const { unitsPerEm, usWinDescent } = metrics.header;
    return ((unitsPerEm + usWinDescent) / unitsPerEm) * fontSizePx;
  }

  // Fallback: approximate. The ratio for most fonts is ~1.2 to 1.35.
  // Calibri is (2048 + 550) / 2048 = 1.268. Use 1.3 as a safe default.
  return fontSizePx * 1.3;
}

/**
 * Calculate the number of wrapped lines for text in a column of given width.
 *
 * Excel wraps text at word boundaries (spaces, hyphens) rather than at
 * arbitrary character positions. If a single word is wider than the column,
 * it overflows (Excel does not break mid-word in normal wrap mode).
 *
 * @param text - The cell text (may contain explicit newlines)
 * @param columnWidthPx - Available column width in pixels (content area, excluding padding)
 * @param font - Cell font
 * @returns Number of lines the text will occupy
 */
export function calculateWrappedLineCount(
  text: string,
  columnWidthPx: number,
  font?: Partial<Font>
): number {
  if (!text || columnWidthPx <= 0) {
    return 1;
  }

  const resolved = resolveFont(font);
  const metrics = getFontMetrics(resolved.name, resolved.bold);
  const lines = text.split(/\r\n|\r|\n/);
  let totalLines = 0;

  for (const line of lines) {
    if (!line) {
      totalLines++;
      continue;
    }

    totalLines += _countWrappedLines(line, columnWidthPx, resolved, metrics);
  }

  return totalLines;
}

/**
 * Measure pixel width of a single character with font adjustments.
 * Shared by both width measurement and wrap calculation.
 */
function _measureCharPx(
  codePoint: number,
  resolved: ResolvedFont,
  metrics: FontMetrics | undefined
): number {
  let charWidth: number;

  if (metrics || (resolved.name === "calibri" && resolved.size === DEFAULT_FONT_SIZE)) {
    const effectiveMetrics = metrics ?? getDefaultFontMetrics();
    charWidth =
      getCharPixelWidth(codePoint, resolved.name, resolved.size, resolved.bold, effectiveMetrics) ??
      Math.round(
        (effectiveMetrics.defaultAdvance / effectiveMetrics.header.unitsPerEm) *
          Math.round((resolved.size / 72) * DPI)
      );
  } else {
    charWidth = Math.ceil(
      (resolved.size / DEFAULT_FONT_SIZE) *
        getMaxDigitWidth({ name: resolved.name, size: resolved.size }) *
        (isWideCharacter(codePoint) ? 2 : 1)
    );
  }

  // Apply bold/italic multipliers
  if (resolved.bold && !hasBoldMetrics(resolved.name)) {
    charWidth = Math.ceil(charWidth * 1.05);
  }
  if (resolved.italic) {
    charWidth = Math.ceil(charWidth * 1.02);
  }

  return charWidth;
}

/**
 * Count wrapped lines for a single line (no explicit newlines) using
 * word-boundary wrapping that matches Excel behavior.
 *
 * Excel breaks at spaces and hyphens. If a single word exceeds the column
 * width, the word overflows on its line (Excel does not mid-word break).
 */
function _countWrappedLines(
  line: string,
  columnWidthPx: number,
  resolved: ResolvedFont,
  metrics: FontMetrics | undefined
): number {
  // Split line into "words" — segments separated by spaces or hyphens.
  // The delimiter (space/hyphen) is included at the end of the preceding word,
  // matching how Excel accounts for space width before breaking.
  const words = _splitIntoWords(line);

  if (words.length === 0) {
    return 1;
  }

  let lineCount = 1;
  let currentWidth = 0;

  for (const word of words) {
    // Measure word width
    let wordWidth = 0;
    for (const char of word) {
      wordWidth += _measureCharPx(char.codePointAt(0)!, resolved, metrics);
    }

    if (currentWidth === 0) {
      // First word on line: always place it (even if wider than column)
      currentWidth = wordWidth;
    } else if (currentWidth + wordWidth > columnWidthPx) {
      // Word doesn't fit on current line: wrap
      lineCount++;
      currentWidth = wordWidth;
    } else {
      // Word fits: accumulate
      currentWidth += wordWidth;
    }
  }

  return lineCount;
}

/**
 * Split a line into words for wrapping purposes.
 * Delimiters (space, hyphen) are kept at the end of the preceding word.
 * This matches Excel's wrapping behavior where the space is consumed
 * before the line break.
 *
 * "Hello World" → ["Hello ", "World"]
 * "one-two-three" → ["one-", "two-", "three"]
 * "a  b" → ["a ", " ", "b"]
 */
function _splitIntoWords(line: string): string[] {
  const words: string[] = [];
  let current = "";

  for (const char of line) {
    current += char;
    // Break after spaces and hyphens
    if (char === " " || char === "-" || char === "\t") {
      words.push(current);
      current = "";
    }
  }
  if (current) {
    words.push(current);
  }

  return words;
}

/**
 * Calculate the auto-fit row height in points for a cell.
 *
 * @param text - Cell display text
 * @param font - Cell font
 * @param alignment - Cell alignment (for wrapText check)
 * @param columnWidthPx - Column content width in pixels (needed for wrapText)
 * @returns Row height in points
 */
export function calculateAutoFitHeight(
  text: string,
  font?: Partial<Font>,
  alignment?: Partial<Alignment>,
  columnWidthPx?: number
): number {
  if (!text) {
    return pixelToPoints(getLineHeightPx(font));
  }

  const lineHeightPx = getLineHeightPx(font);
  let lineCount: number;

  if (alignment?.wrapText && columnWidthPx && columnWidthPx > 0) {
    lineCount = calculateWrappedLineCount(text, columnWidthPx, font);
  } else {
    // Count explicit newlines only
    lineCount = text.split(/\r\n|\r|\n/).length;
  }

  return pixelToPoints(lineHeightPx * lineCount);
}

/**
 * Calculate the auto-fit row height for rich text.
 */
export function calculateRichTextAutoFitHeight(
  richText: RichText[],
  defaultFont?: Partial<Font>,
  alignment?: Partial<Alignment>,
  columnWidthPx?: number
): number {
  // Find the largest font in any run (determines line height)
  let maxFontSize = defaultFont?.size ?? DEFAULT_FONT_SIZE;
  let maxFontForHeight: Partial<Font> | undefined = defaultFont;

  for (const run of richText) {
    const runSize = run.font?.size ?? defaultFont?.size ?? DEFAULT_FONT_SIZE;
    if (runSize > maxFontSize) {
      maxFontSize = runSize;
      maxFontForHeight = run.font ? { ...defaultFont, ...run.font } : defaultFont;
    }
  }

  // Concatenate all text for line counting
  const fullText = richText.map(r => r.text).join("");

  return calculateAutoFitHeight(fullText, maxFontForHeight, alignment, columnWidthPx);
}

/**
 * Get the content area width of a column in pixels (excluding padding).
 *
 * @param charWidth - Column width in Excel character units
 * @param mdw - Max digit width
 * @returns Content width in pixels
 */
export function getColumnContentWidthPx(charWidth: number, mdw: number): number {
  const totalPx = charWidthToPixel(charWidth, mdw);
  const pp = getPixelPadding(mdw);
  return Math.max(0, totalPx - pp);
}

// =============================================================================
// Cell-Level Measurement Helpers
// =============================================================================

/**
 * Minimal cell shape used by cell-level measurement helpers.
 * Avoids importing the full `Cell` class to prevent circular dependencies.
 */
export interface MeasurableCell {
  readonly value: unknown;
  readonly numFmt: string | NumFmt | undefined;
  readonly text: string;
  readonly effectiveType: ValueType;
  readonly font: Partial<Font> | undefined;
  readonly alignment: Partial<Alignment> | undefined;
}

/**
 * Get the pixel width of a cell's display text.
 *
 * Handles all cell value types: string, number (formatted), date (formatted),
 * boolean, formula result, rich text, hyperlink, error.
 */
export function getCellTextWidthPx(cell: MeasurableCell): number {
  const cellType = cell.effectiveType;
  const font = cell.font;

  // Rich text: measure per-run with individual fonts
  if (cellType === ValueType.RichText) {
    const value = cell.value;
    if (value && typeof value === "object" && "richText" in value) {
      return measureRichTextWidthPx((value as { richText: RichText[] }).richText, font);
    }
  }

  // Get the display text (applies number formatting)
  const displayText = getCellDisplayText(cell);
  if (!displayText) {
    return 0;
  }

  return measureTextWidthPx(displayText, font);
}

/**
 * Get the height in points a cell needs.
 *
 * Considers wrapText alignment, indent, and explicit newlines.
 *
 * @param cell           - The cell to measure
 * @param mdw            - Max digit width in pixels
 * @param columnWidthPx  - Column content width in pixels (needed for wrapText cells)
 */
export function getCellHeightPt(cell: MeasurableCell, mdw: number, columnWidthPx?: number): number {
  const font = cell.font;
  const alignment = cell.alignment;
  const cellType = cell.effectiveType;

  // Rich text
  if (cellType === ValueType.RichText) {
    const value = cell.value;
    if (value && typeof value === "object" && "richText" in value) {
      return calculateRichTextAutoFitHeight(
        (value as { richText: RichText[] }).richText,
        font,
        alignment,
        columnWidthPx
      );
    }
  }

  const displayText = getCellDisplayText(cell);
  if (!displayText) {
    return 0;
  }

  const effectiveColWidthPx = alignment?.wrapText ? columnWidthPx : undefined;

  return calculateAutoFitHeight(displayText, font, alignment, effectiveColWidthPx);
}
