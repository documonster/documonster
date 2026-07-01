/**
 * Shared rendering constants used by both the layout engine and page renderer.
 *
 * Keeping these in one place ensures row-height computation and text rendering
 * use exactly the same values, preventing clipped or overlapping content.
 */

/** Horizontal cell padding in points (left + right = 2 × CELL_PADDING_H). */
export const CELL_PADDING_H = 3;

/** Vertical cell padding in points (top + bottom = 2 × CELL_PADDING_V). */
export const CELL_PADDING_V = 2;

/**
 * Line-height multiplier applied to the font size.
 *
 * Excel's default row height for an 11pt font is 15pt, which after removing
 * vertical padding (2 × 2 = 4pt) leaves 11pt × 1.0 — but Excel also adds
 * internal leading. A factor of 1.2 matches standard PDF/typographic practice
 * and keeps text readable without inflating row heights.
 */
export const LINE_HEIGHT_FACTOR = 1.2;

/** Width of one indent level in points (~3 characters at 11pt). */
export const INDENT_WIDTH = 10;

/**
 * Excel column widths are measured in characters of the default font's digit width.
 * For Calibri 11pt (the default), maxDigitWidth ≈ 7 pixels at 96 DPI. The
 * per-column pixel padding (4px text margin + 1px gridline) is derived inside
 * `charWidthToPixel` (`@utils/units`); see that helper for the exact formula.
 * To convert the resulting pixels to PDF points multiply by `PX_TO_PT` (72/96).
 */
export const MAX_DIGIT_WIDTH_PX = 7;
export const PX_TO_PT = 72 / 96; // 0.75
