/**
 * OOXML Measurement Unit Conversions — shared across all modules (Layer 0).
 *
 * Conversion functions between the measurement units used throughout OOXML
 * (DrawingML EMU, twips, points, pixels, inches, centimeters). These are pure
 * numeric helpers with no module dependencies, so any module — Excel, Word,
 * PDF — may import them directly via `@utils/units` instead of re-deriving the
 * same magic constants (9525, 12700, 914400, …) locally.
 */

// =============================================================================
// Base unit constants
// =============================================================================

/** Twips per inch (OOXML page measurements). */
export const TWIPS_PER_INCH = 1440;
/** Twips per point. */
export const TWIPS_PER_POINT = 20;
/**
 * Twips per centimeter, derived exactly from 1 inch = 2.54 cm = 1440 twips
 * (= 566.9291…). Using the exact factor — rather than the rounded 567 — keeps
 * metric page sizes aligned with their canonical twip values, e.g.
 * `cmToTwips(21)` → 11906 and `cmToTwips(29.7)` → 16838 (A4).
 */
export const TWIPS_PER_CM = TWIPS_PER_INCH / 2.54;
/** Twips per millimeter, derived exactly from {@link TWIPS_PER_CM}. */
export const TWIPS_PER_MM = TWIPS_PER_CM / 10;

/** EMU (English Metric Units) per inch — DrawingML coordinate space. */
export const EMU_PER_INCH = 914_400;
/** EMU per centimeter. */
export const EMU_PER_CM = 360_000;
/** EMU per point. */
export const EMU_PER_POINT = 12_700;
/** EMU per pixel at 96 DPI (CSS pixel). */
export const EMU_PER_PX = 9_525;

/** Default chart width in EMU (6 inches). */
export const DEFAULT_CHART_WIDTH_EMU = 6 * EMU_PER_INCH;
/** Default chart height in EMU (3.5 inches for c:chart). */
export const DEFAULT_CHART_HEIGHT_EMU = Math.round(3.5 * EMU_PER_INCH);
/** Default ChartEx height in EMU (4 inches). */
export const DEFAULT_CHART_EX_HEIGHT_EMU = 4 * EMU_PER_INCH;
/** Default wrap margin for inline drawings in EMU (≈0.125 inch). */
export const DEFAULT_WRAP_MARGIN_EMU = 114_300;
/** Default `wp:anchor` `relativeHeight` (z-order) when not specified. */
export const DEFAULT_RELATIVE_HEIGHT = 251_658_240;

// =============================================================================
// Twips conversions (1 inch = 1440 twips, 1 pt = 20 twips, 1 cm = 567 twips)
// =============================================================================

/** Convert inches to twips. */
export function inchesToTwips(inches: number): number {
  return Math.round(inches * TWIPS_PER_INCH);
}

/** Convert twips to inches. */
export function twipsToInches(twips: number): number {
  return twips / TWIPS_PER_INCH;
}

/** Convert centimeters to twips. */
export function cmToTwips(cm: number): number {
  return Math.round(cm * TWIPS_PER_CM);
}

/** Convert twips to centimeters. */
export function twipsToCm(twips: number): number {
  return twips / TWIPS_PER_CM;
}

/** Convert points to twips. */
export function ptToTwips(pt: number): number {
  return Math.round(pt * TWIPS_PER_POINT);
}

/** Convert twips to points. */
export function twipsToPt(twips: number): number {
  return twips / TWIPS_PER_POINT;
}

/** Convert millimeters to twips. */
export function mmToTwips(mm: number): number {
  return Math.round(mm * TWIPS_PER_MM);
}

// =============================================================================
// EMU conversions (1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU)
// =============================================================================

/** Convert inches to EMU. */
export function inchesToEmu(inches: number): number {
  return Math.round(inches * EMU_PER_INCH);
}

/** Convert EMU to inches. */
export function emuToInches(emu: number): number {
  return emu / EMU_PER_INCH;
}

/** Convert centimeters to EMU. */
export function cmToEmu(cm: number): number {
  return Math.round(cm * EMU_PER_CM);
}

/** Convert EMU to centimeters. */
export function emuToCm(emu: number): number {
  return emu / EMU_PER_CM;
}

/** Convert points to EMU. */
export function ptToEmu(pt: number): number {
  return Math.round(pt * EMU_PER_POINT);
}

/** Convert EMU to points. */
export function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}

/** Convert pixels (at 96 DPI) to EMU. */
export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/** Convert EMU to pixels (at 96 DPI). */
export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

// =============================================================================
// Half-point conversions (font sizes: 1 pt = 2 half-points)
// =============================================================================

/** Convert points to half-points (for w:sz). */
export function ptToHalfPoint(pt: number): number {
  return Math.round(pt * 2);
}

/** Convert half-points to points. */
export function halfPointToPt(hp: number): number {
  return hp / 2;
}

// =============================================================================
// Eighth-point conversions (border widths: 1 pt = 8 eighth-points)
// =============================================================================

/** Convert points to eighth-points (for border w:sz). */
export function ptToEighthPoint(pt: number): number {
  return Math.round(pt * 8);
}

/** Convert eighth-points to points. */
export function eighthPointToPt(ep: number): number {
  return ep / 8;
}

// =============================================================================
// Line spacing conversions (auto mode: 240ths of a line)
// =============================================================================

/** Convert line multiplier to line spacing value (auto mode). */
export function lineMultiplierToSpacing(multiplier: number): number {
  return Math.round(multiplier * 240);
}

/** Convert line spacing value to multiplier (auto mode). */
export function spacingToLineMultiplier(spacing: number): number {
  return spacing / 240;
}

// =============================================================================
// Table width percent (fiftieths of a percent: 5000 = 100%)
// =============================================================================

/** Convert percentage (0-100) to table width pct value. */
export function percentToTablePct(percent: number): number {
  return Math.round(percent * 50);
}

/** Convert table width pct value to percentage. */
export function tablePctToPercent(pct: number): number {
  return pct / 50;
}

// =============================================================================
// Excel column-width conversions (ECMA-376 §18.3.1.13)
// =============================================================================
//
// Excel stores column widths in "character units" relative to a font's
// Maximum Digit Width (MDW, in pixels). Converting between character units and
// pixels also involves a per-column Pixel Padding (PP). These helpers are pure
// — the caller supplies the MDW for the workbook's default font (Calibri 11pt
// → 7) — so they live here and are shared by the Excel auto-fit engine and the
// PDF layout engine, which previously each carried their own approximation.

/**
 * Calculate the Pixel Padding (PP) for a given Max Digit Width.
 *
 * Formula: `PP = 2 * CEIL(MDW / 4) + 1`.
 */
export function getPixelPadding(mdw: number): number {
  return 2 * Math.ceil(mdw / 4) + 1;
}

/**
 * Convert an Excel column character width to pixels.
 *
 * The formula differs for widths below one character:
 * - `width < 1`: `pixels = ROUND(width * (MDW + PP))`
 * - `width >= 1`: `pixels = ROUND(width * MDW) + PP`
 *
 * Returns 0 for a non-positive MDW or a zero width.
 */
export function charWidthToPixel(width: number, mdw: number): number {
  if (mdw <= 0 || width === 0) {
    return 0;
  }
  const pp = getPixelPadding(mdw);
  if (width < 1) {
    return Math.round(width * (mdw + pp));
  }
  return Math.round(width * mdw) + pp;
}

/**
 * Convert a pixel width to an Excel column character width (as stored in XLSX).
 *
 * Formula: `TRUNC(pixels / MDW * 256) / 256` — character units with 1/256
 * precision. Returns 0 for a non-positive MDW.
 */
export function pixelToCharWidth(pixels: number, mdw: number): number {
  if (mdw <= 0) {
    return 0;
  }
  return Math.trunc((pixels / mdw) * 256) / 256;
}

/** Convert a pixel length to points at 96 DPI (`points = pixels * 72 / 96`). */
export function pixelToPoints(pixels: number): number {
  return (pixels * 72) / 96;
}

/** Convert a point length to pixels at 96 DPI (`pixels = points * 96 / 72`). */
export function pointsToPixel(points: number): number {
  return (points * 96) / 72;
}
