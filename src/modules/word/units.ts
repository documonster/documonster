/**
 * DOCX Module - Unit Conversion Utilities
 *
 * Conversion functions between common measurement units used in OOXML.
 */

// =============================================================================
// Base unit constants
// =============================================================================

/** Twips per inch (OOXML page measurements). */
export const TWIPS_PER_INCH = 1440;
/** Twips per point. */
export const TWIPS_PER_POINT = 20;
/** Twips per centimeter (approximate). */
export const TWIPS_PER_CM = 567;
/** Twips per millimeter (approximate). */
export const TWIPS_PER_MM = 56.7;

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
