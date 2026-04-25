/**
 * DOCX Module - Unit Conversion Utilities
 *
 * Conversion functions between common measurement units used in OOXML.
 */

// =============================================================================
// Twips conversions (1 inch = 1440 twips, 1 pt = 20 twips, 1 cm = 567 twips)
// =============================================================================

/** Convert inches to twips. */
export function inchesToTwips(inches: number): number {
  return Math.round(inches * 1440);
}

/** Convert twips to inches. */
export function twipsToInches(twips: number): number {
  return twips / 1440;
}

/** Convert centimeters to twips. */
export function cmToTwips(cm: number): number {
  return Math.round(cm * 567);
}

/** Convert twips to centimeters. */
export function twipsToCm(twips: number): number {
  return twips / 567;
}

/** Convert points to twips. */
export function ptToTwips(pt: number): number {
  return Math.round(pt * 20);
}

/** Convert twips to points. */
export function twipsToPt(twips: number): number {
  return twips / 20;
}

/** Convert millimeters to twips. */
export function mmToTwips(mm: number): number {
  return Math.round(mm * 56.7);
}

// =============================================================================
// EMU conversions (1 inch = 914400 EMU, 1 cm = 360000 EMU, 1 pt = 12700 EMU)
// =============================================================================

/** Convert inches to EMU. */
export function inchesToEmu(inches: number): number {
  return Math.round(inches * 914400);
}

/** Convert EMU to inches. */
export function emuToInches(emu: number): number {
  return emu / 914400;
}

/** Convert centimeters to EMU. */
export function cmToEmu(cm: number): number {
  return Math.round(cm * 360000);
}

/** Convert EMU to centimeters. */
export function emuToCm(emu: number): number {
  return emu / 360000;
}

/** Convert points to EMU. */
export function ptToEmu(pt: number): number {
  return Math.round(pt * 12700);
}

/** Convert pixels (at 96 DPI) to EMU. */
export function pxToEmu(px: number): number {
  return Math.round(px * 9525);
}

/** Convert EMU to pixels (at 96 DPI). */
export function emuToPx(emu: number): number {
  return emu / 9525;
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
