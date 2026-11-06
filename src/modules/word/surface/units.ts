/**
 * `Units` namespace surface — measurement unit conversions (twips, EMU,
 * points, etc.).
 *
 * `import { Units } from "documonster/word"` →
 *   `Units.inchesToTwips(1)`, `Units.ptToEmu(12)`, … — tree-shaken via
 *   `export * as Units`.
 *
 * Each conversion is the shared one from `@utils/units`, with its argument
 * checked: a `NaN` or `Infinity` converts to `NaN`, which reached the document
 * unnoticed and was refused only when it was written, naming an XML attribute
 * rather than the call. A finite argument gives exactly the shared result.
 */
import { checkFinite } from "@word/core/argument-checks";
import * as U from "@word/units";

function convert(where: string, field: string, value: number, fn: (v: number) => number): number {
  checkFinite(value, field, `Units.${where}`);
  return fn(value);
}

export function inchesToTwips(inches: number): number {
  return convert("inchesToTwips", "inches", inches, U.inchesToTwips);
}
export function twipsToInches(twips: number): number {
  return convert("twipsToInches", "twips", twips, U.twipsToInches);
}
export function cmToTwips(cm: number): number {
  return convert("cmToTwips", "cm", cm, U.cmToTwips);
}
export function twipsToCm(twips: number): number {
  return convert("twipsToCm", "twips", twips, U.twipsToCm);
}
export function ptToTwips(pt: number): number {
  return convert("ptToTwips", "pt", pt, U.ptToTwips);
}
export function twipsToPt(twips: number): number {
  return convert("twipsToPt", "twips", twips, U.twipsToPt);
}
export function mmToTwips(mm: number): number {
  return convert("mmToTwips", "mm", mm, U.mmToTwips);
}
export function inchesToEmu(inches: number): number {
  return convert("inchesToEmu", "inches", inches, U.inchesToEmu);
}
export function emuToInches(emu: number): number {
  return convert("emuToInches", "emu", emu, U.emuToInches);
}
export function cmToEmu(cm: number): number {
  return convert("cmToEmu", "cm", cm, U.cmToEmu);
}
export function emuToCm(emu: number): number {
  return convert("emuToCm", "emu", emu, U.emuToCm);
}
export function ptToEmu(pt: number): number {
  return convert("ptToEmu", "pt", pt, U.ptToEmu);
}
export function pxToEmu(px: number): number {
  return convert("pxToEmu", "px", px, U.pxToEmu);
}
export function emuToPx(emu: number): number {
  return convert("emuToPx", "emu", emu, U.emuToPx);
}
export function ptToHalfPoint(pt: number): number {
  return convert("ptToHalfPoint", "pt", pt, U.ptToHalfPoint);
}
export function halfPointToPt(hp: number): number {
  return convert("halfPointToPt", "hp", hp, U.halfPointToPt);
}
export function ptToEighthPoint(pt: number): number {
  return convert("ptToEighthPoint", "pt", pt, U.ptToEighthPoint);
}
export function eighthPointToPt(ep: number): number {
  return convert("eighthPointToPt", "ep", ep, U.eighthPointToPt);
}
export function lineMultiplierToSpacing(multiplier: number): number {
  return convert("lineMultiplierToSpacing", "multiplier", multiplier, U.lineMultiplierToSpacing);
}
export function spacingToLineMultiplier(spacing: number): number {
  return convert("spacingToLineMultiplier", "spacing", spacing, U.spacingToLineMultiplier);
}
export function percentToTablePct(percent: number): number {
  return convert("percentToTablePct", "percent", percent, U.percentToTablePct);
}
export function tablePctToPercent(pct: number): number {
  return convert("tablePctToPercent", "pct", pct, U.tablePctToPercent);
}
