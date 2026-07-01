/**
 * `Units` namespace surface — measurement unit conversions (twips, EMU,
 * points, etc.).
 *
 * `import { Units } from "documonster/word"` →
 *   `Units.inchesToTwips(1)`, `Units.ptToEmu(12)`, … — tree-shaken via
 *   `export * as Units`.
 */
export {
  inchesToTwips,
  twipsToInches,
  cmToTwips,
  twipsToCm,
  ptToTwips,
  twipsToPt,
  mmToTwips,
  inchesToEmu,
  emuToInches,
  cmToEmu,
  emuToCm,
  ptToEmu,
  pxToEmu,
  emuToPx,
  ptToHalfPoint,
  halfPointToPt,
  ptToEighthPoint,
  eighthPointToPt,
  lineMultiplierToSpacing,
  spacingToLineMultiplier,
  percentToTablePct,
  tablePctToPercent
} from "@word/units";
