/**
 * What the XLSX writer loads, behind an `import()`, when a sheet has a formula a caller set — see `writeXlsxPackage`.
 *
 * A module of its own rather than an `import()` of `utils/array-evaluation` directly: that module is also imported
 * statically by the streaming and XLSB writers, and as a dynamic entry too it became a chunk of its own that rolldown
 * put its runtime helpers in — so every bundle importing anything with a namespace export loaded the formula parser
 * with them (`verify:treeshake`, `browser /word: Package`).
 */
export { needsArrayEvaluation, unwritableFormulaReason } from "@excel/utils/array-evaluation";
