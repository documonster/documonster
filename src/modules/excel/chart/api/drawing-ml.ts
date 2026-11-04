/**
 * `DrawingMl` namespace surface — DrawingML shape properties (`spPr`) and text
 * properties (`txPr`) as they appear on chart models.
 *
 * `import { DrawingMl } from "documonster/chart"` →
 *   `DrawingMl.parseSpPr(spPr)`, `DrawingMl.getSpPrFillColor(spPr)`,
 *   `DrawingMl.setSpPrFill(spPr, fill)`, `DrawingMl.createTxPr({ size: 1200 })`, …
 *
 * One verb per job, for both families:
 * - `parse*` reads a round-tripped object (which may carry only raw XML) into structured fields;
 * - `create*` builds a fresh structured object, without the raw XML, so later edits are written;
 * - `get*` reads one property; `set*` returns a copy with one property replaced.
 */

export {
  buildSpPr as createSpPr,
  buildTxPr as createTxPr,
  getSpPrFillColor,
  getSpPrGradient,
  getSpPrLine,
  getSpPrPattern,
  getTxPrColor,
  getTxPrFontSize,
  parseSpPr,
  parseTxPr,
  setSpPrFill,
  setSpPrLine
} from "@excel/chart/shared/shape-properties";
