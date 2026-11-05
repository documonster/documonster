/**
 * `DrawingMl` namespace surface — DrawingML shape properties (`spPr`) and text
 * properties (`txPr`) as they appear on chart models.
 *
 * `import { DrawingMl } from "documonster/chart"` →
 *   `DrawingMl.parseSpPr(spPr)`, `DrawingMl.getSpPrFillColor(spPr)`,
 *   `DrawingMl.withSpPrFill(spPr, fill)`, `DrawingMl.createTxPr({ size: 1200 })`, …
 *
 * One verb per job, for both families:
 * - `parse*` reads a round-tripped object (which may carry only raw XML) into structured fields;
 * - `create*` builds a fresh structured object, without the raw XML, so later edits are written;
 * - `get*` reads one property; `with*` returns a copy with one property replaced, leaving its input as it was.
 */

import { checkObjectArgument } from "@excel/chart/api/call";
import type {
  ChartColor,
  ChartFill,
  ChartLine,
  ChartTextProperties,
  ShapeProperties
} from "@excel/chart/model/types";
import {
  buildSpPr,
  buildTxPr,
  getSpPrFillColor as getSpPrFillColorImpl,
  getSpPrGradient as getSpPrGradientImpl,
  getSpPrLine as getSpPrLineImpl,
  getSpPrPattern as getSpPrPatternImpl,
  getTxPrColor as getTxPrColorImpl,
  getTxPrFontSize as getTxPrFontSizeImpl
} from "@excel/chart/shared/shape-properties";

export {
  parseSpPr,
  parseTxPr,
  withSpPrFill,
  withSpPrLine
} from "@excel/chart/shared/shape-properties";

/**
 * A structured shape-properties object from `props`, keeping every field `parseSpPr` produces and none of the raw
 * XML, so later edits to it are written.
 */
export function createSpPr(props: ShapeProperties): ShapeProperties {
  checkObjectArgument("DrawingMl.createSpPr", "props", props);
  return buildSpPr(props);
}

/** A structured text-properties object from `props`, without the raw XML, so later edits to it are written. */
export function createTxPr(props: ChartTextProperties): ChartTextProperties {
  checkObjectArgument("DrawingMl.createTxPr", "props", props);
  return buildTxPr(props);
}

/** The solid fill colour of `spPr`, raw or structured. */
export function getSpPrFillColor(spPr: ShapeProperties): ChartColor | undefined {
  checkObjectArgument("DrawingMl.getSpPrFillColor", "spPr", spPr);
  return getSpPrFillColorImpl(spPr);
}

/** The gradient fill of `spPr`, raw or structured. */
export function getSpPrGradient(spPr: ShapeProperties): ChartFill["gradient"] | undefined {
  checkObjectArgument("DrawingMl.getSpPrGradient", "spPr", spPr);
  return getSpPrGradientImpl(spPr);
}

/** The line (outline) of `spPr`, raw or structured. */
export function getSpPrLine(spPr: ShapeProperties): ChartLine | undefined {
  checkObjectArgument("DrawingMl.getSpPrLine", "spPr", spPr);
  return getSpPrLineImpl(spPr);
}

/** The pattern fill of `spPr`, raw or structured. */
export function getSpPrPattern(spPr: ShapeProperties): ChartFill["pattern"] | undefined {
  checkObjectArgument("DrawingMl.getSpPrPattern", "spPr", spPr);
  return getSpPrPatternImpl(spPr);
}

/** The font colour of `txPr`, raw or structured. */
export function getTxPrColor(txPr: ChartTextProperties): ChartColor | undefined {
  checkObjectArgument("DrawingMl.getTxPrColor", "txPr", txPr);
  return getTxPrColorImpl(txPr);
}

/** The font size of `txPr` in points (`10`, not the `1000` the file stores), raw or structured. */
export function getTxPrFontSize(txPr: ChartTextProperties): number | undefined {
  checkObjectArgument("DrawingMl.getTxPrFontSize", "txPr", txPr);
  return getTxPrFontSizeImpl(txPr);
}
