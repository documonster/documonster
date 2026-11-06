/**
 * What the chart part's writer refuses, checked while the body is written.
 *
 * Apart from `chart-writer.ts` because the body writer runs for every document and the chart part's writer only for one
 * with a chart: a bundler splits chunks by module, so while the check lived beside the renderer, writing a paragraph
 * loaded the whole chart serialiser.
 */
import { unrecognizedValue } from "@utils/unrecognized-values";
import { MODEL_CHECK, WRITTEN, checkOneOf, writtenToken } from "@word/core/argument-checks";
import { drawingInteger } from "@word/core/drawing-geometry";
import type { Chart } from "@word/types";

// -- Data labels position map ---------------------------------------------------

export const DATA_LABEL_POSITION_MAP: Record<string, string> = {
  outsideEnd: "outEnd",
  center: "ctr",
  insideEnd: "inEnd",
  bestFit: "bestFit"
};
export const DATA_LABEL_POSITIONS: ReadonlySet<string> = new Set(
  Object.keys(DATA_LABEL_POSITION_MAP)
);

/**
 * The checks the chart part's writer makes, run while the body is written —
 * where the chart's place in the document is known — so a refused value is
 * named by its path there. Fields are relative to the `Chart`.
 */
export function checkWritableChart(chart: Chart): void {
  if (chart.legend !== "none") {
    writtenToken(chart, "legend", WRITTEN.legendPosition, "legend");
  }
  const dataLabels = chart.dataLabels;
  if (
    dataLabels &&
    unrecognizedValue(dataLabels, "position") === undefined &&
    dataLabels.position
  ) {
    checkOneOf(dataLabels.position, DATA_LABEL_POSITIONS, "dataLabels.position", MODEL_CHECK);
  }
  drawingInteger(chart.width, "width", 0);
  drawingInteger(chart.height, "height", 0);
}
