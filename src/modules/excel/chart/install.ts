/**
 * Opt-in chart-support installer.
 *
 * This file carries the static imports that wire the concrete chart
 * implementation (cache populator, chart builder, ChartEx builder,
 * preset engine, sidecar serialisers, `Chart` class and chart-api
 * helpers) into the excel host. Keep these imports isolated in a
 * separate module so callers who import only the types, the SVG/PNG
 * renderers, or the parser from `./index` do not trigger the entire
 * builder pipeline at bundle time.
 *
 * A single process-wide slot is populated:
 *
 * - The chart-support slot (`@excel/chart-host-registry`) — enables
 *   `worksheet.addChart()`, `worksheet.addLineChart()`, chart cache
 *   population during `Workbook.writeXlsx(workbook)`, and `Chart`
 *   reconstruction during XLSX load.
 *
 * The slot accepts `null` to uninstall. See
 * {@link uninstallChartSupport} for tests that need to exercise the
 * "no chart support" error path.
 */

import { registerChartSupport, type ChartSupport } from "../chart-host-registry";
import { ChartSpaceXform } from "../xlsx/xform/chart/chart-space-xform";
import { fillChartCaches, fillChartExCaches } from "./cache-populator";
import { Chart } from "./chart";
import {
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  seriesFromColumns
} from "./chart-api";
import { buildChartModel, buildComboChartModel } from "./chart-builder";
import { buildChartExModel } from "./chart-ex-builder";
import { parseChartEx } from "./chart-ex-parser";
import {
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  renderChartEx,
  renderChartExLegendXml,
  renderChartExPng,
  rewriteChartExDataRefsToDefinedNames
} from "./chart-ex-renderer";
import { resolvePendingChartImages } from "./chart-images";
import { applyChartExPreset, applyChartPreset } from "./chart-presets";
import { drawChartPdf, renderChartPng } from "./chart-renderer";
import { buildChartColors, buildChartStyle } from "./chart-sidecar";
import { themeIndexToName } from "./chart-utils";

/**
 * Install the chart runtime into the excel host.
 *
 * After calling this once, chart-manipulating APIs on `Worksheet`,
 * `Workbook` and `Chartsheet` run the full chart builder and renderer
 * pipeline instead of throwing. Subsequently-loaded workbooks also
 * reconstruct `Chart` instances for their existing charts so consumer
 * code can inspect / mutate them.
 *
 * Safe to call more than once — the registry accepts the last
 * registration. Calling is idempotent and cheap; no shared state is
 * mutated beyond the host-registry slot.
 */
export function installChartSupport(): void {
  const support: ChartSupport = {
    buildChartModel,
    buildComboChartModel,
    buildChartExModel,
    applyChartPreset,
    applyChartExPreset,
    fillChartCaches: fillChartCaches as ChartSupport["fillChartCaches"],
    fillChartExCaches: fillChartExCaches as ChartSupport["fillChartExCaches"],
    resolvePendingChartImages:
      resolvePendingChartImages as ChartSupport["resolvePendingChartImages"],
    buildChartStyle,
    buildChartColors,
    createChart: (worksheet, ids, range) => new Chart(worksheet, ids, range),
    chartOptionsFromTable: chartOptionsFromTable as ChartSupport["chartOptionsFromTable"],
    chartOptionsFromRows: chartOptionsFromRows as ChartSupport["chartOptionsFromRows"],
    chartExOptionsFromTable: chartExOptionsFromTable as ChartSupport["chartExOptionsFromTable"],
    chartExOptionsFromRows: chartExOptionsFromRows as ChartSupport["chartExOptionsFromRows"],
    seriesFromColumns: seriesFromColumns as ChartSupport["seriesFromColumns"],
    parseChartEx,
    renderChartEx,
    renderChartExLegendXml,
    rewriteChartExDataRefsToDefinedNames,
    themeIndexToName,
    createChartSpaceXform: () => new ChartSpaceXform(),
    drawChartPdf,
    drawChartExPdf,
    renderChartPng,
    renderChartExPng,
    canRenderChartExAsVectorPdf
  };
  registerChartSupport(support);
}

/**
 * Uninstall the chart runtime, restoring the cold-start state where
 * chart-manipulating APIs throw. Mainly useful for tests that exercise
 * the "no chart support" error path. In production, calling this is
 * rarely necessary — subsequent `installChartSupport()` calls simply
 * overwrite the previous registration.
 */
export function uninstallChartSupport(): void {
  registerChartSupport(null);
}
