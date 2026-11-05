import { namedCall } from "@excel/chart/api/call";
import type {
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions,
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  SeriesFromColumnsOptions
} from "@excel/chart/build/chart-api";
import {
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable
} from "@excel/chart/build/chart-api";
/**
 * `Chart` namespace surface — chart creation / management on a worksheet,
 * plus operations on a chart handle.
 *
 * `import { Chart } from "documonster/excel"` →
 *   `Chart.place(ws, "D1:K20", options)`, `Chart.list(ws)`, `Chart.remove(ws, chart)` — creation / management.
 *
 * `Chart.place` is the one way to add a chart: the options say which (classic, combo, Office 2016, pivot), and
 * `Chart.applyPreset` / `createOptionsFromTable` / `createOptionsFromRows` build them. It returns the chart's handle,
 * which `Chart.remove` and every operation below take. The deprecated `add*` family returns the chart's
 * workbook-wide **part number** instead (`chart3.xml` → `3`); that is never a position in `Chart.list(ws)`, which is
 * what `Chart.remove`'s bare-number form takes.
 *   `Chart.toSvg(chart)`, `Chart.mutate(chart, fn)`, `Chart.setStyle(chart, n)`,
 *   `Chart.getSeries(chart, i)`, … — operations on a chart handle.
 *
 * `getSeriesCount(chart, group)` counts the series of one chart-type group;
 * `getTotalSeriesCount(chart)` counts them across every group of a combo chart.
 *
 * Each member is independently tree-shaken (`export * as` over flat named
 * exports): a consumer that never references `Chart` — or references only
 * `Chart.place` but never `Chart.toSvg` — gets the unused chart implementation
 * (renderer, mutation engine, …) dropped from the bundle. Rendering lives in a
 * separate leaf module for the same reason.
 */
import {
  chartAddSeries,
  chartClone,
  chartCopyTo,
  chartMutate,
  chartRemoveSeries,
  chartRemoveUserShapes,
  chartSetSeriesValues,
  chartSetStyle,
  chartSetTitle,
  chartSetUserShapesXml,
  chartSetBuiltInStyle,
  chartSetTitleRichText,
  chartAddSeriesFromOptions,
  chartSetLegend,
  chartSetSeriesCategories,
  chartSetSeriesName,
  chartSetSpPr,
  chartUpdateSeries,
  chartAnchorModel,
  chartAxes,
  chartIsChartEx,
  chartMutateChartEx,
  chartCategoryAxis,
  chartChartExModel,
  chartChartModel,
  chartChartTypes,
  chartLegend,
  chartPlotArea,
  chartSpPr,
  chartTitle,
  chartTitleRichText,
  chartTotalSeriesCount,
  chartUnknownElements,
  chartUserShapesXml,
  chartValueAxis
} from "@excel/chart/chart-handle";
import type { AddChartExOptions } from "@excel/chart/model/chart-ex-types";
import {
  CHART_EX_PRESETS,
  CHART_PRESETS,
  applyChartExPreset,
  applyChartPreset
} from "@excel/chart/model/chart-presets";
import type { ExcelChartExPreset, ExcelChartPreset } from "@excel/chart/model/chart-presets";
import type {
  AddChartOptions,
  AddChartRange,
  AddChartSeriesOptions
} from "@excel/chart/model/types";
import { chartToPNG, chartToSVG } from "@excel/chart/render/chart-render-ops";
import { checkChart } from "@excel/core/feature-handles";
import type { TableData } from "@excel/core/table";
import {
  addChart,
  addColumnChart,
  addBarChart,
  addLineChart,
  addAreaChart,
  addPieChart,
  addDoughnutChart,
  addScatterChart,
  addBubbleChart,
  addRadarChart,
  addStockChart,
  addSurfaceChart,
  addHistogramChart,
  addParetoChart,
  addWaterfallChart,
  addFunnelChart,
  addTreemapChart,
  addSunburstChart,
  addBoxWhiskerChart,
  addRegionMapChart,
  addComboChart,
  addChartEx,
  addPresetChart,
  addPresetChartEx,
  addChartFromTable,
  addChartExFromTable,
  addPivotChart,
  addPivotComboChart,
  placeChart,
  addChartFromRows,
  addColumnChartFromRows,
  addChartExFromRows,
  getCharts,
  isChartExType,
  removeChart,
  seriesFromColumns as seriesFromColumnsImpl
} from "@excel/core/worksheet";
import type { ChartPlacement } from "@excel/core/worksheet";
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { ChartHandle, WorksheetData } from "@excel/core/worksheet-core";
import { ExcelChartOptionsError } from "@excel/errors";
import { checkListTarget } from "@excel/surface/argument-checks";
import {
  checkChartRange,
  checkChartSeries,
  checkLegend,
  checkPlacement
} from "@excel/surface/chart-checks";
import { argumentChecks } from "@excel/utils/arguments";

/**
 * Series options for the columns of a range of the sheet — one series per value column, with the categories and the
 * names taken from the range — for {@link place} or {@link addSeries}.
 */
export function createSeriesFromColumns(
  ws: WorksheetData,
  options: SeriesFromColumnsOptions
): ReturnType<typeof seriesFromColumnsImpl> {
  checkWorksheet(ws);
  return namedCall("Chart.createSeriesFromColumns", () => seriesFromColumnsImpl(ws, options));
}

/** Remove a chart — the handle {@link place} or {@link list} returned, or its index in that list. */
export function remove(ws: WorksheetData, chart: ChartHandle | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Chart.remove", "chart", chart, "a chart");
  return removeChart(ws, chart);
}

/** The sheet's charts, in the order they were placed. */
export function list(ws: WorksheetData): ChartHandle[] {
  checkWorksheet(ws);
  return getCharts(ws);
}

// Every way of adding a chart is checked here before the chart is built: a range that is not one was written as an
// anchor of `NaN`s, which failed only when the workbook was written.
/**
 * @deprecated Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function add(...args: Parameters<typeof addChart>): ReturnType<typeof addChart> {
  checkPlacement("Chart.add", args);
  return addChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "col" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addColumn(
  ...args: Parameters<typeof addColumnChart>
): ReturnType<typeof addColumnChart> {
  checkPlacement("Chart.addColumn", args);
  return addColumnChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "bar" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addBar(...args: Parameters<typeof addBarChart>): ReturnType<typeof addBarChart> {
  checkPlacement("Chart.addBar", args);
  return addBarChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "line" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addLine(...args: Parameters<typeof addLineChart>): ReturnType<typeof addLineChart> {
  checkPlacement("Chart.addLine", args);
  return addLineChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "area" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addArea(...args: Parameters<typeof addAreaChart>): ReturnType<typeof addAreaChart> {
  checkPlacement("Chart.addArea", args);
  return addAreaChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "pie" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPie(...args: Parameters<typeof addPieChart>): ReturnType<typeof addPieChart> {
  checkPlacement("Chart.addPie", args);
  return addPieChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "doughnut" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addDoughnut(
  ...args: Parameters<typeof addDoughnutChart>
): ReturnType<typeof addDoughnutChart> {
  checkPlacement("Chart.addDoughnut", args);
  return addDoughnutChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "scatter" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addScatter(
  ...args: Parameters<typeof addScatterChart>
): ReturnType<typeof addScatterChart> {
  checkPlacement("Chart.addScatter", args);
  return addScatterChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "bubble" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addBubble(
  ...args: Parameters<typeof addBubbleChart>
): ReturnType<typeof addBubbleChart> {
  checkPlacement("Chart.addBubble", args);
  return addBubbleChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "radar" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addRadar(
  ...args: Parameters<typeof addRadarChart>
): ReturnType<typeof addRadarChart> {
  checkPlacement("Chart.addRadar", args);
  return addRadarChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "stock" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addStock(
  ...args: Parameters<typeof addStockChart>
): ReturnType<typeof addStockChart> {
  checkPlacement("Chart.addStock", args);
  return addStockChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "surface" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addSurface(
  ...args: Parameters<typeof addSurfaceChart>
): ReturnType<typeof addSurfaceChart> {
  checkPlacement("Chart.addSurface", args);
  return addSurfaceChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "histogram" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addHistogram(
  ...args: Parameters<typeof addHistogramChart>
): ReturnType<typeof addHistogramChart> {
  checkPlacement("Chart.addHistogram", args);
  return addHistogramChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "pareto" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPareto(
  ...args: Parameters<typeof addParetoChart>
): ReturnType<typeof addParetoChart> {
  checkPlacement("Chart.addPareto", args);
  return addParetoChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "waterfall" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addWaterfall(
  ...args: Parameters<typeof addWaterfallChart>
): ReturnType<typeof addWaterfallChart> {
  checkPlacement("Chart.addWaterfall", args);
  return addWaterfallChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "funnel" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addFunnel(
  ...args: Parameters<typeof addFunnelChart>
): ReturnType<typeof addFunnelChart> {
  checkPlacement("Chart.addFunnel", args);
  return addFunnelChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "treemap" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addTreemap(
  ...args: Parameters<typeof addTreemapChart>
): ReturnType<typeof addTreemapChart> {
  checkPlacement("Chart.addTreemap", args);
  return addTreemapChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "sunburst" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addSunburst(
  ...args: Parameters<typeof addSunburstChart>
): ReturnType<typeof addSunburstChart> {
  checkPlacement("Chart.addSunburst", args);
  return addSunburstChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "boxWhisker" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addBoxWhisker(
  ...args: Parameters<typeof addBoxWhiskerChart>
): ReturnType<typeof addBoxWhiskerChart> {
  checkPlacement("Chart.addBoxWhisker", args);
  return addBoxWhiskerChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, type: "regionMap" })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addRegionMap(
  ...args: Parameters<typeof addRegionMapChart>
): ReturnType<typeof addRegionMapChart> {
  checkPlacement("Chart.addRegionMap", args);
  return addRegionMapChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addCombo(
  ...args: Parameters<typeof addComboChart>
): ReturnType<typeof addComboChart> {
  checkPlacement("Chart.addCombo", args);
  return addComboChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addEx(...args: Parameters<typeof addChartEx>): ReturnType<typeof addChartEx> {
  checkPlacement("Chart.addEx", args);
  return addChartEx(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPreset(
  ...args: Parameters<typeof addPresetChart>
): ReturnType<typeof addPresetChart> {
  checkPlacement("Chart.addPreset", args);
  return addPresetChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPresetEx(
  ...args: Parameters<typeof addPresetChartEx>
): ReturnType<typeof addPresetChartEx> {
  checkPlacement("Chart.addPresetEx", args);
  return addPresetChartEx(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addFromTable(
  ...args: Parameters<typeof addChartFromTable>
): ReturnType<typeof addChartFromTable> {
  checkPlacement("Chart.addFromTable", args);
  return addChartFromTable(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addExFromTable(
  ...args: Parameters<typeof addChartExFromTable>
): ReturnType<typeof addChartExFromTable> {
  checkPlacement("Chart.addExFromTable", args);
  return addChartExFromTable(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPivot(
  ...args: Parameters<typeof addPivotChart>
): ReturnType<typeof addPivotChart> {
  checkPlacement("Chart.addPivot", args);
  return addPivotChart(...args);
}

/**
 * @deprecated Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addPivotCombo(
  ...args: Parameters<typeof addPivotComboChart>
): ReturnType<typeof addPivotComboChart> {
  checkPlacement("Chart.addPivotCombo", args);
  return addPivotComboChart(...args);
}

/**
 * Add a chart to `ws` over `range` and return its handle — the one way to add a chart.
 *
 * `options` describes any chart this library draws: a classic one (`type: "bar"`, `"line"`, `"pie"`, …), an
 * Office 2016 one (`type: "histogram"`, `"treemap"`, `"waterfall"`, …), a combo (`groups`), or a pivot chart
 * (`pivotTable`). {@link applyPreset}, {@link createOptionsFromTable} and {@link createOptionsFromRows} build it from
 * a preset, a table or plain rows.
 *
 * @throws {ExcelInvalidAddressError} when `range` is not a reference.
 * @throws {ExcelArgumentError} when `range` or a series is malformed.
 * @throws {ExcelChartOptionsError} when the chart's options are invalid.
 */
export function place(
  ws: WorksheetData,
  range: AddChartRange,
  options: ChartPlacement
): ChartHandle {
  checkWorksheet(ws);
  const check = argumentChecks("Chart.place");
  check.required("options", options);
  check.record("options", options);
  checkPlacement("Chart.place", [options, range]);
  return namedCall("Chart.place", () => placeChart(ws, options, range));
}

/** What {@link applyPreset} merges a preset into: a classic chart's options or an Office 2016 chart's. */
export type PresetOptions =
  | (Omit<AddChartOptions, "type"> & Partial<Pick<AddChartOptions, "type">>)
  | (Omit<AddChartExOptions, "type"> & Partial<Pick<AddChartExOptions, "type">>);

/** A preset's options: the chart type and grouping Excel's gallery names, merged into `options`. */
export function applyPreset(
  preset: ExcelChartPreset,
  options: Omit<AddChartOptions, "type"> & Partial<Pick<AddChartOptions, "type">>
): AddChartOptions;
export function applyPreset(
  preset: ExcelChartExPreset,
  options: Omit<AddChartExOptions, "type"> & Partial<Pick<AddChartExOptions, "type">>
): AddChartExOptions;
/** A preset chosen at run time, of either family: the options must suit both. */
export function applyPreset(
  preset: ExcelChartPreset | ExcelChartExPreset,
  options: PresetOptions
): AddChartOptions | AddChartExOptions;
export function applyPreset(
  preset: ExcelChartPreset | ExcelChartExPreset,
  options: PresetOptions
): AddChartOptions | AddChartExOptions {
  const check = argumentChecks("Chart.applyPreset");
  check.required("options", options);
  check.record("options", options);
  if (!Object.hasOwn(CHART_PRESETS, preset) && !Object.hasOwn(CHART_EX_PRESETS, preset)) {
    // Naming only the Office 2016 family, as the fallback did, sent a caller who meant a classic preset astray.
    throw new ExcelChartOptionsError(
      "preset",
      `Chart.applyPreset: unknown chart preset ${JSON.stringify(preset)}; expected a classic preset such as ` +
        `"columnClustered" or an Office 2016 one such as "histogram".`
    );
  }
  return namedCall("Chart.applyPreset", () =>
    Object.hasOwn(CHART_PRESETS, preset)
      ? applyChartPreset(preset as ExcelChartPreset, options as never)
      : applyChartExPreset(preset as ExcelChartExPreset, options as never)
  );
}

/** The Office 2016 options {@link createOptionsFromTable} takes: a region map is not drawn from a table. */
export type TableExOptions = AddChartExFromTableOptions & {
  type: Exclude<AddChartExOptions["type"], "regionMap">;
};

/** The Office 2016 options {@link createOptionsFromRows} takes. */
export type RowsExOptions<T extends object> = AddChartExFromRowsOptions<T> & {
  type: Exclude<AddChartExOptions["type"], "regionMap">;
};

/** A chart's options from a table on `ws`: its columns become the series, its first column the categories. */
export function createOptionsFromTable(
  ws: WorksheetData,
  table: string | TableData,
  options: AddChartFromTableOptions
): AddChartOptions;
export function createOptionsFromTable(
  ws: WorksheetData,
  table: string | TableData,
  options: TableExOptions
): AddChartExOptions;
export function createOptionsFromTable(
  ws: WorksheetData,
  table: string | TableData,
  options: AddChartFromTableOptions | TableExOptions
): AddChartOptions | AddChartExOptions;
export function createOptionsFromTable(
  ws: WorksheetData,
  table: string | TableData,
  options: AddChartFromTableOptions | TableExOptions
): AddChartOptions | AddChartExOptions {
  checkWorksheet(ws);
  const check = argumentChecks("Chart.createOptionsFromTable");
  check.required("options", options);
  check.record("options", options);
  return namedCall("Chart.createOptionsFromTable", () =>
    isChartExType(options.type)
      ? chartExOptionsFromTable(ws, table, options as never)
      : chartOptionsFromTable(ws, table, options as AddChartFromTableOptions)
  );
}

/**
 * A chart's options from plain rows, which are written to `ws` for the chart to refer to — from `startCell` (`A1`
 * by default), a header row first unless `includeHeaders` is `false`, over whatever those cells held.
 *
 * Everything is checked before the first cell is written, so a call that throws leaves the sheet as it was. Writing
 * the rows and adding the chart are separate steps: when `Chart.place` then refuses the options, the rows stay.
 */
export function createOptionsFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T>
): AddChartOptions;
export function createOptionsFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: RowsExOptions<T>
): AddChartExOptions;
export function createOptionsFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T> | RowsExOptions<T>
): AddChartOptions | AddChartExOptions;
export function createOptionsFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T> | RowsExOptions<T>
): AddChartOptions | AddChartExOptions {
  checkWorksheet(ws);
  const check = argumentChecks("Chart.createOptionsFromRows");
  check.required("options", options);
  check.record("options", options);
  check.required("rows", rows);
  check.array("rows", rows);
  return namedCall("Chart.createOptionsFromRows", () =>
    isChartExType(options.type)
      ? chartExOptionsFromRows(ws, rows, options as never)
      : chartOptionsFromRows(ws, rows, options as AddChartFromRowsOptions<T>)
  );
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T>,
  range: AddChartRange
): number {
  checkWorksheet(ws);
  checkPlacement("Chart.addFromRows", [ws, rows, options, range]);
  return addChartFromRows(ws, rows, options, range);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, { ...options, type: "bar", barDir: "col" }))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addColumnFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: Omit<AddChartFromRowsOptions<T>, "type" | "barDir">,
  range: AddChartRange
): number {
  checkWorksheet(ws);
  checkPlacement("Chart.addColumnFromRows", [ws, rows, options, range]);
  return addColumnChartFromRows(ws, rows, options, range);
}

/**
 * @deprecated Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this
 * returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.
 */
export function addExFromRows<T extends object>(
  ws: WorksheetData,
  rows: readonly T[],
  options: AddChartExFromRowsOptions<T> & { type: Exclude<AddChartExOptions["type"], "regionMap"> },
  range: AddChartRange
): number {
  checkWorksheet(ws);
  checkPlacement("Chart.addExFromRows", [ws, rows, options, range]);
  return addChartExFromRows(ws, rows, options, range);
}

export type { ChartPlacement as Placement } from "@excel/core/worksheet";

export {
  chartGetSeries as getSeries,
  chartGetSeriesCount as getSeriesCount,
  chartGetAxis as getAxis
} from "@excel/chart/chart-handle";

// The editors run through `namedCall`, so an argument error raised by the shared code they reach starts with the call
// the caller wrote (`Chart.mutate: …`), not with whichever internal function found it.

/** Set or clear (`undefined`) the chart's title: text, rich text, or `{ formula }` naming the cell that holds it. */
export function setTitle(c: ChartHandle, value: Parameters<typeof chartSetTitle>[1]): void {
  namedCall("Chart.setTitle", () => chartSetTitle(c, value));
}

/**
 * Edit a classic chart's model through `mutator`, which receives a copy; the edited copy is applied when it returns.
 * Returns the handle.
 */
export function mutate(
  c: ChartHandle,
  mutator: Parameters<typeof chartMutate>[1],
  options?: Parameters<typeof chartMutate>[2]
): ChartHandle {
  checkChart(c);
  checkMutation("Chart.mutate", mutator, options);
  return namedCall("Chart.mutate", () => chartMutate(c, mutator, options));
}

/** {@link mutate} for an Office 2016 (ChartEx) chart's model. */
export function mutateEx(
  c: ChartHandle,
  mutator: Parameters<typeof chartMutateChartEx>[1],
  options?: Parameters<typeof chartMutateChartEx>[2]
): ChartHandle {
  checkChart(c);
  checkMutation("Chart.mutateEx", mutator, options);
  return namedCall("Chart.mutateEx", () => chartMutateChartEx(c, mutator, options));
}

function checkMutation(where: string, mutator: unknown, options: unknown): void {
  const check = argumentChecks(where);
  check.required("mutator", mutator);
  if (typeof mutator !== "function") {
    // Called as it was given, it failed inside with "mutator is not a function", a `TypeError`.
    check.fail("mutator", "a function that edits the model", mutator);
  }
  if (check.record("options", options)) {
    check.boolean("options.preferRawPatch", options.preferRawPatch);
    check.boolean("options.requireRawPatch", options.requireRawPatch);
  }
}

/** Apply one of Excel's 48 built-in chart styles. Returns the handle. */
export function setStyle(c: ChartHandle, style: number): ChartHandle {
  return namedCall("Chart.setStyle", () => chartSetStyle(c, style));
}

/** Add a series to the chart-type group `groupIndex` (default the first). */
export function addSeries(
  c: ChartHandle,
  series: Parameters<typeof chartAddSeries>[1],
  groupIndex?: number
): void {
  namedCall("Chart.addSeries", () => chartAddSeries(c, series, groupIndex));
}

/** Remove series `index` of the chart-type group `groupIndex` (default the first), and return it. */
export function removeSeries(
  c: ChartHandle,
  index: number,
  groupIndex?: number
): ReturnType<typeof chartRemoveSeries> {
  return namedCall("Chart.removeSeries", () => chartRemoveSeries(c, index, groupIndex));
}

/** Point series `index` at new values (a formula such as `"Sheet1!$B$2:$B$9"`). */
export function setSeriesValues(
  c: ChartHandle,
  index: number,
  values: string,
  groupIndex?: number
): boolean {
  return namedCall("Chart.setSeriesValues", () =>
    chartSetSeriesValues(c, index, values, groupIndex)
  );
}

/** Set or clear (`undefined`) the chart's user-shapes overlay (`c:userShapes` drawing XML). */
export function setUserShapesXml(c: ChartHandle, xml: Uint8Array | string | undefined): void {
  namedCall("Chart.setUserShapesXml", () => chartSetUserShapesXml(c, xml));
}

/** Remove the chart's user-shapes overlay. */
export function removeUserShapes(c: ChartHandle): void {
  namedCall("Chart.removeUserShapes", () => chartRemoveUserShapes(c));
}

/**
 * Copy the chart onto `targetWs` at `range` (by default beside where it sits), and return the copy's part number; see
 * {@link list} for its handle. Pass the chart's own worksheet (`c.worksheet`) to copy it beside itself.
 */
export function copyTo(c: ChartHandle, targetWs: WorksheetData, range?: AddChartRange): number {
  if (range !== undefined) {
    checkChartRange("Chart.copyTo", range);
  }
  return namedCall("Chart.copyTo", () => chartCopyTo(c, targetWs, range));
}

/**
 * Copy the chart onto its own sheet at `range` (by default beside it), and return the copy's part number.
 *
 * @deprecated Use {@link copyTo} with the chart's own worksheet — `Chart.copyTo(c, c.worksheet, range)`. It places the
 * same copy at the same default position; this name suggested an independent copy, as `Document.clone` returns, but it
 * adds a chart to the sheet and returns its part number. Removed in 1.0.
 */
export const clone = cloneChart;

function cloneChart(c: ChartHandle, range?: AddChartRange): number {
  if (range !== undefined) {
    // A range that is not one failed inside with a `TypeError`.
    checkChartRange("Chart.clone", range);
  }
  return namedCall("Chart.clone", () => chartClone(c, range));
}

/** Update common fields on an existing classic chart series. The series' name and categories are checked first. */
export function updateSeries(
  c: ChartHandle,
  index: number,
  options: Partial<AddChartSeriesOptions>,
  groupIndex = 0
): boolean {
  checkChart(c);
  checkChartSeries("Chart.updateSeries", "options", options);
  return namedCall("Chart.updateSeries", () => chartUpdateSeries(c, index, options, groupIndex));
}

/** Add a series from high-level options, matching the target chart type group. */
export function addSeriesFromOptions(
  c: ChartHandle,
  options: AddChartSeriesOptions,
  groupIndex = 0
): boolean {
  checkChart(c);
  checkChartSeries("Chart.addSeriesFromOptions", "options", options);
  return namedCall("Chart.addSeriesFromOptions", () =>
    chartAddSeriesFromOptions(c, options, groupIndex)
  );
}

/** Point a series at new categories. A value that is no reference was written as an empty `<c:cat/>`. */
export function setSeriesCategories(
  c: ChartHandle,
  index: number,
  categories: string,
  groupIndex = 0
): boolean {
  checkChart(c);
  checkChartSeries("Chart.setSeriesCategories", "", { categories });
  return namedCall("Chart.setSeriesCategories", () =>
    chartSetSeriesCategories(c, index, categories, groupIndex)
  );
}

/** Rename a series: text, or `{ formula }` naming the cell that holds it. */
export function setSeriesName(
  c: ChartHandle,
  index: number,
  name: NonNullable<AddChartSeriesOptions["name"]>,
  groupIndex = 0
): boolean {
  checkChart(c);
  argumentChecks("Chart.setSeriesName").required("name", name);
  checkChartSeries("Chart.setSeriesName", "", { name });
  return namedCall("Chart.setSeriesName", () => chartSetSeriesName(c, index, name, groupIndex));
}

/** Set or clear (`undefined`) the chart's legend. */
export function setLegend(c: ChartHandle, value: Parameters<typeof chartSetLegend>[1]): void {
  checkChart(c);
  checkLegend("Chart.setLegend", "value", value);
  namedCall("Chart.setLegend", () => chartSetLegend(c, value));
}

/** Set or clear (`undefined`) the chart area's shape properties. */
export function setSpPr(c: ChartHandle, value: Parameters<typeof chartSetSpPr>[1]): void {
  checkChart(c);
  argumentChecks("Chart.setSpPr").record("value", value);
  namedCall("Chart.setSpPr", () => chartSetSpPr(c, value));
}

export {
  chartIsChartEx as isEx,
  chartChartModel as getModel,
  chartChartExModel as getExModel,
  chartAnchorModel as getAnchorModel,
  chartTitle as getTitle,
  chartTotalSeriesCount as getTotalSeriesCount,
  chartChartTypes as getTypes,
  chartAxes as getAxes,
  chartCategoryAxis as getCategoryAxis,
  chartValueAxis as getValueAxis,
  chartPlotArea as getPlotArea,
  chartLegend as getLegend,
  chartSpPr as getSpPr,
  chartTitleRichText as getTitleRichText,
  chartUnknownElements as getUnknownElements,
  chartUserShapesXml as getUserShapesXml,
  chartToSVG as toSvg,
  chartToPNG as toPng
};

// --- deprecated names ---

/** @deprecated Use {@link setStyle}, which this forwards to. Removed in 1.0. */
export const setBuiltInStyle = chartSetBuiltInStyle;
/** @deprecated Use {@link setTitle}, which takes rich text too. Removed in 1.0. */
export const setTitleRichText = chartSetTitleRichText;

/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const chartModel = chartChartModel;
/** @deprecated Use {@link getExModel}. Removed in 1.0. */
export const chartExModel = chartChartExModel;
/** @deprecated Use {@link isEx}. Removed in 1.0. */
export const isChartEx = chartIsChartEx;
/** @deprecated Use {@link mutateEx}. Removed in 1.0. */
export const mutateChartEx = mutateEx;
/** @deprecated Use {@link createSeriesFromColumns}. Removed in 1.0. */
export const seriesFromColumns = createSeriesFromColumns;
/** @deprecated Use {@link getAnchorModel}. Removed in 1.0. */
export const anchorModel = chartAnchorModel;
/** @deprecated Use {@link getTitle}. Removed in 1.0. */
export const title = chartTitle;
/** @deprecated Use {@link getTotalSeriesCount}. Removed in 1.0. */
export const totalSeriesCount = chartTotalSeriesCount;
/** @deprecated Use {@link getTypes}. Removed in 1.0. */
export const chartTypes = chartChartTypes;
/** @deprecated Use {@link getAxes}. Removed in 1.0. */
export const axes = chartAxes;
/** @deprecated Use {@link getCategoryAxis}. Removed in 1.0. */
export const categoryAxis = chartCategoryAxis;
/** @deprecated Use {@link getValueAxis}. Removed in 1.0. */
export const valueAxis = chartValueAxis;
/** @deprecated Use {@link getPlotArea}. Removed in 1.0. */
export const plotArea = chartPlotArea;
/** @deprecated Use {@link getLegend}. Removed in 1.0. */
export const legend = chartLegend;
/** @deprecated Use {@link getSpPr}. Removed in 1.0. */
export const spPr = chartSpPr;
/** @deprecated Use {@link getTitleRichText}. Removed in 1.0. */
export const titleRichText = chartTitleRichText;
/** @deprecated Use {@link getUnknownElements}. Removed in 1.0. */
export const unknownElements = chartUnknownElements;
/** @deprecated Use {@link getUserShapesXml}. Removed in 1.0. */
export const userShapesXml = chartUserShapesXml;
/** @deprecated Use {@link list}. Removed in 1.0. */
export const get = list;
/** @deprecated Use {@link toSvg}. Removed in 1.0. */
export const toSVG = chartToSVG;
/** @deprecated Use {@link toPng}. Removed in 1.0. */
export const toPNG = chartToPNG;
