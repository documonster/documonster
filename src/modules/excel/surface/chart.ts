/**
 * `Chart` namespace surface — chart creation / management on a worksheet,
 * plus operations on a chart handle.
 *
 * `import { Chart } from "documonster/excel"` →
 *   `Chart.add(ws, opts, range)`, `Chart.addBar(ws, opts, "D1:K20")`,
 *   `Chart.get(ws)`, `Chart.remove(ws, chart)` — creation / management.
 *   `Chart.toSVG(chart)`, `Chart.mutate(chart, fn)`, `Chart.setStyle(chart, n)`,
 *   `Chart.getSeries(chart, i)`, … — operations on a chart handle.
 *
 * Each member is independently tree-shaken (`export * as` over flat named
 * exports): a consumer that never references `Chart` — or references only
 * `Chart.add` but never `Chart.toSVG` — gets the unused chart implementation
 * (renderer, mutation engine, …) dropped from the bundle.
 */
export {
  addChart as add,
  addColumnChart as addColumn,
  addBarChart as addBar,
  addLineChart as addLine,
  addAreaChart as addArea,
  addPieChart as addPie,
  addDoughnutChart as addDoughnut,
  addScatterChart as addScatter,
  addBubbleChart as addBubble,
  addRadarChart as addRadar,
  addStockChart as addStock,
  addSurfaceChart as addSurface,
  addHistogramChart as addHistogram,
  addParetoChart as addPareto,
  addWaterfallChart as addWaterfall,
  addFunnelChart as addFunnel,
  addTreemapChart as addTreemap,
  addSunburstChart as addSunburst,
  addBoxWhiskerChart as addBoxWhisker,
  addRegionMapChart as addRegionMap,
  addComboChart as addCombo,
  addChartEx as addEx,
  addPresetChart as addPreset,
  addPresetChartEx as addPresetEx,
  addChartFromTable as addFromTable,
  addChartFromRows as addFromRows,
  addColumnChartFromRows as addColumnFromRows,
  addChartExFromTable as addExFromTable,
  addChartExFromRows as addExFromRows,
  addPivotChart as addPivot,
  addPivotComboChart as addPivotCombo,
  seriesFromColumns,
  getCharts as get,
  removeChart as remove
} from "@excel/worksheet";

// Chart-handle operations (formerly methods on the `Chart` class). Re-exported
// from the chart implementation under unprefixed namespace names so the public
// form reads `Chart.toSVG(chart)` consistent with `Cell.setValue(ws, …)`.
export {
  chartChartModel as chartModel,
  chartChartExModel as chartExModel,
  chartIsChartEx as isChartEx,
  chartAnchorModel as anchorModel,
  chartTitle as title,
  chartSetTitle as setTitle,
  chartToSVG as toSVG,
  chartToPNG as toPNG,
  chartMutate as mutate,
  chartMutateChartEx as mutateChartEx,
  chartSetStyle as setStyle,
  chartSetBuiltInStyle as setBuiltInStyle,
  chartAddSeries as addSeries,
  chartRemoveSeries as removeSeries,
  chartGetSeries as getSeries,
  chartUpdateSeries as updateSeries,
  chartAddSeriesFromOptions as addSeriesFromOptions,
  chartSetSeriesValues as setSeriesValues,
  chartSetSeriesCategories as setSeriesCategories,
  chartSetSeriesName as setSeriesName,
  chartGetSeriesCount as getSeriesCount,
  chartTotalSeriesCount as totalSeriesCount,
  chartChartTypes as chartTypes,
  chartAxes as axes,
  chartGetAxis as getAxis,
  chartCategoryAxis as categoryAxis,
  chartValueAxis as valueAxis,
  chartPlotArea as plotArea,
  chartLegend as legend,
  chartSetLegend as setLegend,
  chartSpPr as spPr,
  chartSetSpPr as setSpPr,
  chartTitleRichText as titleRichText,
  chartSetTitleRichText as setTitleRichText,
  chartUnknownElements as unknownElements,
  chartUserShapesXml as userShapesXml,
  chartSetUserShapesXml as setUserShapesXml,
  chartRemoveUserShapes as removeUserShapes,
  chartCopyTo as copyTo,
  chartClone as clone
} from "@excel/chart/chart-handle";
