/**
 * `Chart` namespace surface — chart creation / management on a worksheet.
 *
 * `import { Chart } from "documonster/excel"` → `Chart.addColumn(ws, opts, "D1:K20")`,
 * `Chart.add(ws, opts, range)`, `Chart.get(ws)`, `Chart.remove(ws, chart)`.
 *
 * A consumer that never imports `Chart` gets the entire chart implementation
 * tree-shaken out of the bundle.
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
