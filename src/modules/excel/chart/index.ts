/**
 * Public entry for the excelts chart module.
 *
 * Two complementary usage styles are supported:
 *
 * 1. **Functional, zero-side-effect** — tree-shakeable access to
 *    chart builders, renderers, presets and parsers:
 *    ```ts
 *    import { buildChartModel, renderChartSvg } from "@cjnoname/excelts/chart";
 *    const model = buildChartModel(options);
 *    const svg   = renderChartSvg(model);
 *    ```
 *    Bundlers drop everything the consumer does not reference.
 *
 * 2. **Support installation** — enables `worksheet.addChart()`,
 *    `workbook.xlsx.writeFile()` chart-cache population, and chart
 *    reconstruction during XLSX load:
 *    ```ts
 *    import { installChartSupport } from "@cjnoname/excelts/chart";
 *    installChartSupport();                    // once, at startup
 *    worksheet.addChart(options, "A1:D20");    // now works
 *    ```
 *
 * Chart support is **never installed implicitly** — consumers pay for
 * what they ask for. The package's root `sideEffects: false` contract
 * stays intact, so bundles that only use the functional API include
 * exactly the code paths reachable from the exports they reference.
 *
 * `installChartSupport` lives in a separate module (`./install.ts`) so
 * its host-registry wiring and the full chart builder/renderer pipeline
 * don't get pulled in by consumers who only import one or two
 * functional helpers.
 */

export { installChartSupport, uninstallChartSupport } from "./install";

export { Chart, buildChartModel } from "./chart";
export { buildComboChartModel } from "./chart-builder";
export { fillChartCaches, fillChartExCaches, fillNumRef, fillStrRef } from "./cache-populator";
export {
  parseChartColors,
  buildChartColors,
  parseChartStyle,
  buildChartStyle
} from "./chart-sidecar";
export { buildChartExModel } from "./chart-ex-builder";
export { parseChartEx } from "./chart-ex-parser";
export {
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  renderChartEx,
  renderChartExPng,
  renderChartExSvg,
  VECTOR_PDF_CHART_EX_LAYOUT_IDS
} from "./chart-ex-renderer";
export {
  applyChartExPreset,
  applyChartPreset,
  CHART_EX_PRESETS,
  CHART_PRESETS,
  EXCEL_CHART_EX_PRESETS,
  EXCEL_CHART_PRESETS
} from "./chart-presets";
export {
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  seriesFromColumns
} from "./chart-api";
export type {
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions,
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  SeriesFromColumnsOptions
} from "./chart-api";
export type { ExcelChartExPreset, ExcelChartPreset } from "./chart-presets";
export {
  applyAxisTransform,
  buildChartScene,
  buildEffectFilter,
  drawChartPdf,
  renderChartPng,
  renderChartSvg
} from "./chart-renderer";
export type {
  ChartPdfDrawingSurface,
  ChartPdfPathOp,
  ChartRenderOptions,
  ChartScene,
  ChartSceneLegend,
  ChartSceneLine,
  ChartScenePieSlice,
  ChartSceneSeries,
  ChartSceneText,
  PdfChartRenderOptions,
  PdfColor,
  RegionMapDataOptions,
  RegionMapMatchRule
} from "./chart-renderer";
export { resolveTopologyObject } from "./topojson";
export type { ResolvedRing, TopoGeometry, TopoGeometryCollection, TopologyLike } from "./topojson";
export type {
  ChartExModel,
  ChartExSpace,
  ChartExData,
  ChartExDataEntry,
  ChartExStringDimension,
  ChartExNumericDimension,
  ChartExDimensionType,
  ChartExChart,
  ChartExPlotArea,
  ChartExSeries,
  ChartExSeriesType,
  ChartExDataLabels,
  ChartExLayoutProperties,
  ChartExAxis,
  AddChartExOptions,
  AddChartExHistogramOptions,
  AddChartExWaterfallOptions,
  AddChartExBoxWhiskerOptions,
  AddChartExSeriesOptions,
  ChartExType
} from "./chart-ex-types";
export type { ChartExEntry } from "./chart";
export type { ChartEntry, ChartAnchorModel } from "./chart";
export type {
  ChartModel,
  ChartType,
  ChartData,
  PlotArea,
  ChartTypeGroup,
  ChartAxis,
  ChartTitle,
  ChartLegend,
  AddChartOptions,
  AddBarChartOptions,
  AddBarChartSeriesOptions,
  AddPieChartOptions,
  AddPieChartSeriesOptions,
  AddScatterChartOptions,
  AddScatterChartSeriesOptions,
  AddSurfaceChartOptions,
  AddSurfaceChartSeriesOptions,
  AddChartSeriesOptions,
  AddChartRange,
  AddComboChartOptions,
  ComboChartGroupOptions,
  ChartStyleModel,
  ChartColorsModel,
  ChartColorsEntry,
  ShapeProperties,
  ChartTextProperties,
  ChartColor,
  ChartFill,
  ChartLine,
  PivotChartSource,
  SeriesBase,
  AddChartMarkerOptions,
  AddDataLabelsOptions,
  AddTrendlineOptions,
  AddErrorBarsOptions,
  AddDataPointOptions,
  AddAxisOptions,
  // Add* option types
  AddShapeFillOptions,
  AddTitleOptions,
  AddLegendOptions,
  AddPlotAreaOptions,
  AddTrendlineLabelOptions,
  AddDataLabelEntryOptions,
  // Chart type groups
  BarChartGroup,
  LineChartGroup,
  PieChartGroup,
  DoughnutChartGroup,
  AreaChartGroup,
  ScatterChartGroup,
  BubbleChartGroup,
  RadarChartGroup,
  SurfaceChartGroup,
  StockChartGroup,
  OfPieChartGroup,
  // Series types
  BarSeries,
  LineSeries,
  PieSeries,
  AreaSeries,
  ScatterSeries,
  BubbleSeries,
  RadarSeries,
  SurfaceSeries,
  // Axis types
  CategoryAxis,
  ValueAxis,
  DateAxis,
  SeriesAxis,
  // Sub-types
  DataLabels,
  DataLabelEntry,
  DataPoint,
  DataTable,
  Trendline,
  TrendlineLabel,
  ErrorBars,
  ChartMarker,
  ChartRichText,
  ChartBodyProperties,
  ChartParagraphProperties,
  ChartBullet,
  ChartLineSpacing,
  UnderlineStyle,
  StrikeStyle,
  CapStyle,
  ParagraphAlignment,
  EffectList,
  Shadow,
  Scene3D,
  ShapeProperties3D,
  Bevel,
  NumberReference,
  StringReference,
  NumberDataSource,
  AxisDataSource,
  View3D,
  BarGrouping,
  LineGrouping,
  ChartLayout,
  ChartParagraph,
  ChartTextRun,
  DisplayUnits,
  // Layout and entry types
  ManualLayout,
  LegendEntry,
  DataLabelPosition,
  LegendPosition,
  // Enum/union types
  TrendlineType,
  ErrorBarType,
  ErrorBarDirection,
  ErrorBarValueType,
  // Chart-specific types
  BandFormat,
  UpDownBars,
  PictureOptions,
  PictureFormat,
  // Data types
  NumberLiteral,
  StringLiteral,
  MultiLevelStringReference,
  MultiLevelStringCache,
  NumberCache,
  StringCache,
  // Axis enums
  ScatterStyle,
  RadarStyle,
  StockSeries,
  BarDirection,
  BarShape,
  OfPieType,
  SplitType,
  AxisCrosses,
  AxisPosition,
  AxisOrientation,
  TickMark,
  TickLabelPosition,
  TimeUnit,
  LabelAlignment,
  DisplayBlanksAs,
  PrintSettings,
  PivotFormat,
  // Transitively-referenced types that callers need when constructing
  // public option shapes. They used to live behind `@excel/chart/types`
  // only, forcing consumers to either duplicate the interface or reach
  // into the internal module path.
  DataLabelsRange,
  AxisBase,
  PresetGeometry,
  CustomGeometry,
  CustomGeometryPath,
  CustomGeometryCommand,
  ShapeTransform,
  ChartBlipFill,
  ChartColorVariation,
  AddChartPictureFillImage,
  ChartPictureFillImageData,
  PivotChartOptions,
  ChartUnknownElement,
  ChartStyleElement,
  ChartRange
} from "./types";
export {
  parseSpPr,
  parseTxPr,
  getSpPrFillColor,
  getSpPrLine,
  getSpPrGradient,
  getSpPrPattern,
  getTxPrFontSize,
  getTxPrColor,
  buildSpPr,
  buildTxPr,
  setSpPrFill,
  setSpPrLine
} from "./shape-properties";
