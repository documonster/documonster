/**
 * Index file for the chart module.
 */
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
  PivotFormat
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
