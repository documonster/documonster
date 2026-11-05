/**
 * Public entry for the documonster chart module.
 *
 * **Functional, zero-side-effect, fully tree-shakeable.** Three namespaces:
 *
 * - `ChartClassic` — classic charts: build a model, render it to SVG / PNG / a display list
 *   (which `documonster/draw` puts on a PDF page through `Pdf.createDrawSurface`), read and
 *   write the chart style and colour sidecars;
 * - `ChartEx` — the same operations for the modern ChartEx family, plus its XML;
 * - `DrawingMl` — the `spPr` / `txPr` shape and text property helpers.
 *
 * ```ts
 * import { ChartClassic } from "documonster/chart";
 * const model = ChartClassic.createModel(options);
 * const svg   = ChartClassic.toSvg(model);
 * ```
 *
 * No install / registration step exists. The high-level chart APIs
 * (`Chart.place`, chart serialisation on `Workbook.toBuffer`,
 * XLSX chart reconstruction on load) import the chart implementation directly
 * and statically. A consumer that never references any chart API gets the
 * entire chart implementation tree-shaken out by the bundler — the package's
 * root `sideEffects: false` contract keeps this guarantee intact.
 */

import * as ChartExNs from "@excel/chart/api/chart-ex";
import * as ClassicNs from "@excel/chart/api/classic";
import * as DrawingMlNs from "@excel/chart/api/drawing-ml";
import {
  fillChartCaches as fillChartCachesImpl,
  fillChartExCaches as fillChartExCachesImpl,
  fillNumRef as fillNumRefImpl,
  fillStrRef as fillStrRefImpl
} from "@excel/chart/bridge/excel-chart-host";
import type { ChartExModel as ChartExModelType } from "@excel/chart/model/chart-ex-types";
import {
  EXCEL_CHART_EX_PRESETS as EXCEL_CHART_EX_PRESETS_IMPL,
  EXCEL_CHART_PRESETS as EXCEL_CHART_PRESETS_IMPL
} from "@excel/chart/model/chart-presets";
import {
  canRenderChartExAsVectorPdf as canRenderChartExAsVectorPdfImpl,
  checkChartExModel,
  drawChartExPdf as drawChartExPdfImpl,
  VECTOR_PDF_CHART_EX_LAYOUT_IDS as VECTOR_PDF_CHART_EX_LAYOUT_IDS_IMPL
} from "@excel/chart/render/chart-ex-renderer";
import { drawChartPdf as drawChartPdfImpl } from "@excel/chart/render/chart-renderer";
import type {
  ChartPdfDrawingSurface as ChartPdfDrawingSurfaceType,
  ChartPdfPathOp as ChartPdfPathOpType,
  PdfChartRenderOptions as PdfChartRenderOptionsType,
  PdfColor as ChartPdfColorType,
  RegionMapDataOptions as RegionMapDataOptionsType
} from "@excel/chart/render/chart-renderer";
import { checkChartRenderOptions, checkDrawPdfArguments } from "@excel/chart/render/render-options";

// =============================================================================
// Public API — namespaces (tree-shaken via `export * as`)
// =============================================================================

export * as ChartClassic from "@excel/chart/api/classic";
export * as ChartEx from "@excel/chart/api/chart-ex";
export * as DrawingMl from "@excel/chart/api/drawing-ml";
export type { ChartFillNumRefOptions, ChartFillRefOptions } from "@excel/chart/api/types";
export type { ChartThemeColors } from "@excel/chart/shared/chart-utils";

/**
 * Invalid chart options, references, presets or topology throw `ExcelChartOptionsError`, and an
 * argument of the wrong type or shape throws `ExcelArgumentError` — both `ExcelError`s, the same
 * classes `documonster/excel` publishes. `ChartEx.fromXml` reports malformed XML with
 * `documonster/xml`'s `XmlParseError` (see `isXmlError`).
 */
export {
  ExcelArgumentError,
  ExcelChartOptionsError,
  ExcelError,
  isExcelError,
  // Deprecated pre-prefix name; removed in 1.0.
  ChartOptionsError
} from "@excel/errors";
export type { ChartHandle } from "@excel/core/worksheet-core";
export { CHART_EX_PRESETS, CHART_PRESETS } from "@excel/chart/model/chart-presets";
/** @deprecated The preset names are the keys of the table: `Object.keys(CHART_PRESETS)`. Removed in 1.0. */
export const EXCEL_CHART_PRESETS = EXCEL_CHART_PRESETS_IMPL;
/** @deprecated The preset names are the keys of the table: `Object.keys(CHART_EX_PRESETS)`. Removed in 1.0. */
export const EXCEL_CHART_EX_PRESETS = EXCEL_CHART_EX_PRESETS_IMPL;
/**
 * @deprecated It only answers the deprecated {@link canRenderChartExAsVectorPdf}: every Office 2016 chart now has a
 * vector picture through `ChartEx.toDrawList`. Removed in 1.0.
 */
export const VECTOR_PDF_CHART_EX_LAYOUT_IDS = VECTOR_PDF_CHART_EX_LAYOUT_IDS_IMPL;
export type {
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions,
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  SeriesFromColumnsOptions
} from "@excel/chart/build/chart-api";
export type {
  ChartExPresetConfig,
  ChartPresetConfig,
  ExcelChartExPreset,
  ExcelChartPreset
} from "@excel/chart/model/chart-presets";
export type {
  ChartRenderOptions,
  ChartScene,
  ChartSceneLegend,
  ChartSceneLine,
  ChartScenePieSlice,
  ChartSceneSeries,
  ChartSceneText,
  RegionMapDataOptions,
  RegionMapMatchRule
} from "@excel/chart/render/chart-renderer";
export type {
  ResolvedRing,
  TopoGeometry,
  TopoGeometryCollection,
  TopologyLike
} from "@excel/chart/render/topojson";
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
  ChartExType,
  ChartExEntry
} from "@excel/chart/model/chart-ex-types";
export type { ChartEntry, ChartAnchorModel, ChartRelEntry } from "@excel/chart/model/types";
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
  ChartColorTransform,
  ChartColorTransformKind,
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
  // public option shapes. They used to live behind `@excel/chart/model/types`
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
} from "@excel/chart/model/types";
/**
 * The colour a {@link ChartPdfDrawingSurface} is painted with — channels in `0..1`,
 * optional alpha. Prefixed because `documonster/pdf` publishes its own `PdfColor`.
 *
 * @deprecated A type of the deprecated {@link drawChartPdf} / {@link drawChartExPdf}. A chart
 * goes onto a PDF page as a display list: `Draw.render(ChartClassic.toDrawList(model),
 * Pdf.createDrawSurface(page, rect))`. Removed in 1.0.
 */
export type ChartPdfColor = ChartPdfColorType;
/** @deprecated See {@link ChartPdfColor}. Removed in 1.0. */
export type PdfColor = ChartPdfColorType;
/**
 * The surface the deprecated {@link drawChartPdf} / {@link drawChartExPdf} draw onto.
 *
 * @deprecated A chart goes onto a PDF page as a display list, through `Pdf.createDrawSurface`
 * (a `PdfDrawPage`). Removed in 1.0.
 */
export type ChartPdfDrawingSurface = ChartPdfDrawingSurfaceType;
/** @deprecated A path operation of {@link ChartPdfDrawingSurface}. Removed in 1.0. */
export type ChartPdfPathOp = ChartPdfPathOpType;
/**
 * The options of the deprecated {@link drawChartPdf}: the render options and the box's corner.
 *
 * @deprecated Pass `ChartRenderOptions` to `ChartClassic.toDrawList` and the box to
 * `Pdf.createDrawSurface`. Removed in 1.0.
 */
export type PdfChartRenderOptions = PdfChartRenderOptionsType;

// =============================================================================
// Deprecated flat names — forwards to the namespaces above
// =============================================================================

/** @deprecated Use `ChartClassic.createModel`. Removed in 1.0. */
export const buildChartModel = ClassicNs.createModel;
/** @deprecated Use `ChartClassic.createComboModel`. Removed in 1.0. */
export const buildComboChartModel = ClassicNs.createComboModel;
/** @deprecated Use `ChartClassic.fillCaches(model, workbook, { worksheet })`. Removed in 1.0. */
export const fillChartCaches = fillChartCachesImpl;
/** @deprecated Use `ChartEx.fillCaches(model, workbook, { worksheet })`. Removed in 1.0. */
export const fillChartExCaches = fillChartExCachesImpl;
/**
 * @deprecated Use `ChartClassic.fillNumRef(ref, workbook, { date1904, worksheet })`. Removed in 1.0.
 */
export const fillNumRef = fillNumRefImpl;
/**
 * @deprecated Use `ChartClassic.fillStrRef(ref, workbook, { worksheet })`. Removed in 1.0.
 */
export const fillStrRef = fillStrRefImpl;
/** @deprecated Use `ChartClassic.fromColorsXml`. Removed in 1.0. */
export const parseChartColors = ClassicNs.fromColorsXml;
/** @deprecated Use `ChartClassic.toColorsXml`. Removed in 1.0. */
export const buildChartColors = ClassicNs.toColorsXml;
/** @deprecated Use `ChartClassic.fromStyleXml`. Removed in 1.0. */
export const parseChartStyle = ClassicNs.fromStyleXml;
/** @deprecated Use `ChartClassic.toStyleXml`. Removed in 1.0. */
export const buildChartStyle = ClassicNs.toStyleXml;
/** @deprecated Use `ChartEx.createModel`. Removed in 1.0. */
export const buildChartExModel = ChartExNs.createModel;
/** @deprecated Use `ChartEx.fromXml`. Removed in 1.0. */
export const parseChartEx = ChartExNs.fromXml;
/**
 * @deprecated Every Office 2016 chart has a vector picture: `ChartEx.toDrawList(model)` draws
 * any model, and `Pdf.fromChart` picks its path itself. This still says whether each series'
 * layout is one of `VECTOR_PDF_CHART_EX_LAYOUT_IDS`. Removed in 1.0.
 */
export const canRenderChartExAsVectorPdf = canRenderChartExAsVectorPdfImpl;
/**
 * @deprecated Put the chart on the page as a display list, as every producer does:
 * `Draw.render(ChartEx.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.
 * This draws that same display list — the chart and its chart area — onto a surface with
 * `drawRect` / `drawLine` / `drawText`. Removed in 1.0.
 */
export function drawChartExPdf(
  surface: ChartPdfDrawingSurfaceType,
  model: ChartExModelType,
  rect: { x: number; y: number; width: number; height: number },
  options: { title?: string; regionMap?: RegionMapDataOptionsType } = {}
): ChartPdfDrawingSurfaceType {
  const where = "drawChartExPdf";
  checkDrawPdfArguments(where, surface, rect);
  checkChartExModel(where, model);
  checkChartRenderOptions(where, options);
  return drawChartExPdfImpl(surface, model, rect, options);
}
/** @deprecated Use `ChartEx.toPng`. Removed in 1.0. */
export const renderChartExPng = ChartExNs.toPng;
/** @deprecated Use `ChartEx.toSvg`. Removed in 1.0. */
export const renderChartExSvg = ChartExNs.toSvg;
/** @deprecated Use `ChartEx.toXml` — this returns the part's XML, not a picture. Removed in 1.0. */
export const renderChartEx = ChartExNs.toXml;
/** @deprecated Use `ChartEx.applyPreset`. Removed in 1.0. */
export const applyChartExPreset = ChartExNs.applyPreset;
/** @deprecated Use `ChartClassic.applyPreset`. Removed in 1.0. */
export const applyChartPreset = ClassicNs.applyPreset;
/** @deprecated Use `ChartEx.createOptionsFromRows`. Removed in 1.0. */
export const chartExOptionsFromRows = ChartExNs.createOptionsFromRows;
/** @deprecated Use `ChartEx.createOptionsFromTable`. Removed in 1.0. */
export const chartExOptionsFromTable = ChartExNs.createOptionsFromTable;
/** @deprecated Use `ChartClassic.createOptionsFromRows`. Removed in 1.0. */
export const chartOptionsFromRows = ClassicNs.createOptionsFromRows;
/** @deprecated Use `ChartClassic.createOptionsFromTable`. Removed in 1.0. */
export const chartOptionsFromTable = ClassicNs.createOptionsFromTable;
/** @deprecated Use `ChartClassic.createSeriesFromColumns`. Removed in 1.0. */
export const seriesFromColumns = ClassicNs.createSeriesFromColumns;
/** @deprecated Use `ChartClassic.Backend.applyAxisTransform`. Removed in 1.0. */
export const applyAxisTransform = ClassicNs.Backend.applyAxisTransform;
/** @deprecated Use `ChartClassic.Backend.toScene`. Removed in 1.0. */
export const buildChartScene = ClassicNs.Backend.toScene;
/** @deprecated Use `ChartClassic.Backend.toSvgFilter`. Removed in 1.0. */
export const buildEffectFilter = ClassicNs.Backend.toSvgFilter;
/**
 * @deprecated Put the chart on the page as a display list, as every producer does:
 * `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.
 * This draws that same display list — the chart and its chart area — onto a surface with
 * `drawRect` / `drawLine` / `drawText`. Removed in 1.0.
 */
export const drawChartPdf = drawChartPdfImpl;
/** @deprecated Use `ChartClassic.toPng`. Removed in 1.0. */
export const renderChartPng = ClassicNs.toPng;
/** @deprecated Use `ChartClassic.toSvg`. Removed in 1.0. */
export const renderChartSvg = ClassicNs.toSvg;
/** @deprecated Use `ChartClassic.Backend.resolveTopologyObject`. Removed in 1.0. */
export const resolveTopologyObject = ClassicNs.Backend.resolveTopologyObject;
/** @deprecated Use `DrawingMl.parseSpPr`. Removed in 1.0. */
export const parseSpPr = DrawingMlNs.parseSpPr;
/** @deprecated Use `DrawingMl.parseTxPr`. Removed in 1.0. */
export const parseTxPr = DrawingMlNs.parseTxPr;
/** @deprecated Use `DrawingMl.getSpPrFillColor`. Removed in 1.0. */
export const getSpPrFillColor = DrawingMlNs.getSpPrFillColor;
/** @deprecated Use `DrawingMl.getSpPrLine`. Removed in 1.0. */
export const getSpPrLine = DrawingMlNs.getSpPrLine;
/** @deprecated Use `DrawingMl.getSpPrGradient`. Removed in 1.0. */
export const getSpPrGradient = DrawingMlNs.getSpPrGradient;
/** @deprecated Use `DrawingMl.getSpPrPattern`. Removed in 1.0. */
export const getSpPrPattern = DrawingMlNs.getSpPrPattern;
/** @deprecated Use `DrawingMl.getTxPrFontSize`. Removed in 1.0. */
export const getTxPrFontSize = DrawingMlNs.getTxPrFontSize;
/** @deprecated Use `DrawingMl.getTxPrColor`. Removed in 1.0. */
export const getTxPrColor = DrawingMlNs.getTxPrColor;
/** @deprecated Use `DrawingMl.createSpPr`. Removed in 1.0. */
export const buildSpPr = DrawingMlNs.createSpPr;
/** @deprecated Use `DrawingMl.createTxPr`. Removed in 1.0. */
export const buildTxPr = DrawingMlNs.createTxPr;
/** @deprecated Use `DrawingMl.withSpPrFill`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0. */
export const setSpPrFill = DrawingMlNs.withSpPrFill;
/** @deprecated Use `DrawingMl.withSpPrLine`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0. */
export const setSpPrLine = DrawingMlNs.withSpPrLine;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
