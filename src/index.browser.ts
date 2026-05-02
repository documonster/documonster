/**
 * Browser entry point - No Node.js dependencies
 * This version is optimized for browser environments with minimal bundle size
 */

// =============================================================================
// Main Classes (Browser-compatible)
// =============================================================================
// All imports point at the explicit `.browser` variants so the emitted
// .d.ts files reflect the browser-only type surface — without this,
// the types would leak Node-only methods (readFile / writeFile /
// file-path streaming input) via `@excel/workbook` resolving to the
// Node class that extends the browser base. Build-time rewriting
// handles the JS bundle; only the types need this explicit mapping.
export { Workbook } from "@excel/workbook.browser";
export { Worksheet } from "@excel/worksheet";
export { Row } from "@excel/row";
export { Column } from "@excel/column";
export { Cell } from "@excel/cell";
export { Range } from "@excel/range";
export { Image } from "@excel/image";
export * from "@excel/anchor";
export { Table } from "@excel/table";
export { Note } from "@excel/note";
export { DataValidations } from "@excel/data-validations";
export { FormCheckbox } from "@excel/form-control";
// Note: the formula engine lives at the `./formula` subpath so it stays
// out of bundles that only need to read / write XLSX files. Import
// `@cj-tech-master/excelts/formula` to enable `Workbook.calculateFormulas()`
// and automatic recalculation in `excelToPdf()`.

// =============================================================================
// Enums
// =============================================================================
export * from "@excel/enums";

// =============================================================================
// Types
// =============================================================================

// Export all type definitions from types.ts
export * from "@excel/types";

// Watermark image generator utility
export { createTextWatermarkImage } from "@excel/utils/watermark-image";
export type { TextWatermarkImageOptions } from "@excel/utils/watermark-image";

// Export pivot table types (type-only, no runtime dependency)
export type {
  PivotTable,
  PivotTableModel,
  PivotTableValue,
  PivotTableSource,
  CacheField,
  SharedItemValue,
  DataField,
  PivotTableChartFormat,
  PivotTableSubtotal,
  RecordValue,
  ParsedCacheDefinition,
  ParsedCacheRecords
} from "@excel/pivot-table";

// Form control types
export type {
  FormCheckboxModel,
  FormCheckboxOptions,
  FormControlRange,
  FormControlAnchor
} from "@excel/form-control";

// =============================================================================
// Streaming Writer (Browser-compatible)
// Uses cross-platform base implementation without Node.js fs
// =============================================================================

export { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
export { WorkbookReader } from "@excel/stream/workbook-reader.browser";
export { WorksheetWriter } from "@excel/stream/worksheet-writer";
export { WorksheetReader } from "@excel/stream/worksheet-reader";

// =============================================================================
// NOTE: Node.js-only features not available in browser:
// - Reading from a file path is not supported (use Uint8Array/ArrayBuffer/Blob instead)
// - Writing to a file path is not supported (use writeBuffer() / stream output, then save as Blob/download)
// =============================================================================

// =============================================================================
// CSV types and stream classes
// =============================================================================
// CSV types and stream classes
// =============================================================================
export type { CsvOptions, CsvInput } from "@excel/workbook.browser";
export {
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "@csv/stream";

// =============================================================================
// Markdown types
// =============================================================================
export type { MarkdownOptions } from "@markdown/types";
export type {
  MarkdownAlignment,
  MarkdownParseResult,
  MarkdownParseOptions,
  MarkdownColumnConfig,
  MarkdownFormatOptions
} from "@markdown/types";

// =============================================================================
// Additional Classes & Types
// =============================================================================

export { DefinedNames, type DefinedNameModel } from "@excel/defined-names";
export { Chartsheet } from "@excel/chartsheet";
export type {
  AddChartsheetOptions,
  AddPivotChartsheetOptions,
  ChartsheetOptions,
  ChartsheetViewOptions
} from "@excel/chartsheet";
export type { CheckboxState } from "@excel/form-control";
export type { ColumnDefn, ColumnHeaderValue } from "@excel/column";
export type { RangeInput } from "@excel/range";
export type {
  WorkbookModel,
  WorkbookMedia,
  WorkbookProtectionModel,
  ExternalLinkModel,
  ExternalLinkCachedSheet
} from "@excel/workbook.browser";

// =============================================================================
// Utilities
// =============================================================================

// Cell display-text helpers (apply numFmt to produce an Excel-style string).
// Cross-platform — no Node dependencies. Ships in both entry points so
// browser users can access the same `Cell.displayText` behavior via
// the standalone helper.
export { getCellDisplayText, formatCellValue, isDateDisplayFormat } from "@excel/utils/cell-format";

// Cell address encoding/decoding (0-indexed)
export {
  decodeCol,
  encodeCol,
  decodeRow,
  encodeRow,
  decodeCell,
  encodeCell,
  decodeRange,
  encodeRange
} from "@excel/utils/address";
export type { CellAddress, SheetRange, Origin } from "@excel/utils/address";

// Worksheet data conversion option types
export type { SheetToJSONOptions, AddJSONOptions, AddAOAOptions } from "@excel/worksheet";

// Chart helpers and types — full public surface. Keep in sync with the
// Node entry (`src/index.ts`). See that file for documentation.
export {
  applyChartExPreset,
  applyChartPreset,
  buildChartExModel,
  buildComboChartModel,
  buildChartModel,
  buildChartScene,
  buildEffectFilter,
  CHART_EX_PRESETS,
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  CHART_PRESETS,
  drawChartPdf,
  EXCEL_CHART_EX_PRESETS,
  EXCEL_CHART_PRESETS,
  fillChartCaches,
  fillChartExCaches,
  fillNumRef,
  fillStrRef,
  Chart,
  renderChartPng,
  renderChartSvg,
  renderChartEx,
  renderChartExPng,
  renderChartExSvg,
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  VECTOR_PDF_CHART_EX_LAYOUT_IDS,
  parseChartColors,
  buildChartColors,
  parseChartStyle,
  buildChartStyle,
  parseChartEx,
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
  setSpPrLine,
  resolveTopologyObject,
  applyAxisTransform,
  seriesFromColumns
} from "@excel/chart/index";
export type {
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions,
  AddChartExOptions,
  AddChartExHistogramOptions,
  AddChartExWaterfallOptions,
  AddChartExBoxWhiskerOptions,
  AddChartExSeriesOptions,
  ChartExType,
  ExcelChartExPreset,
  ExcelChartPreset,
  SeriesFromColumnsOptions,
  ChartEntry,
  ChartAnchorModel,
  ChartExEntry,
  ChartModel,
  ChartType,
  ChartTypeGroup,
  ChartData,
  PlotArea,
  ChartAxis,
  ChartTitle,
  ChartLegend,
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
  ChartPdfDrawingSurface,
  ChartPdfPathOp,
  ChartRenderOptions,
  ChartScene,
  ChartSceneLegend,
  ChartSceneSeries,
  ChartSceneText,
  ChartSceneLine,
  ChartScenePieSlice,
  PdfChartRenderOptions,
  RegionMapDataOptions,
  RegionMapMatchRule,
  AddChartMarkerOptions,
  AddDataLabelsOptions,
  AddTrendlineOptions,
  AddErrorBarsOptions,
  AddDataPointOptions,
  AddAxisOptions,
  AddShapeFillOptions,
  AddTitleOptions,
  AddLegendOptions,
  AddPlotAreaOptions,
  AddTrendlineLabelOptions,
  AddDataLabelEntryOptions,
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
  BarSeries,
  LineSeries,
  PieSeries,
  AreaSeries,
  ScatterSeries,
  BubbleSeries,
  RadarSeries,
  SurfaceSeries,
  CategoryAxis,
  ValueAxis,
  DateAxis,
  SeriesAxis,
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
  ManualLayout,
  LegendEntry,
  DataLabelPosition,
  LegendPosition,
  TrendlineType,
  ErrorBarType,
  ErrorBarDirection,
  ErrorBarValueType,
  BandFormat,
  UpDownBars,
  PictureOptions,
  PictureFormat,
  NumberLiteral,
  StringLiteral,
  MultiLevelStringReference,
  MultiLevelStringCache,
  NumberCache,
  StringCache,
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
  ResolvedRing,
  TopoGeometry,
  TopoGeometryCollection,
  TopologyLike,
  // Transitively-referenced types callers need when constructing
  // public option shapes. Keep in sync with `src/index.ts` and
  // `src/modules/excel/chart/index.ts`.
  AddChartPictureFillImage,
  AxisBase,
  ChartBlipFill,
  ChartColorVariation,
  ChartPictureFillImageData,
  ChartRange,
  ChartStyleElement,
  ChartUnknownElement,
  CustomGeometry,
  CustomGeometryCommand,
  CustomGeometryPath,
  DataLabelsRange,
  PivotChartOptions,
  PresetGeometry,
  ShapeTransform
} from "@excel/chart/index";

// Date conversion (Excel serial dates <-> JS Date)
export { dateToExcel, excelToDate } from "@utils/utils.base";

// Base64 utilities (cross-platform)
export { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils.base";

// XML utilities
export { xmlEncode, xmlDecode, xmlEncodeAttr, validateXmlName } from "@xml/encode";

// Date parsing/formatting (high-performance, zero-dep)
export { DateParser, DateFormatter, getSupportedFormats, type DateFormat } from "@utils/datetime";

// Error infrastructure
export {
  BaseError,
  type BaseErrorOptions,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@utils/errors";

// Binary utilities (cross-platform)
export {
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString
} from "@utils/binary";

// =============================================================================
// PDF Export (Browser-compatible, zero external dependencies)
// =============================================================================

export {
  pdf,
  excelToPdf,
  PageSizes,
  PdfError,
  PdfRenderError,
  PdfFontError,
  PdfStructureError,
  isPdfError
} from "@pdf/index";
export type {
  PdfExportOptions,
  PdfPageSize,
  PdfOrientation,
  PdfMargins,
  PageSizeName,
  PdfColor
} from "@pdf/index";

// =============================================================================
// Errors
// =============================================================================

export {
  ExcelError,
  isExcelError,
  ExcelFileError,
  ExcelDownloadError,
  ExcelNotSupportedError,
  ExcelStreamStateError,
  InvalidAddressError,
  ColumnOutOfBoundsError,
  RowOutOfBoundsError,
  MergeConflictError,
  InvalidValueTypeError,
  XmlParseError,
  WorksheetNameError,
  PivotTableError,
  TableError,
  ImageError,
  MaxItemsExceededError
} from "@excel/errors";

// Markdown errors
export { MarkdownError, MarkdownParseError } from "@markdown/errors";
