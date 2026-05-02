// =============================================================================
// Main Classes
// =============================================================================

export { Workbook } from "@excel/workbook";
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
// Node.js Only: Streaming Classes
// These can also be accessed via Workbook.createStreamWriter/createStreamReader
// =============================================================================

export { WorkbookWriter } from "@excel/stream/workbook-writer";
export { WorkbookReader } from "@excel/stream/workbook-reader";
export { WorksheetWriter } from "@excel/stream/worksheet-writer";
export { WorksheetReader } from "@excel/stream/worksheet-reader";

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

// Pivot table types
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

// Node.js Only: Streaming reader types
export type {
  WorkbookReaderOptions,
  ParseEvent,
  SharedStringEvent,
  WorksheetReadyEvent,
  HyperlinksEvent
} from "@excel/stream/workbook-reader";

export type {
  WorksheetReaderOptions,
  WorksheetEvent,
  RowEvent,
  HyperlinkEvent,
  WorksheetHyperlink
} from "@excel/stream/worksheet-reader";

// Node.js Only: Streaming writer types
export type {
  WorkbookWriterOptions,
  WorkbookZipOptions,
  ZlibOptions
} from "@excel/stream/workbook-writer";

// CSV types and stream classes
export type { CsvOptions, CsvInput } from "@excel/workbook";
export {
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "@csv/stream";

// Markdown types
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
} from "@excel/workbook";
export type { NodeInput } from "@excel/stream/workbook-reader";

// =============================================================================
// Utilities
// =============================================================================

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

// Chart helpers and types — full public surface from @excel/chart so
// users of the package have programmatic access to the chart pipeline
// (building, parsing, rendering, mutation) without reaching into
// internal module paths via tsconfig aliases.
export {
  // Builder / parser
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
  // Cache population / reference resolution
  fillChartCaches,
  fillChartExCaches,
  fillNumRef,
  fillStrRef,
  // Chart registry / classes
  Chart,
  // Renderers (SVG / PNG / PDF) — classic and ChartEx
  renderChartPng,
  renderChartSvg,
  renderChartEx,
  renderChartExPng,
  renderChartExSvg,
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  VECTOR_PDF_CHART_EX_LAYOUT_IDS,
  // Sidecar (chartStyle / chartColors) helpers
  parseChartColors,
  buildChartColors,
  parseChartStyle,
  buildChartStyle,
  // ChartEx parser + builder
  parseChartEx,
  // Shape-properties helpers (parse / build / accessors)
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
  // TopoJSON resolver (for regionMap data)
  resolveTopologyObject,
  // Scene helpers
  applyAxisTransform,
  seriesFromColumns
} from "@excel/chart/index";
export type {
  // High-level entry options
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
  // Chart model types
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
  // ChartEx model types
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
  // Renderer scene types
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
  // Series-level option types
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
  // Group types
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
  // Data / adornment types
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
  // TopoJSON
  ResolvedRing,
  TopoGeometry,
  TopoGeometryCollection,
  TopologyLike,
  // Transitively-referenced types callers need when constructing
  // public option shapes. Exported from `@excel/chart/index` but
  // previously not forwarded through the root — forcing consumers to
  // reach into the internal module path. Keep this block in sync with
  // `src/index.browser.ts` and `src/modules/excel/chart/index.ts`.
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

// Cell display-text helpers (apply numFmt to produce an Excel-style string)
export { getCellDisplayText, formatCellValue, isDateDisplayFormat } from "@excel/utils/cell-format";

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
// PDF Export
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
