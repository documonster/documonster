/**
 * Chart Builder - Constructs a ChartModel from simplified AddChartOptions.
 *
 * This bridges the high-level API (worksheet.addChart) to the full
 * OOXML chart model that the XForm layer serialises.
 */

import { ChartOptionsError } from "@excel/errors";

import type {
  AddChartOptions,
  AddChartSeriesOptions,
  AddComboChartOptions,
  AddAxisOptions,
  AddDataLabelsOptions,
  AddDataLabelEntryOptions,
  AddTrendlineOptions,
  AddTrendlineLabelOptions,
  AddErrorBarsOptions,
  AddDataPointOptions,
  AddChartMarkerOptions,
  AddShapeFillOptions,
  AddTitleOptions,
  AddLegendOptions,
  AddPlotAreaOptions,
  AxisDataSource,
  BarChartGroup,
  BarSeries,
  BubbleChartGroup,
  BubbleSeries,
  ChartAxis,
  ChartData,
  ChartLine,
  ChartMarker,
  ChartModel,
  ChartRichText,
  ChartTextProperties,
  ChartTypeGroup,
  DataLabels,
  DataLabelEntry,
  DataPoint,
  DataTable,
  DoughnutChartGroup,
  ErrorBars,
  LineChartGroup,
  LineSeries,
  NumberDataSource,
  OfPieChartGroup,
  PieChartGroup,
  PieSeries,
  AreaChartGroup,
  AreaSeries,
  RadarChartGroup,
  RadarSeries,
  ScatterChartGroup,
  ScatterSeries,
  SeriesBase,
  StockChartGroup,
  SurfaceChartGroup,
  SurfaceSeries,
  ShapeProperties,
  ChartColor,
  NumberReference,
  StringReference,
  Trendline,
  TrendlineLabel,
  CategoryAxis,
  DateAxis,
  ValueAxis,
  ChartTitle,
  ChartLegend,
  LegendEntry,
  PlotArea,
  UpDownBars,
  View3D,
  BarGrouping,
  LineGrouping,
  LegendPosition,
  PivotChartOptions,
  PivotChartSource,
  SeriesAxis
} from "./types";

const EMU_PER_POINT = 12700;
const DEFAULT_AXIS_START_ID = 100000000;
const AXIS_CHART_TYPES = new Set<AddChartOptions["type"]>([
  "bar",
  "bar3D",
  "line",
  "line3D",
  "area",
  "area3D",
  "scatter",
  "bubble",
  "radar",
  "stock",
  "surface",
  "surface3D"
]);
const NO_TRENDLINE_CHART_TYPES = new Set<AddChartOptions["type"]>([
  "pie",
  "pie3D",
  "doughnut",
  "ofPie",
  "surface",
  "surface3D"
]);
const PIE_FAMILY_CHART_TYPES = new Set<AddChartOptions["type"]>([
  "pie",
  "pie3D",
  "doughnut",
  "ofPie"
]);

/**
 * Simple axis ID allocator — scoped per buildChartModel call.
 * Axis IDs only need to be unique within a single chart.
 */
class AxIdAllocator {
  private next: number;
  constructor(start = DEFAULT_AXIS_START_ID) {
    this.next = start;
  }
  alloc(): number {
    return this.next++;
  }
}

function makeNumRef(formula: string): NumberReference {
  return { formula, cache: { points: [] } };
}

function makeStrRef(formula: string): StringReference {
  return { formula, cache: { points: [] } };
}

function makeNumData(formula: string): NumberDataSource {
  return { numRef: makeNumRef(formula) };
}

function makeCatData(formula: string): AxisDataSource {
  return { strRef: makeStrRef(formula) };
}

function makeAxisData(input: string | AxisDataSource): AxisDataSource {
  return typeof input === "string" ? makeCatData(input) : input;
}

function makeNumericAxisData(input: string | AxisDataSource): AxisDataSource {
  return typeof input === "string" ? { numRef: makeNumRef(input) } : input;
}

/**
 * Wrap a scatter/bubble `xValues` input as the appropriate
 * {@link AxisDataSource}. `xValueType` disambiguates the two OOXML
 * spellings:
 *
 *   - `"number"` (default) → `numRef` — standard scatter usage
 *   - `"text"`             → `strRef` — labelled scatter / bubble with
 *     categorical x axis (Excel renders it as evenly-spaced labels)
 *
 * When the caller passes a pre-built `AxisDataSource`, the hint is
 * ignored — the structure already carries the intent.
 */
function makeXAxisData(
  input: string | AxisDataSource,
  xValueType: "number" | "text" | undefined
): AxisDataSource {
  if (typeof input !== "string") {
    return input;
  }
  return xValueType === "text" ? makeCatData(input) : makeNumericAxisData(input);
}

function assertChartOptions(condition: unknown, message: string): void {
  if (!condition) {
    throw new ChartOptionsError(message);
  }
}

function assertFiniteNumber(value: unknown, path: string): asserts value is number {
  assertChartOptions(
    typeof value === "number" && Number.isFinite(value),
    `${path} must be a finite number.`
  );
}

function assertIntegerInRange(
  value: unknown,
  path: string,
  min: number,
  max: number
): asserts value is number {
  assertFiniteNumber(value, path);
  assertChartOptions(
    Number.isInteger(value) && value >= min && value <= max,
    `${path} must be an integer between ${min} and ${max}.`
  );
}

function assertNumberInRange(
  value: unknown,
  path: string,
  min: number,
  max: number
): asserts value is number {
  assertFiniteNumber(value, path);
  assertChartOptions(value >= min && value <= max, `${path} must be between ${min} and ${max}.`);
}

function validateChartOptions(opts: AddChartOptions, path = "chart"): void {
  assertChartOptions(!!opts && typeof opts === "object", `${path} options are required.`);
  assertChartOptions(!!opts.type, `${path}.type is required.`);
  validateChartLevelOptions(opts, path);
  const series = opts.series ?? [];
  // A chart with zero series is invalid — Excel will either refuse to
  // open the file or render a broken chart area. Catch it at build
  // time with a precise error instead of deferring to Excel.
  assertChartOptions(series.length > 0, `${path}.series must contain at least one series.`);
  for (let i = 0; i < series.length; i++) {
    validateSeriesOptions(opts.type, series[i], `${path}.series[${i}]`);
  }
}

function validateComboChartOptions(opts: AddComboChartOptions): void {
  assertChartOptions(!!opts && typeof opts === "object", "combo chart options are required.");
  assertChartOptions(
    Array.isArray(opts.groups) && opts.groups.length > 0,
    "combo chart groups must contain at least one group."
  );
  for (let i = 0; i < opts.groups.length; i++) {
    validateChartOptions(opts.groups[i], `groups[${i}]`);
  }
  validateSharedChartOptions(opts, "combo chart");
}

function validateChartLevelOptions(opts: AddChartOptions, path: string): void {
  validateSharedChartOptions(opts, path);
  if (opts.grouping !== undefined) {
    assertChartOptions(
      opts.type === "bar" ||
        opts.type === "bar3D" ||
        opts.type === "line" ||
        opts.type === "line3D" ||
        opts.type === "area" ||
        opts.type === "area3D",
      `${path}.grouping is only valid for bar, line, and area charts.`
    );
  }
  if (opts.barDir !== undefined) {
    assertChartOptions(
      opts.type === "bar" || opts.type === "bar3D",
      `${path}.barDir is only valid for bar and bar3D charts.`
    );
  }
  if (opts.scatterStyle !== undefined) {
    assertChartOptions(
      opts.type === "scatter",
      `${path}.scatterStyle is only valid for scatter charts.`
    );
  }
  if (opts.radarStyle !== undefined) {
    assertChartOptions(opts.type === "radar", `${path}.radarStyle is only valid for radar charts.`);
  }
  if (opts.ofPieType !== undefined) {
    assertChartOptions(opts.type === "ofPie", `${path}.ofPieType is only valid for ofPie charts.`);
  }
  if (opts.wireframe !== undefined) {
    assertChartOptions(
      opts.type === "surface" || opts.type === "surface3D",
      `${path}.wireframe is only valid for surface and surface3D charts.`
    );
  }
  if (opts.bandFormats !== undefined) {
    assertChartOptions(
      opts.type === "surface" || opts.type === "surface3D",
      `${path}.bandFormats is only valid for surface and surface3D charts.`
    );
    for (let i = 0; i < opts.bandFormats.length; i++) {
      assertIntegerInRange(
        opts.bandFormats[i].index,
        `${path}.bandFormats[${i}].index`,
        0,
        Number.MAX_SAFE_INTEGER
      );
    }
  }
  if (!AXIS_CHART_TYPES.has(opts.type)) {
    assertChartOptions(
      opts.categoryAxis === undefined,
      `${path}.categoryAxis is not valid for ${opts.type} charts because they do not have axes.`
    );
    assertChartOptions(
      opts.valueAxis === undefined,
      `${path}.valueAxis is not valid for ${opts.type} charts because they do not have axes.`
    );
  }
  if ((opts.type === "surface" || opts.type === "surface3D") && opts.dataLabels !== undefined) {
    assertChartOptions(false, `${path}.dataLabels is not supported for surface charts.`);
  }
  if (opts.holeSize !== undefined) {
    assertChartOptions(
      opts.type === "doughnut",
      `${path}.holeSize is only valid for doughnut charts.`
    );
    assertIntegerInRange(opts.holeSize, `${path}.holeSize`, 0, 90);
  }
  if (opts.firstSliceAng !== undefined) {
    assertChartOptions(
      PIE_FAMILY_CHART_TYPES.has(opts.type),
      `${path}.firstSliceAng is only valid for pie, doughnut, and ofPie charts.`
    );
    assertIntegerInRange(opts.firstSliceAng, `${path}.firstSliceAng`, 0, 360);
  }
  if (opts.gapWidth !== undefined) {
    assertChartOptions(
      opts.type === "bar" || opts.type === "bar3D" || opts.type === "ofPie",
      `${path}.gapWidth is only valid for bar, bar3D, and ofPie charts.`
    );
    assertIntegerInRange(opts.gapWidth, `${path}.gapWidth`, 0, 500);
  }
  if (opts.overlap !== undefined) {
    assertChartOptions(
      opts.type === "bar" || opts.type === "bar3D",
      `${path}.overlap is only valid for bar and bar3D charts.`
    );
    assertIntegerInRange(opts.overlap, `${path}.overlap`, -100, 100);
  }
  if (opts.bubbleScale !== undefined) {
    assertChartOptions(
      opts.type === "bubble",
      `${path}.bubbleScale is only valid for bubble charts.`
    );
    assertIntegerInRange(opts.bubbleScale, `${path}.bubbleScale`, 0, 300);
  }
  if (opts.showNegBubbles !== undefined) {
    assertChartOptions(
      opts.type === "bubble",
      `${path}.showNegBubbles is only valid for bubble charts.`
    );
  }
  if (opts.sizeRepresents !== undefined) {
    assertChartOptions(
      opts.type === "bubble",
      `${path}.sizeRepresents is only valid for bubble charts.`
    );
  }
  if (opts.splitPos !== undefined) {
    assertChartOptions(opts.type === "ofPie", `${path}.splitPos is only valid for ofPie charts.`);
    assertFiniteNumber(opts.splitPos, `${path}.splitPos`);
  }
  if (opts.secondPieSize !== undefined) {
    assertChartOptions(
      opts.type === "ofPie",
      `${path}.secondPieSize is only valid for ofPie charts.`
    );
    assertIntegerInRange(opts.secondPieSize, `${path}.secondPieSize`, 5, 200);
  }
  if (opts.shape !== undefined) {
    assertChartOptions(opts.type === "bar3D", `${path}.shape is only valid for bar3D charts.`);
  }
  if (opts.showMarker !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "line3D" || opts.type === "radar",
      `${path}.showMarker is only valid for line, line3D, and radar charts.`
    );
  }
  if (opts.smooth !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "line3D" || opts.type === "scatter",
      `${path}.smooth is only valid for line, line3D, and scatter charts.`
    );
  }
  if (opts.hiLowLines !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "line3D" || opts.type === "stock",
      `${path}.hiLowLines is only valid for line, line3D, and stock charts.`
    );
  }
  if (opts.upDownBars !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "line3D" || opts.type === "stock",
      `${path}.upDownBars is only valid for line, line3D, and stock charts.`
    );
  }
  if (opts.dropLines !== undefined) {
    assertChartOptions(
      opts.type === "line" ||
        opts.type === "line3D" ||
        opts.type === "area" ||
        opts.type === "area3D" ||
        opts.type === "stock",
      `${path}.dropLines is only valid for line, area, and stock charts.`
    );
  }
  if (opts.serLines !== undefined) {
    assertChartOptions(
      opts.type === "bar" || opts.type === "ofPie",
      `${path}.serLines is only valid for bar and ofPie charts.`
    );
  }
  validateAxisOptions(opts.categoryAxis, `${path}.categoryAxis`);
  validateAxisOptions(opts.valueAxis, `${path}.valueAxis`);
}

function validateSharedChartOptions(opts: ChartModelOptions, path: string): void {
  if (opts.style !== undefined) {
    assertIntegerInRange(opts.style, `${path}.style`, 1, 48);
  }
}

function validateSeriesOptions(
  chartType: AddChartOptions["type"],
  opts: AddChartSeriesOptions,
  path: string
): void {
  assertChartOptions(!!opts && typeof opts === "object", `${path} must be an object.`);
  assertChartOptions(
    typeof opts.values === "string" && opts.values.length > 0,
    `${path}.values is required and must be a non-empty formula string.`
  );
  if (opts.trendline !== undefined) {
    assertChartOptions(
      !NO_TRENDLINE_CHART_TYPES.has(chartType),
      `${path}.trendline is not valid for ${chartType} charts.`
    );
    const trendlines = Array.isArray(opts.trendline) ? opts.trendline : [opts.trendline];
    for (let i = 0; i < trendlines.length; i++) {
      validateTrendlineOptions(
        trendlines[i],
        `${path}.trendline${trendlines.length > 1 ? `[${i}]` : ""}`
      );
    }
  }
  if (opts.marker?.size !== undefined) {
    assertIntegerInRange(opts.marker.size, `${path}.marker.size`, 2, 72);
  }
  if (opts.explosion !== undefined) {
    assertChartOptions(
      PIE_FAMILY_CHART_TYPES.has(chartType),
      `${path}.explosion is only valid for pie, doughnut, and ofPie charts.`
    );
    assertIntegerInRange(opts.explosion, `${path}.explosion`, 0, 400);
  }
  if (opts.bubble3D !== undefined) {
    assertChartOptions(chartType === "bubble", `${path}.bubble3D is only valid for bubble charts.`);
  }
  if (opts.bubbleSize !== undefined) {
    assertChartOptions(
      chartType === "bubble",
      `${path}.bubbleSize is only valid for bubble charts.`
    );
  }
  if (chartType === "bubble") {
    assertChartOptions(
      opts.xValues !== undefined,
      `${path}.xValues is required for bubble charts. Use a numeric range for x-values.`
    );
    assertChartOptions(
      opts.bubbleSize !== undefined,
      `${path}.bubbleSize is required for bubble charts. Use a numeric range for bubble sizes.`
    );
  }
  if (chartType === "scatter") {
    assertChartOptions(
      opts.xValues !== undefined,
      `${path}.xValues is required for scatter charts. Use a numeric range for x-values.`
    );
  }
  if (opts.xValues !== undefined) {
    assertChartOptions(
      chartType === "scatter" || chartType === "bubble",
      `${path}.xValues is only valid for scatter and bubble charts.`
    );
  }
  if (opts.dataPoints) {
    for (let i = 0; i < opts.dataPoints.length; i++) {
      validateDataPointOptions(chartType, opts.dataPoints[i], `${path}.dataPoints[${i}]`);
    }
  }
  if (opts.errorBars) {
    assertChartOptions(
      !PIE_FAMILY_CHART_TYPES.has(chartType),
      `${path}.errorBars is not valid for ${chartType} charts.`
    );
    const errorBars = Array.isArray(opts.errorBars) ? opts.errorBars : [opts.errorBars];
    for (let i = 0; i < errorBars.length; i++) {
      validateErrorBarsOptions(
        errorBars[i],
        `${path}.errorBars${errorBars.length > 1 ? `[${i}]` : ""}`
      );
    }
  }
}

function validateSeriesPatchOptions(
  chartType: AddChartOptions["type"],
  opts: Partial<AddChartSeriesOptions>,
  path: string
): void {
  if (opts.values !== undefined) {
    assertChartOptions(
      typeof opts.values === "string" && opts.values.length > 0,
      `${path}.values must be a non-empty formula string.`
    );
  }
  const fullOpts: AddChartSeriesOptions = {
    ...opts,
    values: opts.values ?? "__excelts_placeholder__"
  };
  if (chartType === "scatter" && fullOpts.xValues === undefined) {
    fullOpts.xValues = "__excelts_placeholder__";
  }
  if (chartType === "bubble") {
    if (fullOpts.xValues === undefined) {
      fullOpts.xValues = "__excelts_placeholder__";
    }
    if (fullOpts.bubbleSize === undefined) {
      fullOpts.bubbleSize = "__excelts_placeholder__";
    }
  }
  validateSeriesOptions(chartType, fullOpts, path);
}

function validateDataPointOptions(
  chartType: AddChartOptions["type"],
  opts: AddDataPointOptions,
  path: string
): void {
  assertIntegerInRange(opts.index, `${path}.index`, 0, Number.MAX_SAFE_INTEGER);
  if (opts.explosion !== undefined) {
    assertChartOptions(
      PIE_FAMILY_CHART_TYPES.has(chartType),
      `${path}.explosion is only valid for pie, doughnut, and ofPie charts.`
    );
    assertIntegerInRange(opts.explosion, `${path}.explosion`, 0, 400);
  }
  if (opts.marker?.size !== undefined) {
    assertIntegerInRange(opts.marker.size, `${path}.marker.size`, 2, 72);
  }
}

function validateTrendlineOptions(opts: AddTrendlineOptions, path: string): void {
  assertChartOptions(!!opts && typeof opts === "object", `${path} must be an object.`);
  assertChartOptions(!!opts.type, `${path}.type is required.`);
  if (opts.type === "poly") {
    assertIntegerInRange(opts.order, `${path}.order`, 2, 6);
  } else {
    assertChartOptions(
      opts.order === undefined,
      `${path}.order is only valid for polynomial trendlines.`
    );
  }
  if (opts.type === "movingAvg") {
    assertIntegerInRange(opts.period, `${path}.period`, 2, Number.MAX_SAFE_INTEGER);
  } else if (opts.period !== undefined) {
    assertIntegerInRange(opts.period, `${path}.period`, 2, Number.MAX_SAFE_INTEGER);
  }
  if (opts.forward !== undefined) {
    assertNumberInRange(opts.forward, `${path}.forward`, 0, Number.MAX_SAFE_INTEGER);
  }
  if (opts.backward !== undefined) {
    assertNumberInRange(opts.backward, `${path}.backward`, 0, Number.MAX_SAFE_INTEGER);
  }
}

function validateErrorBarsOptions(opts: AddErrorBarsOptions, path: string): void {
  assertChartOptions(!!opts && typeof opts === "object", `${path} must be an object.`);
  assertChartOptions(!!opts.type, `${path}.type is required.`);
  if (opts.type === "cust") {
    assertChartOptions(
      !!opts.plus && !!opts.minus,
      `${path}.plus and ${path}.minus are required when type is "cust".`
    );
  } else {
    assertChartOptions(
      opts.plus === undefined && opts.minus === undefined,
      `${path}.plus and ${path}.minus are only valid when type is "cust".`
    );
  }
  if (opts.value !== undefined) {
    assertNumberInRange(opts.value, `${path}.value`, 0, Number.MAX_SAFE_INTEGER);
  }
}

function validateAxisOptions(opts: AddAxisOptions | undefined, path: string): void {
  if (!opts) {
    return;
  }
  if (opts.min !== undefined) {
    assertFiniteNumber(opts.min, `${path}.min`);
  }
  if (opts.max !== undefined) {
    assertFiniteNumber(opts.max, `${path}.max`);
  }
  if (opts.min !== undefined && opts.max !== undefined) {
    assertChartOptions(opts.min < opts.max, `${path}.min must be less than ${path}.max.`);
  }
  if (opts.majorUnit !== undefined) {
    assertNumberInRange(opts.majorUnit, `${path}.majorUnit`, 0, Number.MAX_SAFE_INTEGER);
    assertChartOptions(opts.majorUnit > 0, `${path}.majorUnit must be greater than 0.`);
  }
  if (opts.minorUnit !== undefined) {
    assertNumberInRange(opts.minorUnit, `${path}.minorUnit`, 0, Number.MAX_SAFE_INTEGER);
    assertChartOptions(opts.minorUnit > 0, `${path}.minorUnit must be greater than 0.`);
  }
  if (opts.logBase !== undefined) {
    assertNumberInRange(opts.logBase, `${path}.logBase`, 2, 1000);
  }
  if (opts.textRotation !== undefined) {
    assertIntegerInRange(opts.textRotation, `${path}.textRotation`, -90, 90);
  }
  if (opts.lblOffset !== undefined) {
    assertIntegerInRange(opts.lblOffset, `${path}.lblOffset`, 0, 1000);
  }
  if (opts.tickLblSkip !== undefined) {
    assertIntegerInRange(opts.tickLblSkip, `${path}.tickLblSkip`, 1, Number.MAX_SAFE_INTEGER);
  }
  if (opts.tickMarkSkip !== undefined) {
    assertIntegerInRange(opts.tickMarkSkip, `${path}.tickMarkSkip`, 1, Number.MAX_SAFE_INTEGER);
  }
  if (opts.crossesAt !== undefined) {
    assertFiniteNumber(opts.crossesAt, `${path}.crossesAt`);
  }
  if (opts.customUnit !== undefined) {
    assertNumberInRange(opts.customUnit, `${path}.customUnit`, 0, Number.MAX_SAFE_INTEGER);
    assertChartOptions(opts.customUnit > 0, `${path}.customUnit must be greater than 0.`);
  }
}

function hexToColor(hex: string): ChartColor {
  return { srgb: hex.replace(/^#/, "").toUpperCase() };
}

function escapeXmlText(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildPivotSourceXml(source: PivotChartSource): string {
  const name = typeof source === "string" ? source : source.name;
  const fmtId = typeof source === "string" ? 0 : (source.fmtId ?? 0);
  if (!name) {
    throw new Error("Pivot chart source name is required.");
  }
  if (!Number.isInteger(fmtId) || fmtId < 0) {
    throw new Error("Pivot chart source fmtId must be a non-negative integer.");
  }
  // Pivot chart options used to be embedded inside `<c:pivotSource>` under a
  // private `xmlns:excelts` namespace; Excel never recognised that and the
  // parser never read it back. Options are now routed into
  // `ChartModel.pivotOptions` (see {@link ChartModel.pivotOptions}) and
  // serialised as MS standard `c14:pivotOptions` inside chartSpace's extLst.
  return `<c:pivotSource><c:name>${escapeXmlText(name)}</c:name><c:fmtId val="${fmtId}"/></c:pivotSource>`;
}

/**
 * Convert a simplified AddShapeFillOptions (or a pre-built ShapeProperties)
 * into a structured ShapeProperties object. Returns undefined when input is
 * undefined.
 */
function toShapeProperties(
  input: ShapeProperties | AddShapeFillOptions | undefined
): ShapeProperties | undefined {
  if (!input) {
    return undefined;
  }
  // If it already looks like a ShapeProperties (has fill/line/effectList/_rawXml
  // but no hex-string "fill" field), pass through.
  if (isShapeProperties(input)) {
    return input;
  }
  const opts = input as AddShapeFillOptions;
  const spPr: ShapeProperties = {};
  if (opts.noFill) {
    spPr.fill = { noFill: true };
  } else if (opts.fill) {
    spPr.fill = { solid: hexToColor(opts.fill) };
  } else if (opts.gradient) {
    spPr.fill = { gradient: opts.gradient };
  } else if (opts.pattern) {
    spPr.fill = { pattern: opts.pattern };
  }
  if (opts.border || opts.borderWidth !== undefined) {
    spPr.line = {};
    if (opts.border) {
      spPr.line.color = hexToColor(opts.border);
    }
    if (opts.borderWidth !== undefined) {
      spPr.line.width = opts.borderWidth * EMU_PER_POINT;
    }
  }
  return Object.keys(spPr).length > 0 ? spPr : undefined;
}

function isShapeProperties(v: unknown): v is ShapeProperties {
  if (!v || typeof v !== "object") {
    return false;
  }
  const o = v as Record<string, unknown>;
  // ShapeProperties has structured `fill` as an object or `line` or `_rawXml`,
  // whereas AddShapeFillOptions.fill is a hex string.
  if ("_rawXml" in o) {
    return true;
  }
  if ("fill" in o && typeof o.fill === "object") {
    return true;
  }
  if ("line" in o && typeof o.line === "object" && !Array.isArray(o.line)) {
    // AddShapeFillOptions does not have `line`; ShapeProperties does.
    return true;
  }
  if ("effectList" in o) {
    return true;
  }
  return false;
}

function makeSolidFill(hex: string): ShapeProperties {
  return { fill: { solid: hexToColor(hex) } };
}

function makeSeriesTx(
  name: string | { formula: string } | undefined
): { strRef?: StringReference; value?: string } | undefined {
  if (!name) {
    return undefined;
  }
  if (typeof name === "string") {
    return { value: name };
  }
  return { strRef: makeStrRef(name.formula) };
}

// ---------------------------------------------------------------------------
// Series-level option builders
// ---------------------------------------------------------------------------

function buildSeriesSpPr(opts: AddChartSeriesOptions): ShapeProperties | undefined {
  // Explicit `spPr` on the options object takes precedence over the sugared form.
  const custom = toShapeProperties(opts.spPr);
  if (custom) {
    return custom;
  }
  let spPr: ShapeProperties | undefined;
  if (opts.fill) {
    spPr = makeSolidFill(opts.fill);
  }
  if (opts.line || opts.lineWidth || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth) {
      line.width = Math.round(opts.lineWidth * EMU_PER_POINT); // pt → EMU
    }
    if (opts.lineDash) {
      line.dash = opts.lineDash;
    }
    spPr = { ...spPr, line };
  }
  return spPr;
}

function buildMarkerFromOpts(opts: AddChartMarkerOptions): ChartMarker {
  const m: ChartMarker = {};
  if (opts.symbol) {
    m.symbol = opts.symbol;
  }
  if (opts.size !== undefined) {
    m.size = opts.size;
  }
  if (opts.fill || opts.border) {
    const mSpPr: ShapeProperties = {};
    if (opts.fill) {
      mSpPr.fill = { solid: hexToColor(opts.fill) };
    }
    if (opts.border) {
      mSpPr.line = { color: hexToColor(opts.border) };
    }
    m.spPr = mSpPr;
  }
  return m;
}

function buildDataLabelsFromOpts(opts: AddDataLabelsOptions): DataLabels {
  const dl: DataLabels = {};
  if (opts.showLegendKey !== undefined) {
    dl.showLegendKey = opts.showLegendKey;
  }
  if (opts.showVal !== undefined) {
    dl.showVal = opts.showVal;
  }
  if (opts.showCatName !== undefined) {
    dl.showCatName = opts.showCatName;
  }
  if (opts.showSerName !== undefined) {
    dl.showSerName = opts.showSerName;
  }
  if (opts.showPercent !== undefined) {
    dl.showPercent = opts.showPercent;
  }
  if (opts.showBubbleSize !== undefined) {
    dl.showBubbleSize = opts.showBubbleSize;
  }
  if (opts.showLeaderLines !== undefined) {
    dl.showLeaderLines = opts.showLeaderLines;
  }
  if (opts.position) {
    dl.position = opts.position;
  }
  if (opts.separator) {
    dl.separator = opts.separator;
  }
  if (opts.numFmt) {
    dl.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  }
  const spPr = toShapeProperties(opts.spPr);
  if (spPr) {
    dl.spPr = spPr;
  }
  if (opts.txPr) {
    dl.txPr = opts.txPr;
  }
  if (opts.entries && opts.entries.length > 0) {
    dl.entries = opts.entries.map(e => buildDataLabelEntryFromOpts(e));
  }
  if (opts.valueFromCells !== undefined) {
    // "Value From Cells" (Excel 2013+). Accept either a bare formula or
    // the full {formula, cache} shape. The cache is filled in later by
    // `fillChartCaches` when the worksheet is available.
    dl.dataLabelsRange =
      typeof opts.valueFromCells === "string"
        ? { formula: opts.valueFromCells }
        : opts.valueFromCells;
  }
  return dl;
}

function buildDataLabelEntryFromOpts(opts: AddDataLabelEntryOptions): DataLabelEntry {
  const entry: DataLabelEntry = { index: opts.index };
  if (opts.delete) {
    entry.delete = true;
    // OOXML: a deleted label carries only its index; skip other fields.
    return entry;
  }
  if (opts.text !== undefined) {
    if (typeof opts.text === "string") {
      entry.text = {
        paragraphs: [{ runs: [{ text: opts.text }] }]
      };
    } else {
      entry.text = opts.text;
    }
  }
  if (opts.position) {
    entry.position = opts.position;
  }
  if (opts.numFmt) {
    entry.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  }
  const spPr = toShapeProperties(opts.spPr);
  if (spPr) {
    entry.spPr = spPr;
  }
  if (opts.txPr) {
    entry.txPr = opts.txPr;
  }
  if (opts.showVal !== undefined) {
    entry.showVal = opts.showVal;
  }
  if (opts.showCatName !== undefined) {
    entry.showCatName = opts.showCatName;
  }
  if (opts.showSerName !== undefined) {
    entry.showSerName = opts.showSerName;
  }
  if (opts.showPercent !== undefined) {
    entry.showPercent = opts.showPercent;
  }
  if (opts.showBubbleSize !== undefined) {
    entry.showBubbleSize = opts.showBubbleSize;
  }
  if (opts.showLegendKey !== undefined) {
    entry.showLegendKey = opts.showLegendKey;
  }
  return entry;
}

function buildTrendlineFromOpts(opts: AddTrendlineOptions): Trendline {
  const t: Trendline = { type: opts.type };
  if (opts.name) {
    t.name = opts.name;
  }
  if (opts.order !== undefined) {
    t.order = opts.order;
  }
  if (opts.period !== undefined) {
    t.period = opts.period;
  }
  if (opts.forward !== undefined) {
    t.forward = opts.forward;
  }
  if (opts.backward !== undefined) {
    t.backward = opts.backward;
  }
  if (opts.intercept !== undefined) {
    t.intercept = opts.intercept;
  }
  if (opts.displayRSqr !== undefined) {
    t.displayRSqr = opts.displayRSqr;
  }
  if (opts.displayEq !== undefined) {
    t.displayEq = opts.displayEq;
  }
  if (opts.line || opts.lineWidth || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth) {
      line.width = Math.round(opts.lineWidth * EMU_PER_POINT);
    }
    if (opts.lineDash) {
      line.dash = opts.lineDash;
    }
    t.spPr = { line };
  }
  if (opts.label) {
    t.trendlineLbl = buildTrendlineLabelFromOpts(opts.label);
  }
  return t;
}

function buildTrendlineLabelFromOpts(opts: AddTrendlineLabelOptions): TrendlineLabel {
  const lbl: TrendlineLabel = {};
  if (opts.text) {
    lbl.text = opts.text;
  }
  if (opts.numFmt) {
    lbl.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  }
  if (opts.layout) {
    lbl.layout = opts.layout;
  }
  const spPr = toShapeProperties(opts.spPr);
  if (spPr) {
    lbl.spPr = spPr;
  }
  if (opts.txPr) {
    lbl.txPr = opts.txPr;
  }
  return lbl;
}

function buildErrorBarsFromOpts(opts: AddErrorBarsOptions): ErrorBars {
  const eb: ErrorBars = {
    barDir: opts.barDir ?? "both",
    errValType: opts.type
  };
  if (opts.direction) {
    eb.errDir = opts.direction;
  }
  if (opts.value !== undefined) {
    eb.val = opts.value;
  }
  if (opts.noEndCap !== undefined) {
    eb.noEndCap = opts.noEndCap;
  }
  if (opts.plus) {
    eb.plus = makeNumData(opts.plus);
  }
  if (opts.minus) {
    eb.minus = makeNumData(opts.minus);
  }
  // Line / shape styling
  const customSpPr = toShapeProperties(opts.spPr);
  if (customSpPr) {
    eb.spPr = customSpPr;
  } else if (opts.line || opts.lineWidth || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth) {
      line.width = Math.round(opts.lineWidth * EMU_PER_POINT);
    }
    if (opts.lineDash) {
      line.dash = opts.lineDash;
    }
    eb.spPr = { line };
  }
  return eb;
}

function buildDataPointFromOpts(opts: AddDataPointOptions): DataPoint {
  const dp: DataPoint = { index: opts.index };
  if (opts.fill || opts.border) {
    const spPr: ShapeProperties = {};
    if (opts.fill) {
      spPr.fill = { solid: hexToColor(opts.fill) };
    }
    if (opts.border) {
      spPr.line = { color: hexToColor(opts.border) };
    }
    dp.spPr = spPr;
  }
  if (opts.explosion !== undefined) {
    dp.explosion = opts.explosion;
  }
  if (opts.bubble3D !== undefined) {
    dp.bubble3D = opts.bubble3D;
  }
  if (opts.marker) {
    dp.marker = buildMarkerFromOpts(opts.marker);
  }
  if (opts.invertIfNegative !== undefined) {
    dp.invertIfNegative = opts.invertIfNegative;
  }
  return dp;
}

/**
 * Apply common series-level options from AddChartSeriesOptions to any series object.
 */
function applySeriesOptions(s: any, opts: AddChartSeriesOptions): void {
  s.spPr = buildSeriesSpPr(opts);
  if (opts.marker) {
    s.marker = buildMarkerFromOpts(opts.marker);
  }
  if (opts.dataLabels) {
    s.dataLabels = buildDataLabelsFromOpts(opts.dataLabels);
  }
  if (opts.trendline) {
    const trendlineOpts = Array.isArray(opts.trendline) ? opts.trendline : [opts.trendline];
    s.trendlines = trendlineOpts.map(buildTrendlineFromOpts);
  }
  if (opts.dataPoints) {
    s.dataPoints = opts.dataPoints.map(buildDataPointFromOpts);
  }
  if (opts.txPr) {
    s.txPr = opts.txPr;
  }
  if (opts.pictureFill) {
    s.pictureOptions = {
      applyToFront: opts.pictureFill.applyToFront,
      applyToSides: opts.pictureFill.applyToSides,
      applyToEnd: opts.pictureFill.applyToEnd,
      pictureFormat: opts.pictureFill.fillMode,
      pictureStackUnit: opts.pictureFill.scale
    };
    // Route the blipFill itself through spPr.fill.blip. The two halves
    // of the OOXML "picture on bars" setting live in different places:
    // `c:pictureOptions` (stretch/stack/scale/apply-to-*) and
    // `a:blipFill` (the actual image rel) — we unify them under the
    // single `pictureFill` option for ergonomics and split them back
    // here.
    if (opts.pictureFill.image !== undefined || opts.pictureFill.relationshipId) {
      s.spPr = s.spPr ?? {};
      s.spPr.fill = s.spPr.fill ?? {};
      s.spPr.fill.blip = {
        fillMode:
          opts.pictureFill.fillMode === "stack" || opts.pictureFill.fillMode === "stackScale"
            ? "tile"
            : "stretch",
        ...(opts.pictureFill.relationshipId
          ? { relationshipId: opts.pictureFill.relationshipId }
          : {}),
        ...(opts.pictureFill.image !== undefined ? { _pendingImage: opts.pictureFill.image } : {})
      };
    }
  }
}

function buildBarSeries(opts: AddChartSeriesOptions, idx: number): BarSeries {
  const s: BarSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  if (opts.invertIfNegative !== undefined) {
    s.invertIfNegative = opts.invertIfNegative;
  }
  if (opts.errorBars) {
    s.errorBars = Array.isArray(opts.errorBars)
      ? buildErrorBarsFromOpts(opts.errorBars[0])
      : buildErrorBarsFromOpts(opts.errorBars);
  }
  return s;
}

function buildLineSeries(opts: AddChartSeriesOptions, idx: number): LineSeries {
  const s: LineSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  if (opts.smooth !== undefined) {
    s.smooth = opts.smooth;
  }
  if (opts.errorBars) {
    s.errorBars = Array.isArray(opts.errorBars)
      ? buildErrorBarsFromOpts(opts.errorBars[0])
      : buildErrorBarsFromOpts(opts.errorBars);
  }
  return s;
}

function buildPieSeries(opts: AddChartSeriesOptions, idx: number): PieSeries {
  const s: PieSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  if (opts.explosion !== undefined) {
    s.explosion = opts.explosion;
  }
  return s;
}

function buildAreaSeries(opts: AddChartSeriesOptions, idx: number): AreaSeries {
  const s: AreaSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  if (opts.errorBars) {
    s.errorBars = Array.isArray(opts.errorBars)
      ? buildErrorBarsFromOpts(opts.errorBars[0])
      : buildErrorBarsFromOpts(opts.errorBars);
  }
  return s;
}

function buildScatterSeries(opts: AddChartSeriesOptions, idx: number): ScatterSeries {
  const s: ScatterSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.xValues) {
    s.xVal = makeXAxisData(opts.xValues, opts.xValueType);
  }
  if (opts.values) {
    s.yVal = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  if (opts.smooth !== undefined) {
    s.smooth = opts.smooth;
  }
  if (opts.errorBars) {
    const ebOpts = Array.isArray(opts.errorBars) ? opts.errorBars : [opts.errorBars];
    s.errorBars = ebOpts.map(buildErrorBarsFromOpts);
  }
  return s;
}

function buildBubbleSeries(opts: AddChartSeriesOptions, idx: number): BubbleSeries {
  const s: BubbleSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.xValues) {
    s.xVal = makeXAxisData(opts.xValues, opts.xValueType);
  }
  if (opts.values) {
    s.yVal = makeNumData(opts.values);
  }
  if (opts.bubbleSize) {
    s.bubbleSize = makeNumData(opts.bubbleSize);
  }
  applySeriesOptions(s, opts);
  if (opts.invertIfNegative !== undefined) {
    s.invertIfNegative = opts.invertIfNegative;
  }
  if (opts.bubble3D !== undefined) {
    s.bubble3D = opts.bubble3D;
  }
  if (opts.errorBars) {
    const ebOpts = Array.isArray(opts.errorBars) ? opts.errorBars : [opts.errorBars];
    s.errorBars = ebOpts.map(buildErrorBarsFromOpts);
  }
  return s;
}

function buildRadarSeries(opts: AddChartSeriesOptions, idx: number): RadarSeries {
  const s: RadarSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  return s;
}

function buildSurfaceSeries(opts: AddChartSeriesOptions, idx: number): SurfaceSeries {
  const s: SurfaceSeries = { index: idx, order: idx };
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
  applySeriesOptions(s, opts);
  return s;
}

export function buildChartSeriesForType(
  chartType: AddChartOptions["type"],
  options: AddChartSeriesOptions,
  index: number
): SeriesBase {
  validateSeriesOptions(chartType, options, "series");
  if (chartType === "bar" || chartType === "bar3D") {
    return buildBarSeries(options, index);
  }
  if (chartType === "line" || chartType === "line3D" || chartType === "stock") {
    return buildLineSeries(options, index);
  }
  if (
    chartType === "pie" ||
    chartType === "pie3D" ||
    chartType === "doughnut" ||
    chartType === "ofPie"
  ) {
    return buildPieSeries(options, index);
  }
  if (chartType === "area" || chartType === "area3D") {
    return buildAreaSeries(options, index);
  }
  if (chartType === "scatter") {
    return buildScatterSeries(options, index);
  }
  if (chartType === "bubble") {
    return buildBubbleSeries(options, index);
  }
  if (chartType === "radar") {
    return buildRadarSeries(options, index);
  }
  if (chartType === "surface" || chartType === "surface3D") {
    return buildSurfaceSeries(options, index);
  }
  return buildBarSeries(options, index);
}

export function applyChartSeriesOptionsPatch(
  series: SeriesBase,
  options: Partial<AddChartSeriesOptions>,
  chartType?: AddChartOptions["type"]
): void {
  if (chartType) {
    validateSeriesPatchOptions(chartType, options, "series");
  }
  if (options.name !== undefined) {
    series.tx = makeSeriesTx(options.name);
  }
  if (options.categories !== undefined) {
    const target = series as SeriesBase & { cat?: AxisDataSource; xVal?: AxisDataSource };
    if (target.xVal && !target.cat) {
      target.xVal = makeNumericAxisData(options.categories);
    } else {
      target.cat = makeAxisData(options.categories);
    }
  }
  if (options.xValues !== undefined) {
    (series as SeriesBase & { xVal?: AxisDataSource }).xVal = makeXAxisData(
      options.xValues,
      options.xValueType
    );
  }
  if (options.values !== undefined) {
    const target = series as SeriesBase & { val?: NumberDataSource; yVal?: NumberDataSource };
    if (target.yVal && !target.val) {
      target.yVal = makeNumData(options.values);
    } else {
      target.val = makeNumData(options.values);
    }
  }
  if (options.bubbleSize !== undefined) {
    (series as SeriesBase & { bubbleSize?: NumberDataSource }).bubbleSize = makeNumData(
      options.bubbleSize
    );
  }
  if (hasSeriesShapePatch(options)) {
    series.spPr = buildSeriesSpPr(options as AddChartSeriesOptions);
  }
  if (options.marker !== undefined) {
    (series as SeriesBase & { marker?: ChartMarker }).marker = buildMarkerFromOpts(options.marker);
  }
  if (options.dataLabels !== undefined) {
    (series as SeriesBase & { dataLabels?: DataLabels }).dataLabels = buildDataLabelsFromOpts(
      options.dataLabels
    );
  }
  if (options.trendline !== undefined) {
    const trendlineOpts = Array.isArray(options.trendline)
      ? options.trendline
      : [options.trendline];
    (series as SeriesBase & { trendlines?: Trendline[] }).trendlines =
      trendlineOpts.map(buildTrendlineFromOpts);
  }
  if (options.dataPoints !== undefined) {
    (series as SeriesBase & { dataPoints?: DataPoint[] }).dataPoints =
      options.dataPoints.map(buildDataPointFromOpts);
  }
  if (options.errorBars !== undefined) {
    const errorBars = Array.isArray(options.errorBars) ? options.errorBars : [options.errorBars];
    const built = errorBars.map(buildErrorBarsFromOpts);
    (series as SeriesBase & { errorBars?: ErrorBars | ErrorBars[] }).errorBars =
      chartType === "scatter" || chartType === "bubble" ? built : built[0];
  }
  if (options.txPr !== undefined) {
    (series as SeriesBase & { txPr?: ChartTextProperties }).txPr = options.txPr;
  }
  if (options.pictureFill !== undefined) {
    (series as SeriesBase & { pictureOptions?: unknown }).pictureOptions = {
      applyToFront: options.pictureFill.applyToFront,
      applyToSides: options.pictureFill.applyToSides,
      applyToEnd: options.pictureFill.applyToEnd,
      pictureFormat: options.pictureFill.fillMode,
      pictureStackUnit: options.pictureFill.scale
    };
  }
  if (options.smooth !== undefined) {
    (series as SeriesBase & { smooth?: boolean }).smooth = options.smooth;
  }
  if (options.invertIfNegative !== undefined) {
    (series as SeriesBase & { invertIfNegative?: boolean }).invertIfNegative =
      options.invertIfNegative;
  }
  if (options.explosion !== undefined) {
    (series as SeriesBase & { explosion?: number }).explosion = options.explosion;
  }
  if (options.bubble3D !== undefined) {
    (series as SeriesBase & { bubble3D?: boolean }).bubble3D = options.bubble3D;
  }
}

function hasSeriesShapePatch(options: Partial<AddChartSeriesOptions>): boolean {
  return (
    options.spPr !== undefined ||
    options.fill !== undefined ||
    options.line !== undefined ||
    options.lineWidth !== undefined ||
    options.lineDash !== undefined
  );
}

function buildTitle(input: string | { formula: string } | ChartRichText): ChartTitle {
  if (typeof input === "string") {
    return {
      text: {
        paragraphs: [{ runs: [{ text: input }] }]
      },
      overlay: false
    };
  }
  if ("formula" in input) {
    return {
      strRef: makeStrRef(input.formula),
      overlay: false
    };
  }
  // ChartRichText
  return {
    text: input,
    overlay: false
  };
}

function buildDataTableFromOpts(opts: AddChartOptions["dataTable"]): DataTable | undefined {
  if (!opts) {
    return undefined;
  }
  if (opts === true) {
    return { showHorzBorder: true, showVertBorder: true, showOutline: true, showKeys: true };
  }
  return {
    showHorzBorder: opts.showHorzBorder,
    showVertBorder: opts.showVertBorder,
    showOutline: opts.showOutline,
    showKeys: opts.showKeys
  };
}

function buildUpDownBarsFromOpts(opts: AddChartOptions["upDownBars"]): UpDownBars | undefined {
  if (!opts) {
    return undefined;
  }
  if (opts === true) {
    return { gapWidth: 150 };
  }
  const udb: UpDownBars = { gapWidth: opts.gapWidth ?? 150 };
  const upBars = toShapeProperties(opts.upBars);
  if (upBars) {
    udb.upBars = upBars;
  }
  const downBars = toShapeProperties(opts.downBars);
  if (downBars) {
    udb.downBars = downBars;
  }
  return udb;
}

function buildLegend(opts: {
  showLegend?: boolean;
  legendPosition?: LegendPosition;
}): ChartLegend | undefined {
  if (opts.showLegend === false) {
    return undefined;
  }
  return {
    legendPos: opts.legendPosition ?? "b",
    overlay: false
  };
}

/**
 * Apply user-specified axis options to a built axis.
 */
function applyAxisOptions(axis: ChartAxis, opts: AddAxisOptions | undefined): void {
  if (!opts) {
    return;
  }
  if (opts.title) {
    axis.title = buildTitle(opts.title);
  }
  if (opts.titleOptions) {
    if (!axis.title) {
      axis.title = { overlay: false };
    }
    applyTitleOptions(axis.title, opts.titleOptions);
  }
  if (opts.numFmt) {
    axis.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  }
  if (opts.min !== undefined || opts.max !== undefined || opts.orientation || opts.logBase) {
    if (!axis.scaling) {
      axis.scaling = {};
    }
    if (opts.min !== undefined) {
      axis.scaling.min = opts.min;
    }
    if (opts.max !== undefined) {
      axis.scaling.max = opts.max;
    }
    if (opts.orientation) {
      axis.scaling.orientation = opts.orientation;
    }
    if (opts.logBase !== undefined) {
      axis.scaling.logBase = opts.logBase;
    }
  }
  if (opts.majorUnit !== undefined && axis.axisType === "val") {
    (axis as ValueAxis).majorUnit = opts.majorUnit;
  }
  if (opts.minorUnit !== undefined && axis.axisType === "val") {
    (axis as ValueAxis).minorUnit = opts.minorUnit;
  }
  if (opts.majorTickMark) {
    axis.majorTickMark = opts.majorTickMark;
  }
  if (opts.minorTickMark) {
    axis.minorTickMark = opts.minorTickMark;
  }
  if (opts.tickLblPos) {
    axis.tickLblPos = opts.tickLblPos;
  }
  // Major gridlines: boolean toggle OR structured style
  const majorGridlinesStyle = toShapeProperties(opts.majorGridlinesStyle);
  if (majorGridlinesStyle) {
    axis.majorGridlines = majorGridlinesStyle;
  } else if (opts.majorGridlines !== undefined) {
    axis.majorGridlines = opts.majorGridlines ? {} : undefined;
  }
  const minorGridlinesStyle = toShapeProperties(opts.minorGridlinesStyle);
  if (minorGridlinesStyle) {
    axis.minorGridlines = minorGridlinesStyle;
  } else if (opts.minorGridlines !== undefined) {
    axis.minorGridlines = opts.minorGridlines ? {} : undefined;
  }
  if (opts.hidden !== undefined) {
    axis.delete = opts.hidden;
  }
  if (opts.crossBetween !== undefined && axis.axisType === "val") {
    (axis as ValueAxis).crossBetween = opts.crossBetween;
  }
  // Text properties — structured txPr takes priority over textRotation
  if (opts.txPr) {
    axis.txPr = opts.txPr;
  } else if (opts.textRotation !== undefined) {
    axis.txPr = { rotation: opts.textRotation * 60000 };
  }
  if (opts.lblAlgn !== undefined && axis.axisType === "cat") {
    (axis as CategoryAxis).lblAlgn = opts.lblAlgn;
  }
  if (opts.lblOffset !== undefined && axis.axisType === "cat") {
    (axis as CategoryAxis).lblOffset = opts.lblOffset;
  }
  if (opts.tickLblSkip !== undefined && axis.axisType === "cat") {
    (axis as CategoryAxis).tickLblSkip = opts.tickLblSkip;
  }
  if (opts.tickMarkSkip !== undefined && axis.axisType === "cat") {
    (axis as CategoryAxis).tickMarkSkip = opts.tickMarkSkip;
  }
  if (opts.crosses !== undefined) {
    axis.crosses = opts.crosses;
  }
  if (opts.crossesAt !== undefined) {
    axis.crossesAt = opts.crossesAt;
  }
  if (axis.axisType === "val") {
    const valAx = axis as ValueAxis;
    if (
      opts.displayUnits !== undefined ||
      opts.customUnit !== undefined ||
      opts.displayUnitsLabel
    ) {
      valAx.dispUnits = valAx.dispUnits ?? {};
      if (opts.displayUnits !== undefined) {
        valAx.dispUnits.builtInUnit = opts.displayUnits;
      }
      if (opts.customUnit !== undefined) {
        valAx.dispUnits.custUnit = opts.customUnit;
      }
      if (opts.displayUnitsLabel !== undefined) {
        valAx.dispUnits.label = buildTitle(opts.displayUnitsLabel);
      }
    }
  }
  // Date-axis-specific units
  if (axis.axisType === "date") {
    const dateAx = axis as DateAxis;
    if (opts.baseTimeUnit) {
      dateAx.baseTimeUnit = opts.baseTimeUnit;
    }
    if (opts.majorTimeUnit) {
      dateAx.majorTimeUnit = opts.majorTimeUnit;
    }
    if (opts.minorTimeUnit) {
      dateAx.minorTimeUnit = opts.minorTimeUnit;
    }
  }
  // Structured spPr takes priority over line-only shortcuts
  const customSpPr = toShapeProperties(opts.spPr);
  if (customSpPr) {
    axis.spPr = customSpPr;
  } else if (opts.lineColor || opts.lineWidth || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.lineColor) {
      line.color = hexToColor(opts.lineColor);
    }
    if (opts.lineWidth) {
      line.width = Math.round(opts.lineWidth * EMU_PER_POINT); // pt to EMU
    }
    if (opts.lineDash) {
      line.dash = opts.lineDash;
    }
    if (!axis.spPr) {
      axis.spPr = {};
    }
    axis.spPr.line = line;
  }
}

function buildCatValAxes(axIds: AxIdAllocator): { catAx: CategoryAxis; valAx: ValueAxis } {
  const catAxId = axIds.alloc();
  const valAxId = axIds.alloc();
  return {
    catAx: {
      axisType: "cat",
      axId: catAxId,
      scaling: { orientation: "minMax" },
      delete: false,
      axPos: "b",
      crossAx: valAxId,
      crosses: "autoZero",
      auto: true,
      lblAlgn: "ctr",
      lblOffset: 100
    },
    valAx: {
      axisType: "val",
      axId: valAxId,
      scaling: { orientation: "minMax" },
      delete: false,
      axPos: "l",
      crossAx: catAxId,
      crosses: "autoZero",
      crossBetween: "between",
      majorGridlines: {}
    }
  };
}

function buildValValAxes(axIds: AxIdAllocator): { xAx: ValueAxis; yAx: ValueAxis } {
  const xAxId = axIds.alloc();
  const yAxId = axIds.alloc();
  return {
    xAx: {
      axisType: "val",
      axId: xAxId,
      scaling: { orientation: "minMax" },
      delete: false,
      axPos: "b",
      crossAx: yAxId,
      crosses: "autoZero"
    },
    yAx: {
      axisType: "val",
      axId: yAxId,
      scaling: { orientation: "minMax" },
      delete: false,
      axPos: "l",
      crossAx: xAxId,
      crosses: "autoZero",
      majorGridlines: {}
    }
  };
}

function buildCatValSerAxes(axIds: AxIdAllocator): {
  catAx: CategoryAxis;
  valAx: ValueAxis;
  serAx: SeriesAxis;
} {
  const { catAx, valAx } = buildCatValAxes(axIds);
  const serAxId = axIds.alloc();
  // serAx crosses the catAx
  const serAx: SeriesAxis = {
    axisType: "ser",
    axId: serAxId,
    scaling: { orientation: "minMax" },
    delete: false,
    axPos: "b",
    crossAx: catAx.axId,
    crosses: "autoZero"
  };
  return { catAx, valAx, serAx };
}

function buildChartTypeGroup(
  opts: AddChartOptions,
  seriesOpts: AddChartSeriesOptions[],
  axIds: AxIdAllocator
): { group: ChartTypeGroup; axes: ChartAxis[] } {
  const type = opts.type;
  let result: { group: ChartTypeGroup; axes: ChartAxis[] };

  switch (type) {
    case "bar": {
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: BarChartGroup = {
        type,
        barDir: opts.barDir ?? "col",
        grouping: (opts.grouping as BarGrouping) ?? "clustered",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildBarSeries),
        gapWidth: opts.gapWidth ?? 150,
        overlap: opts.overlap,
        serLines: opts.serLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId]
      };
      result = { group, axes: [catAx, valAx] };
      break;
    }
    case "bar3D": {
      const { catAx, valAx, serAx } = buildCatValSerAxes(axIds);
      const group: BarChartGroup = {
        type,
        barDir: opts.barDir ?? "col",
        grouping: (opts.grouping as BarGrouping) ?? "clustered",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildBarSeries),
        gapWidth: opts.gapWidth ?? 150,
        overlap: opts.overlap,
        shape: opts.shape,
        axisIds: [catAx.axId, valAx.axId, serAx.axId]
      };
      result = { group, axes: [catAx, valAx, serAx] };
      break;
    }
    case "line": {
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: LineChartGroup = {
        type,
        grouping: (opts.grouping as LineGrouping) ?? "standard",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildLineSeries),
        marker: opts.showMarker ?? true,
        smooth: opts.smooth,
        hiLowLines: opts.hiLowLines ? {} : undefined,
        upDownBars: buildUpDownBarsFromOpts(opts.upDownBars),
        dropLines: opts.dropLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId]
      };
      result = { group, axes: [catAx, valAx] };
      break;
    }
    case "line3D": {
      const { catAx, valAx, serAx } = buildCatValSerAxes(axIds);
      const group: LineChartGroup = {
        type,
        grouping: (opts.grouping as LineGrouping) ?? "standard",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildLineSeries),
        marker: opts.showMarker ?? true,
        smooth: opts.smooth,
        hiLowLines: opts.hiLowLines ? {} : undefined,
        upDownBars: buildUpDownBarsFromOpts(opts.upDownBars),
        dropLines: opts.dropLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId, serAx.axId]
      };
      result = { group, axes: [catAx, valAx, serAx] };
      break;
    }
    case "pie":
    case "pie3D": {
      const group: PieChartGroup = {
        type,
        varyColors: opts.varyColors ?? true,
        series: seriesOpts.map(buildPieSeries),
        firstSliceAng: opts.firstSliceAng ?? 0
      };
      result = { group, axes: [] };
      break;
    }
    case "doughnut": {
      const group: DoughnutChartGroup = {
        type: "doughnut",
        varyColors: opts.varyColors ?? true,
        series: seriesOpts.map(buildPieSeries),
        firstSliceAng: opts.firstSliceAng ?? 0,
        holeSize: opts.holeSize ?? 50
      };
      result = { group, axes: [] };
      break;
    }
    case "area": {
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: AreaChartGroup = {
        type,
        grouping: (opts.grouping as LineGrouping) ?? "standard",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildAreaSeries),
        dropLines: opts.dropLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId]
      };
      result = { group, axes: [catAx, valAx] };
      break;
    }
    case "area3D": {
      const { catAx, valAx, serAx } = buildCatValSerAxes(axIds);
      const group: AreaChartGroup = {
        type,
        grouping: (opts.grouping as LineGrouping) ?? "standard",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildAreaSeries),
        dropLines: opts.dropLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId, serAx.axId]
      };
      result = { group, axes: [catAx, valAx, serAx] };
      break;
    }
    case "scatter": {
      const { xAx, yAx } = buildValValAxes(axIds);
      const group: ScatterChartGroup = {
        type: "scatter",
        scatterStyle: opts.scatterStyle ?? "lineMarker",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildScatterSeries),
        axisIds: [xAx.axId, yAx.axId]
      };
      result = { group, axes: [xAx, yAx] };
      break;
    }
    case "bubble": {
      const { xAx, yAx } = buildValValAxes(axIds);
      const group: BubbleChartGroup = {
        type: "bubble",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildBubbleSeries),
        bubbleScale: opts.bubbleScale,
        showNegBubbles: opts.showNegBubbles,
        sizeRepresents: opts.sizeRepresents,
        axisIds: [xAx.axId, yAx.axId]
      };
      result = { group, axes: [xAx, yAx] };
      break;
    }
    case "radar": {
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: RadarChartGroup = {
        type: "radar",
        radarStyle: opts.radarStyle ?? "marker",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildRadarSeries),
        axisIds: [catAx.axId, valAx.axId]
      };
      result = { group, axes: [catAx, valAx] };
      break;
    }
    case "stock": {
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: StockChartGroup = {
        type: "stock",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildLineSeries),
        hiLowLines: opts.hiLowLines ? {} : undefined,
        upDownBars: buildUpDownBarsFromOpts(opts.upDownBars),
        dropLines: opts.dropLines ? {} : undefined,
        axisIds: [catAx.axId, valAx.axId]
      };
      result = { group, axes: [catAx, valAx] };
      break;
    }
    case "surface":
    case "surface3D": {
      const { catAx, valAx, serAx } = buildCatValSerAxes(axIds);
      const group: SurfaceChartGroup = {
        type,
        wireframe: opts.wireframe,
        series: seriesOpts.map(buildSurfaceSeries),
        axisIds: [catAx.axId, valAx.axId, serAx.axId]
      };
      if (opts.bandFormats && opts.bandFormats.length > 0) {
        group.bandFormats = opts.bandFormats.map(bf => {
          const spPr = toShapeProperties(bf.spPr);
          return { index: bf.index, spPr: spPr ?? {} };
        });
      }
      result = { group, axes: [catAx, valAx, serAx] };
      break;
    }
    case "ofPie": {
      const group: OfPieChartGroup = {
        type: "ofPie",
        ofPieType: opts.ofPieType ?? "pie",
        varyColors: opts.varyColors ?? true,
        series: seriesOpts.map(buildPieSeries),
        gapWidth: opts.gapWidth,
        splitType: opts.splitType,
        splitPos: opts.splitPos,
        secondPieSize: opts.secondPieSize,
        serLines: opts.serLines ? {} : undefined
      };
      result = { group, axes: [] };
      break;
    }
    default: {
      const _exhaustive: never = type;
      throw new Error(`Unsupported chart type: ${_exhaustive}`);
    }
  }

  // Post-process: apply group-level data labels
  if (opts.dataLabels) {
    (result.group as { dataLabels?: DataLabels }).dataLabels = buildDataLabelsFromOpts(
      opts.dataLabels
    );
  }

  // Post-process: apply axis options
  if (result.axes.length > 0) {
    // For cat+val axes, apply categoryAxis options to first cat/date axis,
    // and valueAxis options to first val axis.
    // For scatter/bubble (val+val), categoryAxis → first val (x), valueAxis → second val (y).
    const catAx = result.axes.find(a => a.axisType === "cat" || a.axisType === "date");
    const valAx = result.axes.find(a => a.axisType === "val");
    if (catAx) {
      applyAxisOptions(catAx, opts.categoryAxis);
      const vAx = result.axes.find(a => a.axisType === "val");
      if (vAx) {
        applyAxisOptions(vAx, opts.valueAxis);
      }
    } else if (valAx) {
      // scatter/bubble: x axis = first val, y axis = second val
      applyAxisOptions(valAx, opts.categoryAxis);
      const secondVal = result.axes.find(a => a.axisType === "val" && a.axId !== valAx.axId);
      if (secondVal) {
        applyAxisOptions(secondVal, opts.valueAxis);
      }
    }
  }

  return result;
}

/**
 * Shared options for both single and combo chart builders.
 */
interface ChartModelOptions {
  title?: string | { formula: string } | ChartRichText;
  showLegend?: boolean;
  legendPosition?: LegendPosition;
  displayBlanksAs?: "gap" | "span" | "zero";
  view3D?: View3D;
  style?: number;
  plotVisOnly?: boolean;
  showDLblsOverMax?: boolean;
  titleOptions?: AddTitleOptions;
  legendOptions?: AddLegendOptions;
  plotAreaOptions?: AddPlotAreaOptions;
  pivotSource?: PivotChartSource;
  pivotChartOptions?: PivotChartOptions;
  floor?: ShapeProperties | AddShapeFillOptions;
  sideWall?: ShapeProperties | AddShapeFillOptions;
  backWall?: ShapeProperties | AddShapeFillOptions;
}

/**
 * Finalize a ChartModel from a PlotArea and shared options.
 */
function finalizeChartModel(plotArea: PlotArea, opts: ChartModelOptions): ChartModel {
  const chart: ChartData = {
    plotArea,
    plotVisOnly: opts.plotVisOnly ?? true,
    dispBlanksAs: opts.displayBlanksAs ?? "gap"
  };

  if (opts.showDLblsOverMax !== undefined) {
    chart.showDLblsOverMax = opts.showDLblsOverMax;
  }
  if (opts.pivotSource) {
    if (
      typeof opts.pivotSource === "object" &&
      opts.pivotChartOptions &&
      !opts.pivotSource.options
    ) {
      opts.pivotSource.options = opts.pivotChartOptions;
    }
    chart.pivotFormats = [{ index: 0 }];
  }

  if (opts.title) {
    chart.title = buildTitle(opts.title);
    // Apply title options (layout, overlay, spPr, txPr) if provided
    if (opts.titleOptions) {
      applyTitleOptions(chart.title, opts.titleOptions);
    }
    chart.autoTitleDeleted = false;
  } else if (opts.titleOptions) {
    // Layout/style without title text is unusual but allowed
    chart.title = { overlay: false };
    applyTitleOptions(chart.title, opts.titleOptions);
    chart.autoTitleDeleted = false;
  } else {
    chart.autoTitleDeleted = true;
  }

  const legend = buildLegend(opts);
  if (legend) {
    chart.legend = legend;
    if (opts.legendOptions) {
      applyLegendOptions(legend, opts.legendOptions);
    }
  }

  // Plot area options (layout, background)
  if (opts.plotAreaOptions) {
    if (opts.plotAreaOptions.layout) {
      plotArea.layout = opts.plotAreaOptions.layout;
    }
    const plotSpPr = toShapeProperties(opts.plotAreaOptions.spPr);
    if (plotSpPr) {
      plotArea.spPr = plotSpPr;
    }
  }

  if (opts.view3D) {
    chart.view3D = opts.view3D;
  }
  // 3D walls / floor
  const floor = toShapeProperties(opts.floor);
  if (floor) {
    chart.floor = floor;
  }
  const sideWall = toShapeProperties(opts.sideWall);
  if (sideWall) {
    chart.sideWall = sideWall;
  }
  const backWall = toShapeProperties(opts.backWall);
  if (backWall) {
    chart.backWall = backWall;
  }

  return {
    chart,
    style: opts.style,
    roundedCorners: false,
    lang: "en-US",
    pivotSource: opts.pivotSource ? buildPivotSourceXml(opts.pivotSource) : undefined,
    pivotOptions: extractPivotOptions(opts.pivotSource, opts.pivotChartOptions)
  };
}

/**
 * Resolve the effective {@link PivotChartOptions} for a model built from
 * {@link AddChartOptions}.
 *
 * Accepts two redundant inputs so callers can pass the structured metadata
 * via either the top-level `pivotChartOptions` field or the `pivotSource`
 * object (its `options` sub-field, retained for ergonomic clustering).
 * When both are set the top-level value wins — this matches the precedence
 * used elsewhere in `chart-builder.ts` when the same setting has two
 * spellings.
 *
 * Returns `undefined` when no options are provided so the writer can skip
 * emitting an empty `c14:pivotOptions` element.
 */
function extractPivotOptions(
  source: PivotChartSource | undefined,
  explicit: PivotChartOptions | undefined
): PivotChartOptions | undefined {
  if (explicit) {
    return explicit;
  }
  if (source && typeof source === "object" && source.options) {
    return source.options;
  }
  return undefined;
}

function applyTitleOptions(title: ChartTitle, opts: AddTitleOptions): void {
  if (opts.layout) {
    title.layout = opts.layout;
  }
  if (opts.overlay !== undefined) {
    title.overlay = opts.overlay;
  }
  const spPr = toShapeProperties(opts.spPr);
  if (spPr) {
    title.spPr = spPr;
  }
  if (opts.txPr) {
    title.txPr = opts.txPr;
  }
}

function applyLegendOptions(legend: ChartLegend, opts: AddLegendOptions): void {
  if (opts.layout) {
    legend.layout = opts.layout;
  }
  if (opts.overlay !== undefined) {
    legend.overlay = opts.overlay;
  }
  const spPr = toShapeProperties(opts.spPr);
  if (spPr) {
    legend.spPr = spPr;
  }
  if (opts.txPr) {
    legend.txPr = opts.txPr;
  }
  if (opts.entries && opts.entries.length > 0) {
    legend.legendEntries = opts.entries.map(e => {
      const entry: LegendEntry = { index: e.index };
      if (e.hidden) {
        entry.delete = true;
      }
      if (e.txPr) {
        entry.txPr = e.txPr;
      }
      return entry;
    });
  }
}

/**
 * Build a full ChartModel from the simplified AddChartOptions.
 */
export function buildChartModel(opts: AddChartOptions): ChartModel {
  validateChartOptions(opts);
  const seriesOpts = opts.series ?? [];
  const axIds = new AxIdAllocator();
  const { group: chartTypeGroup, axes } = buildChartTypeGroup(opts, seriesOpts, axIds);

  const plotArea: PlotArea = {
    chartTypes: [chartTypeGroup],
    axes,
    dataTable: buildDataTableFromOpts(opts.dataTable)
  };

  return finalizeChartModel(plotArea, opts);
}

/**
 * Build a combo chart model with multiple chart type groups.
 *
 * Each group can optionally be plotted on secondary axes.
 * The builder creates primary axes for the first group that needs
 * them, and secondary axes for any group with `useSecondaryAxis: true`.
 */
export function buildComboChartModel(opts: AddComboChartOptions): ChartModel {
  validateComboChartOptions(opts);
  const axIds = new AxIdAllocator();
  const chartTypeGroups: ChartTypeGroup[] = [];
  const allAxes: ChartAxis[] = [];

  // Create primary axes (shared by all groups without useSecondaryAxis)
  let primaryCatAx: CategoryAxis | undefined;
  let primaryValAx: ValueAxis | undefined;
  let primarySerAx: SeriesAxis | undefined;
  let primaryXAx: ValueAxis | undefined;
  let primaryYAx: ValueAxis | undefined;

  // Create secondary axes (shared by all groups with useSecondaryAxis)
  let secondaryCatAx: CategoryAxis | undefined;
  let secondaryValAx: ValueAxis | undefined;
  let secondarySerAx: SeriesAxis | undefined;
  let secondaryXAx: ValueAxis | undefined;
  let secondaryYAx: ValueAxis | undefined;

  for (const groupOpts of opts.groups) {
    const seriesOpts = groupOpts.series ?? [];
    const { group, axes } = buildChartTypeGroup(groupOpts, seriesOpts, axIds);

    // Detect serAx (3D chart types produce 3 axes: catAx + valAx + serAx)
    const hasSerAx = axes.length >= 3 && axes[2].axisType === "ser";

    if (axes.length > 0 && groupOpts.useSecondaryAxis) {
      // Secondary axis group: create secondary axes if not yet created
      const isCatVal = axes[0].axisType === "cat";
      if (isCatVal) {
        if (!secondaryCatAx) {
          const sCatId = axIds.alloc();
          const sValId = axIds.alloc();
          secondaryCatAx = {
            axisType: "cat",
            axId: sCatId,
            scaling: { orientation: "minMax" },
            delete: true, // secondary cat axis is typically hidden
            axPos: "b",
            crossAx: sValId,
            crosses: "autoZero",
            auto: true,
            lblAlgn: "ctr",
            lblOffset: 100
          };
          secondaryValAx = {
            axisType: "val",
            axId: sValId,
            scaling: { orientation: "minMax" },
            delete: false,
            axPos: "r", // right side
            crossAx: sCatId,
            crosses: "max", // cross at max of secondary cat axis
            crossBetween: "between"
          };
          allAxes.push(secondaryCatAx, secondaryValAx);
        }
        if (hasSerAx) {
          // 3D secondary: create or reuse secondary serAx
          if (!secondarySerAx) {
            const sSerAxId = axIds.alloc();
            secondarySerAx = {
              axisType: "ser",
              axId: sSerAxId,
              scaling: { orientation: "minMax" },
              delete: false,
              axPos: "b",
              crossAx: secondaryCatAx!.axId,
              crosses: "autoZero"
            };
            allAxes.push(secondarySerAx);
          }
          group.axisIds = [secondaryCatAx!.axId, secondaryValAx!.axId, secondarySerAx.axId];
        } else {
          group.axisIds = [secondaryCatAx!.axId, secondaryValAx!.axId];
        }
      } else {
        // Scatter/bubble secondary axes (valAx + valAx)
        if (!secondaryXAx) {
          const sXId = axIds.alloc();
          const sYId = axIds.alloc();
          secondaryXAx = {
            axisType: "val",
            axId: sXId,
            scaling: { orientation: "minMax" },
            delete: true,
            axPos: "b",
            crossAx: sYId,
            crosses: "autoZero",
            crossBetween: "midCat"
          };
          secondaryYAx = {
            axisType: "val",
            axId: sYId,
            scaling: { orientation: "minMax" },
            delete: false,
            axPos: "r",
            crossAx: sXId,
            crosses: "max",
            crossBetween: "midCat"
          };
          allAxes.push(secondaryXAx, secondaryYAx);
        }
        group.axisIds = [secondaryXAx!.axId, secondaryYAx!.axId];
      }
      // Don't add the auto-generated axes from buildChartTypeGroup
    } else if (axes.length > 0) {
      // Primary axis group
      const isCatVal = axes[0].axisType === "cat";
      if (isCatVal) {
        if (!primaryCatAx) {
          primaryCatAx = axes[0] as CategoryAxis;
          primaryValAx = axes[1] as ValueAxis;
          allAxes.push(primaryCatAx, primaryValAx);
          if (hasSerAx) {
            primarySerAx = axes[2] as SeriesAxis;
            allAxes.push(primarySerAx);
          }
        } else {
          // Reuse primary axes
          if (hasSerAx) {
            // 3D group needs serAx — create one if the primary set didn't have it
            if (!primarySerAx) {
              const serAxId = axIds.alloc();
              primarySerAx = {
                axisType: "ser",
                axId: serAxId,
                scaling: { orientation: "minMax" },
                delete: false,
                axPos: "b",
                crossAx: primaryCatAx.axId,
                crosses: "autoZero"
              };
              allAxes.push(primarySerAx);
            }
            group.axisIds = [primaryCatAx.axId, primaryValAx!.axId, primarySerAx.axId];
          } else {
            group.axisIds = [primaryCatAx.axId, primaryValAx!.axId];
          }
        }
      } else {
        // Scatter/bubble
        if (!primaryXAx) {
          primaryXAx = axes[0] as ValueAxis;
          primaryYAx = axes[1] as ValueAxis;
          allAxes.push(primaryXAx, primaryYAx);
        } else {
          group.axisIds = [primaryXAx.axId, primaryYAx!.axId];
        }
      }
    }

    chartTypeGroups.push(group);
  }

  // Renumber series index/order globally across all groups.
  // OOXML requires c:idx and c:order to be unique within the entire chart.
  // Only renumber if indices are the default per-group values (0, 1, 2...) —
  // i.e. they would collide across groups. If the user has already assigned
  // explicit globally-unique indices, leave them untouched (P4.2 fix).
  if (indicesWouldCollide(chartTypeGroups)) {
    let globalIdx = 0;
    for (const grp of chartTypeGroups) {
      for (const s of grp.series) {
        s.index = globalIdx;
        s.order = globalIdx;
        globalIdx++;
      }
    }
  }

  const plotArea: PlotArea = {
    chartTypes: chartTypeGroups,
    axes: allAxes,
    dataTable: buildDataTableFromOpts(opts.dataTable)
  };

  return finalizeChartModel(plotArea, opts);
}

/**
 * Check whether any two groups have colliding series indices.
 * When true, the combo builder renumbers all series to guarantee uniqueness.
 */
function indicesWouldCollide(groups: ChartTypeGroup[]): boolean {
  const seen = new Set<number>();
  for (const g of groups) {
    for (const s of g.series as SeriesBase[]) {
      if (seen.has(s.index)) {
        return true;
      }
      seen.add(s.index);
    }
  }
  return false;
}
