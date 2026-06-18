/**
 * Chart Builder - Constructs a ChartModel from simplified AddChartOptions.
 *
 * This bridges the high-level API (worksheet.addChart) to the full
 * OOXML chart model that the XForm layer serialises.
 */

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
  ChartTypeGroup,
  DataLabels,
  DataLabelEntry,
  DataLabelPosition,
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
  PictureOptions,
  SeriesAxis
} from "@excel/chart/model/types";
import { escapeXml } from "@excel/chart/shared/chart-utils";
import { ChartOptionsError } from "@excel/errors";

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
 * Valid `c:dLblPos` values Excel accepts per chart type.
 *
 * Although ECMA-376 `ST_DLblPos` technically allows any of the nine
 * values (`b | bestFit | ctr | inBase | inEnd | l | outEnd | r | t`)
 * on any `c:dLbl` / `c:dLbls` element, Excel's reader is stricter:
 * emitting a value outside the per-chart-type allow-list below
 * triggers "Repaired Records: Drawing" warnings on open, and for
 * doughnut charts the offending `drawing*.xml` part is stripped
 * entirely ("Removed Part"). The allow-lists match the options
 * surfaced by Excel 365's "Format Data Labels → Label Position"
 * panel, which is the canonical UI reference.
 *
 * - `doughnut`: Excel's UI exposes no position choices at all, and
 *   any `c:dLblPos` in a doughnut chart's `c:dLbls` causes the
 *   entire drawing part to be removed on open. Use an empty list so
 *   the validator rejects every value.
 * - `bar` / `bar3D`: `inBase` is unique to bar — it anchors the
 *   label to the axis end of a column, useful for negative values.
 * - `pie`, `pie3D`, `ofPie`: share the pie label set. `bestFit` is
 *   Excel's default and the only value that lets Excel place labels
 *   automatically with leader lines.
 * - `line` / `line3D` / `scatter` / `bubble` / `radar` / `stock`:
 *   share the cartesian label set (above / below / left / right /
 *   center).
 * - `area` / `area3D`: Excel only accepts `ctr` for area fills.
 * - `surface` / `surface3D`: data labels are already rejected
 *   wholesale by `validateChartLevelOptions`.
 */
const VALID_DLBL_POSITIONS_BY_TYPE: Partial<
  Record<AddChartOptions["type"], ReadonlySet<DataLabelPosition>>
> = {
  bar: new Set<DataLabelPosition>(["ctr", "inBase", "inEnd", "outEnd"]),
  bar3D: new Set<DataLabelPosition>(["ctr", "inBase", "inEnd", "outEnd"]),
  line: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  line3D: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  scatter: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  bubble: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  radar: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  stock: new Set<DataLabelPosition>(["ctr", "l", "r", "t", "b"]),
  pie: new Set<DataLabelPosition>(["bestFit", "ctr", "inEnd", "outEnd"]),
  pie3D: new Set<DataLabelPosition>(["bestFit", "ctr", "inEnd", "outEnd"]),
  ofPie: new Set<DataLabelPosition>(["bestFit", "ctr", "inEnd", "outEnd"]),
  doughnut: new Set<DataLabelPosition>(),
  area: new Set<DataLabelPosition>(["ctr"]),
  area3D: new Set<DataLabelPosition>(["ctr"])
};

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
    // `BarGrouping` and `LineGrouping` overlap on stacked/percentStacked but
    // diverge on `clustered` (bar only) and `standard` (line/area only).
    // Reject the wrong-family value up front instead of silently emitting
    // invalid OOXML.
    const g = opts.grouping;
    if (opts.type === "bar" || opts.type === "bar3D") {
      assertChartOptions(
        g === "clustered" || g === "stacked" || g === "percentStacked",
        `${path}.grouping=${JSON.stringify(g)} is not valid for bar charts (use "clustered" | "stacked" | "percentStacked").`
      );
    } else {
      assertChartOptions(
        g === "standard" || g === "stacked" || g === "percentStacked",
        `${path}.grouping=${JSON.stringify(g)} is not valid for ${opts.type} charts (use "standard" | "stacked" | "percentStacked").`
      );
    }
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
  if (opts.dataLabels) {
    // Chart-level data labels mirror the series-level ones at the
    // `CT_*Chart/c:dLbls` slot (e.g. `CT_BarChart/c:dLbls`). Excel
    // applies the same per-chart-type restrictions on `c:dLblPos`
    // at this level, so run the same validator here.
    validateDataLabelsOptions(opts.type, opts.dataLabels, `${path}.dataLabels`);
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
  if (opts.gapDepth !== undefined) {
    // `c:gapDepth` is only declared on the 3-D chart types'
    // `CT_*3DChart` definitions (bar3D, line3D, area3D). `pie3D`
    // does NOT have it despite the type being 3-D.
    assertChartOptions(
      opts.type === "bar3D" || opts.type === "line3D" || opts.type === "area3D",
      `${path}.gapDepth is only valid for bar3D, line3D, and area3D charts.`
    );
    assertIntegerInRange(opts.gapDepth, `${path}.gapDepth`, 0, 500);
  }
  if (opts.overlap !== undefined) {
    // `c:overlap` belongs to `CT_BarChart` only — `CT_Bar3DChart`
    // omits it. The writer rejects it on bar3D (see
    // `buildChartTypeGroup:case "bar3D"`); keep the two in sync so
    // the rejection surfaces here at authoring time with a
    // consistent error.
    assertChartOptions(
      opts.type === "bar",
      `${path}.overlap is only valid for 2-D bar charts (bar3D rejects overlap per CT_Bar3DChart).`
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
      opts.type === "line" || opts.type === "radar",
      `${path}.showMarker is only valid for line and radar charts (line3D does not support markers).`
    );
  }
  if (opts.smooth !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "scatter",
      `${path}.smooth is only valid for line and scatter charts (line3D does not support smooth).`
    );
  }
  if (opts.hiLowLines !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "stock",
      `${path}.hiLowLines is only valid for line and stock charts (line3D does not support hiLowLines).`
    );
  }
  if (opts.upDownBars !== undefined) {
    assertChartOptions(
      opts.type === "line" || opts.type === "stock",
      `${path}.upDownBars is only valid for line and stock charts (line3D does not support upDownBars).`
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
  path: string,
  flags: { allowMissingRefs?: boolean } = {}
): void {
  const allowMissingRefs = flags.allowMissingRefs === true;
  assertChartOptions(!!opts && typeof opts === "object", `${path} must be an object.`);
  if (allowMissingRefs) {
    if (opts.values !== undefined) {
      assertChartOptions(
        typeof opts.values === "string" && opts.values.length > 0,
        `${path}.values must be a non-empty formula string.`
      );
    }
  } else {
    assertChartOptions(
      typeof opts.values === "string" && opts.values.length > 0,
      `${path}.values is required and must be a non-empty formula string.`
    );
  }
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
  if (chartType === "bubble" && !allowMissingRefs) {
    assertChartOptions(
      opts.xValues !== undefined,
      `${path}.xValues is required for bubble charts. Use a numeric range for x-values.`
    );
    assertChartOptions(
      opts.bubbleSize !== undefined,
      `${path}.bubbleSize is required for bubble charts. Use a numeric range for bubble sizes.`
    );
  }
  if (chartType === "scatter" && !allowMissingRefs) {
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
  // Note on `categories` for scatter / bubble: the field has no direct
  // home on `ScatterSeries` / `BubbleSeries`, which use `xVal` for the
  // numeric X axis. `buildScatterSeries` / `buildBubbleSeries` silently
  // ignore `opts.categories` at create time (the `xVal` slot is already
  // populated from `opts.xValues`). On the patch path
  // (`applyChartSeriesOptionsPatch`), `options.categories` is routed to
  // `xVal` so callers can switch a scatter's X source to a text axis
  // via `{ categories, xValueType: "text" }`. Accepting the field on
  // creation preserves API symmetry (users can pass the same option
  // bundle across chart types in test fixtures) at the cost of a silent
  // drop for this specific combination.
  if (opts.dataPoints) {
    for (let i = 0; i < opts.dataPoints.length; i++) {
      validateDataPointOptions(chartType, opts.dataPoints[i], `${path}.dataPoints[${i}]`);
    }
  }
  if (opts.dataLabels) {
    validateDataLabelsOptions(chartType, opts.dataLabels, `${path}.dataLabels`);
  }
  if (opts.errorBars) {
    assertChartOptions(
      !PIE_FAMILY_CHART_TYPES.has(chartType),
      `${path}.errorBars is not valid for ${chartType} charts.`
    );
    const errorBars = Array.isArray(opts.errorBars) ? opts.errorBars : [opts.errorBars];
    // Non-scatter/bubble series only support a single error-bar configuration
    // (`ErrorBars`, not an array). Fail fast rather than silently keep only
    // `errorBars[0]` and discard the rest.
    if (
      Array.isArray(opts.errorBars) &&
      opts.errorBars.length > 1 &&
      chartType !== "scatter" &&
      chartType !== "bubble"
    ) {
      assertChartOptions(
        false,
        `${path}.errorBars must be a single configuration for ${chartType} charts; arrays are only valid for scatter and bubble.`
      );
    }
    // Scatter / bubble: CT_ScatterSer / CT_BubbleSer declare
    // `errBars` as `maxOccurs="2"` — one X, one Y. Allowing more
    // than two (or two with the same `c:errDir`) is an OOXML schema
    // violation and causes Excel to repair the chart on open,
    // silently dropping the extras. Fail the author up-front so
    // mistakes like passing an array of "every error-bar type" on a
    // single series are caught rather than producing a file that
    // prompts a repair dialog.
    if (
      (chartType === "scatter" || chartType === "bubble") &&
      Array.isArray(opts.errorBars) &&
      opts.errorBars.length > 2
    ) {
      assertChartOptions(
        false,
        `${path}.errorBars allows at most 2 entries on ${chartType} charts (one for direction "x", one for "y"). Split additional configurations across separate series.`
      );
    }
    if (
      (chartType === "scatter" || chartType === "bubble") &&
      Array.isArray(opts.errorBars) &&
      opts.errorBars.length === 2
    ) {
      const d0 = opts.errorBars[0]?.direction;
      const d1 = opts.errorBars[1]?.direction;
      assertChartOptions(
        d0 !== undefined && d1 !== undefined && d0 !== d1,
        `${path}.errorBars must use distinct directions ("x" and "y") when providing two entries on ${chartType} charts.`
      );
    }
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
  // Patch path: unlike series-creation validation we accept a partial
  // options bag with no `values` / `xValues` / `bubbleSize`. The
  // `allowMissingRefs` flag short-circuits the "required" assertions so we
  // don't need to inject placeholder strings (previous implementation used
  // an `__excelts_placeholder__` sentinel that would have triggered false
  // positives if `values` ever gained a content-level check).
  validateSeriesOptions(chartType, opts as AddChartSeriesOptions, path, { allowMissingRefs: true });
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
  } else {
    // `period` is only meaningful for moving-average trendlines (Excel
    // silently ignores it on other types, which quickly devolves into
    // phantom config bugs). Match the stricter `order` handling above.
    assertChartOptions(
      opts.period === undefined,
      `${path}.period is only valid for movingAvg trendlines.`
    );
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

/**
 * Ensure every `dLblPos` value (group-level and per-entry overrides)
 * is one that Excel will accept for the chart type — see the
 * rationale on {@link VALID_DLBL_POSITIONS_BY_TYPE}. Writing an
 * invalid position causes Excel to flag the drawing as corrupted
 * ("Repaired Records" / "Removed Part") even though the OOXML
 * schema would technically accept it.
 *
 * Rejects at author time so the caller gets a pointer to the exact
 * offending field rather than debugging a Removed Part dialog.
 */
function validateDataLabelsOptions(
  chartType: AddChartOptions["type"],
  opts: AddDataLabelsOptions,
  path: string
): void {
  const allowed = VALID_DLBL_POSITIONS_BY_TYPE[chartType];
  const describeAllowed = (): string => {
    if (!allowed) {
      return "(unknown chart type)";
    }
    if (allowed.size === 0) {
      return `(${chartType} does not support c:dLblPos — Excel rejects any value)`;
    }
    return [...allowed].sort().join(", ");
  };
  if (opts.position !== undefined) {
    assertChartOptions(
      !!allowed && allowed.has(opts.position),
      `${path}.position="${opts.position}" is not valid for ${chartType} charts. Allowed: ${describeAllowed()}.`
    );
  }
  if (opts.entries) {
    for (let i = 0; i < opts.entries.length; i++) {
      const entry = opts.entries[i];
      if (entry?.position !== undefined) {
        assertChartOptions(
          !!allowed && allowed.has(entry.position),
          `${path}.entries[${i}].position="${entry.position}" is not valid for ${chartType} charts. Allowed: ${describeAllowed()}.`
        );
      }
    }
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
  if (opts.majorUnit !== undefined && opts.minorUnit !== undefined) {
    assertChartOptions(
      opts.minorUnit <= opts.majorUnit,
      `${path}.minorUnit must be less than or equal to ${path}.majorUnit.`
    );
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

/**
 * Normalise a user-facing hex colour into a structured {@link ChartColor}.
 * Accepts `"#RRGGBB"` / `"RRGGBB"` and the optional 8-digit
 * `"RRGGBBAA"` form. The alpha byte, when present, is decoded into
 * `color.alpha` on the OOXML 0–100000 scale (0 = fully transparent,
 * 100000 = fully opaque) rather than discarded. Throws
 * `ChartOptionsError` when the input is not a valid hex triplet so the
 * caller sees the mistake at the assignment site rather than via a
 * downstream XML parser rejection.
 */
export function hexToColor(hex: string): ChartColor {
  const cleaned = hex.replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}([0-9A-F]{2})?$/.test(cleaned)) {
    throw new ChartOptionsError(
      `Invalid hex colour: ${JSON.stringify(hex)}. Expected 6-digit (or 8-digit with alpha) hex like "#FF0000".`
    );
  }
  const color: ChartColor = { srgb: cleaned.slice(0, 6) };
  if (cleaned.length === 8) {
    // 8-digit form: trailing 2 bytes encode alpha on the 0–255 scale.
    // `ChartColor.alpha` stores OOXML's 0–100000 integer; convert and
    // round so the wire value is an integer, matching Excel's output.
    const alphaByte = parseInt(cleaned.slice(6, 8), 16);
    color.alpha = Math.round((alphaByte / 255) * 100000);
  }
  return color;
}

/**
 * Escape XML text content while stripping characters forbidden by
 * XML 1.0. Without the strip step, a chart built from user input
 * containing `\b` / `\f` / other C0 control characters would produce
 * an XML-invalid `<c:v>` or `<c:f>` payload that strict readers
 * refuse to parse. Preserved: `\t`, `\n`, `\r` (the only C0 chars
 * XML 1.0 allows in content).
 */
function buildPivotSourceXml(source: PivotChartSource): string {
  const name = typeof source === "string" ? source : source.name;
  const fmtId = typeof source === "string" ? 0 : (source.fmtId ?? 0);
  if (!name) {
    throw new ChartOptionsError("Pivot chart source name is required.");
  }
  if (!Number.isInteger(fmtId) || fmtId < 0) {
    throw new ChartOptionsError("Pivot chart source fmtId must be a non-negative integer.");
  }
  // Pivot chart options used to be embedded inside `<c:pivotSource>` under a
  // private `xmlns:excelts` namespace; Excel never recognised that and the
  // parser never read it back. Options are now routed into
  // `ChartModel.pivotOptions` (see {@link ChartModel.pivotOptions}) and
  // serialised as MS standard `c14:pivotOptions` inside chartSpace's extLst.
  return `<c:pivotSource><c:name>${escapeXml(name)}</c:name><c:fmtId val="${fmtId}"/></c:pivotSource>`;
}

/**
 * Convert a simplified AddShapeFillOptions (or a pre-built ShapeProperties)
 * into a structured ShapeProperties object. Returns undefined when input is
 * undefined.
 */
/**
 * Normalise a user-facing shape-fill option bundle (hex-string fill,
 * hex-string border, borderWidth in points, gradient, pattern, …) into a
 * structured {@link ShapeProperties}. Passes through already-structured
 * shapes unchanged so callers can mix the two forms freely.
 *
 * Exported so `chart-ex-builder` can reuse the exact same normalisation
 * for ChartEx `spPr` options — previously ChartEx only accepted the
 * fully-structured form, which was an API asymmetry.
 */
export function toShapeProperties(
  input: ShapeProperties | AddShapeFillOptions | undefined
): ShapeProperties | undefined {
  if (!input) {
    return undefined;
  }
  // If it already looks like a ShapeProperties (has fill/line/effectList/_rawXml
  // but no hex-string "fill" field), shallow-clone rather than return
  // the caller's reference. Every downstream consumer treats the
  // returned object as owned (e.g. `applyChartSeriesOptionsPatch`
  // deletes `_rawXml` on it; combo-group builders overwrite inner
  // fields), and leaking the caller's object into the model would
  // mean subsequent patches on one chart silently mutate the
  // caller's options blob — and if that blob was reused across
  // multiple `addChart(...)` calls, every later chart would see the
  // stripped / mutated state. The clone is shallow because `fill` /
  // `line` sub-trees are themselves replaced wholesale by the
  // downstream patchers when they change; deeper aliasing is not
  // currently a hazard.
  if (isShapeProperties(input)) {
    return { ...input };
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
      // OOXML `a:ln/@w` is `ST_LineWidth` = `xsd:int`. A fractional
      // `borderWidth` (e.g. 0.825pt, or any point value that doesn't
      // round-trip through `n / 12700` cleanly) would interpolate as
      // `"10477.5"`, which strict readers reject. Rounding matches the
      // sibling builders (`buildSeriesSpPr`, etc.).
      spPr.line.width = Math.round(opts.borderWidth * EMU_PER_POINT);
    } else if (opts.border) {
      // Mirror chart-ex-builder: when the caller sets a border colour
      // without an explicit width, fall back to 9525 EMU (0.75pt).
      // Without this DrawingML readers treat `<a:ln>` as hairline,
      // which typically disappears on screen.
      spPr.line.width = 9525;
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
  // Check the remaining `ShapeProperties` fields that distinguish it
  // from `AddShapeFillOptions`. `AddShapeFillOptions` has only
  // `fill` (as hex-string) and `line` (as hex-string); anything else
  // is a `ShapeProperties`-only field. Previously the sniffer missed
  // the 3D/geometry fields entirely — a user passing
  // `{ sp3d: {...} }` or `{ transform: {...} }` alone was
  // misclassified and the fields silently dropped.
  if ("scene3d" in o || "sp3d" in o) {
    return true;
  }
  if ("transform" in o || "presetGeometry" in o || "customGeometry" in o) {
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
  if (name === undefined) {
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
  if (opts.line || opts.lineWidth !== undefined || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth !== undefined) {
      line.width = Math.round(opts.lineWidth * EMU_PER_POINT); // pt → EMU
    } else if (opts.line) {
      // Mirror chart-ex-builder: when the caller sets a line colour
      // without an explicit width, fall back to 9525 EMU (0.75pt —
      // Excel's default for chart series borders). Without this
      // DrawingML readers treat `<a:ln>` as hairline, which disappears
      // at typical screen DPI and is never what the user means by
      // `line: "#FF0000"`.
      line.width = 9525;
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
      // Default to 9525 EMU (0.75pt) — without an explicit width
      // DrawingML treats the outline as hairline and the marker ring
      // typically vanishes at on-screen DPI. Matches the other
      // builder paths that default to the same width.
      mSpPr.line = { color: hexToColor(opts.border), width: 9525 };
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
  if (opts.separator !== undefined) {
    dl.separator = opts.separator;
  }
  if (opts.numFmt !== undefined) {
    dl.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  } else if (opts.numFmtLinked !== undefined) {
    // Allow `numFmtLinked` without an explicit `numFmt` to re-link the
    // data-label number format to the source cell. OOXML's `CT_NumFmt`
    // requires `formatCode`, so default it to `General` (the Excel
    // default when `sourceLinked="1"` is emitted without an author-
    // specified format). Previously this branch silently dropped the
    // flag: callers had no way to opt into source-linking on data
    // labels without also supplying a redundant format code.
    dl.numFmt = { formatCode: "General", sourceLinked: opts.numFmtLinked };
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
  } else if (opts.numFmtLinked !== undefined) {
    // Allow standalone `numFmtLinked` — see `buildDataLabelsFromOpts`
    // for the matching semantics.
    entry.numFmt = { formatCode: "General", sourceLinked: opts.numFmtLinked };
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
  if (opts.name !== undefined) {
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
  if (opts.line || opts.lineWidth !== undefined || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth !== undefined) {
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
  if (opts.text !== undefined) {
    lbl.text = opts.text;
  }
  if (opts.numFmt !== undefined) {
    lbl.numFmt = { formatCode: opts.numFmt, sourceLinked: opts.numFmtLinked };
  } else if (opts.numFmtLinked !== undefined) {
    // Allow standalone `numFmtLinked` — see `buildDataLabelsFromOpts`
    // for the matching semantics.
    lbl.numFmt = { formatCode: "General", sourceLinked: opts.numFmtLinked };
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
  } else if (opts.line || opts.lineWidth !== undefined || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.line) {
      line.color = hexToColor(opts.line);
    }
    if (opts.lineWidth !== undefined) {
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
      spPr.line = { color: hexToColor(opts.border), width: 9525 };
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
/**
 * Structural super-type accepted by {@link applySeriesOptions}. All
 * series types ({@link BarSeries}, {@link LineSeries}, …) satisfy this
 * shape because the fields we set here are either common to every
 * series (spPr / txPr / dataLabels / trendlines / dataPoints) or are
 * only written when the caller opts in via the corresponding option
 * key. Using this type instead of `any` keeps the common fields
 * type-checked while still letting the concrete caller pass any of
 * the series variants.
 */
type MutableSeriesWithOptions = SeriesBase & {
  marker?: ChartMarker;
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  dataPoints?: DataPoint[];
  pictureOptions?: PictureOptions;
};

function applySeriesOptions(
  s: MutableSeriesWithOptions,
  opts: AddChartSeriesOptions,
  context: { supportsPictureOptions?: boolean } = {}
): void {
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
  // NOTE: classic chart series have no `txPr` slot in OOXML, so the
  // field was removed from `AddChartSeriesOptions`. Trendline labels,
  // axis labels, data labels, and titles all have their own `txPr`
  // entry points — route run-level text styling through those.
  if (opts.pictureFill) {
    applyPictureFillToSeries(s, opts.pictureFill, {
      supportsPictureOptions: context.supportsPictureOptions
    });
  }
}

/**
 * Split an `AddShapeFillOptions`-style `pictureFill` option bundle into the
 * two OOXML artefacts that render the feature:
 *   - `c:pictureOptions` (stretch/stack/apply-to-*) — stored on
 *     `series.pictureOptions`
 *   - `a:blipFill` (the actual image rel) — stored on
 *     `series.spPr.fill.blip`
 *
 * Both are updated in a single place so the new-series path
 * ({@link applySeriesOptions}) and the patch path
 * ({@link applyChartSeriesOptionsPatch}) cannot diverge.
 */
function applyPictureFillToSeries(
  series: SeriesBase,
  pictureFill: NonNullable<AddChartSeriesOptions["pictureFill"]>,
  options: { supportsPictureOptions?: boolean } = {}
): void {
  // `c:pictureOptions` (stretch/stack/applyTo*/scale) is only declared
  // on `CT_BarSer` per ECMA-376 §21.2.2.162. Non-bar callers get the
  // `<a:blipFill>` on their `spPr.fill.blip` (which is legal on every
  // shape), but the caller-supplied `applyToFront` / `fillMode` /
  // `scale` hints are silently dropped — we would otherwise emit a
  // `<c:pictureOptions>` that schema validators reject. Surface the
  // mismatch if the caller set bar-only fields on a non-bar series
  // rather than letting the data quietly disappear.
  const barOnlyFieldsUsed =
    pictureFill.applyToFront !== undefined ||
    pictureFill.applyToSides !== undefined ||
    pictureFill.applyToEnd !== undefined ||
    pictureFill.scale !== undefined;
  if (options.supportsPictureOptions) {
    (series as SeriesBase & { pictureOptions?: PictureOptions }).pictureOptions = {
      applyToFront: pictureFill.applyToFront,
      applyToSides: pictureFill.applyToSides,
      applyToEnd: pictureFill.applyToEnd,
      pictureFormat: pictureFill.fillMode,
      pictureStackUnit: pictureFill.scale
    };
  } else if (barOnlyFieldsUsed) {
    throw new ChartOptionsError(
      "pictureFill.applyToFront / applyToSides / applyToEnd / scale are only valid for bar / bar3D series (mapped to c:pictureOptions in CT_BarSer). Use the plain `fill` / `spPr.fill.blip` form for other series types, or omit the bar-only hints."
    );
  }
  if (pictureFill.image !== undefined || pictureFill.relationshipId) {
    // Parsed-from-XML series carry their `spPr` as a dual
    // representation: the structured `fill` / `line` slots are
    // populated AND `_rawXml` holds the original DrawingML bytes.
    // The writer short-circuits to `_rawXml` (see
    // `chart-space-xform._renderSpPr`) whenever it's present, which
    // means a downstream mutation like this one — which sets
    // `series.spPr.fill.blip` but can't touch the raw bytes — gets
    // silently overwritten at save time. The chart rel points at the
    // new image, the chart XML still carries the old fill. Strip
    // `_rawXml` here so the structured patch wins and the writer
    // emits the fresh `<a:blipFill>`.
    //
    // We also need to repopulate enough of the structured model that
    // the dropped `_rawXml` doesn't take authored line / geometry /
    // effect properties down with it. `parseSpPr` is already called
    // at load time and stores the structured fields alongside the
    // raw bytes, so the structured slots still carry whatever the
    // loaded file had. Nothing extra needed on the clone side.
    if (series.spPr && typeof series.spPr === "object") {
      delete (series.spPr as { _rawXml?: string })._rawXml;
    }
    series.spPr = series.spPr ?? {};
    series.spPr.fill = series.spPr.fill ?? {};
    // Also clear a raw representation that might live on the
    // existing `fill` object (rare — reserved for `<a:solidFill>`
    // / `<a:gradFill>` raw captures some parsers produce). The
    // writer emits a single `<a:*Fill>` child per `<a:spPr>`, so
    // the new blip must be the only fill in play.
    delete (series.spPr.fill as { solid?: unknown }).solid;
    delete (series.spPr.fill as { gradient?: unknown }).gradient;
    delete (series.spPr.fill as { pattern?: unknown }).pattern;
    delete (series.spPr.fill as { noFill?: unknown }).noFill;
    series.spPr.fill.blip = {
      fillMode:
        pictureFill.fillMode === "stack" || pictureFill.fillMode === "stackScale"
          ? "tile"
          : "stretch",
      ...(pictureFill.relationshipId ? { relationshipId: pictureFill.relationshipId } : {}),
      ...(pictureFill.image !== undefined ? { _pendingImage: pictureFill.image } : {})
    };
  }
}

/**
 * Populates common category-value series fields (tx, cat, val) shared by
 * bar, line, pie, area, radar, and surface series builders.
 */
function populateCatValBase(
  s: { tx?: unknown; cat?: unknown; val?: unknown },
  opts: AddChartSeriesOptions
): void {
  s.tx = makeSeriesTx(opts.name);
  if (opts.categories) {
    s.cat = makeAxisData(opts.categories);
  }
  if (opts.values) {
    s.val = makeNumData(opts.values);
  }
}

/**
 * Build a single error-bars entry from options, normalising the
 * array-vs-single input shape used by non-scatter/bubble series.
 */
function buildSingleErrorBars(
  opts: AddChartSeriesOptions["errorBars"]
): ReturnType<typeof buildErrorBarsFromOpts> | undefined {
  if (!opts) {
    return undefined;
  }
  if (Array.isArray(opts)) {
    return opts.length > 0 ? buildErrorBarsFromOpts(opts[0]) : undefined;
  }
  return buildErrorBarsFromOpts(opts);
}

function buildBarSeries(opts: AddChartSeriesOptions, idx: number): BarSeries {
  const s: BarSeries = { index: idx, order: idx };
  populateCatValBase(s, opts);
  applySeriesOptions(s, opts, { supportsPictureOptions: true });
  if (opts.invertIfNegative !== undefined) {
    s.invertIfNegative = opts.invertIfNegative;
  }
  s.errorBars = buildSingleErrorBars(opts.errorBars);
  return s;
}

function buildLineSeries(opts: AddChartSeriesOptions, idx: number): LineSeries {
  const s: LineSeries = { index: idx, order: idx };
  populateCatValBase(s, opts);
  applySeriesOptions(s, opts);
  if (opts.smooth !== undefined) {
    s.smooth = opts.smooth;
  }
  s.errorBars = buildSingleErrorBars(opts.errorBars);
  return s;
}

function buildPieSeries(opts: AddChartSeriesOptions, idx: number): PieSeries {
  const s: PieSeries = { index: idx, order: idx };
  populateCatValBase(s, opts);
  applySeriesOptions(s, opts);
  if (opts.explosion !== undefined) {
    s.explosion = opts.explosion;
  }
  return s;
}

function buildAreaSeries(opts: AddChartSeriesOptions, idx: number): AreaSeries {
  const s: AreaSeries = { index: idx, order: idx };
  populateCatValBase(s, opts);
  applySeriesOptions(s, opts);
  s.errorBars = buildSingleErrorBars(opts.errorBars);
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
  populateCatValBase(s, opts);
  applySeriesOptions(s, opts);
  return s;
}

function buildSurfaceSeries(opts: AddChartSeriesOptions, idx: number): SurfaceSeries {
  const s: SurfaceSeries = { index: idx, order: idx };
  populateCatValBase(s, opts);
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
  // Exhaustiveness check: every value of `AddChartOptions["type"]` must be
  // handled above. Falling back to `buildBarSeries` silently would mis-map
  // any new chart type introduced in the future.
  const _exhaustive: never = chartType;
  throw new ChartOptionsError(`Unsupported chart type: ${String(_exhaustive)}.`);
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
    // Scatter / bubble series have no category axis — their X axis is
    // `xVal`. When a patch targets such a series (identified by an
    // already-populated `xVal` and no `cat`), route `categories`
    // through `makeXAxisData` so `options.xValueType: "text"` flips
    // the series from a numeric `xVal` to a labelled one. For every
    // other series type the patch goes to the structural `cat` slot.
    if (target.xVal && !target.cat) {
      target.xVal = makeXAxisData(options.categories, options.xValueType);
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
    const patchShape = buildSeriesSpPr(options as AddChartSeriesOptions);
    if (options.spPr !== undefined) {
      // Explicit `spPr` replaces the whole shape — the caller opted in
      // to a full structured override.
      series.spPr = patchShape;
    } else {
      // Sugared patches (`fill` / `line` / `lineWidth` / `lineDash`)
      // should update only the affected sub-object AND merge inside the
      // sub-object so narrow patches (e.g. `lineDash` alone) don't wipe
      // out the adjacent fields.
      //
      // Previously the merge stopped at the top level (`{ ...existing,
      // line: patchShape.line }`) — so `updateSeries(0, { lineDash: "dash" })`
      // on a series with `{ line: "#FF0000", lineWidth: 2 }` dropped the
      // colour AND width because `buildSeriesSpPr({ lineDash: "dash" })`
      // only sets `line.dash`, and the top-level spread replaced the
      // whole `line` sub-object with `{ dash: "dash" }`. Deep-merge the
      // sub-objects so each field survives until the caller explicitly
      // overwrites it.
      const existing = series.spPr ?? {};
      const next: ShapeProperties = { ...existing };
      if (patchShape?.fill !== undefined) {
        next.fill = patchShape.fill;
      }
      if (patchShape?.line !== undefined) {
        next.line = { ...(existing.line ?? {}), ...patchShape.line };
      }
      series.spPr = next;
    }
    // CRITICAL: clear `_rawXml` so the writer serialises from the
    // structured fields we just patched. The chart-space writer emits
    // `_rawXml` verbatim when present (`_renderSpPr` at the top of the
    // function body in `chart-space-xform.ts`), which meant every
    // structured mutation was silently overridden by the stale raw
    // bytes captured at parse time. After this clear the next save
    // re-serialises the full shape tree from the structured fields.
    if (series.spPr && "_rawXml" in series.spPr) {
      delete (series.spPr as { _rawXml?: string })._rawXml;
    }
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
    // Validate each error-bar config independently. Previously the
    // patch path relied on the outer `validateSeriesPatchOptions` call
    // at the top of this function, but that skip when `chartType` is
    // omitted. `CT_ErrBars` requires `errValType`, and without a
    // validator an options object missing `type` would silently emit
    // `<c:errValType/>` (no attribute) → schema-invalid output.
    for (const eb of errorBars) {
      validateErrorBarsOptions(eb, "series.errorBars");
    }
    const built = errorBars.map(buildErrorBarsFromOpts);
    (series as SeriesBase & { errorBars?: ErrorBars | ErrorBars[] }).errorBars =
      chartType === "scatter" || chartType === "bubble" ? built : built[0];
  }
  // NOTE: classic chart series have no `txPr` slot in OOXML — see the
  // `AddChartSeriesOptions` type comment. The field was removed from
  // the options bag so this patch branch is gone as well.
  if (options.pictureFill !== undefined) {
    applyPictureFillToSeries(series, options.pictureFill, {
      supportsPictureOptions: chartType === "bar" || chartType === "bar3D"
    });
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
  if (opts.title !== undefined) {
    // Replace the title text but preserve any previously-applied
    // `titleOptions` (layout, overlay, spPr, txPr) attached to
    // `axis.title`. Combo charts call `applyAxisOptions` repeatedly on
    // a shared axis (see `buildChartTypeGroup` reuse paths), and a
    // wholesale replacement here discarded every field except the new
    // text. We graft the new title runs onto the existing frame.
    const fresh = buildTitle(opts.title);
    axis.title = axis.title
      ? {
          ...axis.title,
          strRef: fresh.strRef,
          text: fresh.text,
          rawTx: fresh.rawTx
        }
      : fresh;
  }
  if (opts.titleOptions) {
    if (!axis.title) {
      axis.title = { overlay: false };
    }
    applyTitleOptions(axis.title, opts.titleOptions);
  }
  if (opts.numFmt !== undefined) {
    // Merge onto any prior `numFmt` so a later call that supplies
    // `opts.numFmt` without `opts.numFmtLinked` doesn't reset
    // `sourceLinked` back to `undefined`. Same class of bug as the
    // title replacement above — triggered when combo-chart flows
    // apply options to a shared axis in multiple passes.
    axis.numFmt = {
      ...axis.numFmt,
      formatCode: opts.numFmt,
      ...(opts.numFmtLinked !== undefined ? { sourceLinked: opts.numFmtLinked } : {})
    };
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
  // Major gridlines: the explicit `majorGridlines` boolean is an on/off
  // switch that must win over any `majorGridlinesStyle`. Callers who
  // pass both `{ majorGridlines: false, majorGridlinesStyle: { … } }`
  // mean "hide them, even though I've authored a style for the ON
  // state" — dropping the style and emitting nothing is the right
  // call. The previous code checked the boolean in an `else if`, so
  // an explicit `false` was silently ignored whenever a style was
  // supplied and the gridlines stayed drawn.
  const majorGridlinesStyle = toShapeProperties(opts.majorGridlinesStyle);
  if (opts.majorGridlines === false) {
    axis.majorGridlines = undefined;
  } else if (majorGridlinesStyle) {
    axis.majorGridlines = majorGridlinesStyle;
  } else if (opts.majorGridlines === true) {
    axis.majorGridlines = {};
  }
  const minorGridlinesStyle = toShapeProperties(opts.minorGridlinesStyle);
  if (opts.minorGridlines === false) {
    axis.minorGridlines = undefined;
  } else if (minorGridlinesStyle) {
    axis.minorGridlines = minorGridlinesStyle;
  } else if (opts.minorGridlines === true) {
    axis.minorGridlines = {};
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
    // `a:bodyPr/@rot` is `ST_Angle` = `xsd:int`. Round so a caller
    // passing fractional degrees (or a value like 1.05° that IEEE 754
    // turns into 63000.000000000004) doesn't inject `"NaN"` /
    // `"63000.00000000001"` as an attribute value.
    axis.txPr = { rotation: Math.round(opts.textRotation * 60000) };
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
      opts.displayUnitsLabel !== undefined
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
    if (opts.baseTimeUnit !== undefined) {
      dateAx.baseTimeUnit = opts.baseTimeUnit;
    }
    if (opts.majorTimeUnit !== undefined) {
      dateAx.majorTimeUnit = opts.majorTimeUnit;
    }
    if (opts.minorTimeUnit !== undefined) {
      dateAx.minorTimeUnit = opts.minorTimeUnit;
    }
  }
  // Structured spPr takes priority over line-only shortcuts
  const customSpPr = toShapeProperties(opts.spPr);
  if (customSpPr) {
    axis.spPr = customSpPr;
  } else if (opts.lineColor || opts.lineWidth !== undefined || opts.lineDash) {
    const line: ChartLine = {};
    if (opts.lineColor) {
      line.color = hexToColor(opts.lineColor);
    }
    if (opts.lineWidth !== undefined) {
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
      const barDir = opts.barDir ?? "col";
      // Horizontal bar charts (`barDir="bar"`) swap the axis
      // directions: the category axis is on the left (vertical),
      // value axis at the bottom (horizontal). Excel itself emits
      // `axPos="l"` / `axPos="b"` respectively in that case.
      // `buildCatValAxes` produces the column-chart defaults
      // (`catAx.axPos="b"`, `valAx.axPos="l"`), which were left
      // unchanged for horizontal bar charts — the resulting XML
      // still rendered in Excel because Excel infers orientation
      // from `c:barDir`, but the renderer's `pickAxis` (position-
      // based) picked the wrong axis for gridlines / tick labels on
      // horizontal bars built via `addBarChart`.
      //
      // We deliberately do NOT move `majorGridlines` between axes —
      // `<c:majorGridlines/>` on a value axis means "draw gridlines
      // perpendicular to the value axis at every tick". When the
      // value axis is horizontal (bottom), the renderer correctly
      // projects those as vertical strokes because it picks up
      // gridlines from the X-axis slot (by `axPos`), not by
      // axis-type.
      if (barDir === "bar") {
        catAx.axPos = "l";
        valAx.axPos = "b";
      }
      const group: BarChartGroup = {
        type,
        barDir,
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
      const barDir = opts.barDir ?? "col";
      if (barDir === "bar") {
        catAx.axPos = "l";
        valAx.axPos = "b";
      }
      // `CT_Bar3DChart` does NOT accept `overlap` or `serLines` — both
      // are 2-D-only. Reject them loud here so the options validator
      // catches the mistake before the writer silently drops them.
      if (opts.overlap !== undefined) {
        throw new ChartOptionsError(
          'bar3D charts do not support `overlap` (valid only on 2-D `bar`). Remove the field or switch to `type: "bar"`.'
        );
      }
      if (opts.serLines !== undefined) {
        throw new ChartOptionsError(
          'bar3D charts do not support `serLines` (valid only on 2-D `bar`). Remove the field or switch to `type: "bar"`.'
        );
      }
      const group: BarChartGroup = {
        type,
        barDir,
        grouping: (opts.grouping as BarGrouping) ?? "clustered",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildBarSeries),
        gapWidth: opts.gapWidth ?? 150,
        gapDepth: opts.gapDepth,
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
      // `CT_Line3DChart` does NOT accept `marker`, `smooth`,
      // `hiLowLines`, or `upDownBars` — all 2-D-only.
      // Rejected upstream in validateChartLevelOptions.
      const group: LineChartGroup = {
        type,
        grouping: (opts.grouping as LineGrouping) ?? "standard",
        varyColors: opts.varyColors,
        series: seriesOpts.map(buildLineSeries),
        dropLines: opts.dropLines ? {} : undefined,
        gapDepth: opts.gapDepth,
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
        // `CT_Area3DChart` carries `gapDepth`; pass it through from
        // options. The equivalent lines for `bar3D` / `line3D` already
        // did this — the omission here silently dropped the option.
        gapDepth: opts.gapDepth,
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
      // `CT_StockChart` has no `varyColors` attribute per schema — see
      // `StockChartGroup` in types.ts. Reject the option here rather
      // than silently dropping it at emit time so the mistake
      // surfaces at authoring.
      if (opts.varyColors !== undefined) {
        throw new ChartOptionsError(
          "stock charts do not support `varyColors` (not in CT_StockChart). Remove the field or switch to a line / bar / area chart."
        );
      }
      const { catAx, valAx } = buildCatValAxes(axIds);
      const group: StockChartGroup = {
        type: "stock",
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
      throw new ChartOptionsError(`Unsupported chart type: ${String(_exhaustive)}.`);
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
  title?: string | { formula: string } | ChartRichText | null;
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
    // Don't mutate caller-owned `opts.pivotSource` — `extractPivotOptions`
    // already handles the merge priority between
    // `opts.pivotChartOptions` (explicit) and `pivotSource.options`
    // (embedded). Writing back here produced surprising side-effects
    // when callers reused the same `pivotSource` object across charts.
    chart.pivotFormats = [{ index: 0 }];
  }

  // Title handling has three mutually exclusive shapes:
  //
  //   1. `title === null`       → explicit suppression. Emit
  //      `autoTitleDeleted="1"` so Excel records the user removed the
  //      auto-title. Do NOT build a title frame even if `titleOptions`
  //      was also provided — the explicit `null` wins over layout /
  //      style hints.
  //   2. `title` truthy         → build the title, optionally apply
  //      `titleOptions` on top.
  //   3. `title` absent, `titleOptions` set → layout / style for the
  //      auto-generated title. Uncommon but valid.
  //   4. Everything else        → leave `autoTitleDeleted` undefined so
  //      the writer omits it, matching Excel's behaviour for a fresh
  //      unmodified chart.
  //
  // Checking `title === null` before `if (opts.title)` is important —
  // `if (null)` is falsy, so reversing the order landed an explicit
  // `null` in the `titleOptions`-only branch that built an empty title
  // frame.
  if (opts.title === null) {
    chart.autoTitleDeleted = true;
  } else if (opts.title) {
    chart.title = buildTitle(opts.title);
    if (opts.titleOptions) {
      applyTitleOptions(chart.title, opts.titleOptions);
    }
    chart.autoTitleDeleted = false;
  } else if (opts.titleOptions) {
    // Layout/style without title text is unusual but allowed
    chart.title = { overlay: false };
    applyTitleOptions(chart.title, opts.titleOptions);
    chart.autoTitleDeleted = false;
  }
  // else: leave autoTitleDeleted undefined so the writer omits it,
  // matching Excel's behaviour for a fresh unmodified chart — the
  // automatic title appears unless the user explicitly removes it.

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
        // Apply the group's axis options onto the shared secondary
        // axes. Without this, the `categoryAxis` / `valueAxis` fields
        // on a second (or third, …) secondary-axis group were silently
        // discarded along with the auto-built axes — `numFmt`,
        // `majorUnit`, `title`, `orientation`, etc. never reached the
        // shared axis object.
        applyAxisOptions(secondaryCatAx, groupOpts.categoryAxis);
        applyAxisOptions(secondaryValAx!, groupOpts.valueAxis);
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
        // Scatter/bubble secondary axes (valAx + valAx).
        //
        // `crossBetween` is only valid when the value axis crosses a
        // *category* axis — it specifies the tick-label alignment
        // relative to category boundaries (`between` / `midCat`). On a
        // val/val pair it has no defined meaning and strict OOXML
        // validators flag it. The primary scatter axes in
        // `buildValValAxes` deliberately omit it, so emit the secondary
        // scatter axes the same way to stay round-trip consistent.
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
            crosses: "autoZero"
          };
          secondaryYAx = {
            axisType: "val",
            axId: sYId,
            scaling: { orientation: "minMax" },
            delete: false,
            axPos: "r",
            crossAx: sXId,
            crosses: "max"
          };
          allAxes.push(secondaryXAx, secondaryYAx);
        }
        // `categoryAxis` / `valueAxis` map to scatter x / y respectively
        // (mirroring the primary scatter path in `buildChartTypeGroup`).
        applyAxisOptions(secondaryXAx, groupOpts.categoryAxis);
        applyAxisOptions(secondaryYAx!, groupOpts.valueAxis);
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
          // Reuse primary axes — but apply the current group's axis
          // options onto the shared objects so its customisations
          // aren't silently discarded with the throw-away auto-built
          // axes. Previously only the FIRST group's `categoryAxis` /
          // `valueAxis` reached the output; every subsequent group's
          // overrides were dropped on the floor.
          applyAxisOptions(primaryCatAx, groupOpts.categoryAxis);
          applyAxisOptions(primaryValAx!, groupOpts.valueAxis);
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
          // Same rationale as the cat/val reuse path above — apply this
          // group's scatter axis options onto the shared val/val pair.
          applyAxisOptions(primaryXAx, groupOpts.categoryAxis);
          applyAxisOptions(primaryYAx!, groupOpts.valueAxis);
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
