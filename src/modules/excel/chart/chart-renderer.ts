import { measureTextWidthPx } from "@excel/utils/text-metrics";

import { parseTxPr, getSpPrLine, getTxPrFontSize } from "./shape-properties";
import type {
  ChartAxis,
  ChartColor,
  ChartLegend,
  ChartMarker,
  ChartModel,
  ChartTextProperties,
  ChartTitle,
  ChartTypeGroup,
  DataLabelPosition,
  DataLabels,
  DataTable,
  EffectList,
  ErrorBars,
  LegendPosition,
  SeriesBase,
  Trendline
} from "./types";

/**
 * Options for the built-in deterministic chart preview renderer.
 *
 * The renderer is intentionally lightweight and self-contained; it produces a
 * stable SVG/PNG/PDF preview of chart data and core styling, not an
 * Excel-identical layout or rasterization.
 *
 * Notable scope boundaries:
 *
 * - **Text metrics** come from `@excel/utils/text-metrics` (Calibri/Arial/
 *   Times per-character widths + ~230 category fallbacks); legend and title
 *   layouts adapt to real label widths.
 * - **Data labels** honour every `DataLabelPosition` (see `positionDataLabel`).
 *   Pie/doughnut `outEnd`/`bestFit` layouts emit leader lines with a simple
 *   per-hemisphere greedy vertical nudge (`layoutPieLabels`). Bar/line labels
 *   drop or stack when `resolveLabelCollisions` detects bbox overlap.
 * - **PDF bridge** is feature-matched with the SVG path: labels (with
 *   anchor + color + fontFamily/bold/italic from `txPr`), markers
 *   (square/diamond/triangle/x/plus/circle), error bars, trendlines
 *   (with dash), leader lines, rotated axis titles, and real alpha on
 *   area/bubble/radar/bar3D fills flowing through `PdfColor.a` →
 *   `/ExtGState`. Surfaces that ignore the new optional fields receive
 *   the legacy call shape with pre-anchored coordinates and opaque
 *   colours, so pre-alpha consumers keep working unchanged.
 * - **3D charts** — `bar3D` renders as a true extruded box (top + front
 *   + right faces) whose axonometric projection is driven by
 *   `view3D.rotX` / `view3D.rotY` / `view3D.rAngAx`. See
 *   {@link resolveBar3DProjection}. The other 3D variants
 *   (`line3D`/`pie3D`/`area3D`/`surface3D`) render as their 2D
 *   equivalents — OOXML `view3D` / `Scene3D` / `ShapeProperties3D`
 *   metadata is preserved in the model but not consumed for those
 *   types; the renderer returns a deterministic preview rather than a
 *   full 3D scene.
 */
export interface ChartRenderOptions {
  width?: number;
  height?: number;
  title?: string;
  /** Background fill for SVG/PNG previews. Set to "transparent" for transparent PNG output. */
  backgroundColor?: string;
  /** Output scale multiplier for PNG previews. Useful for high-DPI exports. */
  scale?: number;
  /** PNG output DPI metadata. Stored as a pHYs chunk when provided. */
  dpi?: number;
  /**
   * Optional geographic data source for ChartEx `regionMap` previews.
   * When present and the matched features cover all labels, the
   * renderer draws country (or other region) outlines from the user-
   * supplied TopoJSON instead of falling back to the built-in
   * centroid-dot preview. Purely opt-in so the default library has
   * zero bundled geographic assets — supply e.g. `world-atlas`
   * `countries-110m.json` via `topology`.
   *
   * See {@link RegionMapDataOptions} and
   * `src/modules/excel/chart/topojson.ts`.
   */
  regionMap?: RegionMapDataOptions;
}

/**
 * Individual match rule for {@link RegionMapDataOptions.match}. A rule
 * is either the literal string `"id"` (compare against `feature.id`)
 * or `` `property:${propertyKey}` `` (compare against
 * `feature.properties[propertyKey]`). Multiple rules are combined via
 * an ordered fall-back array on the options object itself — see
 * {@link RegionMapDataOptions.match}.
 */
export type RegionMapMatchRule = "id" | `property:${string}`;

/**
 * User-supplied geographic data for `regionMap` rendering. The caller
 * is responsible for loading / caching the TopoJSON file — the library
 * is strictly zero-dependency and ships no map data.
 */
export interface RegionMapDataOptions {
  /**
   * Parsed TopoJSON topology. Compatible with the output of
   * `world-atlas` bundles or any `topojson-server` emitter.
   */
  topology: unknown;
  /**
   * Name of the geometry collection to draw — typically `"countries"`
   * for `world-atlas`. Throws at render time if missing from the
   * topology.
   */
  objectName: string;
  /**
   * How to match each chart category label to a TopoJSON feature.
   *
   * - `"id"` — compare against `feature.id`.
   * - `"property:<key>"` — compare against `feature.properties[key]`.
   *   Common for world-atlas: `"property:name"`.
   *
   * Accepts either a single rule (back-compat) or an ordered fall-back
   * list (`matchers`-style). When a list is supplied the renderer tries
   * each rule in order per feature and keeps the first one that finds a
   * data value for this feature; this is the pattern Natural-Earth-
   * derived topologies need when the categories are localised (e.g.
   * try `property:name_zh` first, then `property:name_en`, then fall
   * back to `id`). Comparison stays case-insensitive and
   * whitespace-trimmed for every rule. Default: `"id"`.
   */
  match?: RegionMapMatchRule | RegionMapMatchRule[];
  /**
   * Projection to use. Overrides `series.layoutPr.projection`. Supports
   * the same set the built-in renderer implements (`mercator`, `miller`,
   * `albers`, `robinson`).
   */
  projection?: "mercator" | "miller" | "albers" | "robinson";
  /**
   * Optional stroke colour for region borders. Default `"#FFFFFF"`.
   */
  strokeColor?: string;
}

export interface PdfChartRenderOptions extends ChartRenderOptions {
  x: number;
  y: number;
  /** Optional deterministic sink for regression tests; production drawing APIs can ignore it. */
  trace?: string[];
}

export interface ChartPdfDrawingSurface {
  drawRect(options: {
    x: number;
    y: number;
    width: number;
    height: number;
    fill?: PdfColor;
    stroke?: PdfColor;
    /** Stroke width when `stroke` is set. Surfaces that ignore it fall back to 1 px. */
    lineWidth?: number;
  }): this;
  drawLine(options: {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    color?: PdfColor;
    /** Stroke width. Surfaces that ignore it fall back to 1 px. */
    lineWidth?: number;
    /**
     * PDF dash pattern (even entries are "on" lengths, odd entries are "off"
     * lengths, in points). Omit for a solid stroke. Surfaces without dash
     * support may ignore this silently.
     */
    dashPattern?: number[];
  }): this;
  /**
   * Draw a single line of text. `anchor` and `rotation` are honoured by
   * surfaces that support them; those that do not are expected to
   * treat the request as `start` / `0°` (i.e. legacy `drawText` behaviour).
   * Chart callers should not assume rotation is available and should
   * still pre-compute anchored x coordinates when a visual fallback is
   * required (see {@link drawPdfText}).
   */
  drawText(
    text: string,
    options: {
      x: number;
      y: number;
      fontSize?: number;
      color?: PdfColor;
      /**
       * Degrees, clockwise. Optional; only the PDF bridge built on
       * `@pdf/builder` honours it today.
       */
      rotation?: number;
      /**
       * Horizontal alignment around `x`. Optional; chart code supplies an
       * explicit pre-anchored `x` as a fallback so surfaces that ignore
       * this still render at the correct position.
       */
      anchor?: "start" | "middle" | "end";
      bold?: boolean;
      italic?: boolean;
      fontFamily?: string;
    }
  ): this;
  drawCircle?(options: {
    cx: number;
    cy: number;
    r: number;
    fill?: PdfColor;
    stroke?: PdfColor;
    lineWidth?: number;
  }): this;
  drawPath?(
    ops: ChartPdfPathOp[],
    options?: {
      fill?: PdfColor;
      stroke?: PdfColor;
      closePath?: boolean;
      lineWidth?: number;
      dashPattern?: number[];
    }
  ): this;
}

export type ChartPdfPathOp =
  | { op: "move"; x: number; y: number }
  | { op: "line"; x: number; y: number }
  | { op: "curve"; x1: number; y1: number; x2: number; y2: number; x3: number; y3: number }
  | { op: "close" };

/**
 * RGB(A) colour triple used by the chart PDF bridge. `a` is optional and
 * defaults to 1 (fully opaque); surfaces that implement transparency
 * (e.g. `@pdf/builder` `PdfPageBuilder`) materialise `a < 1` as an
 * `/ExtGState` resource and emit the corresponding `gs` operator. Older
 * surfaces that ignore `a` render as opaque, which matches the
 * pre-alpha behaviour exactly.
 */
export interface PdfColor {
  r: number;
  g: number;
  b: number;
  a?: number;
}

export interface ChartScene {
  width: number;
  height: number;
  title?: ChartSceneText;
  plot: ChartSceneRect;
  axes: { x: ChartSceneLine; y: ChartSceneLine; x2?: ChartSceneLine; y2?: ChartSceneLine };
  gridlines: ChartSceneLine[];
  xLabels: ChartSceneText[];
  yLabels: ChartSceneText[];
  secondaryXLabels: ChartSceneText[];
  secondaryYLabels: ChartSceneText[];
  axisTitles: ChartSceneText[];
  series: ChartSceneSeries[];
  legend: ChartSceneLegend;
  /**
   * Data-table overlay drawn below the plot area when
   * `model.chart.plotArea.dataTable` is set. The preview writes a
   * compact grid using the same deterministic text-metrics pipeline as
   * the rest of the scene. Full OOXML `c:dTable` styling is honoured at
   * the XML-round-trip level; this preview covers the four display
   * switches (`showHorzBorder`, `showVertBorder`, `showOutline`,
   * `showKeys`) and the typography derived from `txPr`. When the data
   * table is present, the primary x-axis labels are suppressed — Excel
   * does the same so category names only appear once.
   */
  dataTable?: ChartSceneDataTable;
  /**
   * SVG filter definitions referenced by individual series via
   * `effectFilterId`. One entry per unique `a:effectLst` observed on
   * the normalised series. Rendered into `<defs>` by
   * {@link renderChartSvg}; the PDF surface ignores them because SVG
   * filters don't map directly onto PDF graphics state.
   */
  effectFilters: ChartSceneEffectFilter[];
}

/**
 * Layout-resolved representation of a `c:dTable` element for the
 * preview pipeline. Geometry is already in pixel space so SVG/PDF/PNG
 * bridges can consume it uniformly.
 */
export interface ChartSceneDataTable {
  /** Outer rectangle containing the entire table. */
  rect: ChartSceneRect;
  /**
   * Column boundaries (x coordinates) including the left and right
   * edges. `columns[0]` is the left edge of the series-name column,
   * `columns[1]` starts the first category cell, and so on.
   */
  columns: number[];
  /**
   * Row boundaries (y coordinates) including the top and bottom edges.
   * `rows[0]` is the header row (category names) if it's non-empty,
   * `rows[1..]` are the per-series rows.
   */
  rows: number[];
  /** Text nodes for every cell — series name plus each value. */
  cells: ChartSceneText[];
  /**
   * Legend-key swatches drawn to the left of each series name when
   * `showKeys` is enabled. Colors match the series `color`.
   */
  legendSwatches: Array<ChartSceneRect & { color: string }>;
  /** Border strokes derived from `showHorzBorder` / `showVertBorder` / `showOutline`. */
  borders: ChartSceneLine[];
}

export interface ChartSceneEffectFilter {
  /** Stable id the series references in `filter="url(#<id>)"`. */
  id: string;
  /** Full `<filter>…</filter>` XML produced by {@link buildEffectFilter}. */
  xml: string;
}

export interface ChartSceneLegend {
  items: Array<{ label: string; color: string }>;
  rect: ChartSceneRect;
  visible: boolean;
  position?: LegendPosition;
  orientation: "horizontal" | "vertical";
  /**
   * Font-related overrides derived from `model.chart.legend.txPr`
   * (see {@link textStyleFromTxPr}). When populated these flow into
   * both the SVG legend emit and the PDF legend renderer so legend
   * labels match the typography the author requested in the chart
   * XML. `undefined` fields keep the renderer defaults (10 pt Arial).
   */
  textStyle?: {
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    color?: string;
  };
}

export type ChartSceneSeries =
  | ChartSceneBarSeries
  | ChartSceneAreaSeries
  | ChartSceneLineSeries
  | ChartSceneBubbleSeries
  | ChartScenePieSeries
  | ChartSceneRadarSeries
  | ChartSceneStockSeries
  | ChartSceneSurfaceSeries;

export interface ChartSceneAdornment {
  labels?: ChartSceneText[];
  markers?: ChartSceneMarker[];
  trendlines?: ChartSceneTrendline[];
  errorBars?: ChartSceneErrorBar[];
  /**
   * Leader lines connecting data labels to their data points. Currently
   * emitted only by pie/doughnut series when the effective
   * `DataLabelPosition` places labels outside the slice (`outEnd`,
   * `bestFit`). The renderer treats them as decorative strokes and does
   * not attempt collision avoidance beyond the greedy fan-out already
   * applied by {@link buildDataLabels}.
   */
  leaderLines?: ChartSceneLine[];
  /**
   * SVG `<filter>` id assigned by {@link buildChartScene} when this
   * series carries a DrawingML `a:effectLst` (shadow, glow, reflection,
   * soft-edge, blur, inner-shadow). `renderChartSvg` emits the matching
   * `<filter>` in the SVG `<defs>` block and the series' primary shape
   * references it via `filter="url(#<id>)"`.
   *
   * Only present on series whose `spPr.effectList` was non-empty; other
   * series render without a filter attribute so the shared `<defs>` does
   * not bloat the output.
   */
  effectFilterId?: string;
}

export interface ChartSceneBarSeries extends ChartSceneAdornment {
  type: "bar";
  color: string;
  bars: ChartSceneRect[];
  label?: string;
  horizontal?: boolean;
  /**
   * Decorative depth hint (pixels) used only by the SVG/PNG preview to
   * distinguish `bar3D` from `bar`. When `projection3D` is also present
   * the renderer uses this depth along with the projection deltas to
   * emit a proper extruded box (top + front + right side faces, with
   * shading). For plain `bar` this stays `0` and only the front rect
   * is drawn.
   */
  depth?: number;
  /**
   * Axonometric projection deltas derived from OOXML `view3D.rotX` /
   * `view3D.rotY`. Populated for `bar3D` series so the renderer can
   * extrude each bar into a true 3D box: the back face sits at
   * `(bar.x + dx, bar.y - dy)` and back-right edge at
   * `(bar.x + bar.width + dx, …)`. `undefined` for plain `bar` series
   * where no 3D transform applies.
   */
  projection3D?: { dx: number; dy: number };
}

export interface ChartSceneAreaSeries extends ChartSceneAdornment {
  type: "area";
  color: string;
  points: ChartScenePoint[];
  lowerPoints?: ChartScenePoint[];
  baselineY: number;
  label?: string;
  closed?: boolean;
}

export interface ChartSceneLineSeries extends ChartSceneAdornment {
  type: "line" | "scatter";
  color: string;
  points: ChartScenePoint[];
  label?: string;
  smooth?: boolean;
  showLine?: boolean;
}

export interface ChartSceneBubbleSeries extends ChartSceneAdornment {
  type: "bubble";
  color: string;
  bubbles: ChartSceneBubble[];
  label?: string;
}

export interface ChartScenePieSeries extends ChartSceneAdornment {
  type: "pie" | "doughnut" | "ofPie";
  slices: ChartScenePieSlice[];
  secondarySlices?: ChartScenePieSlice[];
  connectors?: ChartSceneLine[];
  label?: string;
}

export interface ChartSceneRadarSeries extends ChartSceneAdornment {
  type: "radar";
  color: string;
  points: ChartScenePoint[];
  center: ChartScenePoint;
  radius: number;
  filled?: boolean;
  label?: string;
}

export interface ChartSceneStockSeries extends ChartSceneAdornment {
  type: "stock";
  color: string;
  candles: ChartSceneStockCandle[];
  label?: string;
}

export interface ChartSceneSurfaceSeries extends ChartSceneAdornment {
  type: "surface";
  cells: ChartSceneSurfaceCell[];
  wireframe?: boolean;
  label?: string;
}

export interface ChartSceneBubble extends ChartScenePoint {
  radius: number;
}

export interface ChartSceneMarker extends ChartScenePoint {
  color: string;
  size: number;
  symbol?: NonNullable<ChartMarker["symbol"]>;
}

export interface ChartSceneTrendline {
  color: string;
  width?: number;
  dash?: string;
  points: ChartScenePoint[];
  label?: ChartSceneText;
}

export interface ChartSceneErrorBar {
  line: ChartSceneLine;
  cap1?: ChartSceneLine;
  cap2?: ChartSceneLine;
}

export interface ChartSceneStockCandle {
  x: number;
  highY: number;
  lowY: number;
  openY?: number;
  closeY?: number;
  width: number;
  up: boolean;
}

export interface ChartSceneSurfaceCell extends ChartSceneRect {
  color: string;
}

interface NormalizedSeries {
  group: ChartTypeGroup;
  groupIndex: number;
  series: SeriesBase;
  seriesIndex: number;
  globalIndex: number;
  label: string;
  color: string;
  values: number[];
  categories?: string[];
  xValues?: number[];
  bubbleSizes?: number[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars[];
  marker?: ChartMarker;
}

interface ValueRange {
  min: number;
  max: number;
}

interface ChartAxisContext {
  axesById: Map<number, ChartAxis>;
  primaryXAxis?: ChartAxis;
  primaryYAxis?: ChartAxis;
  secondaryXAxis?: ChartAxis;
  secondaryYAxis?: ChartAxis;
  yRangesByAxisId: Map<number, ValueRange>;
  xRangesByAxisId: Map<number, ValueRange>;
  defaultYRange: ValueRange;
  defaultXRange: ValueRange;
}

export interface ChartScenePieSlice {
  color: string;
  cx: number;
  cy: number;
  radius: number;
  innerRadius: number;
  startAngle: number;
  endAngle: number;
}

export interface ChartSceneLine {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  width?: number;
}

export interface ChartSceneText {
  x: number;
  y: number;
  text: string;
  fontSize: number;
  color: string;
  anchor?: "start" | "middle" | "end";
  rotate?: number;
  /**
   * Font family typeface, populated from OOXML `a:latin/@typeface` when
   * the originating `txPr` carries a font. `undefined` keeps the
   * renderer defaults: `"Arial"` for SVG, `"Helvetica"` for PDF. The
   * SVG and PDF bridges both honour this — the PDF `FontManager`
   * handles unknown families by silently mapping to Helvetica.
   */
  fontFamily?: string;
  /** Bold from `a:rPr/@b` / `a:defRPr/@b`. */
  bold?: boolean;
  /** Italic from `a:rPr/@i` / `a:defRPr/@i`. */
  italic?: boolean;
}

export interface ChartSceneRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ChartScenePoint {
  x: number;
  y: number;
}

const DEFAULT_WIDTH = 640;
const DEFAULT_HEIGHT = 360;
const COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
const AXIS_COLOR = "#444444";
const GRID_COLOR = "#D9D9D9";

export function buildChartScene(model: ChartModel, options: ChartRenderOptions = {}): ChartScene {
  const width = options.width ?? DEFAULT_WIDTH;
  const height = options.height ?? DEFAULT_HEIGHT;
  const titleText = options.title ?? extractTitle(model);
  const groups = model.chart.plotArea.chartTypes;
  const normalized = normalizeSeries(groups, model);
  const seriesValues = normalized.map(s => s.values);
  const categories =
    normalized.find(s => s.categories && s.categories.length > 0)?.categories ??
    seriesValues[0]?.map((_, i) => String(i + 1)) ??
    [];
  const legend = buildSceneLegend(model.chart.legend, normalized, width, height, !!titleText);
  // Data tables sit below the plot area. Pre-compute their vertical
  // footprint so `getPlotRect` can reserve space before series-level
  // geometry is built. Sizing uses the same text-metrics pipeline as
  // the rest of the scene for consistent wrap behaviour.
  const dataTableSpec = model.chart.plotArea.dataTable;
  const dataTableHeight = dataTableSpec
    ? computeDataTableHeight(dataTableSpec, normalized, categories)
    : 0;
  const plot = getPlotRect(width, height, !!titleText, legend, model, dataTableHeight);
  const axisContext = buildAxisContext(model, normalized);
  const sceneSeries = buildSceneSeries(
    groups,
    normalized,
    plot,
    axisContext,
    categories,
    model.chart.view3D
  );
  // Harvest any `a:effectLst` DrawingML from each series' `spPr` and
  // translate it into SVG `<filter>` markup. Two series with identical
  // effect lists share a single filter definition — the cache key is
  // a stable JSON stringify of the effect tree. The writer (renderChartSvg)
  // embeds the filters in `<defs>` and each series references its
  // filter via the `effectFilterId` we thread onto the scene.
  const effectFilters = assignEffectFilters(normalized, sceneSeries);
  const dataTable = dataTableSpec
    ? buildSceneDataTable(dataTableSpec, normalized, categories, plot, dataTableHeight)
    : undefined;
  const primaryYRange = axisContext.primaryYAxis
    ? (axisContext.yRangesByAxisId.get(axisContext.primaryYAxis.axId) ?? axisContext.defaultYRange)
    : axisContext.defaultYRange;
  const secondaryYRange = axisContext.secondaryYAxis
    ? (axisContext.yRangesByAxisId.get(axisContext.secondaryYAxis.axId) ??
      axisContext.defaultYRange)
    : undefined;
  const primaryXRange = axisContext.primaryXAxis
    ? (axisContext.xRangesByAxisId.get(axisContext.primaryXAxis.axId) ?? axisContext.defaultXRange)
    : axisContext.defaultXRange;
  const secondaryXRange = axisContext.secondaryXAxis
    ? (axisContext.xRangesByAxisId.get(axisContext.secondaryXAxis.axId) ??
      axisContext.defaultXRange)
    : undefined;
  const secondaryXCategories = categoriesForAxis(
    normalized,
    axisContext.secondaryXAxis,
    categories
  );

  return {
    width,
    height,
    title: titleText
      ? {
          x: width / 2,
          y: 26,
          text: titleText,
          fontSize: 18,
          // Default chart-title colour matches the pre-textStyle
          // renderer; authors who supplied an explicit txPr colour win.
          color:
            (model.chart.title
              ? colorFromChartTextProperties(model.chart.title.txPr)
              : undefined) ?? "#222222",
          anchor: "middle",
          ...textStyleFromTxPr(model.chart.title?.txPr)
        }
      : undefined,
    plot,
    gridlines: buildGridlines(plot, axisContext.primaryYAxis),
    // When a data table is drawn, Excel suppresses the primary x-axis
    // category labels because categories already appear as the header
    // row of the table. We keep gridlines and the axis line itself so
    // the plot boundary stays intact.
    xLabels: dataTable
      ? []
      : buildXLabels(categories, plot, axisContext.primaryXAxis, primaryXRange),
    yLabels: buildYLabels(
      primaryYRange.min,
      primaryYRange.max,
      plot,
      axisContext.primaryYAxis,
      false
    ),
    secondaryXLabels: axisContext.secondaryXAxis
      ? buildXLabels(secondaryXCategories, plot, axisContext.secondaryXAxis, secondaryXRange, true)
      : [],
    secondaryYLabels:
      axisContext.secondaryYAxis && secondaryYRange
        ? buildYLabels(
            secondaryYRange.min,
            secondaryYRange.max,
            plot,
            axisContext.secondaryYAxis,
            true
          )
        : [],
    axisTitles: buildAxisTitles(axisContext, plot),
    series: sceneSeries,
    legend,
    dataTable,
    effectFilters,
    axes: {
      x: applyLineStyle(
        {
          x1: plot.x,
          y1: plot.y + plot.height,
          x2: plot.x + plot.width,
          y2: plot.y + plot.height,
          color: AXIS_COLOR
        },
        axisContext.primaryXAxis
      ),
      y: applyLineStyle(
        {
          x1: plot.x,
          y1: plot.y,
          x2: plot.x,
          y2: plot.y + plot.height,
          color: AXIS_COLOR
        },
        axisContext.primaryYAxis
      ),
      x2: axisContext.secondaryXAxis
        ? applyLineStyle(
            {
              x1: plot.x,
              y1: plot.y,
              x2: plot.x + plot.width,
              y2: plot.y,
              color: AXIS_COLOR
            },
            axisContext.secondaryXAxis
          )
        : undefined,
      y2: axisContext.secondaryYAxis
        ? applyLineStyle(
            {
              x1: plot.x + plot.width,
              y1: plot.y,
              x2: plot.x + plot.width,
              y2: plot.y + plot.height,
              color: AXIS_COLOR
            },
            axisContext.secondaryYAxis
          )
        : undefined
    }
  };
}

export function renderChartSvg(model: ChartModel, options: ChartRenderOptions = {}): string {
  const scene = buildChartScene(model, options);
  const parts: string[] = [];
  const backgroundColor = options.backgroundColor ?? "#fff";
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${scene.width}" height="${scene.height}" viewBox="0 0 ${scene.width} ${scene.height}">`
  );
  parts.push("<!-- deterministic preview; not an Excel-identical layout -->");
  if (backgroundColor !== "transparent") {
    parts.push(`<rect width="100%" height="100%" fill="${escapeXmlAttr(backgroundColor)}"/>`);
  }
  // Emit filter definitions before any drawing so shapes can reference
  // them via `filter="url(#...)"`. Keeping `<defs>` at the top of the
  // SVG matches the convention used by Inkscape/Illustrator and plays
  // nicely with rasteriser pre-scan passes.
  if (scene.effectFilters.length > 0) {
    parts.push("<defs>");
    for (const f of scene.effectFilters) {
      parts.push(f.xml);
    }
    parts.push("</defs>");
  }
  if (scene.title) {
    parts.push(renderSvgText(scene.title));
  }
  for (const gridline of scene.gridlines) {
    parts.push(renderSvgLine(gridline));
  }
  parts.push(renderSvgLine(scene.axes.x));
  parts.push(renderSvgLine(scene.axes.y));
  if (scene.axes.x2) {
    parts.push(renderSvgLine(scene.axes.x2));
  }
  if (scene.axes.y2) {
    parts.push(renderSvgLine(scene.axes.y2));
  }
  for (const label of scene.xLabels) {
    parts.push(renderSvgText(label));
  }
  for (const label of scene.yLabels) {
    parts.push(renderSvgText(label));
  }
  for (const label of scene.secondaryXLabels) {
    parts.push(renderSvgText(label));
  }
  for (const label of scene.secondaryYLabels) {
    parts.push(renderSvgText(label));
  }
  for (const title of scene.axisTitles) {
    parts.push(renderSvgText(title));
  }
  for (const s of scene.series) {
    renderSvgSeries(parts, s);
  }
  if (scene.dataTable) {
    renderSvgDataTable(parts, scene.dataTable);
  }
  renderSvgLegend(parts, scene.legend);
  parts.push("</svg>");
  return parts.join("");
}

export async function renderChartPng(
  model: ChartModel,
  options: ChartRenderOptions = {}
): Promise<Uint8Array> {
  const width = options.width ?? DEFAULT_WIDTH;
  const height = options.height ?? DEFAULT_HEIGHT;
  const scale = normalizePngScale(options.scale);
  const svg = renderChartSvg(model, { ...options, width, height });
  return renderSvgToPng(svg, { width, height, scale, dpi: options.dpi });
}

export async function renderSvgToPng(
  svg: string,
  options: Required<Pick<ChartRenderOptions, "width" | "height">> &
    Pick<ChartRenderOptions, "scale" | "dpi">
): Promise<Uint8Array> {
  const width = Math.max(1, Math.round(options.width));
  const height = Math.max(1, Math.round(options.height));
  const scale = normalizePngScale(options.scale);
  const outputWidth = Math.max(1, Math.round(width * scale));
  const outputHeight = Math.max(1, Math.round(height * scale));
  if (typeof document === "undefined" || typeof Image === "undefined") {
    return renderSvgToBasicPng(svg, width, height, scale, options.dpi);
  }
  const blob = new Blob([svg], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    const image = await loadImage(url);
    const canvas = document.createElement("canvas");
    canvas.width = outputWidth;
    canvas.height = outputHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      throw new Error("Canvas 2D context is unavailable");
    }
    ctx.drawImage(image, 0, 0, outputWidth, outputHeight);
    const pngBlob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(result => {
        if (result) {
          resolve(result);
        } else {
          reject(new Error("Canvas PNG encoding failed"));
        }
      }, "image/png");
    });
    return new Uint8Array(await pngBlob.arrayBuffer());
  } finally {
    URL.revokeObjectURL(url);
  }
}

function renderSvgToBasicPng(
  svg: string,
  width: number,
  height: number,
  scale: number,
  dpi: number | undefined
): Uint8Array {
  const outputWidth = Math.max(1, Math.round(width * scale));
  const outputHeight = Math.max(1, Math.round(height * scale));
  const canvas = new BasicRasterCanvas(outputWidth, outputHeight, scale);
  const tagRe = /<(rect|line|circle|polyline|polygon|path)\b[^>]*>|<text\b[^>]*>[\s\S]*?<\/text>/g;
  let match: RegExpExecArray | null;
  while ((match = tagRe.exec(svg)) !== null) {
    const tag = match[0];
    const name = tag.startsWith("<text") ? "text" : match[1];
    const attrs = parseSvgAttrs(tag);
    if (name === "rect") {
      const x = numAttr(attrs, "x") * scale;
      const y = numAttr(attrs, "y") * scale;
      const rectWidth = numAttr(attrs, "width", 0, width) * scale;
      const rectHeight = numAttr(attrs, "height", 0, height) * scale;
      const strokeWidth = numAttr(attrs, "stroke-width", 1) * scale;
      if (attrs.fill !== undefined) {
        canvas.fillRect(x, y, rectWidth, rectHeight, attrs.fill);
      }
      if (attrs.stroke !== undefined && strokeWidth > 0) {
        canvas.strokeRect(x, y, rectWidth, rectHeight, attrs.stroke, strokeWidth);
      }
    } else if (name === "line") {
      canvas.drawLine(
        numAttr(attrs, "x1") * scale,
        numAttr(attrs, "y1") * scale,
        numAttr(attrs, "x2") * scale,
        numAttr(attrs, "y2") * scale,
        attrs.stroke,
        numAttr(attrs, "stroke-width", 1) * scale
      );
    } else if (name === "circle") {
      canvas.fillCircle(
        numAttr(attrs, "cx") * scale,
        numAttr(attrs, "cy") * scale,
        numAttr(attrs, "r") * scale,
        attrs.fill
      );
      canvas.strokeCircle(
        numAttr(attrs, "cx") * scale,
        numAttr(attrs, "cy") * scale,
        numAttr(attrs, "r") * scale,
        attrs.stroke,
        numAttr(attrs, "stroke-width", 1) * scale
      );
    } else if (name === "polyline") {
      canvas.drawPolyline(
        parseSvgPoints(attrs.points, scale),
        attrs.stroke,
        numAttr(attrs, "stroke-width", 1) * scale
      );
    } else if (name === "polygon") {
      const points = parseSvgPoints(attrs.points, scale);
      canvas.fillPolygon(points, attrs.fill);
      if (points.length > 0) {
        canvas.drawPolyline(
          [...points, points[0]],
          attrs.stroke,
          numAttr(attrs, "stroke-width", 1) * scale
        );
      }
    } else if (name === "path") {
      const points = parsePathPoints(attrs.d, scale);
      canvas.fillPolygon(points, attrs.fill);
      if (points.length > 0) {
        canvas.drawPolyline(
          [...points, points[0]],
          attrs.stroke,
          numAttr(attrs, "stroke-width", 1) * scale
        );
      }
    } else if (name === "text") {
      const rotation = parseSvgRotateTransform(attrs.transform);
      canvas.drawText(
        numAttr(attrs, "x") * scale,
        numAttr(attrs, "y") * scale,
        decodeSvgText(tag.match(/<text\b[^>]*>([\s\S]*?)<\/text>/)?.[1] ?? ""),
        numAttr(attrs, "font-size", 10) * scale,
        attrs.fill,
        attrs["text-anchor"],
        rotation
          ? {
              angle: rotation.angle,
              originX: rotation.originX * scale,
              originY: rotation.originY * scale
            }
          : undefined
      );
    }
  }
  return encodePng(outputWidth, outputHeight, canvas.data, dpi);
}

class BasicRasterCanvas {
  readonly data: Uint8Array;

  constructor(
    readonly width: number,
    readonly height: number,
    private readonly scale = 1
  ) {
    this.data = new Uint8Array(width * height * 4);
  }

  fillRect(x: number, y: number, width: number, height: number, color: string | undefined): void {
    const rgba = parseSvgColor(color);
    if (!rgba || width <= 0 || height <= 0) {
      return;
    }
    const x0 = clampInt(Math.floor(x), 0, this.width);
    const y0 = clampInt(Math.floor(y), 0, this.height);
    const x1 = clampInt(Math.ceil(x + width), 0, this.width);
    const y1 = clampInt(Math.ceil(y + height), 0, this.height);
    for (let yy = y0; yy < y1; yy++) {
      for (let xx = x0; xx < x1; xx++) {
        this.setPixel(xx, yy, rgba);
      }
    }
  }

  strokeRect(
    x: number,
    y: number,
    width: number,
    height: number,
    color: string | undefined,
    strokeWidth = 1
  ): void {
    this.drawLine(x, y, x + width, y, color, strokeWidth);
    this.drawLine(x + width, y, x + width, y + height, color, strokeWidth);
    this.drawLine(x + width, y + height, x, y + height, color, strokeWidth);
    this.drawLine(x, y + height, x, y, color, strokeWidth);
  }

  drawPolyline(points: ChartScenePoint[], color: string | undefined, width = 1): void {
    if (points.length < 2) {
      return;
    }
    for (let i = 1; i < points.length; i++) {
      this.drawLine(points[i - 1].x, points[i - 1].y, points[i].x, points[i].y, color, width);
    }
  }

  drawLine(
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    color: string | undefined,
    width = 1
  ): void {
    const rgba = parseSvgColor(color);
    if (!rgba) {
      return;
    }
    let x0 = Math.round(x1);
    let y0 = Math.round(y1);
    const xEnd = Math.round(x2);
    const yEnd = Math.round(y2);
    const dx = Math.abs(xEnd - x0);
    const sx = x0 < xEnd ? 1 : -1;
    const dy = -Math.abs(yEnd - y0);
    const sy = y0 < yEnd ? 1 : -1;
    let err = dx + dy;
    const radius = Math.max(0, Math.floor(width / 2));
    while (true) {
      for (let yy = y0 - radius; yy <= y0 + radius; yy++) {
        for (let xx = x0 - radius; xx <= x0 + radius; xx++) {
          this.setPixel(xx, yy, rgba);
        }
      }
      if (x0 === xEnd && y0 === yEnd) {
        break;
      }
      const e2 = 2 * err;
      if (e2 >= dy) {
        err += dy;
        x0 += sx;
      }
      if (e2 <= dx) {
        err += dx;
        y0 += sy;
      }
    }
  }

  fillCircle(cx: number, cy: number, r: number, color: string | undefined): void {
    const rgba = parseSvgColor(color);
    if (!rgba || r <= 0) {
      return;
    }
    const x0 = Math.floor(cx - r);
    const x1 = Math.ceil(cx + r);
    const y0 = Math.floor(cy - r);
    const y1 = Math.ceil(cy + r);
    const rr = r * r;
    for (let y = y0; y <= y1; y++) {
      for (let x = x0; x <= x1; x++) {
        if ((x - cx) ** 2 + (y - cy) ** 2 <= rr) {
          this.setPixel(x, y, rgba);
        }
      }
    }
  }

  strokeCircle(cx: number, cy: number, r: number, color: string | undefined, width = 1): void {
    const points: ChartScenePoint[] = [];
    const steps = Math.max(12, Math.ceil(r * 2));
    for (let i = 0; i <= steps; i++) {
      const a = (i / steps) * Math.PI * 2;
      points.push({ x: cx + Math.cos(a) * r, y: cy + Math.sin(a) * r });
    }
    this.drawPolyline(points, color, width);
  }

  fillPolygon(points: ChartScenePoint[], color: string | undefined): void {
    const rgba = parseSvgColor(color);
    if (!rgba || points.length < 3) {
      return;
    }
    const minY = clampInt(Math.floor(Math.min(...points.map(p => p.y))), 0, this.height - 1);
    const maxY = clampInt(Math.ceil(Math.max(...points.map(p => p.y))), 0, this.height - 1);
    for (let y = minY; y <= maxY; y++) {
      const intersections: number[] = [];
      for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
        const a = points[i];
        const b = points[j];
        if (a.y > y !== b.y > y) {
          intersections.push(((b.x - a.x) * (y - a.y)) / (b.y - a.y) + a.x);
        }
      }
      intersections.sort((a, b) => a - b);
      for (let i = 0; i < intersections.length; i += 2) {
        const x0 = clampInt(Math.floor(intersections[i]), 0, this.width - 1);
        const x1 = clampInt(Math.ceil(intersections[i + 1] ?? intersections[i]), 0, this.width - 1);
        for (let x = x0; x <= x1; x++) {
          this.setPixel(x, y, rgba);
        }
      }
    }
  }

  drawText(
    x: number,
    y: number,
    text: string,
    fontSize: number,
    color: string | undefined,
    anchor: string | undefined,
    rotation?: { angle: number; originX: number; originY: number }
  ): void {
    if (!text) {
      return;
    }
    // Use the same font-metrics engine as the SVG path so legend/title
    // anchoring rasterises at the right offsets. Fall back to the old
    // approximation if the measurement returns zero (defensive; the
    // metrics engine always returns a positive value for non-empty text).
    const measured = estimateTextWidth(text, fontSize);
    const textWidth = measured > 0 ? measured : Math.max(1, fontSize * 0.5) * text.length;
    const charWidth = textWidth / text.length;
    const startX = anchor === "middle" ? x - textWidth / 2 : anchor === "end" ? x - textWidth : x;
    const top = y - fontSize * 0.75;
    const stroke = Math.max(1, Math.ceil(fontSize * 0.12));
    if (!rotation || rotation.angle === 0) {
      for (let i = 0; i < text.length; i++) {
        const xx = startX + i * charWidth;
        this.fillRect(xx, top, Math.max(1, charWidth * 0.55), stroke, color);
        this.fillRect(xx, top + fontSize * 0.45, Math.max(1, charWidth * 0.45), stroke, color);
      }
      return;
    }
    // Rasterise each pseudo-glyph rectangle's pixels and apply the SVG-style
    // rotation around (originX, originY). This preserves the visible
    // direction of rotated axis-title and tick-label text in the Node PNG
    // fallback, matching what the SVG already emits via
    // `transform="rotate(angle x y)"`. The renderer still uses the same
    // coarse two-bar glyphs as the unrotated path — the goal here is
    // "correct orientation", not "real typography".
    const theta = (rotation.angle * Math.PI) / 180;
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    const ox = rotation.originX;
    const oy = rotation.originY;
    const rotatePixel = (px: number, py: number): [number, number] => {
      const dx = px - ox;
      const dy = py - oy;
      return [ox + dx * cos - dy * sin, oy + dx * sin + dy * cos];
    };
    const fillRotatedRect = (rx: number, ry: number, rw: number, rh: number): void => {
      const rgba = parseSvgColor(color);
      if (!rgba) {
        return;
      }
      // Iterate over the axis-aligned source rectangle's pixel grid and
      // plot each rotated point. Using Math.ceil on dimensions keeps
      // single-pixel strokes visible after rotation.
      const w = Math.max(1, Math.ceil(rw));
      const h = Math.max(1, Math.ceil(rh));
      for (let j = 0; j < h; j++) {
        for (let i = 0; i < w; i++) {
          const [tx, ty] = rotatePixel(rx + i, ry + j);
          this.setPixel(Math.round(tx), Math.round(ty), rgba);
        }
      }
    };
    for (let i = 0; i < text.length; i++) {
      const xx = startX + i * charWidth;
      fillRotatedRect(xx, top, Math.max(1, charWidth * 0.55), stroke);
      fillRotatedRect(xx, top + fontSize * 0.45, Math.max(1, charWidth * 0.45), stroke);
    }
  }

  private setPixel(x: number, y: number, rgba: [number, number, number, number]): void {
    if (x < 0 || y < 0 || x >= this.width || y >= this.height) {
      return;
    }
    const i = (y * this.width + x) * 4;
    const alpha = rgba[3] / 255;
    const inverse = 1 - alpha;
    this.data[i] = Math.round(rgba[0] * alpha + this.data[i] * inverse);
    this.data[i + 1] = Math.round(rgba[1] * alpha + this.data[i + 1] * inverse);
    this.data[i + 2] = Math.round(rgba[2] * alpha + this.data[i + 2] * inverse);
    this.data[i + 3] = Math.round(rgba[3] + this.data[i + 3] * inverse);
  }
}

function parseSvgAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const attrRe = /([\w:-]+)="([^"]*)"/g;
  let match: RegExpExecArray | null;
  while ((match = attrRe.exec(tag)) !== null) {
    attrs[match[1]] = match[2];
  }
  return attrs;
}

/**
 * Parse the `rotate(angle [x y])` form of an SVG transform attribute.
 *
 * The chart renderer emits rotation on text exclusively as
 * `transform="rotate(angle x y)"` (see renderSvgText), so a minimal parser
 * is sufficient. Returns `undefined` for any other transform shape — the
 * Node PNG fallback is deliberately narrow in scope, not a general SVG
 * engine. Missing origin coordinates default to 0/0 per SVG semantics.
 */
function parseSvgRotateTransform(
  transform: string | undefined
): { angle: number; originX: number; originY: number } | undefined {
  if (!transform) {
    return undefined;
  }
  const match =
    /rotate\(\s*(-?\d+(?:\.\d+)?)\s*(?:[,\s]\s*(-?\d+(?:\.\d+)?)\s*[,\s]\s*(-?\d+(?:\.\d+)?)\s*)?\)/.exec(
      transform
    );
  if (!match) {
    return undefined;
  }
  const angle = Number.parseFloat(match[1]);
  if (!Number.isFinite(angle) || angle === 0) {
    return undefined;
  }
  const originX = match[2] !== undefined ? Number.parseFloat(match[2]) : 0;
  const originY = match[3] !== undefined ? Number.parseFloat(match[3]) : 0;
  return { angle, originX, originY };
}

function numAttr(
  attrs: Record<string, string>,
  name: string,
  fallback = 0,
  percentBase?: number
): number {
  const value = attrs[name];
  if (value === undefined) {
    return fallback;
  }
  if (value.endsWith("%")) {
    const percent = Number.parseFloat(value);
    return Number.isFinite(percent) && percentBase !== undefined
      ? (percent / 100) * percentBase
      : fallback;
  }
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseSvgPoints(input: string | undefined, scale = 1): ChartScenePoint[] {
  if (!input) {
    return [];
  }
  const values = input
    .trim()
    .split(/[\s,]+/)
    .map(Number)
    .filter(Number.isFinite);
  const points: ChartScenePoint[] = [];
  for (let i = 0; i + 1 < values.length; i += 2) {
    points.push({ x: values[i] * scale, y: values[i + 1] * scale });
  }
  return points;
}

function parsePathPoints(input: string | undefined, scale = 1): ChartScenePoint[] {
  if (!input) {
    return [];
  }
  const tokens = input.match(/[MLAZ]|-?\d+(?:\.\d+)?/gi) ?? [];
  const points: ChartScenePoint[] = [];
  let i = 0;
  let current: ChartScenePoint | undefined;
  let start: ChartScenePoint | undefined;
  while (i < tokens.length) {
    const command = tokens[i++].toUpperCase();
    if (command === "M" || command === "L") {
      const point = readPathPoint(tokens, i);
      if (!point) {
        break;
      }
      point.x *= scale;
      point.y *= scale;
      i += 2;
      current = point;
      start ??= point;
      points.push(point);
    } else if (command === "A") {
      if (!current) {
        break;
      }
      const arc = readPathArc(tokens, i, current, scale);
      if (!arc) {
        break;
      }
      i += 7;
      points.push(...arc.points);
      current = arc.end;
    } else if (command === "Z") {
      if (start) {
        points.push(start);
      }
    } else {
      const numeric = Number.parseFloat(command);
      if (!Number.isFinite(numeric) || i >= tokens.length) {
        break;
      }
      const y = Number.parseFloat(tokens[i++]);
      if (!Number.isFinite(y)) {
        break;
      }
      current = { x: numeric * scale, y: y * scale };
      start ??= current;
      points.push(current);
    }
  }
  return points;
}

function readPathPoint(tokens: string[], index: number): ChartScenePoint | undefined {
  const x = Number.parseFloat(tokens[index]);
  const y = Number.parseFloat(tokens[index + 1]);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : undefined;
}

function readPathArc(
  tokens: string[],
  index: number,
  from: ChartScenePoint,
  scale: number
): { end: ChartScenePoint; points: ChartScenePoint[] } | undefined {
  const rx = Number.parseFloat(tokens[index]);
  const ry = Number.parseFloat(tokens[index + 1]);
  const largeArc = Number.parseFloat(tokens[index + 3]) === 1;
  const sweep = Number.parseFloat(tokens[index + 4]) === 1;
  const end = readPathPoint(tokens, index + 5);
  if (!Number.isFinite(rx) || !Number.isFinite(ry) || !end) {
    return undefined;
  }
  end.x *= scale;
  end.y *= scale;
  return {
    end,
    points: approximateArcPoints(from, end, Math.max(rx, ry) * scale, largeArc, sweep)
  };
}

function approximateArcPoints(
  start: ChartScenePoint,
  end: ChartScenePoint,
  radius: number,
  largeArc: boolean,
  sweep: boolean
): ChartScenePoint[] {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const chord = Math.hypot(dx, dy);
  if (!Number.isFinite(radius) || radius <= 0 || chord === 0 || chord > radius * 2) {
    return [end];
  }
  const mx = (start.x + end.x) / 2;
  const my = (start.y + end.y) / 2;
  const halfChord = chord / 2;
  const h = Math.sqrt(Math.max(0, radius * radius - halfChord * halfChord));
  const nx = -dy / chord;
  const ny = dx / chord;
  const candidates = [
    { x: mx + nx * h, y: my + ny * h },
    { x: mx - nx * h, y: my - ny * h }
  ];
  const selected =
    candidates.find(center => {
      const delta = arcDelta(center, start, end, sweep);
      return Math.abs(delta) > Math.PI === largeArc;
    }) ?? candidates[0];
  const startAngle = Math.atan2(start.y - selected.y, start.x - selected.x);
  const delta = arcDelta(selected, start, end, sweep);
  const steps = clampInt(Math.ceil((Math.abs(delta) * radius) / 8), 4, 90);
  const points: ChartScenePoint[] = [];
  for (let i = 1; i <= steps; i++) {
    const angle = startAngle + (delta * i) / steps;
    points.push({
      x: selected.x + Math.cos(angle) * radius,
      y: selected.y + Math.sin(angle) * radius
    });
  }
  return points;
}

function arcDelta(
  center: ChartScenePoint,
  start: ChartScenePoint,
  end: ChartScenePoint,
  sweep: boolean
): number {
  const startAngle = Math.atan2(start.y - center.y, start.x - center.x);
  const endAngle = Math.atan2(end.y - center.y, end.x - center.x);
  let delta = endAngle - startAngle;
  if (sweep && delta < 0) {
    delta += Math.PI * 2;
  } else if (!sweep && delta > 0) {
    delta -= Math.PI * 2;
  }
  return delta;
}

function parseSvgColor(color: string | undefined): [number, number, number, number] | undefined {
  if (!color || color === "none" || color === "transparent") {
    return undefined;
  }
  const normalized = color.startsWith("#") ? color.slice(1) : color;
  if (/^[0-9a-fA-F]{3}$/.test(normalized)) {
    return [
      Number.parseInt(normalized[0] + normalized[0], 16),
      Number.parseInt(normalized[1] + normalized[1], 16),
      Number.parseInt(normalized[2] + normalized[2], 16),
      255
    ];
  }
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return [
      Number.parseInt(normalized.slice(0, 2), 16),
      Number.parseInt(normalized.slice(2, 4), 16),
      Number.parseInt(normalized.slice(4, 6), 16),
      255
    ];
  }
  return undefined;
}

function decodeSvgText(value: string): string {
  return value
    .replace(/<[^>]*>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function clampInt(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function encodePng(width: number, height: number, rgba: Uint8Array, dpi?: number): Uint8Array {
  const scanlines = new Uint8Array((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) {
    const src = y * width * 4;
    const dst = y * (width * 4 + 1);
    scanlines[dst] = 0;
    scanlines.set(rgba.subarray(src, src + width * 4), dst + 1);
  }
  const chunks: Uint8Array[] = [];
  chunks.push(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  chunks.push(
    pngChunk("IHDR", concatBytes([u32be(width), u32be(height), new Uint8Array([8, 6, 0, 0, 0])]))
  );
  if (dpi !== undefined) {
    const pixelsPerMeter = Math.max(1, Math.round(dpi / 0.0254));
    chunks.push(
      pngChunk(
        "pHYs",
        concatBytes([u32be(pixelsPerMeter), u32be(pixelsPerMeter), new Uint8Array([1])])
      )
    );
  }
  chunks.push(pngChunk("IDAT", zlibStored(scanlines)));
  chunks.push(pngChunk("IEND", new Uint8Array()));
  return concatBytes(chunks);
}

function normalizePngScale(scale: number | undefined): number {
  if (scale === undefined) {
    return 1;
  }
  if (!Number.isFinite(scale) || scale <= 0) {
    throw new Error("chart render scale must be a positive finite number");
  }
  return Math.min(8, scale);
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = new TextEncoder().encode(type);
  const crcInput = concatBytes([typeBytes, data]);
  return concatBytes([u32be(data.length), typeBytes, data, u32be(crc32Bytes(crcInput))]);
}

function zlibStored(data: Uint8Array): Uint8Array {
  const blocks: Uint8Array[] = [new Uint8Array([0x78, 0x01])];
  let offset = 0;
  while (offset < data.length) {
    const len = Math.min(0xffff, data.length - offset);
    const final = offset + len >= data.length ? 1 : 0;
    blocks.push(
      new Uint8Array([final, len & 0xff, (len >>> 8) & 0xff, ~len & 0xff, (~len >>> 8) & 0xff])
    );
    blocks.push(data.subarray(offset, offset + len));
    offset += len;
  }
  blocks.push(u32be(adler32(data)));
  return concatBytes(blocks);
}

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function u32be(value: number): Uint8Array {
  return new Uint8Array([
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff
  ]);
}

function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  for (const byte of data) {
    a = (a + byte) % 65521;
    b = (b + a) % 65521;
  }
  return ((b << 16) | a) >>> 0;
}

function crc32Bytes(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export function drawChartPdf(
  page: ChartPdfDrawingSurface,
  model: ChartModel,
  options: PdfChartRenderOptions
): ChartPdfDrawingSurface {
  const scene = translateScene(buildChartScene(model, options), options.x, options.y, true);
  const trace = options.trace;
  trace?.push(
    `canvas:${fmt(scene.width)}x${fmt(scene.height)}@${fmt(options.x)},${fmt(options.y)}`
  );
  page.drawRect({
    x: options.x,
    y: options.y,
    width: scene.width,
    height: scene.height,
    fill: { r: 1, g: 1, b: 1 },
    stroke: { r: 0.8, g: 0.8, b: 0.8 }
  });
  if (scene.title) {
    trace?.push(`text:title:${scene.title.text}:${fmt(scene.title.x)},${fmt(scene.title.y)}`);
    drawPdfText(page, scene.title, { anchorOverride: "middle" });
  }
  for (const gridline of scene.gridlines) {
    trace?.push(
      `line:grid:${fmt(gridline.x1)},${fmt(gridline.y1)}-${fmt(gridline.x2)},${fmt(gridline.y2)}`
    );
    page.drawLine({ ...gridline, color: hexToPdfColor(gridline.color) });
  }
  trace?.push(
    `line:x:${fmt(scene.axes.x.x1)},${fmt(scene.axes.x.y1)}-${fmt(scene.axes.x.x2)},${fmt(scene.axes.x.y2)}`
  );
  page.drawLine({ ...scene.axes.x, color: hexToPdfColor(scene.axes.x.color) });
  trace?.push(
    `line:y:${fmt(scene.axes.y.x1)},${fmt(scene.axes.y.y1)}-${fmt(scene.axes.y.x2)},${fmt(scene.axes.y.y2)}`
  );
  page.drawLine({ ...scene.axes.y, color: hexToPdfColor(scene.axes.y.color) });
  if (scene.axes.x2) {
    trace?.push(
      `line:x2:${fmt(scene.axes.x2.x1)},${fmt(scene.axes.x2.y1)}-${fmt(scene.axes.x2.x2)},${fmt(scene.axes.x2.y2)}`
    );
    page.drawLine({ ...scene.axes.x2, color: hexToPdfColor(scene.axes.x2.color) });
  }
  if (scene.axes.y2) {
    trace?.push(
      `line:y2:${fmt(scene.axes.y2.x1)},${fmt(scene.axes.y2.y1)}-${fmt(scene.axes.y2.x2)},${fmt(scene.axes.y2.y2)}`
    );
    page.drawLine({ ...scene.axes.y2, color: hexToPdfColor(scene.axes.y2.color) });
  }
  for (const label of [
    ...scene.xLabels,
    ...scene.yLabels,
    ...scene.secondaryXLabels,
    ...scene.secondaryYLabels,
    ...scene.axisTitles
  ]) {
    trace?.push(`text:label:${label.text}:${fmt(label.x)},${fmt(label.y)}`);
    drawPdfText(page, label);
  }
  for (const s of scene.series) {
    trace?.push(`series:${s.type}`);
    drawPdfSeries(page, s);
    drawPdfAdornments(page, s, trace);
  }
  if (scene.dataTable) {
    drawPdfDataTable(page, scene.dataTable, trace);
  }
  trace?.push(
    `legend:${scene.legend.visible ? "visible" : "hidden"}:${scene.legend.items.map(item => item.label).join("|")}`
  );
  drawPdfLegend(page, scene.legend);
  return page;
}

/**
 * Draw a {@link ChartSceneText} through the PDF surface.
 *
 * Translates the scene's `anchor` / `rotate` / `color` fields into the
 * extended `ChartPdfDrawingSurface.drawText` parameters, and for surfaces
 * that ignore those new fields pre-shifts `x` using the same font
 * metrics the SVG path uses (`estimateTextWidth`). The result is that
 * legacy `drawText(text, { x, y, fontSize })` implementations continue
 * to render text at the correct horizontal position, while an upgraded
 * surface (such as `PdfPageBuilder`) can perform a proper measured
 * alignment with rotation.
 *
 * `anchorOverride` lets callers (e.g. the title) force a particular
 * anchor regardless of what the scene produced — useful where the
 * layout code already centres `x` around the midpoint but we still
 * want the text glyphs centred.
 */
function drawPdfText(
  page: ChartPdfDrawingSurface,
  text: ChartSceneText,
  extra: { anchorOverride?: "start" | "middle" | "end" } = {}
): void {
  const anchor = extra.anchorOverride ?? text.anchor ?? "start";
  // `estimateTextWidth` uses `@excel/utils/text-metrics`; for fallback
  // surfaces that ignore `anchor` we pre-shift x with a width estimate
  // that reflects the actual font whenever a family is declared (the
  // engine falls back to Arial metrics otherwise). Surfaces that honour
  // `anchor` re-resolve using their own `measureText` in `drawText`, so
  // supplying both is safe.
  const width = estimateTextWidth(text.text, text.fontSize, {
    bold: text.bold,
    italic: text.italic,
    fontName: text.fontFamily
  });
  const shiftedX = anchor === "start" ? text.x : text.x - width * (anchor === "middle" ? 0.5 : 1);
  page.drawText(text.text, {
    x: shiftedX,
    y: text.y,
    fontSize: text.fontSize,
    color: text.color ? hexToPdfColor(text.color) : undefined,
    rotation: text.rotate,
    anchor,
    fontFamily: text.fontFamily,
    bold: text.bold,
    italic: text.italic
  });
}

/**
 * Draw series-level adornments: leader lines, error bars, trendlines,
 * markers, data labels, and the bar3D depth hint. The SVG path handles
 * all of these via {@link renderSvgAdornments}; this is the matching
 * PDF implementation. Optional `trace` entries mirror the SVG tag names
 * so golden assertions can verify presence without binding to exact
 * coordinates.
 */
function drawPdfAdornments(
  page: ChartPdfDrawingSurface,
  series: ChartSceneSeries,
  trace?: string[]
): void {
  const color = "color" in series ? hexToPdfColor(series.color) : { r: 0.3, g: 0.3, b: 0.3 };
  // Leader lines (pie/doughnut external labels). Emit before the labels
  // themselves so the glyphs sit on top visually.
  for (const leader of series.leaderLines ?? []) {
    trace?.push(`leader:${fmt(leader.x1)},${fmt(leader.y1)}-${fmt(leader.x2)},${fmt(leader.y2)}`);
    page.drawLine({
      x1: leader.x1,
      y1: leader.y1,
      x2: leader.x2,
      y2: leader.y2,
      color: hexToPdfColor(leader.color),
      lineWidth: leader.width ?? 1
    });
  }
  // Error bars
  for (const eb of series.errorBars ?? []) {
    trace?.push(
      `errorbar:${fmt(eb.line.x1)},${fmt(eb.line.y1)}-${fmt(eb.line.x2)},${fmt(eb.line.y2)}`
    );
    page.drawLine({
      x1: eb.line.x1,
      y1: eb.line.y1,
      x2: eb.line.x2,
      y2: eb.line.y2,
      color: hexToPdfColor(eb.line.color),
      lineWidth: eb.line.width ?? 1
    });
    for (const cap of [eb.cap1, eb.cap2]) {
      if (cap) {
        page.drawLine({
          x1: cap.x1,
          y1: cap.y1,
          x2: cap.x2,
          y2: cap.y2,
          color: hexToPdfColor(cap.color),
          lineWidth: cap.width ?? 1
        });
      }
    }
  }
  // Trendlines — polyline over the precomputed points, plus an optional
  // label at the end. The SVG variant uses `stroke-dasharray="4 3"`
  // whenever `dash` is set; mirror that as a simple on/off dash pattern.
  for (const trend of series.trendlines ?? []) {
    trace?.push(`trendline:${trend.points.length}pts`);
    const dashPattern = trend.dash ? [4, 3] : undefined;
    for (let i = 1; i < trend.points.length; i++) {
      const p0 = trend.points[i - 1];
      const p1 = trend.points[i];
      page.drawLine({
        x1: p0.x,
        y1: p0.y,
        x2: p1.x,
        y2: p1.y,
        color: hexToPdfColor(trend.color),
        lineWidth: trend.width ?? 1.5,
        dashPattern
      });
    }
    if (trend.label) {
      drawPdfText(page, trend.label);
    }
  }
  // Markers — geometric symbols at each point. Mirrors renderSvgMarker.
  for (const marker of series.markers ?? []) {
    drawPdfMarker(page, marker);
  }
  // Data labels (post-collision-resolution, post-leader-line layout).
  for (const label of series.labels ?? []) {
    trace?.push(`label:${label.text}:${fmt(label.x)},${fmt(label.y)}`);
    drawPdfText(page, label);
  }
  // bar3D's decorative depth hint: two parallelograms projected along
  // (+depth, -depth) for every bar. The SVG path emits them inline on
  // each bar (`renderBarDepth`); the PDF path does the same thing here
  // using `drawPath` when available. Surfaces without `drawPath` fall
  // back to drawing the two parallelograms as pairs of lines, which
  // preserves the 3D illusion even without filled polygons.
  if (series.type === "bar" && series.depth && series.depth > 0) {
    for (const bar of series.bars) {
      if (series.projection3D) {
        drawPdfBar3DBox(page, bar, series.projection3D, color, series.horizontal);
      } else {
        drawPdfBarDepth(page, bar, series.depth, color);
      }
    }
  }
  // Radar's translucent fill is emitted by `drawPdfSeries` itself (the
  // filled radar branch `chart-renderer.ts:4139-4146`) using
  // `hexToPdfColorWithAlpha`, so no additional adornment pass is needed
  // here. Surfaces without drawPath degrade to stroke-only, which the
  // same branch also handles.
}

function drawPdfMarker(page: ChartPdfDrawingSurface, marker: ChartSceneMarker): void {
  const r = marker.size / 2;
  const fill = hexToPdfColor(marker.color);
  const symbol = marker.symbol ?? "circle";
  if (symbol === "square") {
    page.drawRect({
      x: marker.x - r,
      y: marker.y - r,
      width: marker.size,
      height: marker.size,
      fill
    });
    return;
  }
  if (symbol === "diamond" && page.drawPath) {
    page.drawPath(
      [
        { op: "move", x: marker.x, y: marker.y - r },
        { op: "line", x: marker.x + r, y: marker.y },
        { op: "line", x: marker.x, y: marker.y + r },
        { op: "line", x: marker.x - r, y: marker.y },
        { op: "close" }
      ],
      { fill }
    );
    return;
  }
  if (symbol === "triangle" && page.drawPath) {
    page.drawPath(
      [
        { op: "move", x: marker.x, y: marker.y - r },
        { op: "line", x: marker.x + r, y: marker.y + r },
        { op: "line", x: marker.x - r, y: marker.y + r },
        { op: "close" }
      ],
      { fill }
    );
    return;
  }
  if (symbol === "x") {
    page.drawLine({
      x1: marker.x - r,
      y1: marker.y - r,
      x2: marker.x + r,
      y2: marker.y + r,
      color: fill,
      lineWidth: 2
    });
    page.drawLine({
      x1: marker.x + r,
      y1: marker.y - r,
      x2: marker.x - r,
      y2: marker.y + r,
      color: fill,
      lineWidth: 2
    });
    return;
  }
  if (symbol === "plus") {
    page.drawLine({
      x1: marker.x - r,
      y1: marker.y,
      x2: marker.x + r,
      y2: marker.y,
      color: fill,
      lineWidth: 2
    });
    page.drawLine({
      x1: marker.x,
      y1: marker.y - r,
      x2: marker.x,
      y2: marker.y + r,
      color: fill,
      lineWidth: 2
    });
    return;
  }
  // Default / circle / dash / dot / star / picture / auto — fall back to
  // a filled circle so at least the point is visible.
  if (page.drawCircle) {
    page.drawCircle({ cx: marker.x, cy: marker.y, r, fill });
  } else {
    page.drawRect({
      x: marker.x - r,
      y: marker.y - r,
      width: marker.size,
      height: marker.size,
      fill
    });
  }
}

function drawPdfBarDepth(
  page: ChartPdfDrawingSurface,
  bar: ChartSceneRect,
  depth: number,
  baseColor: PdfColor
): void {
  // Match the SVG path's `withAlpha(series.color, 0.75)` top/right
  // parallelograms. `PdfColor.a` now flows through to `/ExtGState` when
  // the surface supports it, so the PDF and SVG look identical on
  // capable viewers; legacy surfaces ignore `a` and fall back to an
  // opaque parallelogram which is the pre-alpha rendering.
  const shadeColor: PdfColor = {
    r: baseColor.r,
    g: baseColor.g,
    b: baseColor.b,
    a: 0.75
  };
  const topOps: ChartPdfPathOp[] = [
    { op: "move", x: bar.x, y: bar.y },
    { op: "line", x: bar.x + depth, y: bar.y - depth },
    { op: "line", x: bar.x + bar.width + depth, y: bar.y - depth },
    { op: "line", x: bar.x + bar.width, y: bar.y },
    { op: "close" }
  ];
  const rightOps: ChartPdfPathOp[] = [
    { op: "move", x: bar.x + bar.width, y: bar.y },
    { op: "line", x: bar.x + bar.width + depth, y: bar.y - depth },
    { op: "line", x: bar.x + bar.width + depth, y: bar.y - depth + bar.height },
    { op: "line", x: bar.x + bar.width, y: bar.y + bar.height },
    { op: "close" }
  ];
  if (page.drawPath) {
    page.drawPath(topOps, { fill: shadeColor });
    page.drawPath(rightOps, { fill: shadeColor });
    return;
  }
  // Surfaces without drawPath fall back to outlining the parallelograms
  // with strokes so the 3D hint is still recognisable.
  for (const ops of [topOps, rightOps]) {
    for (let i = 1; i < ops.length; i++) {
      const a = ops[i - 1];
      const b = ops[i];
      if (a.op === "close" || b.op === "close") {
        continue;
      }
      // Only move/line ops have x/y; curve ops never appear in these
      // parallelogram paths but the type narrowing needs the guard.
      if ("x" in a && "y" in a && "x" in b && "y" in b) {
        page.drawLine({ x1: a.x, y1: a.y, x2: b.x, y2: b.y, color: shadeColor });
      }
    }
  }
}

/**
 * Render a bar3D column as a true extruded box via the PDF surface.
 * Mirrors {@link renderBar3DBox} for SVG — three visible faces with
 * light/dark shading. Surfaces without `drawPath` fall back to stroked
 * outlines so the 3D cue is preserved even without filled polygons.
 */
function drawPdfBar3DBox(
  page: ChartPdfDrawingSurface,
  bar: ChartSceneRect,
  proj: { dx: number; dy: number },
  baseColor: PdfColor,
  horizontal: boolean | undefined
): void {
  const dx = proj.dx;
  const dy = proj.dy;
  const right = bar.x + bar.width;
  const top = bar.y;
  const bottom = bar.y + bar.height;
  const topFill: PdfColor = { r: baseColor.r, g: baseColor.g, b: baseColor.b, a: 0.92 };
  const rightFill: PdfColor = { r: baseColor.r, g: baseColor.g, b: baseColor.b, a: 0.75 };
  const topFace: ChartPdfPathOp[] = [
    { op: "move", x: bar.x, y: top },
    { op: "line", x: bar.x + dx, y: top - dy },
    { op: "line", x: right + dx, y: top - dy },
    { op: "line", x: right, y: top },
    { op: "close" }
  ];
  const rightFace: ChartPdfPathOp[] = [
    { op: "move", x: right, y: top },
    { op: "line", x: right + dx, y: top - dy },
    { op: "line", x: right + dx, y: bottom - dy },
    { op: "line", x: right, y: bottom },
    { op: "close" }
  ];
  void horizontal; // same geometry regardless — kept for future tuning
  if (page.drawPath) {
    page.drawPath(topFace, { fill: topFill });
    page.drawPath(rightFace, { fill: rightFill });
    page.drawRect({ x: bar.x, y: top, width: bar.width, height: bar.height, fill: baseColor });
    return;
  }
  for (const ops of [topFace, rightFace]) {
    for (let i = 1; i < ops.length; i++) {
      const a = ops[i - 1];
      const b = ops[i];
      if (a.op === "close" || b.op === "close") {
        continue;
      }
      if ("x" in a && "y" in a && "x" in b && "y" in b) {
        page.drawLine({ x1: a.x, y1: a.y, x2: b.x, y2: b.y, color: rightFill });
      }
    }
  }
  page.drawRect({ x: bar.x, y: top, width: bar.width, height: bar.height, fill: baseColor });
}

function getPlotRect(
  width: number,
  height: number,
  hasTitle: boolean,
  legend: ChartSceneLegend,
  model: ChartModel,
  dataTableHeight = 0
): ChartSceneRect {
  const axes = model.chart.plotArea.axes;
  const leftAxis = axes.find(axis => !axis.delete && axis.axPos === "l");
  const rightAxis = axes.find(axis => !axis.delete && axis.axPos === "r");
  const topAxis = axes.find(axis => !axis.delete && axis.axPos === "t");
  const bottomAxis = axes.find(axis => !axis.delete && axis.axPos === "b");
  // Legend padding is derived from the real scene rectangle so long series
  // names push the plot rectangle inwards instead of being clipped by it.
  // `legend.rect.width` was sized by `legendRect` from actual label widths.
  const leftLegendPad = legend.visible && legend.position === "l" ? legend.rect.width + 12 : 0;
  const rightLegendPad =
    legend.visible && (legend.position === "r" || legend.position === "tr")
      ? legend.rect.width + 16
      : 0;
  const left = 58 + (leftAxis?.title ? 18 : 0) + leftLegendPad;
  const right = 24 + (rightAxis ? 42 : 0) + (rightAxis?.title ? 18 : 0) + rightLegendPad;
  const top =
    (hasTitle ? 52 : 24) +
    (topAxis ? 22 : 0) +
    (topAxis?.title ? 16 : 0) +
    (legend.visible && legend.position === "t" ? 30 : 0);
  // When a data table is drawn below the plot, the legend placed at `b`
  // still needs room underneath it. The axis's x-labels get hidden
  // (handled by `buildChartScene`) so we drop their 22 px contribution
  // when the data table replaces them.
  const bottomAxisLabelSpace = dataTableHeight > 0 ? 0 : 0; // kept explicit for readability
  const bottom =
    46 +
    bottomAxisLabelSpace +
    (bottomAxis?.title ? 18 : 0) +
    (legend.visible && legend.position === "b" ? 28 : 0) +
    dataTableHeight;
  const auto: ChartSceneRect = {
    x: left,
    y: top,
    width: Math.max(10, width - left - right),
    height: Math.max(10, height - top - bottom)
  };
  const manual = model.chart.plotArea.layout?.manualLayout;
  if (!manual) {
    return auto;
  }
  return applyManualPlotLayout(auto, manual, width, height);
}

/**
 * Apply a plotArea `c:layout/c:manualLayout` override to the auto-computed
 * plot rectangle. Only `edge`-mode positioning/sizing is honoured; `factor`
 * mode is documented by the OOXML spec as a ratio relative to the default
 * Excel layout, which this deterministic preview does not compute because
 * the default itself is approximate (`auto`). Undefined axes fall back to
 * the auto value so partial overrides (e.g. custom width only) work.
 * `layoutTarget` is treated as `inner` in all cases — the preview renderer
 * does not yet draw axis labels outside the plot rectangle, so `outer`
 * behaves identically.
 */
function applyManualPlotLayout(
  auto: ChartSceneRect,
  manual: NonNullable<NonNullable<ChartModel["chart"]["plotArea"]["layout"]>["manualLayout"]>,
  chartWidth: number,
  chartHeight: number
): ChartSceneRect {
  const resolveX = (v: number | undefined, mode: "edge" | "factor" | undefined): number =>
    v !== undefined && mode !== "factor" ? clamp01(v) * chartWidth : auto.x;
  const resolveY = (v: number | undefined, mode: "edge" | "factor" | undefined): number =>
    v !== undefined && mode !== "factor" ? clamp01(v) * chartHeight : auto.y;
  const resolveW = (v: number | undefined, mode: "edge" | "factor" | undefined): number =>
    v !== undefined && mode !== "factor" ? clamp01(v) * chartWidth : auto.width;
  const resolveH = (v: number | undefined, mode: "edge" | "factor" | undefined): number =>
    v !== undefined && mode !== "factor" ? clamp01(v) * chartHeight : auto.height;

  const x = resolveX(manual.x, manual.xMode);
  const y = resolveY(manual.y, manual.yMode);
  const w = Math.max(10, resolveW(manual.w, manual.wMode));
  const h = Math.max(10, resolveH(manual.h, manual.hMode));
  // Clamp against the chart canvas so pathological values don't draw off-screen.
  const clampedX = Math.max(0, Math.min(x, chartWidth - 10));
  const clampedY = Math.max(0, Math.min(y, chartHeight - 10));
  const clampedW = Math.min(w, chartWidth - clampedX);
  const clampedH = Math.min(h, chartHeight - clampedY);
  return { x: clampedX, y: clampedY, width: clampedW, height: clampedH };
}

function clamp01(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  if (value < 0) {
    return 0;
  }
  if (value > 1) {
    return 1;
  }
  return value;
}

// ============================================================================
// Data table (c:dTable) layout
// ============================================================================

const DATA_TABLE_ROW_HEIGHT_DEFAULT = 18;
const DATA_TABLE_FONT_SIZE_DEFAULT = 10;
const DATA_TABLE_SWATCH_SIZE = 10;
const DATA_TABLE_NAME_COL_PADDING = 16;
const DATA_TABLE_CELL_PADDING_X = 6;
// Gap between plot baseline and the top of the data table.
const DATA_TABLE_TOP_GAP = 8;

/**
 * Vertical footprint of the data-table overlay, used by
 * {@link getPlotRect} to reserve bottom padding before series geometry
 * is computed. Counts the header row (categories) plus one row per
 * normalised series, each at {@link DATA_TABLE_ROW_HEIGHT_DEFAULT}
 * unless the author's `txPr` explicitly raises the font size.
 */
function computeDataTableHeight(
  table: DataTable,
  normalized: NormalizedSeries[],
  categories: string[]
): number {
  if (normalized.length === 0 || categories.length === 0) {
    return 0;
  }
  const fontSize = resolveDataTableFontSize(table);
  const rowHeight = Math.max(DATA_TABLE_ROW_HEIGHT_DEFAULT, Math.round(fontSize * 1.6));
  // Header (categories) + one row per series.
  const rowCount = 1 + normalized.length;
  return rowCount * rowHeight + DATA_TABLE_TOP_GAP;
}

function resolveDataTableFontSize(table: DataTable): number {
  const explicit = getTxPrFontSizePx(table.txPr);
  return explicit ?? DATA_TABLE_FONT_SIZE_DEFAULT;
}

function getTxPrFontSizePx(txPr: ChartTextProperties | undefined): number | undefined {
  if (!txPr) {
    return undefined;
  }
  // `getTxPrFontSize` returns points; convert to CSS pixels using the
  // 1 pt = 1.333 px rule so the result matches the rest of the
  // renderer's text sizing.
  const points = getTxPrFontSize(txPr);
  if (typeof points !== "number") {
    return undefined;
  }
  return points * (4 / 3);
}

/**
 * Build the scene-level layout for `c:dTable`. The table hugs the
 * plot's x-range so columns line up with category ticks even when the
 * legend or axis titles push the plot inward. The first column is a
 * series-name column sized to the widest label (+ swatch when
 * `showKeys` is on). Remaining columns are equal-width category cells.
 */
function buildSceneDataTable(
  table: DataTable,
  normalized: NormalizedSeries[],
  categories: string[],
  plot: ChartSceneRect,
  height: number
): ChartSceneDataTable | undefined {
  if (normalized.length === 0 || categories.length === 0 || height <= 0) {
    return undefined;
  }
  const fontSize = resolveDataTableFontSize(table);
  const rowHeight = (height - DATA_TABLE_TOP_GAP) / (1 + normalized.length);
  const color = colorFromChartTextProperties(table.txPr) ?? "#444444";
  const textStyle = textStyleFromTxPr(table.txPr);
  const showKeys = table.showKeys !== false;
  // Series-name column width — widest label + swatch + padding.
  const swatchWidth = showKeys ? DATA_TABLE_SWATCH_SIZE + 6 : 0;
  const maxNameWidth = normalized.reduce(
    (acc, s) =>
      Math.max(
        acc,
        estimateTextWidth(s.label || "", fontSize, {
          bold: textStyle.bold,
          italic: textStyle.italic,
          fontName: textStyle.fontFamily
        })
      ),
    0
  );
  const nameColWidth = Math.ceil(maxNameWidth + swatchWidth + DATA_TABLE_NAME_COL_PADDING);
  // Table's left edge sits just below the y-axis, so the name column
  // lives in the left margin. When the name column doesn't fit there
  // we still draw everything inside the chart width — overlapping a
  // little with the y-axis is preferable to clipping.
  const tableLeft = Math.max(4, plot.x - nameColWidth);
  const tableRight = plot.x + plot.width;
  const tableTop = plot.y + plot.height + DATA_TABLE_TOP_GAP;
  const rect: ChartSceneRect = {
    x: tableLeft,
    y: tableTop,
    width: tableRight - tableLeft,
    height: height - DATA_TABLE_TOP_GAP
  };
  // Column boundaries — name column + one per category evenly dividing
  // the plot's x range so ticks line up.
  const columns: number[] = [tableLeft, plot.x];
  const catWidth = plot.width / categories.length;
  for (let i = 1; i <= categories.length; i++) {
    columns.push(plot.x + catWidth * i);
  }
  // Row boundaries — header + one per series.
  const rows: number[] = [tableTop];
  for (let r = 1; r <= 1 + normalized.length; r++) {
    rows.push(tableTop + rowHeight * r);
  }

  const cells: ChartSceneText[] = [];
  const legendSwatches: Array<ChartSceneRect & { color: string }> = [];

  // Header row — category names, centred in each category column.
  for (let c = 0; c < categories.length; c++) {
    const cellX = columns[c + 1] + catWidth / 2;
    const cellY = rows[0] + rowHeight / 2 + fontSize / 3;
    cells.push({
      x: cellX,
      y: cellY,
      text: categories[c] ?? "",
      fontSize,
      color,
      anchor: "middle",
      ...textStyle
    });
  }
  // Series rows — series name (left-aligned, optionally preceded by a
  // swatch) + one value per category.
  for (let s = 0; s < normalized.length; s++) {
    const series = normalized[s];
    const rowY = rows[s + 1] + rowHeight / 2 + fontSize / 3;
    if (showKeys) {
      const swatchX = tableLeft + DATA_TABLE_CELL_PADDING_X;
      const swatchY = rows[s + 1] + (rowHeight - DATA_TABLE_SWATCH_SIZE) / 2;
      legendSwatches.push({
        x: swatchX,
        y: swatchY,
        width: DATA_TABLE_SWATCH_SIZE,
        height: DATA_TABLE_SWATCH_SIZE,
        color: series.color
      });
    }
    const nameX = tableLeft + DATA_TABLE_CELL_PADDING_X + swatchWidth;
    cells.push({
      x: nameX,
      y: rowY,
      text: series.label || "",
      fontSize,
      color,
      anchor: "start",
      ...textStyle
    });
    for (let c = 0; c < categories.length; c++) {
      const v = series.values[c];
      const cellX = columns[c + 1] + catWidth / 2;
      cells.push({
        x: cellX,
        y: rowY,
        text: formatDataTableValue(v),
        fontSize,
        color,
        anchor: "middle",
        ...textStyle
      });
    }
  }

  // Borders — Excel's three toggles map to three primitive strokes:
  // * showOutline      → outer rectangle
  // * showHorzBorder   → horizontal lines between rows
  // * showVertBorder   → vertical lines between columns
  // Any `spPr` on the dTable element is used for the stroke colour so
  // themed borders carry through from the source file.
  const borders: ChartSceneLine[] = [];
  const borderColor = getSpPrLineColor(table.spPr) ?? "#888888";
  const borderWidth = 0.75;
  const pushLine = (x1: number, y1: number, x2: number, y2: number): void => {
    borders.push({ x1, y1, x2, y2, color: borderColor, width: borderWidth });
  };
  if (table.showOutline !== false) {
    pushLine(rect.x, rect.y, rect.x + rect.width, rect.y);
    pushLine(rect.x, rect.y + rect.height, rect.x + rect.width, rect.y + rect.height);
    pushLine(rect.x, rect.y, rect.x, rect.y + rect.height);
    pushLine(rect.x + rect.width, rect.y, rect.x + rect.width, rect.y + rect.height);
  }
  if (table.showHorzBorder !== false) {
    for (let r = 1; r < rows.length - 1; r++) {
      pushLine(rect.x, rows[r], rect.x + rect.width, rows[r]);
    }
  }
  if (table.showVertBorder !== false) {
    // columns[0] and last are part of the outline.
    for (let c = 1; c < columns.length - 1; c++) {
      pushLine(columns[c], rect.y, columns[c], rect.y + rect.height);
    }
  }

  return { rect, columns, rows, cells, legendSwatches, borders };
}

function formatDataTableValue(value: number | undefined): string {
  if (value === undefined || !Number.isFinite(value)) {
    return "";
  }
  // Match Excel's default "General" formatting: no trailing zeroes,
  // up to 6 significant digits.
  if (Number.isInteger(value)) {
    return value.toString();
  }
  return Number.parseFloat(value.toPrecision(6)).toString();
}

/**
 * Unit-depth cabinet projection used by bar3D. `view3D.rotX` tilts the
 * view forward (positive → viewer sees the top of the column), and
 * `view3D.rotY` spins it to the right (positive → viewer sees the right
 * side). The two Excel defaults — `rotX=15`, `rotY=20`,
 * `rAngAx=true` (right-angle axes) — stay in sync with this projection
 * when `view3D` is absent. Returns unit deltas (dx/dy per pixel of
 * depth) so callers can scale by their own bar-width heuristic.
 */
function resolveBar3DProjection(view3D: ChartModel["chart"]["view3D"] | undefined): {
  dx: number;
  dy: number;
} {
  const rotX = toRad(view3D?.rotX ?? 15);
  const rotY = toRad(view3D?.rotY ?? 20);
  // Cabinet-style: depth vector in screen space. Multiply by sign so
  // a positive rotY pushes the back of the bar to the right and a
  // positive rotX pushes the back upward. `rAngAx` (right-angle axes)
  // skips Y rotation entirely; when it's explicitly off we keep the
  // 3D spin.
  const rAngAx = view3D?.rAngAx !== false;
  const xFactor = rAngAx ? 0 : Math.cos(rotY);
  return {
    dx: 0.6 * (rAngAx ? Math.cos(rotY) : xFactor),
    dy: 0.6 * Math.sin(rotX)
  };
}

function toRad(deg: number): number {
  return (deg * Math.PI) / 180;
}

function getSpPrLineColor(spPr: DataTable["spPr"]): string | undefined {
  if (!spPr) {
    return undefined;
  }
  // Import is kept inline so we can reuse the structured accessor without
  // perturbing the top-level import block. `getSpPrLine` handles both
  // structured ShapeProperties and `_rawXml` passthroughs.
  const line = getSpPrLine(spPr);
  const srgb = line?.color?.srgb;
  return typeof srgb === "string" ? `#${srgb.replace(/^#/, "")}` : undefined;
}

function buildSceneSeries(
  groups: ChartTypeGroup[],
  normalized: NormalizedSeries[],
  plot: ChartSceneRect,
  axisContext: ChartAxisContext,
  categories: string[],
  view3D?: ChartModel["chart"]["view3D"]
): ChartSceneSeries[] {
  const result: ChartSceneSeries[] = [];
  for (const group of groups) {
    const groupSeries = normalized.filter(s => s.group === group);
    if (groupSeries.length === 0) {
      continue;
    }
    if (group.type === "pie" || group.type === "pie3D" || group.type === "doughnut") {
      const first = groupSeries[0];
      const yRange = getSeriesYRange(first, axisContext);
      result.push(
        withAdornments(
          buildPieSeries(
            group.type === "doughnut" ? "doughnut" : "pie",
            first.values,
            plot,
            group.type === "doughnut",
            first
          ),
          first,
          plot,
          yRange.min,
          yRange.max,
          categories
        )
      );
      continue;
    }
    if (group.type === "ofPie") {
      const first = groupSeries[0];
      const yRange = getSeriesYRange(first, axisContext);
      result.push(
        withAdornments(
          buildOfPieSeries(first.values, plot, group.secondPieSize),
          first,
          plot,
          yRange.min,
          yRange.max,
          categories
        )
      );
      continue;
    }
    if (group.type === "radar") {
      for (const s of groupSeries) {
        const yRange = getSeriesYRange(s, axisContext);
        result.push(
          withAdornments(
            buildRadarSeries(s, plot, yRange.min, yRange.max, group.radarStyle === "filled"),
            s,
            plot,
            yRange.min,
            yRange.max,
            categories
          )
        );
      }
      continue;
    }
    if (group.type === "stock") {
      const yRange = getGroupYRange(group, axisContext);
      result.push(buildStockSeries(groupSeries, plot, yRange.min, yRange.max, categories));
      continue;
    }
    if (group.type === "surface" || group.type === "surface3D") {
      const yRange = getGroupYRange(group, axisContext);
      result.push(buildSurfaceSeries(groupSeries, plot, yRange.min, yRange.max, group.wireframe));
      continue;
    }
    if (group.type === "area" || group.type === "area3D") {
      const groupRange = getGroupYRange(group, axisContext);
      const baselineY = valueToY(0, groupRange.min, groupRange.max, plot);
      const stacked = group.grouping === "stacked" || group.grouping === "percentStacked";
      const percent = group.grouping === "percentStacked";
      for (const s of groupSeries) {
        const yRange = percent ? { min: 0, max: 1 } : getSeriesYRange(s, axisContext);
        const band = stacked
          ? buildStackedAreaBand(groupSeries, s.seriesIndex, plot, yRange.min, yRange.max, percent)
          : { upper: buildLinePoints(s.values, plot, yRange.min, yRange.max), lower: undefined };
        result.push(
          withAdornments(
            {
              type: "area",
              color: s.color,
              label: s.label,
              points: band.upper,
              lowerPoints: band.lower,
              baselineY,
              closed: group.type === "area3D"
            },
            s,
            plot,
            yRange.min,
            yRange.max,
            categories
          )
        );
      }
      continue;
    }
    if (group.type === "bubble") {
      for (const s of groupSeries) {
        const yRange = getSeriesYRange(s, axisContext);
        const xRange = getSeriesXRange(s, axisContext);
        result.push(
          withAdornments(
            {
              type: "bubble",
              color: s.color,
              label: s.label,
              bubbles: buildBubbles(s, plot, yRange.min, yRange.max, xRange.min, xRange.max)
            },
            s,
            plot,
            yRange.min,
            yRange.max,
            categories
          )
        );
      }
      continue;
    }
    if (group.type === "line" || group.type === "line3D" || group.type === "scatter") {
      const sceneType = group.type === "scatter" ? "scatter" : "line";
      for (const s of groupSeries) {
        const yRange = getSeriesYRange(s, axisContext);
        const xRange = getSeriesXRange(s, axisContext);
        const points =
          group.type === "scatter"
            ? buildScatterPoints(s, plot, yRange.min, yRange.max, xRange.min, xRange.max)
            : buildLinePoints(s.values, plot, yRange.min, yRange.max);
        result.push(
          withAdornments(
            {
              type: sceneType,
              color: s.color,
              label: s.label,
              points,
              smooth: "smooth" in group ? group.smooth || s.marker?.symbol === "auto" : false,
              showLine: group.type !== "scatter" || group.scatterStyle !== "marker"
            },
            s,
            plot,
            yRange.min,
            yRange.max,
            categories
          )
        );
      }
      continue;
    }
    if (group.type === "bar" || group.type === "bar3D") {
      const stacked = group.grouping === "stacked" || group.grouping === "percentStacked";
      const percent = group.grouping === "percentStacked";
      // Cabinet-ish axonometric projection for bar3D. The OOXML default
      // view is rotX=15°, rotY=20°, depthPercent=100; when authors leave
      // `view3D` unset we fall back to those values so the preview still
      // reads as 3D. The scalar `0.35` shrinks the depth to ~one-third
      // the bar width so a column of typical bars doesn't dominate the
      // plot area — Excel itself biases the default projection this way.
      const proj = group.type === "bar3D" ? resolveBar3DProjection(view3D) : undefined;
      for (const s of groupSeries) {
        const horizontal = group.barDir === "bar";
        const yRange = percent ? { min: 0, max: 1 } : getSeriesYRange(s, axisContext);
        const bars = stacked
          ? buildStackedBars(
              groupSeries,
              s.seriesIndex,
              horizontal,
              plot,
              yRange.min,
              yRange.max,
              percent
            )
          : horizontal
            ? buildHorizontalBars(
                s.values,
                s.seriesIndex,
                groupSeries.length,
                categories,
                plot,
                yRange.min,
                yRange.max
              )
            : buildBars(
                s.values,
                s.seriesIndex,
                groupSeries.length,
                categories,
                plot,
                yRange.min,
                yRange.max
              );
        // Scale bar3D depth to roughly 30% of the narrowest bar width
        // (or 20% of the plot width for stacked bars) so the 3D effect
        // is visible but doesn't crush the series. A hard floor of 6px
        // keeps tiny bars from losing their depth cue entirely.
        const bar3DDepth = proj
          ? Math.max(
              6,
              Math.min(
                bars.reduce(
                  (m, r) => Math.min(m, horizontal ? r.height : r.width),
                  Number.POSITIVE_INFINITY
                ) * 0.3,
                plot.width * 0.12
              )
            )
          : 0;
        result.push(
          withAdornments(
            {
              type: "bar",
              color: s.color,
              label: s.label,
              horizontal,
              depth: group.type === "bar3D" ? bar3DDepth : 0,
              projection3D: proj
                ? {
                    dx: proj.dx * bar3DDepth,
                    dy: proj.dy * bar3DDepth
                  }
                : undefined,
              bars
            },
            s,
            plot,
            yRange.min,
            yRange.max,
            categories
          )
        );
      }
    }
  }
  return result;
}

function buildBars(
  values: number[],
  seriesIndex: number,
  seriesCount: number,
  categories: string[],
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartSceneRect[] {
  const count = Math.max(1, categories.length, values.length);
  const groupWidth = plot.width / count;
  const barWidth = (groupWidth * 0.72) / Math.max(1, seriesCount);
  const zero = valueToY(0, min, max, plot);
  return values.map((value, i) => {
    const y = valueToY(value, min, max, plot);
    return {
      x: plot.x + i * groupWidth + groupWidth * 0.14 + seriesIndex * barWidth,
      y: Math.min(y, zero),
      width: barWidth,
      height: Math.abs(zero - y)
    };
  });
}

function buildHorizontalBars(
  values: number[],
  seriesIndex: number,
  seriesCount: number,
  categories: string[],
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartSceneRect[] {
  const count = Math.max(1, categories.length, values.length);
  const groupHeight = plot.height / count;
  const barHeight = (groupHeight * 0.72) / Math.max(1, seriesCount);
  const zero = valueToX(0, min, max, plot);
  return values.map((value, i) => {
    const x = valueToX(value, min, max, plot);
    return {
      x: Math.min(x, zero),
      y: plot.y + i * groupHeight + groupHeight * 0.14 + seriesIndex * barHeight,
      width: Math.abs(x - zero),
      height: barHeight
    };
  });
}

function buildStackedBars(
  groupSeries: NormalizedSeries[],
  seriesIndex: number,
  horizontal: boolean,
  plot: ChartSceneRect,
  min: number,
  max: number,
  percent: boolean
): ChartSceneRect[] {
  const count = Math.max(1, ...groupSeries.map(s => s.values.length));
  const slot = horizontal ? plot.height / count : plot.width / count;
  const thickness = slot * 0.72;
  const zeroY = valueToY(0, min, max, plot);
  const zeroX = valueToX(0, min, max, plot);
  const totals = stackedTotals(groupSeries, count, percent);
  return Array.from({ length: count }, (_, i) => {
    const start = stackedValueAt(groupSeries, seriesIndex, i, totals, false);
    const end = stackedValueAt(groupSeries, seriesIndex, i, totals, true);
    if (horizontal) {
      const x1 = valueToX(start, min, max, plot);
      const x2 = valueToX(end, min, max, plot);
      return {
        x: Math.min(x1, x2, zeroX),
        y: plot.y + i * slot + slot * 0.14,
        width: Math.abs(x2 - x1),
        height: thickness
      };
    }
    const y1 = valueToY(start, min, max, plot);
    const y2 = valueToY(end, min, max, plot);
    return {
      x: plot.x + i * slot + slot * 0.14,
      y: Math.min(y1, y2, zeroY),
      width: thickness,
      height: Math.abs(y2 - y1)
    };
  });
}

function buildStackedAreaBand(
  groupSeries: NormalizedSeries[],
  seriesIndex: number,
  plot: ChartSceneRect,
  min: number,
  max: number,
  percent: boolean
): { upper: ChartScenePoint[]; lower: ChartScenePoint[] } {
  const count = Math.max(1, ...groupSeries.map(s => s.values.length));
  const totals = stackedTotals(groupSeries, count, percent);
  const lowerValues = Array.from({ length: count }, (_, i) =>
    stackedValueAt(groupSeries, seriesIndex, i, totals, false)
  );
  const upperValues = Array.from({ length: count }, (_, i) =>
    stackedValueAt(groupSeries, seriesIndex, i, totals, true)
  );
  return {
    lower: buildLinePoints(lowerValues, plot, min, max),
    upper: buildLinePoints(upperValues, plot, min, max)
  };
}

function stackedTotals(groupSeries: NormalizedSeries[], count: number, percent: boolean): number[] {
  return Array.from({ length: count }, (_, i) => {
    const sum = groupSeries.reduce((total, s) => total + Math.max(0, s.values[i] ?? 0), 0);
    return percent ? sum || 1 : 1;
  });
}

function stackedValueAt(
  groupSeries: NormalizedSeries[],
  seriesIndex: number,
  pointIndex: number,
  totals: number[],
  includeCurrent: boolean
): number {
  const end = includeCurrent ? seriesIndex : seriesIndex - 1;
  let sum = 0;
  for (let i = 0; i <= end; i++) {
    sum += Math.max(0, groupSeries[i]?.values[pointIndex] ?? 0);
  }
  return sum / totals[pointIndex];
}

function buildScatterPoints(
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number,
  xMin?: number,
  xMax?: number
): ChartScenePoint[] {
  const xValues =
    series.xValues && series.xValues.length > 0
      ? series.xValues
      : series.values.map((_, i) => i + 1);
  const computedXMin = Math.min(0, ...xValues.filter(Number.isFinite));
  const computedXMax = Math.max(1, ...xValues.filter(Number.isFinite));
  const effectiveXMin = xMin ?? computedXMin;
  const effectiveXMax = xMax ?? computedXMax;
  return series.values.map((value, i) => ({
    x: valueToX(
      xValues[i] ?? i + 1,
      effectiveXMin,
      effectiveXMax <= effectiveXMin ? effectiveXMin + 1 : effectiveXMax,
      plot
    ),
    y: valueToY(value, min, max, plot)
  }));
}

function buildLinePoints(
  values: number[],
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartScenePoint[] {
  const step = values.length > 1 ? plot.width / (values.length - 1) : plot.width;
  return values.map((value, i) => ({ x: plot.x + i * step, y: valueToY(value, min, max, plot) }));
}

function buildBubbles(
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number,
  xMin?: number,
  xMax?: number
): ChartSceneBubble[] {
  const xValues =
    series.xValues && series.xValues.length > 0
      ? series.xValues
      : series.values.map((_, i) => i + 1);
  const yValues = series.values;
  const sizes = series.bubbleSizes ?? [];
  const sizeMax = Math.max(1, ...sizes.map(v => Math.abs(v)));
  const count = Math.max(xValues.length, yValues.length);
  const computedXMin = Math.min(0, ...xValues.filter(Number.isFinite));
  const computedXMax = Math.max(1, ...xValues.filter(Number.isFinite));
  const effectiveXMin = xMin ?? computedXMin;
  const effectiveXMax = xMax ?? computedXMax;
  return Array.from({ length: count }, (_, i) => ({
    x: valueToX(
      xValues[i] ?? i + 1,
      effectiveXMin,
      effectiveXMax <= effectiveXMin ? effectiveXMin + 1 : effectiveXMax,
      plot
    ),
    y: valueToY(yValues[i] ?? 0, min, max, plot),
    radius: 4 + (Math.sqrt(Math.abs(sizes[i] ?? 1)) / Math.sqrt(sizeMax)) * 16
  }));
}

function buildPieSeries(
  type: "pie" | "doughnut",
  values: number[],
  plot: ChartSceneRect,
  doughnut: boolean,
  series?: NormalizedSeries
): ChartScenePieSeries {
  const radius = Math.min(plot.width, plot.height) / 2.35;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const total = values.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
  let angle = -Math.PI / 2;
  const slices = values.map((value, i) => {
    const next = angle + (Math.max(0, value) / total) * Math.PI * 2;
    const slice = {
      color: COLORS[i % COLORS.length],
      cx,
      cy,
      radius,
      innerRadius: doughnut ? radius * 0.45 : 0,
      startAngle: angle,
      endAngle: next
    };
    angle = next;
    return slice;
  });
  return { type, slices, label: series?.label };
}

function buildOfPieSeries(
  values: number[],
  plot: ChartSceneRect,
  secondPieSize = 75
): ChartScenePieSeries {
  const split = Math.max(1, Math.floor(values.length * 0.7));
  const primaryValues = values.slice(0, split);
  const secondaryValues = values.slice(split);
  const primary = buildPieSeries(
    "pie",
    primaryValues,
    { ...plot, width: plot.width * 0.62 },
    false
  );
  const radius =
    (Math.min(plot.width, plot.height) / 2.35) * Math.max(0.25, Math.min(2, secondPieSize / 100));
  const cx = plot.x + plot.width * 0.78;
  const cy = plot.y + plot.height / 2;
  let angle = -Math.PI / 2;
  const total = secondaryValues.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
  const secondarySlices = secondaryValues.map((value, i) => {
    const next = angle + (Math.max(0, value) / total) * Math.PI * 2;
    const slice = {
      color: COLORS[(split + i) % COLORS.length],
      cx,
      cy,
      radius,
      innerRadius: 0,
      startAngle: angle,
      endAngle: next
    };
    angle = next;
    return slice;
  });
  const connectors: ChartSceneLine[] = [
    {
      x1: primary.slices[0]?.cx ?? plot.x,
      y1: cy - radius,
      x2: cx - radius,
      y2: cy - radius,
      color: GRID_COLOR
    },
    {
      x1: primary.slices[0]?.cx ?? plot.x,
      y1: cy + radius,
      x2: cx - radius,
      y2: cy + radius,
      color: GRID_COLOR
    }
  ];
  return { type: "ofPie", slices: primary.slices, secondarySlices, connectors };
}

function buildRadarSeries(
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number,
  filled: boolean
): ChartSceneRadarSeries {
  const center = { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2 };
  const radius = Math.min(plot.width, plot.height) / 2.25;
  const count = Math.max(3, series.values.length);
  const points = series.values.map((value, i) => {
    const angle = -Math.PI / 2 + (i / count) * Math.PI * 2;
    const r = ((value - min) / (max - min)) * radius;
    return { x: center.x + Math.cos(angle) * r, y: center.y + Math.sin(angle) * r };
  });
  return {
    type: "radar",
    color: series.color,
    points,
    center,
    radius,
    filled,
    label: series.label
  };
}

function buildStockSeries(
  groupSeries: NormalizedSeries[],
  plot: ChartSceneRect,
  min: number,
  max: number,
  categories: string[]
): ChartSceneStockSeries {
  const count = Math.max(1, categories.length, ...groupSeries.map(s => s.values.length));
  const groupWidth = plot.width / count;
  const candleWidth = Math.max(3, groupWidth * 0.45);
  const firstValue = (seriesIndex: number, pointIndex: number): number | undefined =>
    groupSeries[seriesIndex]?.values[pointIndex];
  const useVolume = groupSeries.length >= 5;
  const offset = useVolume ? 1 : 0;
  const hasOpen = groupSeries.length - offset >= 4;
  const candles: ChartSceneStockCandle[] = [];
  for (let i = 0; i < count; i++) {
    const open = hasOpen ? firstValue(offset, i) : undefined;
    const high = firstValue(offset + (hasOpen ? 1 : 0), i) ?? open ?? 0;
    const low = firstValue(offset + (hasOpen ? 2 : 1), i) ?? high;
    const close = firstValue(offset + (hasOpen ? 3 : 2), i) ?? low;
    candles.push({
      x: plot.x + i * groupWidth + groupWidth / 2,
      highY: valueToY(Math.max(high, low, open ?? high, close), min, max, plot),
      lowY: valueToY(Math.min(high, low, open ?? low, close), min, max, plot),
      openY: open === undefined ? undefined : valueToY(open, min, max, plot),
      closeY: valueToY(close, min, max, plot),
      width: candleWidth,
      up: open === undefined || close >= open
    });
  }
  return { type: "stock", color: COLORS[0], candles, label: groupSeries[0]?.label };
}

function buildSurfaceSeries(
  groupSeries: NormalizedSeries[],
  plot: ChartSceneRect,
  min: number,
  max: number,
  wireframe?: boolean
): ChartSceneSurfaceSeries {
  const rows = Math.max(1, groupSeries.length);
  const cols = Math.max(1, ...groupSeries.map(s => s.values.length));
  const cellWidth = plot.width / cols;
  const cellHeight = plot.height / rows;
  const cells: ChartSceneSurfaceCell[] = [];
  for (let r = 0; r < rows; r++) {
    const values = groupSeries[r]?.values ?? [];
    for (let c = 0; c < cols; c++) {
      const value = values[c] ?? min;
      const t = max <= min ? 0 : (value - min) / (max - min);
      cells.push({
        x: plot.x + c * cellWidth,
        y: plot.y + r * cellHeight,
        width: cellWidth,
        height: cellHeight,
        color: interpolateColor("#5B9BD5", "#ED7D31", Math.max(0, Math.min(1, t)))
      });
    }
  }
  return { type: "surface", cells, wireframe, label: groupSeries[0]?.label };
}

function buildGridlines(plot: ChartSceneRect, axis?: ChartAxis): ChartSceneLine[] {
  if (axis?.majorGridlines === undefined) {
    return [];
  }
  const color = colorFromShapeLine(axis.majorGridlines.line) ?? GRID_COLOR;
  const lines: ChartSceneLine[] = [];
  for (let i = 1; i < 5; i++) {
    const y = plot.y + (plot.height * i) / 5;
    lines.push({ x1: plot.x, y1: y, x2: plot.x + plot.width, y2: y, color });
  }
  return lines;
}

function withAdornments<T extends ChartSceneSeries>(
  sceneSeries: T,
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number,
  categories: string[]
): T {
  const points = representativePoints(sceneSeries);
  const values = series.values;
  let labels = buildDataLabels(points, values, categories, series, plot);
  const markers = buildMarkers(points, series);
  const trendlines = buildTrendlines(points, values, series, plot, min, max);
  let leaderLines: ChartSceneLine[] | undefined;
  if (
    labels.length > 0 &&
    (sceneSeries.type === "pie" || sceneSeries.type === "doughnut" || sceneSeries.type === "ofPie")
  ) {
    const mergedLabels = mergeDataLabels(series.group, series.series);
    const position = mergedLabels?.position ?? "outEnd";
    if (position === "outEnd" || position === "bestFit") {
      const layout = layoutPieLabels(sceneSeries as ChartScenePieSeries, labels, plot);
      labels = layout.labels;
      leaderLines = layout.leaderLines;
    }
  } else if (labels.length > 1) {
    // For non-pie series, apply a generic greedy collision pass that
    // nudges overlapping labels vertically and hides labels that cannot
    // be separated even after nudging. Without this, dense bar/line
    // charts emit overlapping <text> glyphs at the same coordinates.
    labels = resolveLabelCollisions(labels, plot);
  }
  const errorBars = buildErrorBars(points, values, series, plot, min, max);
  return {
    ...sceneSeries,
    labels: labels.length > 0 ? labels : undefined,
    markers: markers.length > 0 ? markers : undefined,
    trendlines: trendlines.length > 0 ? trendlines : undefined,
    errorBars: errorBars.length > 0 ? errorBars : undefined,
    leaderLines: leaderLines && leaderLines.length > 0 ? leaderLines : undefined
  };
}

/**
 * Resolve label-to-label collisions for non-pie series (bar, line, scatter,
 * area). Labels are sorted left-to-right, pushed upward when the next
 * label's bounding box overlaps the previous one's, and dropped entirely
 * when even the maximum nudge would still overlap (so the user sees a
 * readable subset instead of an unreadable pile).
 *
 * The bounding box uses `estimateTextWidth` (real glyph metrics via
 * `@excel/utils/text-metrics`) so shorter labels deliberately leave room
 * for neighbours without forcing line breaks. This is the same algorithm
 * Chart.js uses for its default label layout — fast, deterministic, good
 * enough for 95% of bar/line previews without shipping a constraint
 * solver.
 */
function resolveLabelCollisions(labels: ChartSceneText[], plot: ChartSceneRect): ChartSceneText[] {
  const padding = 2;
  const topBound = plot.y + 4;
  interface Entry {
    label: ChartSceneText;
    left: number;
    right: number;
    width: number;
    height: number;
    kept: boolean;
  }
  const entries: Entry[] = labels.map(label => {
    const width = estimateTextWidth(label.text, label.fontSize);
    const height = label.fontSize;
    const anchor = label.anchor ?? "start";
    const left =
      anchor === "middle" ? label.x - width / 2 : anchor === "end" ? label.x - width : label.x;
    return {
      label,
      left,
      right: left + width,
      width,
      height,
      kept: true
    };
  });

  // Sort by horizontal centre so nudging cascades left-to-right; stable on
  // ties preserves input order for deterministic output.
  const ordered = entries
    .map((entry, index) => ({ entry, index }))
    .sort((a, b) => {
      const ax = a.entry.left + a.entry.width / 2;
      const bx = b.entry.left + b.entry.width / 2;
      return ax === bx ? a.index - b.index : ax - bx;
    });

  for (let i = 1; i < ordered.length; i++) {
    const curr = ordered[i].entry;
    // Find the nearest preceding kept neighbour whose bbox horizontally
    // overlaps ours; if any of its top edges clash vertically, push ours
    // above it. Iterate backwards so stacked neighbours chain.
    for (let j = i - 1; j >= 0; j--) {
      const prev = ordered[j].entry;
      if (!prev.kept) {
        continue;
      }
      const horizontallyApart = curr.left >= prev.right || prev.left >= curr.right;
      if (horizontallyApart) {
        continue;
      }
      const prevTop = prev.label.y - prev.height;
      const currTop = curr.label.y - curr.height;
      if (currTop + curr.height + padding > prevTop) {
        // Attempt to move current label above the previous one.
        const newY = prevTop - padding;
        if (newY - curr.height < topBound) {
          // No room left — drop this label entirely instead of overlapping.
          curr.kept = false;
          break;
        }
        curr.label = { ...curr.label, y: newY };
      }
    }
  }

  return entries.filter(entry => entry.kept).map(entry => entry.label);
}

/**
 * Move pie/doughnut data labels to the outside of each slice and emit
 * leader lines connecting the slice arc to the label anchor.
 *
 * Excel's `outEnd` (and `bestFit`, which resolves to the same thing for
 * pie-style charts) places each label along the angular bisector of its
 * slice, just past the outer radius. To keep neighbouring labels from
 * colliding we apply a greedy vertical nudge on each side of the pie:
 * labels are grouped by `left` (π/2 < θ < 3π/2) vs `right` hemisphere,
 * sorted by y, and then shifted downward one-by-one when the next label
 * overlaps its predecessor. The leader line ends on the slice arc (using
 * the original inward anchor) and kinks once at the outer ring so the
 * shape matches Excel's standard "radial + bend" callout rather than a
 * single diagonal.
 *
 * The nudge is intentionally simple — it resolves 95% of common collision
 * cases without pulling in a constraint solver. Extreme datasets (many
 * thin slices packed on one side) may still overlap; those belong to
 * the external oracle gap, not to this preview.
 */
function layoutPieLabels(
  series: ChartScenePieSeries,
  rawLabels: ChartSceneText[],
  plot: ChartSceneRect
): { labels: ChartSceneText[]; leaderLines: ChartSceneLine[] } {
  const allSlices = [...series.slices, ...(series.secondarySlices ?? [])];
  if (allSlices.length !== rawLabels.length) {
    // Defensive: if the adornment builder produced a different label count
    // (e.g. future dataPoint overrides) just return the input unchanged so
    // callers fall back to the interior placement that was working before.
    return { labels: rawLabels, leaderLines: [] };
  }

  const labelOffset = 14;
  interface Entry {
    index: number;
    slice: ChartScenePieSlice;
    angle: number;
    side: "left" | "right";
    anchorX: number;
    anchorY: number;
    x: number;
    y: number;
    text: string;
    fontSize: number;
    color: string;
    textAnchor: "start" | "end";
  }
  const entries: Entry[] = allSlices.map((slice, i) => {
    const angle = (slice.startAngle + slice.endAngle) / 2;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    const anchorX = slice.cx + cos * slice.radius;
    const anchorY = slice.cy + sin * slice.radius;
    // Right hemisphere: label aligned start; left hemisphere: aligned end.
    const side: "left" | "right" = cos >= 0 ? "right" : "left";
    const x = anchorX + cos * labelOffset;
    const y = anchorY + sin * labelOffset;
    const raw = rawLabels[i];
    return {
      index: i,
      slice,
      angle,
      side,
      anchorX,
      anchorY,
      x,
      y,
      text: raw.text,
      fontSize: raw.fontSize,
      color: raw.color,
      textAnchor: side === "left" ? "end" : "start"
    };
  });

  // Greedy collision avoidance per hemisphere: sort top→bottom and push
  // each label below its predecessor's baseline + fontSize.
  const nudge = (hemisphere: Entry[]): void => {
    hemisphere.sort((a, b) => a.y - b.y);
    const topBound = plot.y + 4;
    const bottomBound = plot.y + plot.height - 4;
    for (let i = 0; i < hemisphere.length; i++) {
      const e = hemisphere[i];
      if (i === 0) {
        e.y = Math.max(e.y, topBound);
        continue;
      }
      const prev = hemisphere[i - 1];
      const minY = prev.y + prev.fontSize + 2;
      if (e.y < minY) {
        e.y = minY;
      }
      if (e.y > bottomBound) {
        e.y = bottomBound;
      }
    }
  };
  nudge(entries.filter(e => e.side === "left"));
  nudge(entries.filter(e => e.side === "right"));

  // Clamp horizontal position so labels don't escape the plot rectangle.
  for (const e of entries) {
    if (e.side === "right") {
      e.x = Math.min(e.x, plot.x + plot.width - 4);
    } else {
      e.x = Math.max(e.x, plot.x + 4);
    }
  }

  const labels: ChartSceneText[] = entries.map(e => ({
    x: e.x,
    y: e.y,
    text: e.text,
    fontSize: e.fontSize,
    color: e.color,
    anchor: e.textAnchor === "end" ? "end" : "start"
  }));
  const leaderLines: ChartSceneLine[] = entries.map(e => ({
    x1: e.anchorX,
    y1: e.anchorY,
    x2: e.x - (e.side === "right" ? 4 : -4),
    y2: e.y - e.fontSize / 3,
    color: "#808080",
    width: 1
  }));
  return { labels, leaderLines };
}

function representativePoints(series: ChartSceneSeries): ChartScenePoint[] {
  if (series.type === "bar") {
    return series.bars.map(bar => ({ x: bar.x + bar.width / 2, y: bar.y }));
  }
  if (
    series.type === "area" ||
    series.type === "line" ||
    series.type === "scatter" ||
    series.type === "radar"
  ) {
    return series.points;
  }
  if (series.type === "bubble") {
    return series.bubbles.map(b => ({ x: b.x, y: b.y }));
  }
  if (series.type === "pie" || series.type === "doughnut" || series.type === "ofPie") {
    return [...series.slices, ...(series.secondarySlices ?? [])].map(slice => {
      const a = (slice.startAngle + slice.endAngle) / 2;
      const r = (slice.radius + slice.innerRadius) / 2;
      return { x: slice.cx + Math.cos(a) * r, y: slice.cy + Math.sin(a) * r };
    });
  }
  if (series.type === "stock") {
    return series.candles.map(c => ({ x: c.x, y: c.closeY ?? c.lowY }));
  }
  if (series.type === "surface") {
    return series.cells.map(cell => ({ x: cell.x + cell.width / 2, y: cell.y + cell.height / 2 }));
  }
  return [];
}

function buildDataLabels(
  points: ChartScenePoint[],
  values: number[],
  categories: string[],
  series: NormalizedSeries,
  plot: ChartSceneRect
): ChartSceneText[] {
  const labels = mergeDataLabels(series.group, series.series);
  if (!labels?.showVal && !labels?.showCatName && !labels?.showSerName && !labels?.showPercent) {
    return [];
  }
  const total = values.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
  const position = labels.position ?? "outEnd";
  const labelStyle = textStyleFromTxPr(labels.txPr);
  const labelColor = colorFromChartTextProperties(labels.txPr) ?? "#333333";
  return points.map((point, i) => {
    const entry = labels.entries?.find(e => e.index === i);
    const effectivePosition = entry?.position ?? position;
    const { x, y, anchor } = positionDataLabel(point, effectivePosition, plot);
    return {
      x,
      y,
      text: makeDataLabelText(labels, series, categories[i], values[i] ?? 0, total),
      fontSize: 10,
      color: labelColor,
      anchor,
      ...labelStyle
    };
  });
}

/**
 * Resolve a logical `DataLabelPosition` into concrete SVG text coordinates.
 *
 * The preview renderer does not know the full geometry of every mark (bars
 * vs. markers vs. slices), so this function operates on the scene `point`
 * (which is already the mark's centre or outer-end anchor, depending on how
 * the series type built it). The mapping matches Excel's user-visible
 * positions as closely as possible for deterministic previews:
 *
 * - `t` / `outEnd` → above the point, centred
 * - `b` / `inBase` → below the point, centred
 * - `l` → left of the point, right-anchored
 * - `r` → right of the point, left-anchored
 * - `ctr` / `inEnd` / `bestFit` → at the point, centred
 *
 * `plot` is used to keep the label inside the plot rectangle for edge cases.
 */
function positionDataLabel(
  point: ChartScenePoint,
  position: DataLabelPosition,
  plot: ChartSceneRect
): { x: number; y: number; anchor: "start" | "middle" | "end" } {
  const offset = 6;
  const minY = plot.y + 10;
  const maxY = plot.y + plot.height - 4;
  const minX = plot.x + 4;
  const maxX = plot.x + plot.width - 4;
  switch (position) {
    case "t":
    case "outEnd":
      return {
        x: point.x,
        y: Math.max(minY, point.y - offset),
        anchor: "middle"
      };
    case "b":
    case "inBase":
      return {
        x: point.x,
        y: Math.min(maxY, point.y + offset + 10),
        anchor: "middle"
      };
    case "l":
      return {
        x: Math.max(minX, point.x - offset),
        y: point.y + 3,
        anchor: "end"
      };
    case "r":
      return {
        x: Math.min(maxX, point.x + offset),
        y: point.y + 3,
        anchor: "start"
      };
    case "ctr":
    case "inEnd":
    case "bestFit":
    default:
      return {
        x: point.x,
        y: point.y + 3,
        anchor: "middle"
      };
  }
}

function buildMarkers(points: ChartScenePoint[], series: NormalizedSeries): ChartSceneMarker[] {
  const marker = series.marker;
  if (!marker || marker.symbol === "none") {
    return [];
  }
  return points.map(point => ({
    ...point,
    color: series.color,
    size: marker.size ?? 6,
    symbol: marker.symbol ?? "circle"
  }));
}

function buildTrendlines(
  points: ChartScenePoint[],
  values: number[],
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartSceneTrendline[] {
  if (!series.trendlines || points.length < 2) {
    return [];
  }
  return series.trendlines.map(trendline => {
    const trendPoints =
      trendline.type === "movingAvg"
        ? movingAveragePoints(points, values, trendline.period ?? 2, plot, min, max)
        : linearTrendlinePoints(points);
    return {
      color: colorFromShapeLine(trendline.spPr?.line) ?? "#666666",
      width: trendline.spPr?.line?.width ? trendline.spPr.line.width / 12700 : 1.5,
      dash: trendline.spPr?.line?.dash,
      points: trendPoints,
      label: trendline.name
        ? {
            x: trendPoints[trendPoints.length - 1].x,
            y: trendPoints[trendPoints.length - 1].y - 8,
            text: trendline.name,
            fontSize: 10,
            color: "#555555",
            anchor: "end"
          }
        : undefined
    };
  });
}

function buildErrorBars(
  points: ChartScenePoint[],
  values: number[],
  series: NormalizedSeries,
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartSceneErrorBar[] {
  if (!series.errorBars || points.length === 0) {
    return [];
  }
  const bars: ChartSceneErrorBar[] = [];
  for (const err of series.errorBars) {
    const color = colorFromShapeLine(err.spPr?.line) ?? "#555555";
    for (let i = 0; i < points.length; i++) {
      const p = points[i];
      const value = values[i] ?? 0;
      const amount = errorAmount(err, value, values);
      const plusY = valueToY(value + amount, min, max, plot);
      const minusY = valueToY(value - amount, min, max, plot);
      bars.push({
        line: { x1: p.x, y1: plusY, x2: p.x, y2: minusY, color },
        cap1: err.noEndCap ? undefined : { x1: p.x - 4, y1: plusY, x2: p.x + 4, y2: plusY, color },
        cap2: err.noEndCap ? undefined : { x1: p.x - 4, y1: minusY, x2: p.x + 4, y2: minusY, color }
      });
    }
  }
  return bars;
}

function buildXLabels(
  categories: string[],
  plot: ChartSceneRect,
  axis?: ChartAxis,
  range?: ValueRange,
  top = false
): ChartSceneText[] {
  if (axis?.delete || axis?.tickLblPos === "none") {
    return [];
  }
  if (axis?.axisType === "val") {
    const { min, max } = range ?? { min: 0, max: 1 };
    return buildValueXLabels(min, max, plot, axis, top);
  }
  const count = Math.max(1, categories.length);
  const groupWidth = plot.width / count;
  const skip = axis?.axisType === "cat" || axis?.axisType === "ser" ? (axis.tickLblSkip ?? 1) : 1;
  const visible = categories.slice(0, 12).filter((_, i) => i % Math.max(1, skip) === 0);
  const axisStyle = textStyleFromTxPr(axis?.txPr);
  return visible.map((label, visibleIndex) => {
    const i = visibleIndex * Math.max(1, skip);
    return {
      x: plot.x + i * groupWidth + groupWidth / 2,
      y: top ? plot.y - 10 : plot.y + plot.height + 18,
      text: truncateLabel(label),
      fontSize: 10,
      color: tickLabelColor(axis),
      anchor: "middle",
      ...axisStyle
    };
  });
}

function buildValueXLabels(
  min: number,
  max: number,
  plot: ChartSceneRect,
  axis: ChartAxis | undefined,
  top: boolean
): ChartSceneText[] {
  const labels: ChartSceneText[] = [];
  const color = tickLabelColor(axis);
  const axisStyle = textStyleFromTxPr(axis?.txPr);
  for (let i = 0; i <= 5; i++) {
    const value = min + ((max - min) * i) / 5;
    labels.push({
      x: valueToX(value, min, max, plot),
      y: top ? plot.y - 10 : plot.y + plot.height + 18,
      text: formatAxisNumber(value),
      fontSize: 10,
      color,
      anchor: "middle",
      ...axisStyle
    });
  }
  return labels;
}

function buildYLabels(
  min: number,
  max: number,
  plot: ChartSceneRect,
  axis?: ChartAxis,
  right = false
): ChartSceneText[] {
  if (axis?.delete || axis?.tickLblPos === "none") {
    return [];
  }
  const labels: ChartSceneText[] = [];
  const color = tickLabelColor(axis);
  const axisStyle = textStyleFromTxPr(axis?.txPr);
  for (let i = 0; i <= 5; i++) {
    const value = min + ((max - min) * i) / 5;
    labels.push({
      x: right ? plot.x + plot.width + 8 : plot.x - 8,
      y: valueToY(value, min, max, plot) + 3,
      text: formatAxisNumber(value),
      fontSize: 10,
      color,
      anchor: right ? "start" : "end",
      ...axisStyle
    });
  }
  return labels;
}

function normalizeSeries(groups: ChartTypeGroup[], model?: ChartModel): NormalizedSeries[] {
  // Pre-compute axis log transforms once: `logBase` lives on the value
  // axis (`c:valAx/c:scaling/c:logBase`). We apply the transform to
  // every series' `values` and `xValues` here so every downstream
  // consumer (range calculation, valueToY, trendlines, error bars)
  // sees the already-mapped coordinates without needing axis context.
  //
  // This is a "best-effort" log axis: OOXML forbids non-positive
  // values on log axes, but the renderer is preview-grade and refusing
  // to render is worse than placing them at the axis floor. Negative
  // values pass through unchanged (see {@link applyAxisTransform}).
  const yLogBase = extractValueAxisLogBase(model, "y");
  const xLogBase = extractValueAxisLogBase(model, "x");

  const normalized: NormalizedSeries[] = [];
  let globalIndex = 0;
  for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
    const group = groups[groupIndex];
    const series = collectSeries(group);
    for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex++) {
      const s = series[seriesIndex];
      const rawValues = collectValues(s);
      const rawXValues = collectAxisValues((s as { xVal?: unknown }).xVal);
      normalized.push({
        group,
        groupIndex,
        series: s,
        seriesIndex,
        globalIndex,
        label: collectSeriesLabel(s, globalIndex),
        color: seriesColor(s, globalIndex),
        values: yLogBase
          ? rawValues.map(v => (typeof v === "number" ? applyAxisTransform(v, yLogBase) : v))
          : rawValues,
        categories: collectCategories(s),
        xValues: xLogBase
          ? rawXValues.map(v => (typeof v === "number" ? applyAxisTransform(v, xLogBase) : v))
          : rawXValues,
        bubbleSizes: collectNumberValues(
          (s as { bubbleSize?: { numRef?: unknown } }).bubbleSize?.numRef
        ),
        dataLabels: (s as { dataLabels?: DataLabels }).dataLabels,
        trendlines: (s as { trendlines?: Trendline[] }).trendlines,
        errorBars: normalizeErrorBars((s as { errorBars?: ErrorBars | ErrorBars[] }).errorBars),
        marker: (s as { marker?: ChartMarker }).marker
      });
      globalIndex++;
    }
  }
  return normalized;
}

/**
 * Find the log base configured on the first value axis matching
 * `axisKind`. Returns `undefined` when no axis of that kind exists or
 * when the axis is linear.
 *
 * We treat `axisKind === "y"` as any value axis whose position is
 * `l` / `r` (left / right — the usual vertical placements) and
 * `axisKind === "x"` as a value axis placed at `b` / `t` (the
 * scatter/bubble x axis). Category axes never carry a log scale in
 * OOXML so we ignore them.
 */
function extractValueAxisLogBase(
  model: ChartModel | undefined,
  axisKind: "y" | "x"
): number | undefined {
  const axes = model?.chart?.plotArea?.axes;
  if (!axes) {
    return undefined;
  }
  const hPos = new Set(["b", "t"]);
  const vPos = new Set(["l", "r"]);
  const positions = axisKind === "y" ? vPos : hPos;
  for (const ax of axes) {
    // Only value axes carry logBase. The discriminator is the presence
    // of numeric scaling fields, which `ValueAxis` / `DateAxis` expose.
    const axisPosition = (ax as { axPos?: string }).axPos;
    if (!axisPosition || !positions.has(axisPosition)) {
      continue;
    }
    const logBase = (ax as { scaling?: { logBase?: number } }).scaling?.logBase;
    if (logBase && logBase > 0 && logBase !== 1) {
      return logBase;
    }
  }
  return undefined;
}

function seriesColor(series: SeriesBase, index: number): string {
  return colorFromShapeFill((series as { spPr?: any }).spPr?.fill) ?? COLORS[index % COLORS.length];
}

function normalizeErrorBars(value: ErrorBars | ErrorBars[] | undefined): ErrorBars[] | undefined {
  if (!value) {
    return undefined;
  }
  return Array.isArray(value) ? value : [value];
}

function buildAxisContext(model: ChartModel, normalized: NormalizedSeries[]): ChartAxisContext {
  const axes = model.chart.plotArea.axes;
  const axesById = new Map(axes.map(axis => [axis.axId, axis]));
  const axisIdsByGroup = new Map<ChartTypeGroup, number[]>();
  for (const group of model.chart.plotArea.chartTypes) {
    axisIdsByGroup.set(group, (group as { axisIds?: number[] }).axisIds ?? []);
  }
  const defaultYRange = getValueRange(
    normalized.map(s => s.values),
    undefined
  );
  const defaultXRange = getValueRange(
    normalized.map(s => scatterXValues(s)),
    undefined,
    { includeZero: false }
  );
  const yValuesByAxisId = new Map<number, number[][]>();
  const xValuesByAxisId = new Map<number, number[][]>();
  for (const series of normalized) {
    const ids = axisIdsByGroup.get(series.group) ?? [];
    const yAxisId = getYAxisIdForGroup(series.group, ids, axesById);
    const xAxisId = getXAxisIdForGroup(series.group, ids, axesById);
    if (yAxisId !== undefined) {
      addAxisValues(yValuesByAxisId, yAxisId, series.values);
    }
    if (xAxisId !== undefined && isValueValueGroup(series.group)) {
      addAxisValues(xValuesByAxisId, xAxisId, scatterXValues(series));
    }
  }
  const yRangesByAxisId = new Map<number, ValueRange>();
  for (const [axisId, values] of yValuesByAxisId) {
    yRangesByAxisId.set(axisId, getValueRange(values, axesById.get(axisId)));
  }
  const xRangesByAxisId = new Map<number, ValueRange>();
  for (const [axisId, values] of xValuesByAxisId) {
    xRangesByAxisId.set(
      axisId,
      getValueRange(values, axesById.get(axisId), { includeZero: false })
    );
  }
  return {
    axesById,
    primaryXAxis: pickAxis(axes, "x", false),
    primaryYAxis: pickAxis(axes, "y", false),
    secondaryXAxis: pickAxis(axes, "x", true),
    secondaryYAxis: pickAxis(axes, "y", true),
    yRangesByAxisId,
    xRangesByAxisId,
    defaultYRange,
    defaultXRange
  };
}

function addAxisValues(map: Map<number, number[][]>, axisId: number, values: number[]): void {
  const existing = map.get(axisId);
  if (existing) {
    existing.push(values);
  } else {
    map.set(axisId, [values]);
  }
}

function pickAxis(
  axes: ChartAxis[],
  direction: "x" | "y",
  secondary: boolean
): ChartAxis | undefined {
  const positions = direction === "x" ? ["t", "b"] : ["r", "l"];
  const preferred = secondary ? positions[0] : positions[1];
  const fallback = secondary ? positions[1] : positions[0];
  return (
    axes.find(axis => !axis.delete && axis.axPos === preferred) ??
    (!secondary ? axes.find(axis => !axis.delete && axis.axPos === fallback) : undefined)
  );
}

function getSeriesYRange(series: NormalizedSeries, context: ChartAxisContext): ValueRange {
  const axisId = getYAxisIdForGroup(
    series.group,
    (series.group as { axisIds?: number[] }).axisIds ?? [],
    context.axesById
  );
  return axisId !== undefined
    ? (context.yRangesByAxisId.get(axisId) ?? context.defaultYRange)
    : context.defaultYRange;
}

function getSeriesXRange(series: NormalizedSeries, context: ChartAxisContext): ValueRange {
  const axisId = getXAxisIdForGroup(
    series.group,
    (series.group as { axisIds?: number[] }).axisIds ?? [],
    context.axesById
  );
  return axisId !== undefined
    ? (context.xRangesByAxisId.get(axisId) ?? context.defaultXRange)
    : context.defaultXRange;
}

function getGroupYRange(group: ChartTypeGroup, context: ChartAxisContext): ValueRange {
  const axisId = getYAxisIdForGroup(
    group,
    (group as { axisIds?: number[] }).axisIds ?? [],
    context.axesById
  );
  return axisId !== undefined
    ? (context.yRangesByAxisId.get(axisId) ?? context.defaultYRange)
    : context.defaultYRange;
}

function getYAxisIdForGroup(
  group: ChartTypeGroup,
  axisIds: number[],
  axesById: Map<number, ChartAxis>
): number | undefined {
  if (axisIds.length === 0) {
    return undefined;
  }
  if (isValueValueGroup(group)) {
    return (
      axisIds.find(id => axesById.get(id)?.axPos === "l" || axesById.get(id)?.axPos === "r") ??
      axisIds[1]
    );
  }
  return axisIds.find(id => axesById.get(id)?.axisType === "val") ?? axisIds[1];
}

function getXAxisIdForGroup(
  group: ChartTypeGroup,
  axisIds: number[],
  axesById: Map<number, ChartAxis>
): number | undefined {
  if (axisIds.length === 0) {
    return undefined;
  }
  if (isValueValueGroup(group)) {
    return (
      axisIds.find(id => axesById.get(id)?.axPos === "b" || axesById.get(id)?.axPos === "t") ??
      axisIds[0]
    );
  }
  return axisIds.find(id => axesById.get(id)?.axisType !== "val") ?? axisIds[0];
}

function isValueValueGroup(group: ChartTypeGroup): boolean {
  return group.type === "scatter" || group.type === "bubble";
}

function scatterXValues(series: NormalizedSeries): number[] {
  return series.xValues && series.xValues.length > 0
    ? series.xValues
    : series.values.map((_, i) => i + 1);
}

function buildSceneLegend(
  legend: ChartLegend | undefined,
  series: NormalizedSeries[],
  width: number,
  height: number,
  hasTitle: boolean
): ChartSceneLegend {
  const visible = legend !== undefined && series.length > 0;
  const position = legend?.legendPos ?? "r";
  const orientation = position === "b" || position === "t" ? "horizontal" : "vertical";
  const deletedEntries = new Set(
    legend?.legendEntries?.filter(entry => entry.delete).map(entry => entry.index) ?? []
  );
  const visibleLabels = series.filter((_, index) => !deletedEntries.has(index)).map(s => s.label);
  const rect = legendRect(position, orientation, width, height, hasTitle, visibleLabels);
  const txStyle = textStyleFromTxPr(legend?.txPr);
  const txColor = legend?.txPr ? colorFromChartTextProperties(legend.txPr) : undefined;
  const textStyle =
    txStyle.fontFamily !== undefined ||
    txStyle.bold ||
    txStyle.italic ||
    txStyle.fontSize !== undefined ||
    txColor
      ? { ...txStyle, color: txColor }
      : undefined;
  return {
    rect,
    visible,
    position,
    orientation,
    items: series
      .filter((_, index) => !deletedEntries.has(index))
      .map(s => ({
        label: s.label,
        color: s.color
      })),
    textStyle
  };
}

function legendRect(
  position: LegendPosition,
  orientation: "horizontal" | "vertical",
  width: number,
  height: number,
  hasTitle: boolean,
  labels: string[]
): ChartSceneRect {
  // Each legend entry renders as a 12-px colour swatch + 4 px gap + label
  // + 16 px inter-item gap. Using real text metrics here (rather than a
  // hardcoded `86` per item) lets longer series names push the legend
  // wider instead of being truncated by the viewport clipping region.
  const entryPadding = 32;
  const swatchWidth = 16;
  const legendFontSize = 11;
  const entryWidths = labels.map(
    label => swatchWidth + estimateTextWidth(label, legendFontSize) + entryPadding
  );
  if (orientation === "horizontal") {
    const totalEntries = entryWidths.reduce((sum, w) => sum + w, 0);
    const legendWidth = Math.min(width - 48, Math.max(96, Math.ceil(totalEntries)));
    return {
      x: (width - legendWidth) / 2,
      y: position === "t" ? (hasTitle ? 48 : 20) : height - 26,
      width: legendWidth,
      height: 18
    };
  }
  const longestLabelWidth = entryWidths.reduce((max, w) => (w > max ? w : max), 0);
  // Vertical legends stack entries vertically, so width is governed by the
  // widest label, not the total; ensure a sensible minimum so short names
  // don't produce an absurdly narrow legend column.
  const legendColumnWidth = Math.min(
    Math.max(96, Math.ceil(longestLabelWidth)),
    Math.max(96, width - 64)
  );
  const legendHeight = Math.min(height - 48, Math.max(18, labels.length * 18));
  return {
    x: position === "l" ? 20 : width - legendColumnWidth - 20,
    y:
      position === "tr"
        ? hasTitle
          ? 44
          : 20
        : Math.max(hasTitle ? 52 : 24, (height - legendHeight) / 2),
    width: legendColumnWidth,
    height: legendHeight
  };
}

function buildAxisTitles(context: ChartAxisContext, plot: ChartSceneRect): ChartSceneText[] {
  const titles: ChartSceneText[] = [];
  addAxisTitle(titles, context.primaryXAxis, plot);
  addAxisTitle(titles, context.primaryYAxis, plot);
  addAxisTitle(titles, context.secondaryXAxis, plot);
  addAxisTitle(titles, context.secondaryYAxis, plot);
  return titles;
}

function addAxisTitle(
  titles: ChartSceneText[],
  axis: ChartAxis | undefined,
  plot: ChartSceneRect
): void {
  if (!axis || axis.delete || !axis.title) {
    return;
  }
  const text = titleToText(axis.title);
  if (!text) {
    return;
  }
  const base = {
    text,
    fontSize: 11,
    color: titleColor(axis.title),
    anchor: "middle" as const,
    ...textStyleFromTxPr(axis.title.txPr)
  };
  if (axis.axPos === "b") {
    titles.push({ ...base, x: plot.x + plot.width / 2, y: plot.y + plot.height + 38 });
  } else if (axis.axPos === "t") {
    titles.push({ ...base, x: plot.x + plot.width / 2, y: plot.y - 26 });
  } else if (axis.axPos === "l") {
    titles.push({ ...base, x: plot.x - 44, y: plot.y + plot.height / 2, rotate: -90 });
  } else if (axis.axPos === "r") {
    titles.push({ ...base, x: plot.x + plot.width + 44, y: plot.y + plot.height / 2, rotate: 90 });
  }
}

function categoriesForAxis(
  series: NormalizedSeries[],
  axis: ChartAxis | undefined,
  fallback: string[]
): string[] {
  if (!axis) {
    return fallback;
  }
  const matched = series.find(s => {
    const ids = (s.group as { axisIds?: number[] }).axisIds ?? [];
    return ids[0] === axis.axId;
  });
  return matched?.categories && matched.categories.length > 0 ? matched.categories : fallback;
}

function titleToText(title: ChartTitle): string | undefined {
  return (
    title.text?.paragraphs.map(p => (p.runs ?? []).map(r => r.text).join("")).join("\n") ??
    title.strRef?.cache?.points?.map(p => p.value).join("")
  );
}

function axisColor(axis: ChartAxis | undefined): string {
  return colorFromShapeLine(axis?.spPr?.line) ?? AXIS_COLOR;
}

function tickLabelColor(axis: ChartAxis | undefined): string {
  return colorFromChartTextProperties(axis?.txPr) ?? "#555555";
}

function titleColor(title: ChartTitle): string {
  return colorFromChartTextProperties(title.txPr) ?? "#333333";
}

function colorFromChartTextProperties(
  textProperties: { color?: { srgb?: string } } | undefined
): string | undefined {
  const srgb = textProperties?.color?.srgb;
  return typeof srgb === "string" ? `#${srgb.replace(/^#/, "")}` : undefined;
}

/**
 * Extract the font-related fields ({@link ChartSceneText}) from an
 * OOXML `txPr`-shaped record. Used to thread font family / bold /
 * italic from the chart model into the scene so SVG, PNG and PDF paths
 * all render with the authored typography. `color` / `fontSize` stay
 * at the call site — each text kind has its own default for those
 * (e.g. tick labels default to `#555555`, titles to `#333333`).
 *
 * Accepts both already-structured `ChartTextProperties` and the raw-XML
 * passthrough shape (`{ _rawXml: "..." }`) the xform/parser produces.
 * In the raw case we lazily invoke `parseTxPr` so consumers don't need
 * to pre-process loaded chart parts themselves.
 */
function textStyleFromTxPr(
  textProperties:
    | {
        _rawXml?: string;
        fontFamily?: string;
        bold?: boolean;
        italic?: boolean;
        size?: number;
      }
    | undefined
): { fontFamily?: string; bold?: boolean; italic?: boolean; fontSize?: number } {
  if (!textProperties) {
    return {};
  }
  // If the record arrived as raw XML (from a loaded chart part), decode
  // it once via the shape-properties parser. This mirrors how
  // `titleColor` already resolves colour from either form.
  const resolved =
    typeof textProperties._rawXml === "string" && textProperties.fontFamily === undefined
      ? (parseTxPr(textProperties as unknown as ChartTextProperties) as typeof textProperties)
      : textProperties;
  const out: { fontFamily?: string; bold?: boolean; italic?: boolean; fontSize?: number } = {};
  if (typeof resolved.fontFamily === "string" && resolved.fontFamily.length > 0) {
    out.fontFamily = resolved.fontFamily;
  }
  if (resolved.bold) {
    out.bold = true;
  }
  if (resolved.italic) {
    out.italic = true;
  }
  // `txPr.size` is in hundredths of a point (OOXML convention); convert
  // to points. Ignore zero/NaN so callers keep their default fontSize.
  if (typeof resolved.size === "number" && Number.isFinite(resolved.size) && resolved.size > 0) {
    out.fontSize = resolved.size / 100;
  }
  return out;
}

function mergeDataLabels(group: ChartTypeGroup, series: SeriesBase): DataLabels | undefined {
  const groupLabels = (group as { dataLabels?: DataLabels }).dataLabels;
  const seriesLabels = (series as { dataLabels?: DataLabels }).dataLabels;
  return groupLabels || seriesLabels
    ? { ...(groupLabels ?? {}), ...(seriesLabels ?? {}) }
    : undefined;
}

function makeDataLabelText(
  labels: DataLabels,
  series: NormalizedSeries,
  category: string | undefined,
  value: number,
  total: number
): string {
  const sep = labels.separator ?? ", ";
  const parts: string[] = [];
  if (labels.showSerName) {
    parts.push(series.label);
  }
  if (labels.showCatName && category !== undefined) {
    parts.push(category);
  }
  if (labels.showVal) {
    parts.push(formatAxisNumber(value));
  }
  if (labels.showPercent) {
    parts.push(`${Math.round((Math.max(0, value) / total) * 100)}%`);
  }
  return parts.join(sep);
}

function linearTrendlinePoints(points: ChartScenePoint[]): ChartScenePoint[] {
  if (points.length < 2) {
    return points;
  }
  const n = points.length;
  const meanX = points.reduce((sum, p) => sum + p.x, 0) / n;
  const meanY = points.reduce((sum, p) => sum + p.y, 0) / n;
  const denominator = points.reduce((sum, p) => sum + (p.x - meanX) ** 2, 0) || 1;
  const slope = points.reduce((sum, p) => sum + (p.x - meanX) * (p.y - meanY), 0) / denominator;
  const intercept = meanY - slope * meanX;
  const x1 = points[0].x;
  const x2 = points[points.length - 1].x;
  return [
    { x: x1, y: slope * x1 + intercept },
    { x: x2, y: slope * x2 + intercept }
  ];
}

function movingAveragePoints(
  points: ChartScenePoint[],
  values: number[],
  period: number,
  plot: ChartSceneRect,
  min: number,
  max: number
): ChartScenePoint[] {
  const safePeriod = Math.max(1, Math.floor(period));
  return points.map((point, i) => {
    const start = Math.max(0, i - safePeriod + 1);
    const window = values.slice(start, i + 1);
    const avg = window.reduce((sum, v) => sum + v, 0) / Math.max(1, window.length);
    return { x: point.x, y: valueToY(avg, min, max, plot) };
  });
}

function errorAmount(err: ErrorBars, value: number, values: number[]): number {
  if (err.errValType === "fixedVal") {
    return err.val ?? 1;
  }
  if (err.errValType === "percentage") {
    return Math.abs(value) * ((err.val ?? 5) / 100);
  }
  if (err.errValType === "stdDev") {
    const mean = values.reduce((sum, v) => sum + v, 0) / Math.max(1, values.length);
    const variance =
      values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / Math.max(1, values.length);
    return Math.sqrt(variance) * (err.val ?? 1);
  }
  if (err.errValType === "stdErr") {
    const mean = values.reduce((sum, v) => sum + v, 0) / Math.max(1, values.length);
    const variance =
      values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / Math.max(1, values.length);
    return Math.sqrt(variance) / Math.sqrt(Math.max(1, values.length));
  }
  return err.val ?? 1;
}

function colorFromShapeFill(fill: any): string | undefined {
  const srgb = fill?.solid?.srgb;
  return typeof srgb === "string" ? `#${srgb.replace(/^#/, "")}` : undefined;
}

function colorFromShapeLine(line: any): string | undefined {
  const srgb = line?.color?.srgb;
  return typeof srgb === "string" ? `#${srgb.replace(/^#/, "")}` : undefined;
}

function lineWidthFromShapeLine(line: any): number | undefined {
  return typeof line?.width === "number" ? Math.max(0.5, line.width / 12700) : undefined;
}

function getValueRange(
  seriesValues: number[][],
  axis?: ChartAxis,
  options: { includeZero?: boolean } = {}
): ValueRange {
  const values = seriesValues.flat().filter(Number.isFinite);
  const includeZero = options.includeZero !== false;
  const rawMin = values.length > 0 ? Math.min(...(includeZero ? [0, ...values] : values)) : 0;
  const rawMax = values.length > 0 ? Math.max(1, ...values) : 1;
  const min = axis?.scaling?.min ?? rawMin;
  const max = axis?.scaling?.max ?? rawMax;
  return max <= min ? { min, max: min + 1 } : { min, max };
}

function valueToY(value: number, min: number, max: number, plot: ChartSceneRect): number {
  return plot.y + plot.height - ((value - min) / (max - min)) * plot.height;
}

function valueToX(value: number, min: number, max: number, plot: ChartSceneRect): number {
  return plot.x + ((value - min) / (max - min)) * plot.width;
}

/**
 * Apply a logarithmic axis transform when the given axis is configured
 * with `scaling.logBase`. Caller uses the transformed value / min / max
 * with the plain {@link valueToY} / {@link valueToX} helpers so the
 * rest of the renderer does not need to know about the log scale.
 *
 * Returns the input unchanged when `logBase` is absent / invalid or
 * when the value is non-positive (log scales can't plot ≤0 values; we
 * fall back to the raw value rather than `-Infinity` so the point
 * still shows up somewhere near the bottom).
 */
export function applyAxisTransform(value: number, logBase: number | undefined): number {
  if (!logBase || logBase <= 0 || logBase === 1) {
    return value;
  }
  if (!Number.isFinite(value) || value <= 0) {
    // OOXML: "Values less than or equal to 0 are not allowed on
    // logarithmic axes." We can't refuse to render, so anchor them at
    // the axis minimum by returning a sentinel the caller may clamp
    // later. For now, skip the transform — the caller sees the raw
    // value and the renderer places it at the axis bottom.
    return value;
  }
  return Math.log(value) / Math.log(logBase);
}

/**
 * Generate SVG `<filter>` definitions for any series in `normalized`
 * that carries an `a:effectLst`, and thread the filter id onto the
 * matching `sceneSeries` entry so the renderer can reference it.
 *
 * Effect lists with identical structure share a single filter
 * definition (cache keyed by JSON stringify) so a chart with five
 * series all using the same drop shadow emits only one `<filter>`.
 *
 * Mutates `sceneSeries[i].effectFilterId` in place and returns the
 * list of unique filter definitions. Safe to call on a scene that
 * has no effect lists — returns an empty array in that case.
 */
function assignEffectFilters(
  normalized: NormalizedSeries[],
  sceneSeries: ChartSceneSeries[]
): ChartSceneEffectFilter[] {
  const filters: ChartSceneEffectFilter[] = [];
  const keyToId = new Map<string, string>();
  const limit = Math.min(normalized.length, sceneSeries.length);
  for (let i = 0; i < limit; i++) {
    const effects = (normalized[i].series as { spPr?: { effectList?: EffectList } }).spPr
      ?.effectList;
    if (!effects) {
      continue;
    }
    // Stable cache key so duplicate effect trees share a filter. We
    // intentionally ignore field ordering differences — JSON stringify
    // is deterministic for the flat record shape `EffectList` uses.
    const key = JSON.stringify(effects);
    let id = keyToId.get(key);
    if (!id) {
      const xml = buildEffectFilter(`excelts-fx-${filters.length + 1}`, effects);
      if (!xml) {
        continue;
      }
      id = `excelts-fx-${filters.length + 1}`;
      keyToId.set(key, id);
      filters.push({ id, xml });
    }
    (sceneSeries[i] as ChartSceneAdornment).effectFilterId = id;
  }
  return filters;
}

/**
 * Translate a DrawingML {@link EffectList} into an SVG `<filter>`
 * element. The translation is lossy — SVG filters don't express every
 * Excel effect pixel-for-pixel — but they do capture the four effects
 * that cover the overwhelming majority of real-world usage:
 *
 *   - `outerShadow` / `presetShadow` → `<feGaussianBlur>` +
 *     `<feOffset>` + `<feMerge>` (Excel's outer-shadow dist/dir/blur
 *     maps directly onto SVG's offset + blur-stdDeviation).
 *   - `innerShadow`                  → `<feComposite operator="arithmetic">`
 *     carving a shadow inside the shape.
 *   - `glow`                         → blurred coloured duplicate
 *     under the source.
 *   - `blur`                         → a single `<feGaussianBlur>`.
 *   - `softEdge`                     → a `<feGaussianBlur>` feeding
 *     into the source alpha.
 *
 * Reflection is modelled as a semi-transparent y-flipped duplicate;
 * Excel's actual implementation uses a much more complicated
 * gradient-masked clone. The preview is recognisable as a reflection
 * but does not match Excel pixel-for-pixel.
 *
 * Callers are expected to embed the returned string inside an SVG
 * `<defs>` block and reference it via `filter="url(#<id>)"` on the
 * shapes they want to affect. Returns `""` when the effect list is
 * empty — callers can skip emission entirely.
 */
export function buildEffectFilter(id: string, effects: EffectList | undefined): string {
  if (!effects) {
    return "";
  }
  const prims: string[] = [];
  let inLayer = "SourceGraphic";

  if (effects.blur) {
    prims.push(
      `<feGaussianBlur in="${inLayer}" stdDeviation="${emuToPx(effects.blur.radius ?? 0)}" result="blur0"/>`
    );
    inLayer = "blur0";
  }

  if (effects.softEdge) {
    const sd = emuToPx(effects.softEdge.radius);
    prims.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${sd}" result="softAlpha"/>`,
      `<feComposite in="${inLayer}" in2="softAlpha" operator="in" result="soft"/>`
    );
    inLayer = "soft";
  }

  const shadow = effects.outerShadow ?? effects.presetShadow;
  if (shadow) {
    // `presetShadow` lacks `blurRadius`; fall back to a gentle default.
    const blur = emuToPx((shadow as { blurRadius?: number }).blurRadius ?? 38100);
    const [dx, dy] = polarOffset(shadow.distance ?? 0, shadow.direction ?? 0);
    const colour = colourToHex(shadow.color) ?? "#000000";
    const alpha = alphaFromColor(shadow.color);
    prims.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${blur}" result="shadowBlur"/>`,
      `<feOffset in="shadowBlur" dx="${dx}" dy="${dy}" result="shadowOffset"/>`,
      `<feFlood flood-color="${colour}" flood-opacity="${alpha}" result="shadowColour"/>`,
      `<feComposite in="shadowColour" in2="shadowOffset" operator="in" result="shadowOut"/>`,
      `<feMerge><feMergeNode in="shadowOut"/><feMergeNode in="${inLayer}"/></feMerge>`
    );
    inLayer = "shadowOut-merged";
  }

  if (effects.glow) {
    const blur = emuToPx(effects.glow.radius);
    const colour = colourToHex(effects.glow.color) ?? "#ffff66";
    prims.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${blur}" result="glowBlur"/>`,
      `<feFlood flood-color="${colour}" result="glowColour"/>`,
      `<feComposite in="glowColour" in2="glowBlur" operator="in" result="glowOut"/>`,
      `<feMerge><feMergeNode in="glowOut"/><feMergeNode in="${inLayer}"/></feMerge>`
    );
  }

  if (effects.innerShadow) {
    const blur = emuToPx(effects.innerShadow.blurRadius ?? 0);
    const [dx, dy] = polarOffset(
      effects.innerShadow.distance ?? 0,
      effects.innerShadow.direction ?? 0
    );
    const colour = colourToHex(effects.innerShadow.color) ?? "#000000";
    const alpha = alphaFromColor(effects.innerShadow.color);
    prims.push(
      `<feGaussianBlur in="SourceAlpha" stdDeviation="${blur}" result="innerBlur"/>`,
      `<feOffset in="innerBlur" dx="${dx}" dy="${dy}" result="innerOffset"/>`,
      `<feComposite in="SourceAlpha" in2="innerOffset" operator="out" result="innerClipped"/>`,
      `<feFlood flood-color="${colour}" flood-opacity="${alpha}" result="innerColour"/>`,
      `<feComposite in="innerColour" in2="innerClipped" operator="in" result="innerOut"/>`,
      `<feMerge><feMergeNode in="${inLayer}"/><feMergeNode in="innerOut"/></feMerge>`
    );
  }

  if (prims.length === 0) {
    return "";
  }
  // Filter region is widened past the shape bounds so blurred shadows
  // don't clip at the element edges. SVG's default is -10%..-10%..120%..120%
  // which already covers most of what we need, but Excel effects can
  // be dramatic.
  return `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%" filterUnits="objectBoundingBox">${prims.join(
    ""
  )}</filter>`;
}

/** Convert an Excel EMU distance to approximate SVG px (12700 EMU ≈ 1 pt ≈ 1.333 px). */
function emuToPx(emu: number): number {
  return emu / 12700;
}

/**
 * Decompose Excel's polar shadow (`dist` in EMU, `dir` in 1/60000°)
 * into SVG-native (dx, dy) offsets in px.
 */
function polarOffset(distanceEmu: number, direction1_60000: number): [number, number] {
  const dist = emuToPx(distanceEmu);
  const rad = (direction1_60000 / 60000) * (Math.PI / 180);
  return [dist * Math.cos(rad), dist * Math.sin(rad)];
}

function colourToHex(color: ChartColor | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  if (color.srgb) {
    const s = color.srgb.length === 8 ? color.srgb.slice(2) : color.srgb;
    return `#${s}`;
  }
  return undefined;
}

function alphaFromColor(color: ChartColor | undefined): number {
  if (!color || color.alpha === undefined) {
    return 1;
  }
  // Excel alpha is 0..100000 where 100000 == fully opaque.
  return Math.max(0, Math.min(1, color.alpha / 100000));
}

function applyLineStyle(line: ChartSceneLine, axis: ChartAxis | undefined): ChartSceneLine {
  return {
    ...line,
    color: axisColor(axis),
    width: lineWidthFromShapeLine(axis?.spPr?.line) ?? line.width
  };
}

function renderSvgSeries(parts: string[], series: ChartSceneSeries): void {
  // When the series carries an `a:effectLst`, wrap its primary shapes
  // in a `<g filter="url(#...)">` so the DrawingML effect (shadow /
  // glow / reflection / soft-edge / blur / inner-shadow) applies
  // uniformly. Adornments (labels, markers, trendlines, error bars)
  // intentionally render outside the group so they don't inherit the
  // filter — matching Excel's convention where markers are drawn
  // sharp over a blurred series.
  const filterId = (series as ChartSceneAdornment).effectFilterId;
  if (filterId) {
    parts.push(`<g filter="url(#${filterId})">`);
  }
  if (series.type === "bar") {
    for (const bar of series.bars) {
      if (series.depth && series.projection3D) {
        // Proper extruded box: top face + front face + right face with
        // shading so the user can read the axonometric projection. The
        // back face is hidden behind the front rect by definition, so
        // we only paint three visible faces.
        parts.push(renderBar3DBox(bar, series.projection3D, series.color, series.horizontal));
      } else if (series.depth) {
        parts.push(renderBarDepth(bar, series.depth, withAlpha(series.color, 0.75)));
        parts.push(
          `<rect x="${fmt(bar.x)}" y="${fmt(bar.y)}" width="${fmt(bar.width)}" height="${fmt(bar.height)}" fill="${series.color}"/>`
        );
      } else {
        parts.push(
          `<rect x="${fmt(bar.x)}" y="${fmt(bar.y)}" width="${fmt(bar.width)}" height="${fmt(bar.height)}" fill="${series.color}"/>`
        );
      }
    }
  } else if (series.type === "area") {
    const points = [
      ...series.points,
      ...(series.lowerPoints ?? series.points.map(p => ({ x: p.x, y: series.baselineY })))
        .slice()
        .reverse()
    ];
    parts.push(
      `<polygon points="${points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="${withAlpha(series.color, 0.35)}"/>`
    );
    parts.push(
      `<polyline points="${series.points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="none" stroke="${series.color}" stroke-width="2"/>`
    );
  } else if (series.type === "line" || series.type === "scatter") {
    if (series.showLine !== false) {
      parts.push(
        `<polyline points="${series.points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="none" stroke="${series.color}" stroke-width="2"${series.smooth ? ' stroke-linejoin="round" stroke-linecap="round"' : ""}/>`
      );
    }
    for (const point of series.points) {
      parts.push(
        `<circle cx="${fmt(point.x)}" cy="${fmt(point.y)}" r="3" fill="${series.color}"/>`
      );
    }
  } else if (series.type === "bubble") {
    for (const bubble of series.bubbles) {
      parts.push(
        `<circle cx="${fmt(bubble.x)}" cy="${fmt(bubble.y)}" r="${fmt(bubble.radius)}" fill="${withAlpha(series.color, 0.55)}" stroke="${series.color}"/>`
      );
    }
  } else if (series.type === "pie" || series.type === "doughnut") {
    for (const slice of series.slices) {
      parts.push(renderSvgPieSlice(slice));
    }
  } else if (series.type === "ofPie") {
    for (const slice of series.slices) {
      parts.push(renderSvgPieSlice(slice));
    }
    for (const line of series.connectors ?? []) {
      parts.push(renderSvgLine(line));
    }
    for (const slice of series.secondarySlices ?? []) {
      parts.push(renderSvgPieSlice(slice));
    }
  } else if (series.type === "radar") {
    const points = series.points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ");
    parts.push(
      `<circle cx="${fmt(series.center.x)}" cy="${fmt(series.center.y)}" r="${fmt(series.radius)}" fill="none" stroke="${GRID_COLOR}"/>`
    );
    parts.push(
      `<polygon points="${points}" fill="${series.filled ? withAlpha(series.color, 0.35) : "none"}" stroke="${series.color}" stroke-width="2"/>`
    );
  } else if (series.type === "stock") {
    for (const candle of series.candles) {
      parts.push(
        `<line x1="${fmt(candle.x)}" y1="${fmt(candle.highY)}" x2="${fmt(candle.x)}" y2="${fmt(candle.lowY)}" stroke="#555"/>`
      );
      if (candle.openY !== undefined) {
        const y = Math.min(candle.openY, candle.closeY ?? candle.openY);
        const h = Math.max(1, Math.abs((candle.closeY ?? candle.openY) - candle.openY));
        parts.push(
          `<rect x="${fmt(candle.x - candle.width / 2)}" y="${fmt(y)}" width="${fmt(candle.width)}" height="${fmt(h)}" fill="${candle.up ? "#70AD47" : "#C00000"}" stroke="#555"/>`
        );
      }
    }
  } else if (series.type === "surface") {
    for (const cell of series.cells) {
      parts.push(
        `<rect x="${fmt(cell.x)}" y="${fmt(cell.y)}" width="${fmt(cell.width)}" height="${fmt(cell.height)}" fill="${cell.color}" stroke="${series.wireframe ? "#555" : cell.color}" stroke-width="${series.wireframe ? 1 : 0}"/>`
      );
    }
  }
  if (filterId) {
    parts.push(`</g>`);
  }
  renderSvgAdornments(parts, series);
}

function renderSvgAdornments(parts: string[], series: ChartSceneAdornment): void {
  for (const errorBar of series.errorBars ?? []) {
    parts.push(renderSvgLine(errorBar.line));
    if (errorBar.cap1) {
      parts.push(renderSvgLine(errorBar.cap1));
    }
    if (errorBar.cap2) {
      parts.push(renderSvgLine(errorBar.cap2));
    }
  }
  for (const trendline of series.trendlines ?? []) {
    parts.push(
      `<polyline points="${trendline.points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="none" stroke="${trendline.color}" stroke-width="${trendline.width ?? 1.5}"${trendline.dash ? ' stroke-dasharray="4 3"' : ""}/>`
    );
    if (trendline.label) {
      parts.push(renderSvgText(trendline.label));
    }
  }
  for (const marker of series.markers ?? []) {
    parts.push(renderSvgMarker(marker));
  }
  // Leader lines render *before* labels so the label glyphs visually sit on
  // top of the line rather than being bisected by it.
  for (const leader of series.leaderLines ?? []) {
    parts.push(renderSvgLine(leader));
  }
  for (const label of series.labels ?? []) {
    parts.push(renderSvgText(label));
  }
}

function renderSvgMarker(marker: ChartSceneMarker): string {
  const r = marker.size / 2;
  if (marker.symbol === "square") {
    return `<rect x="${fmt(marker.x - r)}" y="${fmt(marker.y - r)}" width="${fmt(marker.size)}" height="${fmt(marker.size)}" fill="${marker.color}"/>`;
  }
  if (marker.symbol === "diamond") {
    return `<polygon points="${fmt(marker.x)},${fmt(marker.y - r)} ${fmt(marker.x + r)},${fmt(marker.y)} ${fmt(marker.x)},${fmt(marker.y + r)} ${fmt(marker.x - r)},${fmt(marker.y)}" fill="${marker.color}"/>`;
  }
  if (marker.symbol === "triangle") {
    return `<polygon points="${fmt(marker.x)},${fmt(marker.y - r)} ${fmt(marker.x + r)},${fmt(marker.y + r)} ${fmt(marker.x - r)},${fmt(marker.y + r)}" fill="${marker.color}"/>`;
  }
  if (marker.symbol === "x" || marker.symbol === "plus") {
    const diagonal = marker.symbol === "x";
    const lines = diagonal
      ? [
          `<line x1="${fmt(marker.x - r)}" y1="${fmt(marker.y - r)}" x2="${fmt(marker.x + r)}" y2="${fmt(marker.y + r)}" stroke="${marker.color}" stroke-width="2"/>`,
          `<line x1="${fmt(marker.x + r)}" y1="${fmt(marker.y - r)}" x2="${fmt(marker.x - r)}" y2="${fmt(marker.y + r)}" stroke="${marker.color}" stroke-width="2"/>`
        ]
      : [
          `<line x1="${fmt(marker.x - r)}" y1="${fmt(marker.y)}" x2="${fmt(marker.x + r)}" y2="${fmt(marker.y)}" stroke="${marker.color}" stroke-width="2"/>`,
          `<line x1="${fmt(marker.x)}" y1="${fmt(marker.y - r)}" x2="${fmt(marker.x)}" y2="${fmt(marker.y + r)}" stroke="${marker.color}" stroke-width="2"/>`
        ];
    return lines.join("");
  }
  return `<circle cx="${fmt(marker.x)}" cy="${fmt(marker.y)}" r="${fmt(r)}" fill="${marker.color}"/>`;
}

function renderBarDepth(bar: ChartSceneRect, depth: number, color: string): string {
  const right = bar.x + bar.width;
  const top = bar.y;
  const bottom = bar.y + bar.height;
  return [
    `<polygon points="${fmt(bar.x)},${fmt(top)} ${fmt(bar.x + depth)},${fmt(top - depth)} ${fmt(right + depth)},${fmt(top - depth)} ${fmt(right)},${fmt(top)}" fill="${color}"/>`,
    `<polygon points="${fmt(right)},${fmt(top)} ${fmt(right + depth)},${fmt(top - depth)} ${fmt(right + depth)},${fmt(bottom - depth)} ${fmt(right)},${fmt(bottom)}" fill="${withAlpha(color, 0.85)}"/>`
  ].join("");
}

/**
 * Render a bar as an extruded 3D box using the supplied projection
 * deltas. Three visible faces:
 *
 *   1. **Top** — quadrilateral whose back edge sits at
 *      `(x+dx, y-dy)` / `(x+w+dx, y-dy)`. Painted lighter than the
 *      front so the viewer reads "light from above".
 *   2. **Right** — quadrilateral from front-right edge to back-right
 *      edge. Painted slightly darker than the front.
 *   3. **Front** — the original `bar` rectangle, unchanged.
 *
 * The shading uses the series' base color with fixed tints so plain
 * `#4472C4` (the default Office colour) ends up with a believable 3D
 * cue without loading a full lighting model.
 */
function renderBar3DBox(
  bar: ChartSceneRect,
  proj: { dx: number; dy: number },
  color: string,
  horizontal: boolean | undefined
): string {
  const right = bar.x + bar.width;
  const top = bar.y;
  const bottom = bar.y + bar.height;
  const dx = proj.dx;
  const dy = proj.dy;
  // Back corners.
  const bx1 = bar.x + dx;
  const by1 = top - dy;
  const bx2 = right + dx;
  const by2 = top - dy;
  const bxBottom = right + dx;
  const byBottom = bottom - dy;
  const topFill = withAlpha(color, 0.92);
  const rightFill = withAlpha(color, 0.75);
  const parts: string[] = [];
  // For horizontal bars we draw the same three faces but the "top" is
  // actually the top edge of the horizontal bar (still a parallelogram
  // along the depth vector).
  if (horizontal) {
    // Top edge + right (front-facing when horizontal) + front face.
    parts.push(
      `<polygon points="${fmt(bar.x)},${fmt(top)} ${fmt(bx1)},${fmt(by1)} ${fmt(bx2)},${fmt(by2)} ${fmt(right)},${fmt(top)}" fill="${topFill}"/>`,
      `<polygon points="${fmt(right)},${fmt(top)} ${fmt(bx2)},${fmt(by2)} ${fmt(bxBottom)},${fmt(byBottom)} ${fmt(right)},${fmt(bottom)}" fill="${rightFill}"/>`,
      `<rect x="${fmt(bar.x)}" y="${fmt(bar.y)}" width="${fmt(bar.width)}" height="${fmt(bar.height)}" fill="${color}"/>`
    );
  } else {
    parts.push(
      // Top face — front-top-left → back-top-left → back-top-right → front-top-right
      `<polygon points="${fmt(bar.x)},${fmt(top)} ${fmt(bx1)},${fmt(by1)} ${fmt(bx2)},${fmt(by2)} ${fmt(right)},${fmt(top)}" fill="${topFill}"/>`,
      // Right face — front-top-right → back-top-right → back-bottom-right → front-bottom-right
      `<polygon points="${fmt(right)},${fmt(top)} ${fmt(bx2)},${fmt(by2)} ${fmt(bxBottom)},${fmt(byBottom)} ${fmt(right)},${fmt(bottom)}" fill="${rightFill}"/>`,
      // Front face — the original rect.
      `<rect x="${fmt(bar.x)}" y="${fmt(bar.y)}" width="${fmt(bar.width)}" height="${fmt(bar.height)}" fill="${color}"/>`
    );
  }
  return parts.join("");
}

function renderSvgPieSlice(slice: ChartScenePieSlice): string {
  const large = slice.endAngle - slice.startAngle > Math.PI ? 1 : 0;
  const x1 = slice.cx + Math.cos(slice.startAngle) * slice.radius;
  const y1 = slice.cy + Math.sin(slice.startAngle) * slice.radius;
  const x2 = slice.cx + Math.cos(slice.endAngle) * slice.radius;
  const y2 = slice.cy + Math.sin(slice.endAngle) * slice.radius;
  if (slice.innerRadius > 0) {
    const ix1 = slice.cx + Math.cos(slice.endAngle) * slice.innerRadius;
    const iy1 = slice.cy + Math.sin(slice.endAngle) * slice.innerRadius;
    const ix2 = slice.cx + Math.cos(slice.startAngle) * slice.innerRadius;
    const iy2 = slice.cy + Math.sin(slice.startAngle) * slice.innerRadius;
    return `<path d="M ${fmt(x1)} ${fmt(y1)} A ${fmt(slice.radius)} ${fmt(slice.radius)} 0 ${large} 1 ${fmt(x2)} ${fmt(y2)} L ${fmt(ix1)} ${fmt(iy1)} A ${fmt(slice.innerRadius)} ${fmt(slice.innerRadius)} 0 ${large} 0 ${fmt(ix2)} ${fmt(iy2)} Z" fill="${slice.color}"/>`;
  }
  return `<path d="M ${fmt(slice.cx)} ${fmt(slice.cy)} L ${fmt(x1)} ${fmt(y1)} A ${fmt(slice.radius)} ${fmt(slice.radius)} 0 ${large} 1 ${fmt(x2)} ${fmt(y2)} Z" fill="${slice.color}"/>`;
}

function renderSvgDataTable(parts: string[], table: ChartSceneDataTable): void {
  // Borders first so cell text paints on top. Order matters for the
  // BasicRasterCanvas too — that canvas scans the SVG top-to-bottom.
  for (const line of table.borders) {
    parts.push(renderSvgLine(line));
  }
  for (const swatch of table.legendSwatches) {
    parts.push(
      `<rect x="${fmt(swatch.x)}" y="${fmt(swatch.y)}" width="${fmt(swatch.width)}" height="${fmt(swatch.height)}" fill="${swatch.color}"/>`
    );
  }
  for (const cell of table.cells) {
    parts.push(renderSvgText(cell));
  }
}

function renderSvgLegend(parts: string[], legend: ChartSceneLegend): void {
  if (!legend.visible || legend.items.length === 0) {
    return;
  }
  const fontSize = legend.textStyle?.fontSize ?? 10;
  const fontFamily = legend.textStyle?.fontFamily ?? "Arial";
  const color = legend.textStyle?.color ?? "#555";
  const weightAttr = legend.textStyle?.bold ? ' font-weight="bold"' : "";
  const styleAttr = legend.textStyle?.italic ? ' font-style="italic"' : "";
  legend.items.forEach((item, i) => {
    const itemX =
      legend.orientation === "horizontal"
        ? legend.rect.x + i * (legend.rect.width / legend.items.length)
        : legend.rect.x;
    const y = legend.orientation === "horizontal" ? legend.rect.y : legend.rect.y + i * 18;
    parts.push(
      `<rect x="${fmt(itemX)}" y="${fmt(y)}" width="10" height="10" fill="${item.color}"/>`
    );
    parts.push(
      `<text x="${fmt(itemX + 14)}" y="${fmt(y + 9)}" font-family="${escapeXmlAttr(fontFamily)}" font-size="${fontSize}" fill="${color}"${weightAttr}${styleAttr}>${escapeXml(item.label)}</text>`
    );
  });
}

function renderSvgLine(line: ChartSceneLine): string {
  return `<line x1="${fmt(line.x1)}" y1="${fmt(line.y1)}" x2="${fmt(line.x2)}" y2="${fmt(line.y2)}" stroke="${line.color}" stroke-width="${line.width ?? 1}"/>`;
}

function renderSvgText(text: ChartSceneText): string {
  const transform = text.rotate
    ? ` transform="rotate(${fmt(text.rotate)} ${fmt(text.x)} ${fmt(text.y)})"`
    : "";
  const fontFamily = text.fontFamily ?? "Arial";
  const weightAttr = text.bold ? ' font-weight="bold"' : "";
  const styleAttr = text.italic ? ' font-style="italic"' : "";
  return `<text x="${fmt(text.x)}" y="${fmt(text.y)}" text-anchor="${text.anchor ?? "start"}" font-family="${escapeXmlAttr(fontFamily)}" font-size="${text.fontSize}" fill="${text.color}"${weightAttr}${styleAttr}${transform}>${escapeXml(text.text)}</text>`;
}

function drawPdfSeries(page: ChartPdfDrawingSurface, series: ChartSceneSeries): void {
  if (series.type === "bar") {
    for (const bar of series.bars) {
      page.drawRect({ ...bar, fill: hexToPdfColor(series.color) });
    }
  } else if (series.type === "area") {
    if (page.drawPath && series.points.length > 0) {
      const lower = series.lowerPoints ?? series.points.map(p => ({ x: p.x, y: series.baselineY }));
      const ops: ChartPdfPathOp[] = [
        { op: "move", x: lower[0].x, y: lower[0].y },
        ...series.points.map(p => ({ op: "line" as const, x: p.x, y: p.y })),
        ...lower
          .slice()
          .reverse()
          .map(p => ({ op: "line" as const, x: p.x, y: p.y })),
        { op: "close" }
      ];
      // Match the SVG path's `withAlpha(color, 0.35)` fill so stacked
      // areas behind the current one remain visible through the
      // translucent polygon. Opaque fallback for surfaces that ignore
      // `PdfColor.a` — same degradation policy as everywhere else.
      page.drawPath(ops, { fill: hexToPdfColorWithAlpha(series.color, 0.35) });
    }
    for (let i = 1; i < series.points.length; i++) {
      page.drawLine({
        x1: series.points[i - 1].x,
        y1: series.points[i - 1].y,
        x2: series.points[i].x,
        y2: series.points[i].y,
        color: hexToPdfColor(series.color)
      });
    }
  } else if (series.type === "line" || series.type === "scatter") {
    for (let i = 1; i < series.points.length; i++) {
      page.drawLine({
        x1: series.points[i - 1].x,
        y1: series.points[i - 1].y,
        x2: series.points[i].x,
        y2: series.points[i].y,
        color: hexToPdfColor(series.color)
      });
    }
    for (const point of series.points) {
      page.drawRect({
        x: point.x - 2,
        y: point.y - 2,
        width: 4,
        height: 4,
        fill: hexToPdfColor(series.color)
      });
    }
  } else if (series.type === "bubble") {
    for (const bubble of series.bubbles) {
      // Bubble SVG uses `withAlpha(color, 0.55)`; emit real alpha in
      // PDF so overlapping bubbles remain distinguishable.
      const fill = hexToPdfColorWithAlpha(series.color, 0.55);
      if (page.drawCircle) {
        page.drawCircle({
          cx: bubble.x,
          cy: bubble.y,
          r: bubble.radius,
          fill
        });
      } else {
        page.drawRect({
          x: bubble.x - bubble.radius,
          y: bubble.y - bubble.radius,
          width: bubble.radius * 2,
          height: bubble.radius * 2,
          fill
        });
      }
    }
  } else if (series.type === "pie" || series.type === "doughnut") {
    if (page.drawPath) {
      for (const slice of series.slices) {
        page.drawPath(pieSliceToPath(slice), { fill: hexToPdfColor(slice.color) });
      }
    } else {
      // Surfaces without drawPath degrade to slice-outline strokes so the
      // chart is still recognisable. The filled area is lost (we have
      // no way to fill a polygon without drawPath), but the slice
      // boundaries and doughnut ring read clearly.
      for (const slice of series.slices) {
        strokePieSliceOutline(page, slice, hexToPdfColor(slice.color));
      }
    }
  } else if (series.type === "ofPie") {
    if (page.drawPath) {
      for (const slice of [...series.slices, ...(series.secondarySlices ?? [])]) {
        page.drawPath(pieSliceToPath(slice), { fill: hexToPdfColor(slice.color) });
      }
    } else {
      for (const slice of [...series.slices, ...(series.secondarySlices ?? [])]) {
        strokePieSliceOutline(page, slice, hexToPdfColor(slice.color));
      }
    }
    for (const line of series.connectors ?? []) {
      page.drawLine({ ...line, color: hexToPdfColor(line.color) });
    }
  } else if (series.type === "radar") {
    // Filled radar: the SVG path draws a `withAlpha(color, 0.35)` polygon
    // before the stroke loop. Mirror that here when `drawPath` is
    // available; when absent the stroke loop below alone preserves the
    // polygon shape, which is the best degradation a drawLine-only
    // surface can offer.
    if (series.filled && page.drawPath && series.points.length > 0) {
      const ops: ChartPdfPathOp[] = [
        { op: "move", x: series.points[0].x, y: series.points[0].y },
        ...series.points.slice(1).map(p => ({ op: "line" as const, x: p.x, y: p.y })),
        { op: "close" }
      ];
      page.drawPath(ops, { fill: hexToPdfColorWithAlpha(series.color, 0.35) });
    }
    for (let i = 0; i < series.points.length; i++) {
      const next = series.points[(i + 1) % series.points.length];
      page.drawLine({
        x1: series.points[i].x,
        y1: series.points[i].y,
        x2: next.x,
        y2: next.y,
        color: hexToPdfColor(series.color),
        lineWidth: 2
      });
    }
  } else if (series.type === "stock") {
    for (const candle of series.candles) {
      page.drawLine({
        x1: candle.x,
        y1: candle.highY,
        x2: candle.x,
        y2: candle.lowY,
        color: hexToPdfColor("#555555")
      });
      if (candle.openY !== undefined) {
        page.drawRect({
          x: candle.x - candle.width / 2,
          y: Math.min(candle.openY, candle.closeY ?? candle.openY),
          width: candle.width,
          height: Math.max(1, Math.abs((candle.closeY ?? candle.openY) - candle.openY)),
          fill: hexToPdfColor(candle.up ? "#70AD47" : "#C00000")
        });
      }
    }
  } else if (series.type === "surface") {
    for (const cell of series.cells) {
      page.drawRect({ ...cell, fill: hexToPdfColor(cell.color) });
    }
  }
}

function drawPdfDataTable(
  page: ChartPdfDrawingSurface,
  table: ChartSceneDataTable,
  trace: string[] | undefined
): void {
  for (const line of table.borders) {
    trace?.push(`line:dTable:${fmt(line.x1)},${fmt(line.y1)}-${fmt(line.x2)},${fmt(line.y2)}`);
    page.drawLine({ ...line, color: hexToPdfColor(line.color) });
  }
  for (const swatch of table.legendSwatches) {
    page.drawRect({
      x: swatch.x,
      y: swatch.y,
      width: swatch.width,
      height: swatch.height,
      fill: hexToPdfColor(swatch.color)
    });
  }
  for (const cell of table.cells) {
    trace?.push(`text:dTable:${cell.text}:${fmt(cell.x)},${fmt(cell.y)}`);
    drawPdfText(page, cell);
  }
}

function drawPdfLegend(page: ChartPdfDrawingSurface, legend: ChartSceneLegend): void {
  if (!legend.visible || legend.items.length === 0) {
    return;
  }
  const legendFontSize = legend.textStyle?.fontSize ?? 10;
  const fontFamily = legend.textStyle?.fontFamily;
  const bold = legend.textStyle?.bold;
  const italic = legend.textStyle?.italic;
  const textColor = legend.textStyle?.color ? hexToPdfColor(legend.textStyle.color) : undefined;
  const swatchSize = 10;
  const swatchToLabelGap = 4;
  const interItemGap = 16;
  // Walk items left-to-right (horizontal) / top-to-bottom (vertical),
  // measuring real label widths so long names do not collide with the
  // next swatch. This mirrors the SVG emit path (see renderSvgLegend).
  let cursorX = legend.rect.x;
  legend.items.forEach((item, i) => {
    const swatchX = legend.orientation === "horizontal" ? cursorX : legend.rect.x;
    const y = legend.orientation === "horizontal" ? legend.rect.y : legend.rect.y + i * 18;
    page.drawRect({
      x: swatchX,
      y,
      width: swatchSize,
      height: swatchSize,
      fill: hexToPdfColor(item.color)
    });
    page.drawText(item.label, {
      x: swatchX + swatchSize + swatchToLabelGap,
      y: y + 1,
      fontSize: legendFontSize,
      anchor: "start",
      fontFamily,
      bold,
      italic,
      color: textColor
    });
    if (legend.orientation === "horizontal") {
      const labelWidth = estimateTextWidth(item.label, legendFontSize, {
        bold,
        italic,
        fontName: fontFamily
      });
      cursorX += swatchSize + swatchToLabelGap + labelWidth + interItemGap;
    }
  });
}

function translateScene(
  scene: ChartScene,
  offsetX: number,
  offsetY: number,
  flipY: boolean
): ChartScene {
  const convertY = (y: number) => (flipY ? offsetY + scene.height - y : offsetY + y);
  const mapPoint = (p: ChartScenePoint): ChartScenePoint => ({
    x: offsetX + p.x,
    y: convertY(p.y)
  });
  const mapRect = (r: ChartSceneRect): ChartSceneRect => {
    const y = flipY ? offsetY + scene.height - r.y - r.height : offsetY + r.y;
    return { x: offsetX + r.x, y, width: r.width, height: r.height };
  };
  const mapLine = (line: ChartSceneLine): ChartSceneLine => ({
    ...line,
    x1: offsetX + line.x1,
    y1: convertY(line.y1),
    x2: offsetX + line.x2,
    y2: convertY(line.y2)
  });
  const mapText = (text: ChartSceneText): ChartSceneText => ({
    ...text,
    x: offsetX + text.x,
    y: convertY(text.y)
  });
  return {
    ...scene,
    title: scene.title ? mapText(scene.title) : undefined,
    plot: mapRect(scene.plot),
    axes: {
      x: mapLine(scene.axes.x),
      y: mapLine(scene.axes.y),
      x2: scene.axes.x2 ? mapLine(scene.axes.x2) : undefined,
      y2: scene.axes.y2 ? mapLine(scene.axes.y2) : undefined
    },
    gridlines: scene.gridlines.map(mapLine),
    xLabels: scene.xLabels.map(mapText),
    yLabels: scene.yLabels.map(mapText),
    secondaryXLabels: scene.secondaryXLabels.map(mapText),
    secondaryYLabels: scene.secondaryYLabels.map(mapText),
    axisTitles: scene.axisTitles.map(mapText),
    legend: { ...scene.legend, rect: mapRect(scene.legend.rect) },
    dataTable: scene.dataTable
      ? {
          rect: mapRect(scene.dataTable.rect),
          // Column/row arrays are plain x/y scalars — columns is an array
          // of x (no flip) and rows is an array of y (flip-aware via
          // convertY). We remap them so downstream consumers can still
          // use them in the translated frame.
          columns: scene.dataTable.columns.map(x => offsetX + x),
          rows: scene.dataTable.rows.map(y => convertY(y)),
          cells: scene.dataTable.cells.map(mapText),
          legendSwatches: scene.dataTable.legendSwatches.map(s => ({
            ...mapRect(s),
            color: s.color
          })),
          borders: scene.dataTable.borders.map(mapLine)
        }
      : undefined,
    series: scene.series.map(s => translateSeries(s, mapPoint, mapRect, mapLine, mapText))
  };
}

function translateSeries(
  series: ChartSceneSeries,
  mapPoint: (point: ChartScenePoint) => ChartScenePoint,
  mapRect: (rect: ChartSceneRect) => ChartSceneRect,
  mapLine: (line: ChartSceneLine) => ChartSceneLine,
  mapText: (text: ChartSceneText) => ChartSceneText
): ChartSceneSeries {
  if (series.type === "bar") {
    return translateAdornments(
      { ...series, bars: series.bars.map(mapRect) },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "line" || series.type === "scatter") {
    return translateAdornments(
      { ...series, points: series.points.map(mapPoint) },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "radar") {
    return translateAdornments(
      {
        ...series,
        points: series.points.map(mapPoint),
        center: mapPoint(series.center)
      },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "area") {
    return translateAdornments(
      {
        ...series,
        points: series.points.map(mapPoint),
        lowerPoints: series.lowerPoints?.map(mapPoint),
        baselineY: mapPoint({ x: 0, y: series.baselineY }).y
      },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "bubble") {
    return translateAdornments(
      { ...series, bubbles: series.bubbles.map(b => ({ ...mapPoint(b), radius: b.radius })) },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "pie" || series.type === "doughnut") {
    return translateAdornments(
      { ...series, slices: series.slices.map(slice => translatePieSlice(slice, mapPoint)) },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "ofPie") {
    return translateAdornments(
      {
        ...series,
        slices: series.slices.map(slice => translatePieSlice(slice, mapPoint)),
        secondarySlices: series.secondarySlices?.map(slice => translatePieSlice(slice, mapPoint)),
        connectors: series.connectors?.map(line => ({
          ...line,
          ...lineFromPoints(
            mapPoint({ x: line.x1, y: line.y1 }),
            mapPoint({ x: line.x2, y: line.y2 })
          )
        }))
      },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "stock") {
    return translateAdornments(
      {
        ...series,
        candles: series.candles.map(c => ({
          ...c,
          x: mapPoint({ x: c.x, y: 0 }).x,
          highY: mapPoint({ x: 0, y: c.highY }).y,
          lowY: mapPoint({ x: 0, y: c.lowY }).y,
          openY: c.openY === undefined ? undefined : mapPoint({ x: 0, y: c.openY }).y,
          closeY: c.closeY === undefined ? undefined : mapPoint({ x: 0, y: c.closeY }).y
        }))
      },
      mapPoint,
      mapLine,
      mapText
    );
  }
  if (series.type === "surface") {
    return translateAdornments(
      { ...series, cells: series.cells.map(cell => ({ ...cell, ...mapRect(cell) })) },
      mapPoint,
      mapLine,
      mapText
    );
  }
  return series;
}

function translateAdornments<T extends ChartSceneSeries>(
  series: T,
  mapPoint: (point: ChartScenePoint) => ChartScenePoint,
  mapLine: (line: ChartSceneLine) => ChartSceneLine,
  mapText: (text: ChartSceneText) => ChartSceneText
): T {
  return {
    ...series,
    labels: series.labels?.map(mapText),
    markers: series.markers?.map(m => ({ ...m, ...mapPoint(m) })),
    trendlines: series.trendlines?.map(t => ({
      ...t,
      points: t.points.map(mapPoint),
      label: t.label ? mapText(t.label) : undefined
    })),
    errorBars: series.errorBars?.map(e => ({
      line: mapLine(e.line),
      cap1: e.cap1 ? mapLine(e.cap1) : undefined,
      cap2: e.cap2 ? mapLine(e.cap2) : undefined
    })),
    leaderLines: series.leaderLines?.map(mapLine)
  };
}

function translatePieSlice(
  slice: ChartScenePieSlice,
  mapPoint: (point: ChartScenePoint) => ChartScenePoint
): ChartScenePieSlice {
  const center = mapPoint({ x: slice.cx, y: slice.cy });
  return {
    ...slice,
    cx: center.x,
    cy: center.y,
    startAngle: -slice.endAngle,
    endAngle: -slice.startAngle
  };
}

function lineFromPoints(
  p1: ChartScenePoint,
  p2: ChartScenePoint
): Pick<ChartSceneLine, "x1" | "y1" | "x2" | "y2"> {
  return { x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y };
}

/**
 * Stroke the outline of a pie / doughnut slice with `drawLine` calls.
 *
 * Used by {@link drawPdfSeries} as a fallback when the target surface
 * does not implement `drawPath` — without it, pie charts would be
 * completely invisible on minimal surfaces. The outline uses the same
 * polyline approximation of the arc (`arcPolyline`) that
 * {@link pieSliceToPath} builds for the filled version, so the shape
 * matches pixel-for-pixel modulo fill/no-fill.
 */
function strokePieSliceOutline(
  page: ChartPdfDrawingSurface,
  slice: ChartScenePieSlice,
  color: PdfColor
): void {
  const ops = pieSliceToPath(slice);
  let last: { x: number; y: number } | undefined;
  let first: { x: number; y: number } | undefined;
  for (const op of ops) {
    if (op.op === "move") {
      last = { x: op.x, y: op.y };
      first = last;
    } else if (op.op === "line" && last) {
      page.drawLine({ x1: last.x, y1: last.y, x2: op.x, y2: op.y, color });
      last = { x: op.x, y: op.y };
    } else if (op.op === "close" && last && first) {
      page.drawLine({ x1: last.x, y1: last.y, x2: first.x, y2: first.y, color });
      last = undefined;
      first = undefined;
    }
  }
}

function pieSliceToPath(slice: ChartScenePieSlice): ChartPdfPathOp[] {
  const ops: ChartPdfPathOp[] = [];
  const startOuter = polarPoint(slice.cx, slice.cy, slice.radius, slice.startAngle);
  const endOuter = polarPoint(slice.cx, slice.cy, slice.radius, slice.endAngle);
  if (slice.innerRadius > 0) {
    const endInner = polarPoint(slice.cx, slice.cy, slice.innerRadius, slice.endAngle);
    ops.push({ op: "move", ...startOuter });
    ops.push(...arcPolyline(slice.cx, slice.cy, slice.radius, slice.startAngle, slice.endAngle));
    ops.push({ op: "line", ...endInner });
    ops.push(
      ...arcPolyline(slice.cx, slice.cy, slice.innerRadius, slice.endAngle, slice.startAngle)
    );
    ops.push({ op: "close" });
    return ops;
  }
  ops.push({ op: "move", x: slice.cx, y: slice.cy });
  ops.push({ op: "line", ...startOuter });
  ops.push(...arcPolyline(slice.cx, slice.cy, slice.radius, slice.startAngle, slice.endAngle));
  ops.push({ op: "line", ...endOuter });
  ops.push({ op: "close" });
  return ops;
}

function arcPolyline(
  cx: number,
  cy: number,
  radius: number,
  startAngle: number,
  endAngle: number
): ChartPdfPathOp[] {
  const sweep = endAngle - startAngle;
  const segments = Math.max(2, Math.ceil(Math.abs(sweep) / (Math.PI / 12)));
  const ops: ChartPdfPathOp[] = [];
  for (let i = 1; i <= segments; i++) {
    ops.push({ op: "line", ...polarPoint(cx, cy, radius, startAngle + (sweep * i) / segments) });
  }
  return ops;
}

function polarPoint(cx: number, cy: number, radius: number, angle: number): ChartScenePoint {
  return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
}

function collectSeries(group: ChartTypeGroup | undefined): SeriesBase[] {
  return ((group as { series?: SeriesBase[] } | undefined)?.series ?? []) as SeriesBase[];
}

function collectValues(series: SeriesBase): number[] {
  const ref =
    (series as { val?: any; yVal?: any }).val?.numRef ?? (series as { yVal?: any }).yVal?.numRef;
  return collectNumberValues(ref);
}

function collectNumberValues(ref: any): number[] {
  return (ref?.cache?.points ?? []).map((p: { value: number | null }) =>
    typeof p.value === "number" ? p.value : 0
  );
}

function collectAxisValues(axisData: any): number[] {
  return collectNumberValues(axisData?.numRef);
}

function collectCategories(series: SeriesBase | undefined): string[] | undefined {
  const cat = (series as { cat?: any } | undefined)?.cat;
  const ref = cat?.strRef ?? cat?.multiLvlStrRef?.cache?.levels?.[0];
  return (
    ref?.cache?.points?.map((p: { value: string }) => p.value) ??
    ref?.points?.map((p: { value: string }) => p.value)
  );
}

function collectSeriesLabel(series: SeriesBase | undefined, index: number): string {
  const tx = (
    series as
      | { tx?: { value?: string; strRef?: { cache?: { points?: Array<{ value: string }> } } } }
      | undefined
  )?.tx;
  return (
    tx?.value ?? tx?.strRef?.cache?.points?.map(p => p.value).join("") ?? `Series ${index + 1}`
  );
}

function extractTitle(model: ChartModel): string | undefined {
  return model.chart.title?.text?.paragraphs
    .map(p => (p.runs ?? []).map(r => r.text).join(""))
    .join("\n");
}

function formatAxisNumber(value: number): string {
  if (Math.abs(value) >= 1000) {
    return value.toFixed(0);
  }
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function truncateLabel(label: string): string {
  return label.length > 12 ? label.slice(0, 11) + "..." : label;
}

/**
 * Measure the rendered pixel width of a label using the Excel module's
 * built-in font metrics engine (`@excel/utils/text-metrics`). This replaces
 * the previous `text.length * fontSize * 0.55` approximation with a real
 * per-character advance-width lookup, which matters wherever the renderer
 * must allocate space for a label (legend width, title centring, PDF
 * anchoring). Chart labels default to Arial to match the SVG emit path
 * (`font-family="Arial"` in renderSvgText), and `bold`/`italic` are
 * forwarded when present — the Excel engine already knows how to degrade
 * to a category-average factor for unrecognised faces.
 */
function estimateTextWidth(
  text: string,
  fontSize: number,
  options: { bold?: boolean; italic?: boolean; fontName?: string } = {}
): number {
  if (!text) {
    return 0;
  }
  return measureTextWidthPx(text, {
    name: options.fontName ?? "arial",
    size: fontSize,
    bold: options.bold,
    italic: options.italic
  });
}

function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeXmlAttr(value: string): string {
  return escapeXml(value).replace(/"/g, "&quot;");
}

function hexToPdfColor(hex: string): PdfColor {
  const clean = hex.replace(/^#/, "");
  return {
    r: parseInt(clean.slice(0, 2), 16) / 255,
    g: parseInt(clean.slice(2, 4), 16) / 255,
    b: parseInt(clean.slice(4, 6), 16) / 255
  };
}

/**
 * Like {@link hexToPdfColor} but attaches an alpha value. Callers use
 * this to mirror the SVG path's `withAlpha(color, 0.35)` pattern on
 * the PDF bridge: the hex itself stays opaque, `a` carries the
 * transparency the SVG would paint by white-blending. Surfaces that
 * honour `PdfColor.a` (notably `PdfPageBuilder` via `/ExtGState`)
 * produce real transparency; those that don't render opaque, which is
 * the pre-alpha behaviour.
 */
function hexToPdfColorWithAlpha(hex: string, alpha: number): PdfColor {
  return { ...hexToPdfColor(hex), a: Math.max(0, Math.min(1, alpha)) };
}

function interpolateColor(a: string, b: string, t: number): string {
  const ca = a.replace(/^#/, "");
  const cb = b.replace(/^#/, "");
  const mix = (i: number) => {
    const av = parseInt(ca.slice(i, i + 2), 16);
    const bv = parseInt(cb.slice(i, i + 2), 16);
    return Math.round(av + (bv - av) * t)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${mix(0)}${mix(2)}${mix(4)}`;
}

function withAlpha(hex: string, alpha: number): string {
  const clean = hex.replace(/^#/, "");
  const mix = (component: string) => {
    const value = parseInt(component, 16);
    return Math.round(value * alpha + 255 * (1 - alpha))
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${mix(clean.slice(0, 2))}${mix(clean.slice(2, 4))}${mix(clean.slice(4, 6))}`;
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to load chart SVG image"));
    image.src = url;
  });
}

function fmt(value: number): string {
  return value.toFixed(2).replace(/\.00$/, "");
}
