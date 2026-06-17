/**
 * ChartEx (Office 2016+ extended charts) type definitions.
 *
 * These types mirror the ECMA-376 Part 1, §21.3 `cx:` namespace for the eight
 * modern chart types: sunburst, treemap, waterfall, funnel, histogram
 * (and Pareto), box-whisker, and region map.
 *
 * Classic charts live in `c:` (ChartML); chartEx uses its own `cx:` schema
 * with shared formatting primitives (spPr/txPr) but very different data model.
 */

import type {
  AddShapeFillOptions,
  ChartColorsModel,
  ChartLayout,
  ChartLegend,
  ChartRelEntry,
  ChartRichText,
  ChartStyleModel,
  ChartTextProperties,
  ChartTitle,
  ShapeProperties
} from "./types";

// ============================================================================
// Top-level chartEx model
// ============================================================================

/**
 * A chartEx chart (one `chartEx{N}.xml` file).
 */
export interface ChartExModel {
  /** Root chart space containing data, series, and plot area */
  chartSpace: ChartExSpace;
  /** Raw XML (for passthrough of existing cx:chart files) */
  rawXml?: string;
  /** Chart rels — preserved for round-trip */
  rels?: unknown[];
  /** Referenced external workbook/package parts used by cx:externalData relationships. */
  externalParts?: Record<string, Uint8Array>;
  /** Optional structured style sidecar metadata for chartEx consumers. */
  style?: ChartStyleModel;
  /** Optional structured colors sidecar metadata for chartEx consumers. */
  colors?: ChartColorsModel;
  /**
   * Child elements the parser did not recognise at well-known locations.
   *
   * Populated by {@link parseChartEx} when walking a loaded chartEx part. Each
   * entry records a `parent/child` breadcrumb so `strict` template mode can
   * refuse to silently drop extension/vendor XML when the structured writer
   * has to rebuild the part (i.e. no raw XML passthrough and no safe raw
   * patch available). Purely informational in the default `preserve` mode.
   */
  unknownElements?: ChartExUnknownElement[];
}

/**
 * Describes one unstructured child element discovered while parsing a
 * chartEx part. `path` uses `/` as the separator and is relative to the
 * `cx:chartSpace` root.
 */
export interface ChartExUnknownElement {
  /** Fully-qualified element name (e.g. `cx:unknownTag`). */
  name: string;
  /** Slash-separated breadcrumb, e.g. `cx:chartSpace/cx:chart/cx:unknownTag`. */
  path: string;
}

/**
 * `cx:chartSpace` — root element.
 */
export interface ChartExSpace {
  /** Chart data (each entry is one data item for series to reference) */
  chartData: ChartExData;
  /** Chart element (title, plotArea, legend) */
  chart: ChartExChart;
  /** Clr map override (optional) */
  clrMapOvr?: string;
  /** ChartSpace-level shape properties (frame / background). */
  spPr?: ShapeProperties;
  /** ChartSpace-level default text properties. */
  txPr?: ChartTextProperties;
  /**
   * `cx:protection` — raw XML preserved for round-trip. The structured
   * model does not currently expose locking flags; consumers who need
   * to edit protection should mutate the raw string.
   */
  protection?: string;
  /**
   * `cx:externalData` — references to external workbook / package
   * parts. Per Chart2014 / `CT_ChartSpace` schema this is a child of
   * `cx:chartSpace` (alongside `cx:chart`, `cx:spPr`, …), NOT a
   * child of `cx:chartData`. Previous versions of this library placed
   * it inside `cx:chartData` — the parser now accepts both locations
   * (legacy on-disk shapes still load) and the writer always emits
   * at the chartSpace level.
   */
  externalData?: Array<{ id: string; autoUpdate?: boolean }>;
  /**
   * `cx:printSettings` — raw XML preserved for round-trip (headerFooter,
   * pageMargins, pageSetup). Same "raw-only" contract as `protection`.
   */
  printSettings?: string;
  /** Extension list */
  extLst?: string;
}

export interface ChartExData {
  /** Chart data (numeric/string arrays referenced by series) */
  data: ChartExDataEntry[];
}

/**
 * A `cx:data` entry, representing a column of values that one or more series
 * reference by `id`.
 */
export interface ChartExDataEntry {
  /** 0-based id, referenced by `cx:dataId` in series */
  id: number;
  /** String dimension (typically categories) */
  strDim?: ChartExStringDimension;
  /** Numeric dimension (typically values) */
  numDim?: ChartExNumericDimension;
}

export interface ChartExStringDimension {
  /** Dimension type attribute: "cat" / "val" / "x" / "y" / "size" / etc. */
  type: ChartExDimensionType;
  /** Formula reference (e.g. "Sheet1!$A$1:$A$10") */
  formula?: string;
  /** Level array (each level is one row of string points) */
  levels?: Array<{ ptCount?: number; points: Array<{ index: number; value: string }> }>;
  /**
   * @internal Skip `fillChartExCaches` population. Set on hierarchical
   * (treemap / sunburst) `<cx:strDim>` entries whose `<cx:f>` points to
   * a contiguous multi-column range — Excel expects NO `<cx:lvl>` cache
   * on those dimensions and re-reads the cells on open. Caching the
   * flattened point list into a single `<cx:lvl>` makes the hierarchy
   * renderer paint an empty plot area.
   */
  _skipCache?: boolean;
}

export interface ChartExNumericDimension {
  /** Dimension type */
  type: ChartExDimensionType;
  /** Formula reference */
  formula?: string;
  /** Cached points */
  levels?: Array<{
    ptCount?: number;
    formatCode?: string;
    points: Array<{ index: number; value: number }>;
  }>;
  /** @internal See `ChartExStringDimension._skipCache`. */
  _skipCache?: boolean;
}

export type ChartExDimensionType =
  | "cat"
  | "val"
  | "x"
  | "y"
  | "size"
  | "colorVal"
  | "from"
  | "to"
  | "classification";

// ============================================================================
// Chart element
// ============================================================================

export interface ChartExChart {
  title?: ChartTitle;
  plotArea: ChartExPlotArea;
  legend?: ChartLegend;
  /** Whether the chart has an automatic title */
  autoTitleDeleted?: boolean;
  /**
   * Extension list — raw XML preserved for round-trip of future
   * `CT_Chart` extensions (e.g. `c15:` annotations).
   */
  extLst?: string;
}

export interface ChartExPlotArea {
  /** Plot area region (layout + spPr) */
  plotAreaRegion?: {
    layout?: ChartLayout;
    plotSurface?: ShapeProperties;
    /** Plot area series */
    series: ChartExSeries[];
    /**
     * `<cx:extLst>` inside `<cx:plotAreaRegion>` — preserved verbatim
     * for round-trip of Chart2014 extension blocks (e.g. the
     * `cx14:` markers Excel writes for pivot-backed charts). Parsed
     * and re-emitted as raw XML; the library never synthesises this
     * field.
     */
    extLst?: string;
  };
  /** Axes (value/category/etc.) */
  axis?: ChartExAxis[];
  /** Direct access to series (short-hand if no plotAreaRegion) */
  series?: ChartExSeries[];
  spPr?: ShapeProperties;
  /**
   * `<cx:extLst>` on the plot area itself (sibling of
   * `plotAreaRegion`). Preserved for round-trip.
   */
  extLst?: string;
}

// ============================================================================
// Series
// ============================================================================

export type ChartExSeriesType =
  | "sunburst"
  | "treemap"
  | "waterfall"
  | "funnel"
  | "clusteredColumn"
  | "boxWhisker"
  | "paretoLine"
  | "regionMap";

export interface ChartExSeries {
  layoutId: ChartExSeriesType;
  /**
   * Original `@layoutId` attribute as it appeared in the source XML,
   * when the value did not match any of the {@link ChartExSeriesType}
   * enum members. The parser falls back to `"clusteredColumn"` so the
   * renderer has a shape it understands, but the writer re-emits
   * `rawLayoutId` verbatim when present so round-trips through
   * Excel-authored files don't lose a future / vendor-extended
   * layoutId the consumer never asked us to interpret.
   *
   * The renderer emits `rawLayoutId` only when `layoutId` is still
   * the neutral `"clusteredColumn"` fallback; assigning any other
   * structured layoutId causes the writer to use the new value and
   * ignore `rawLayoutId`. This means a caller who wants to explicitly
   * downgrade a vendor-extended series to the canonical
   * `"clusteredColumn"` layout — and stop preserving the original
   * raw attribute — must also clear `rawLayoutId` directly
   * (`series.rawLayoutId = undefined`). Plain builders never set
   * this field, so freshly created series always emit the canonical
   * enum form.
   * @internal
   */
  rawLayoutId?: string;
  /**
   * Display name (tx) — string literal, formula reference (strRef), or
   * a structured rich-text block. `rich` is a {@link ChartRichText};
   * when set, the writer emits `<cx:rich>…</cx:rich>` with one or more
   * `<a:p>` paragraphs, matching what Excel produces for bold/coloured
   * series labels.
   *
   * `strRef` mirrors the classic {@link StringReference} shape: the
   * raw formula plus an optional cached resolved value. Previous
   * versions stored only the formula string; that form is still
   * accepted via the `string` alternate, but round-trip of Excel-
   * authored files now preserves the `<cx:v>` cached label so
   * re-opens display the series name without recalculation.
   */
  tx?: {
    rich?: ChartRichText;
    strRef?: string | { formula: string; cached?: string };
    value?: string;
  };
  /** Hidden attribute */
  hidden?: boolean;
  /** Ownership — "primary" or "standard" */
  ownerIdx?: number;
  /** Shape properties */
  spPr?: ShapeProperties;
  /**
   * Text properties (`<cx:txPr>`). Per `CT_Series` in the Chart2014
   * schema, series can carry their own DrawingML text body properties
   * which propagate to all series-level text (axis tick labels for the
   * series, data-label defaults, legend entry text) unless overridden
   * by a more specific `txPr` further down the tree.
   */
  txPr?: ChartTextProperties;
  /** Data references — each series references one or more cx:data entries */
  dataPt?: Array<{ idx: number; spPr?: ShapeProperties }>;
  /** Data labels */
  dataLabels?: ChartExDataLabels;
  /** Data refs — each { axis, dataId } */
  dataRefs?: Array<{ /* cx:axisId or cx:dataId */ dataId?: number; axisId?: number }>;
  /** Layout properties (per-layoutId settings) */
  layoutPr?: ChartExLayoutProperties;
  /** Axis bindings */
  axisId?: number[];
  /**
   * `<cx:valueColors>` — raw XML preserved for round-trip. This
   * sub-element carries a colour-by-value palette (gradient stops keyed
   * to value buckets) used by region-map and treemap charts. The
   * structured model does not yet interpret the stops.
   */
  valueColors?: string;
  /**
   * `<cx:valueColorPositions>` — raw XML preserved for round-trip.
   * Companion to {@link valueColors}: lists the value/position pairs
   * the palette maps onto.
   */
  valueColorPositions?: string;
  /** Extension list */
  extLst?: string;
}

export interface ChartExDataLabels {
  visibility?: {
    seriesName?: boolean;
    categoryName?: boolean;
    value?: boolean;
    numFmt?: boolean;
  };
  position?: string;
  separator?: string;
  numFmt?: string;
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
}

/**
 * Per-series layout properties. Only fields relevant to the series type are used.
 */
export interface ChartExLayoutProperties {
  /** Raw cx:layoutPr XML preserved from loaded files. */
  _rawXml?: string;
  // Sunburst / Treemap
  parentLabelLayout?: "banner" | "overlapping" | "none";
  // Waterfall
  subtotals?: Array<{ idx: number }>;
  connectorLines?: boolean;
  /**
   * Preview-only colour overrides for the three waterfall bar kinds.
   * These are NOT part of the Chart2014 schema — Excel stores per-bar
   * colours on `<cx:dataPt>` elements referenced by `subtotal` indices.
   * Setting one affects only the SVG/PDF renderer; the XML writer
   * does not emit them, so round-trip of a waterfall chart relies on
   * `<cx:dataPt>` colours authored by Excel being preserved by the
   * series-level `dataPt` path.
   */
  increaseSpPr?: ShapeProperties;
  decreaseSpPr?: ShapeProperties;
  totalSpPr?: ShapeProperties;
  // Funnel
  // (funnel has no extra layout props at this level — uses plot area)
  // Histogram / Pareto
  binning?: {
    binSize?: number;
    binCount?: number;
    binType?: "auto" | "binCount" | "binSize" | "categories" | "manual";
    intervalClosed?: "l" | "r";
    underflow?: number;
    overflow?: number;
  };
  /** Pareto charts render a cumulative line over histogram columns. */
  paretoLine?: boolean;
  // BoxWhisker
  quartileMethod?: "inclusive" | "exclusive";
  showMeanLine?: boolean;
  showMeanMarker?: boolean;
  showInnerPoints?: boolean;
  showOutlierPoints?: boolean;
  // RegionMap
  projection?: "mercator" | "albers" | "miller" | "robinson";
  regionLabels?: "none" | "bestFit" | "showAll";
  geoMappingLevel?: "automatic" | "country" | "state" | "county" | "postalCode";
  /** Extension list */
  extLst?: string;
}

// ============================================================================
// Axis
// ============================================================================

export interface ChartExAxis {
  axisId: number;
  /** Category or value axis */
  type: "cat" | "val";
  /** Axis title */
  title?: ChartTitle;
  /** Major tick mark */
  majorTickMark?: "none" | "inside" | "outside" | "cross";
  /** Minor tick mark */
  minorTickMark?: "none" | "inside" | "outside" | "cross";
  /**
   * `<cx:majorGridlines>` — styled gridlines drawn at each major
   * tick. Per Chart2014 `CT_Axis`, the element wraps a single
   * `<cx:spPr>` child; `undefined` means "no major gridlines" while
   * an empty object means "default-styled gridlines". Previously
   * this field was absent from the type, and Excel-authored charts
   * with styled gridlines had them dropped on round-trip.
   */
  majorGridlines?: ShapeProperties;
  /** `<cx:minorGridlines>` — same semantics at minor tick positions. */
  minorGridlines?: ShapeProperties;
  /**
   * `<cx:tickLabels>` — tick-label rendering flag. Excel emits an
   * empty `<cx:tickLabels/>` on every axis by default; without it
   * the tick labels are suppressed entirely on load. Treated the
   * same way as `majorGridlines`: `undefined` means "omit",
   * presence (even as an empty object) means "emit `<cx:tickLabels/>`".
   */
  tickLabels?: Record<string, never> | { rawXml?: string };
  /** Number format */
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  /** Shape properties */
  spPr?: ShapeProperties;
  /** Text properties */
  txPr?: ChartTextProperties;
  /** Hidden */
  hidden?: boolean;
  /** Category axis scaling */
  catScaling?: { gapWidth?: number };
  /** Value axis scaling — `min` / `max` / `majorUnit` / `minorUnit`. */
  valScaling?: { max?: number; min?: number; majorUnit?: number; minorUnit?: number };
  /**
   * `<cx:units>` — display-unit scaling for value axes (thousand /
   * million / custom). Stored as the raw XML slice so it round-trips
   * verbatim; the structured model doesn't yet interpret it, but the
   * writer must not silently drop this element.
   */
  units?: string;
  /** Extension list */
  extLst?: string;
}

// ============================================================================
// High-level (addChartEx) options
// ============================================================================

/**
 * Top-level options for creating a chartEx programmatically.
 */
export interface AddChartExOptions {
  /** Chart type */
  type: ChartExType;
  /** Series configuration */
  series: AddChartExSeriesOptions[];
  /** Category values reference */
  categories?: string;
  /**
   * Chart title. Accepts the same three forms as classic charts:
   *
   *   - `string` — plain title text
   *   - `{ formula: "Sheet1!$A$1" }` — a worksheet formula reference; the
   *     rendered title is the live cell value
   *   - {@link ChartRichText} — fully-structured rich text with per-run
   *     formatting (colour / font / bold / italic)
   *
   * Pass `null` to explicitly suppress the title (sets
   * `autoTitleDeleted="1"` so Excel will NOT auto-generate a
   * single-series title). Omit the option entirely to let Excel
   * decide whether to auto-title the chart — this is the default
   * behaviour and produces the same output Excel would emit for a
   * chart authored with no title via its UI.
   */
  title?: string | { formula: string } | ChartRichText | null;
  /** Show legend */
  showLegend?: boolean;
  /** Legend position */
  legendPosition?: "b" | "l" | "r" | "t" | "tr";
  /** Type-specific layout overrides */
  layout?: ChartExLayoutProperties;
  /** Histogram/pareto binning shortcut; merged into `layout.binning`. */
  binning?: ChartExLayoutProperties["binning"];
  /**
   * ChartEx chart frame styling. Accepts either a fully-structured
   * {@link ShapeProperties} (the OOXML-shape representation used
   * throughout the chart module) or the ergonomic
   * {@link AddShapeFillOptions} bag with hex colour / border / gradient /
   * pattern shortcuts — matching the `floor` / `sideWall` / `backWall`
   * options on classic `AddChartOptions` so the two APIs stay symmetric.
   */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Optional sidecar-style metadata retained on the structured model. */
  chartStyle?: ChartStyleModel;
  /** Optional sidecar-color metadata retained on the structured model. */
  chartColors?: ChartColorsModel;
}

export interface AddChartExHistogramOptions extends Omit<AddChartExOptions, "type"> {
  /** `histogram` for frequency columns, `pareto` for frequency columns plus cumulative line. */
  type?: "histogram" | "pareto";
  layout?: Pick<ChartExLayoutProperties, "binning">;
}

export interface AddChartExWaterfallOptions extends Omit<AddChartExOptions, "type"> {
  type?: "waterfall";
  layout?: Pick<ChartExLayoutProperties, "subtotals">;
}

export interface AddChartExBoxWhiskerOptions extends Omit<AddChartExOptions, "type"> {
  type?: "boxWhisker";
  layout?: Pick<
    ChartExLayoutProperties,
    "quartileMethod" | "showMeanLine" | "showMeanMarker" | "showInnerPoints" | "showOutlierPoints"
  >;
}

/**
 * Supported chartEx high-level types.
 */
export type ChartExType =
  | "sunburst"
  | "treemap"
  | "waterfall"
  | "funnel"
  | "histogram"
  | "pareto"
  | "boxWhisker"
  | "regionMap";

export interface AddChartExSeriesOptions {
  /**
   * Display name for the series. Accepts:
   *   - `string` — literal caption (e.g. `"Quarterly sales"`).
   *   - `{ formula: string }` — worksheet cell reference resolved at
   *     read time (e.g. `{ formula: "Sheet1!$B$1" }`). Matches the
   *     classic chart-builder `AddChartSeriesOptions.name` shape so
   *     applications can share typings across the two builders.
   *   - `ChartRichText` — structured rich-text for per-run formatting.
   *
   * Previously only the string form was accepted; the formula /
   * rich-text forms were silently dropped via type narrowing.
   */
  name?: string | { formula: string } | ChartRichText;
  /** Values reference (e.g. "Sheet1!$B$2:$B$10") */
  values: string;
  /** Literal cached values for headless charts that are not backed by worksheet formulas. */
  literalValues?: number[];
  /** Literal cached categories for headless charts that are not backed by worksheet formulas. */
  literalCategories?: string[];
  /** Per-type extra references:
   *  - histogram: not used (categories carries the bin source)
   *  - waterfall: subtotal indices
   *  - pareto: line series uses the cumulative values automatically
   */
  /** Sub-category hierarchy levels (sunburst/treemap) */
  hierarchy?: string[];
  /** Literal cached hierarchy levels for headless sunburst/treemap previews. */
  literalHierarchy?: string[][];
  /** Waterfall subtotal indices (0-based) */
  subtotals?: number[];
  /** Waterfall subtotal marker objects, useful when callers already use OOXML-shaped config. */
  subtotalPoints?: Array<{ idx: number }>;
  /** Fill color (hex) */
  fill?: string;
  /** Border color (hex) */
  border?: string;
  /**
   * Series-level shape properties. Accepts either a structured
   * {@link ShapeProperties} or the shorthand {@link AddShapeFillOptions}
   * bag (same ergonomic hex-colour / gradient / pattern form classic
   * charts already support). Takes precedence over the
   * {@link AddChartExSeriesOptions.fill} / {@link AddChartExSeriesOptions.border}
   * convenience fields when both are provided.
   */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Data labels */
  dataLabels?: {
    showValue?: boolean;
    showCategory?: boolean;
    showSeriesName?: boolean;
    /**
     * Emit `<cx:visibility numFmt="1"/>` so the data label shows the
     * formatted number alongside the other visibility flags. Maps to
     * `ChartExDataLabels.visibility.numFmt`. Defaults to `undefined`
     * (Excel's own behaviour: the attribute is omitted, readers pick
     * a default based on the layout).
     */
    showNumFmt?: boolean;
    position?: string;
    separator?: string;
    numFmt?: string;
    /**
     * Shape properties for the data-label fills / borders. Accepts
     * the same structured {@link ShapeProperties} form as other
     * `spPr` slots in this file. Propagated to the internal
     * `ChartExDataLabels.spPr` so the renderer's `<cx:spPr>` emit
     * carries it. Previously the internal type exposed this slot but
     * the public options dropped it, leaving programmatic ChartEx
     * authors without a way to style data-label backgrounds.
     */
    spPr?: ShapeProperties | AddShapeFillOptions;
    /**
     * Run-level text properties for the label text. Routed to the
     * internal `ChartExDataLabels.txPr`, which the renderer emits as
     * `<cx:txPr>`. See {@link ChartTextProperties} for the field
     * shape (font family, size, bold/italic, colour …).
     */
    txPr?: ChartTextProperties;
  };
}

/**
 * Stored entry for a structured ChartEx (Office 2016+ extended chart).
 * When a ChartEx is created programmatically via `addChartEx()`, a structured
 * model is stored here and serialised through the builder/renderer on write.
 * When a ChartEx is round-tripped, raw bytes are used instead (stored under
 * `workbook._chartExEntries`). Pure data — stored in
 * `WorkbookData._chartExStructuredEntries`.
 */
export interface ChartExEntry {
  /** 1-based chartEx number (matches chartEx{N}.xml) */
  chartExNumber: number;
  /** Structured model (built from addChartEx options) */
  model: ChartExModel;
  /** Original chartEx XML bytes from a loaded workbook, used for clean round-trip passthrough */
  rawData?: Uint8Array;
  /** JSON snapshot of `model` taken when `rawData` was parsed */
  modelSnapshot?: string;
  /** True once a high-level API mutates the parsed chartEx model */
  dirty?: boolean;
  /** When true, simple high-level mutations may patch raw ChartEx XML instead of full re-render. */
  preferRawPatch?: boolean;
  /** When true, writing fails instead of re-rendering if raw ChartEx XML cannot be safely patched. */
  requireRawPatch?: boolean;
  /** ChartEx rels — preserved for round-trip */
  rels?: ChartRelEntry[];
}
