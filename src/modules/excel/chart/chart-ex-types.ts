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

import type { ChartLayout, ChartLegend, ChartTitle, ShapeProperties } from "./types";

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
  /** Extension list */
  extLst?: string;
}

export interface ChartExData {
  /** Externally referenced data (each entry has an id + data source) */
  externalData?: Array<{ id: string; autoUpdate?: boolean }>;
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
}

export interface ChartExPlotArea {
  /** Plot area region (layout + spPr) */
  plotAreaRegion?: {
    layout?: ChartLayout;
    plotSurface?: ShapeProperties;
    /** Plot area series */
    series: ChartExSeries[];
  };
  /** Axes (value/category/etc.) */
  axis?: ChartExAxis[];
  /** Direct access to series (short-hand if no plotAreaRegion) */
  series?: ChartExSeries[];
  spPr?: ShapeProperties;
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
  /** 0-based series index */
  seriesIndex?: number;
  /** Display name (tx) — string or formula */
  tx?: { rich?: unknown; strRef?: string; value?: string };
  /** Hidden attribute */
  hidden?: boolean;
  /** Ownership — "primary" or "standard" */
  ownerIdx?: number;
  /** Shape properties */
  spPr?: ShapeProperties;
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
  txPr?: unknown;
}

/**
 * Per-series layout properties. Only fields relevant to the series type are used.
 */
export interface ChartExLayoutProperties {
  // Sunburst / Treemap
  parentLabelLayout?: "banner" | "overlapping" | "none";
  // Waterfall
  subtotals?: Array<{ idx: number }>;
  // Funnel
  // (funnel has no extra layout props at this level — uses plot area)
  // Histogram / Pareto
  binning?: {
    binSize?: number;
    binCount?: number;
    binType?: "binCount" | "binSize" | "categories" | "manual";
    intervalClosed?: "l" | "r";
    underflow?: number;
    overflow?: number;
  };
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
  /** Scaling (for value axes) */
  scaling?: { min?: number; max?: number; orientation?: "minMax" | "maxMin" };
  /** Major tick mark */
  majorTickMark?: "none" | "inside" | "outside" | "cross";
  /** Minor tick mark */
  minorTickMark?: "none" | "inside" | "outside" | "cross";
  /** Major/minor unit */
  majorUnit?: number;
  minorUnit?: number;
  /** Number format */
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  /** Shape properties */
  spPr?: ShapeProperties;
  /** Text properties */
  txPr?: unknown;
  /** Hidden */
  hidden?: boolean;
  /** Label alignment */
  lblAlgn?: "ctr" | "l" | "r";
  /** Label offset */
  lblOffset?: number;
  /** Category axis scaling */
  catScaling?: { gapWidth?: number };
  /** Value axis scaling */
  valScaling?: { max?: number; min?: number; majorUnit?: number; minorUnit?: number };
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
  /** Chart title */
  title?: string;
  /** Show legend */
  showLegend?: boolean;
  /** Legend position */
  legendPosition?: "b" | "l" | "r" | "t" | "tr";
  /** Type-specific layout overrides */
  layout?: ChartExLayoutProperties;
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
  name?: string;
  /** Values reference (e.g. "Sheet1!$B$2:$B$10") */
  values: string;
  /** Per-type extra references:
   *  - histogram: not used (categories carries the bin source)
   *  - waterfall: subtotal indices
   *  - pareto: line series uses the cumulative values automatically
   */
  /** Sub-category hierarchy levels (sunburst/treemap) */
  hierarchy?: string[];
  /** Waterfall subtotal indices (0-based) */
  subtotals?: number[];
  /** Fill color (hex) */
  fill?: string;
  /** Border color (hex) */
  border?: string;
  /** Series-level shape properties */
  spPr?: ShapeProperties;
  /** Data labels */
  dataLabels?: {
    showValue?: boolean;
    showCategory?: boolean;
    showSeriesName?: boolean;
    position?: string;
    separator?: string;
    numFmt?: string;
  };
}
