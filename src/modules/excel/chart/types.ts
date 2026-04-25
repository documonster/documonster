/**
 * Chart type definitions for ExcelTS
 *
 * Covers the full OOXML DrawingML Chart specification (c: namespace).
 * Every chart in an XLSX file is a `c:chartSpace` containing a `c:chart`
 * which holds a `c:plotArea` with one or more chart type groups.
 */

import type { AnchorModel } from "@excel/anchor";

// ============================================================================
// Chart Type Enums
// ============================================================================

/**
 * All chart types supported by OOXML DrawingML Charts.
 */
export type ChartType =
  | "bar"
  | "bar3D"
  | "line"
  | "line3D"
  | "pie"
  | "pie3D"
  | "doughnut"
  | "area"
  | "area3D"
  | "scatter"
  | "bubble"
  | "radar"
  | "stock"
  | "surface"
  | "surface3D"
  | "ofPie"; // pie-of-pie / bar-of-pie

/**
 * Bar/column chart grouping.
 */
export type BarGrouping = "clustered" | "stacked" | "percentStacked" | "standard";

/**
 * Bar chart direction.
 */
export type BarDirection = "bar" | "col";

/**
 * Line/area chart grouping.
 */
export type LineGrouping = "standard" | "stacked" | "percentStacked";

/**
 * Radar chart style.
 */
export type RadarStyle = "standard" | "marker" | "filled";

/**
 * Scatter chart style.
 */
export type ScatterStyle = "lineMarker" | "line" | "marker" | "smooth" | "smoothMarker";

/**
 * Stock chart arrangement (HLC, OHLC, VHLC, VOHLC).
 */
// Stock charts are identified by their series count, not an explicit style attribute.

/**
 * Of-pie chart type.
 */
export type OfPieType = "pie" | "bar";

/**
 * Shape for 3D bar charts.
 */
export type BarShape = "box" | "cone" | "coneToMax" | "cylinder" | "pyramid" | "pyramidToMax";

/**
 * Axis cross position.
 */
export type AxisCrosses = "autoZero" | "max" | "min";

/**
 * Axis position.
 */
export type AxisPosition = "b" | "l" | "r" | "t";

/**
 * Axis orientation.
 */
export type AxisOrientation = "minMax" | "maxMin";

/**
 * Tick mark type.
 */
export type TickMark = "none" | "inside" | "outside" | "cross";

/**
 * Tick label position.
 */
export type TickLabelPosition = "high" | "low" | "nextTo" | "none";

/**
 * Time unit for date axes.
 */
export type TimeUnit = "days" | "months" | "years";

/**
 * Label alignment for category axes.
 */
export type LabelAlignment = "ctr" | "l" | "r";

/**
 * Legend position.
 */
export type LegendPosition = "b" | "l" | "r" | "t" | "tr";

/**
 * Display blanks as.
 */
export type DisplayBlanksAs = "gap" | "span" | "zero";

/**
 * Label position for data labels.
 */
export type DataLabelPosition =
  | "bestFit"
  | "b"
  | "ctr"
  | "inBase"
  | "inEnd"
  | "l"
  | "outEnd"
  | "r"
  | "t";

/**
 * Trendline type.
 */
export type TrendlineType = "exp" | "linear" | "log" | "movingAvg" | "poly" | "power";

/**
 * Error bar type.
 */
export type ErrorBarType = "both" | "minus" | "plus";

/**
 * Error bar direction (x or y axis).
 */
export type ErrorBarDirection = "x" | "y";

/**
 * Error bar value type.
 */
export type ErrorBarValueType = "cust" | "fixedVal" | "percentage" | "stdDev" | "stdErr";

/**
 * Split type for of-pie charts.
 */
export type SplitType = "auto" | "cust" | "percent" | "pos" | "val";

/**
 * Picture format for picture fills in markers/bars.
 */
export type PictureFormat = "stretch" | "stack" | "stackScale";

// ============================================================================
// Color and Style Types
// ============================================================================

/**
 * DrawingML solid color fill.
 */
export interface ChartColor {
  /** sRGB hex color (e.g. "FF0000") */
  srgb?: string;
  /** Theme color index */
  theme?: number;
  /** System color (e.g. "windowText", "window") */
  sysClr?: string;
  /** Preset color name (e.g. "black", "white") */
  prstClr?: string;
  /** Tint modifier (0 to 1.0, fraction — stored as 0–100000 in OOXML). E.g., 0.5 = 50% tint. */
  tint?: number;
  /** Shade (OOXML percentage, 0-100000) */
  shade?: number;
  /** Saturation modification (OOXML percentage) */
  satMod?: number;
  /** Luminance modification (OOXML percentage, 0-100000) */
  lumMod?: number;
  /** Luminance offset (OOXML percentage, 0-100000) */
  lumOff?: number;
  /** Alpha (OOXML percentage, 0-100000; 100000 = fully opaque) */
  alpha?: number;
}

/**
 * Line (outline) properties.
 */
export interface ChartLine {
  /** Width in EMU */
  width?: number;
  /** Solid fill color */
  color?: ChartColor;
  /** Dash style */
  dash?:
    | "solid"
    | "dot"
    | "dash"
    | "lgDash"
    | "dashDot"
    | "lgDashDot"
    | "lgDashDotDot"
    | "sysDash"
    | "sysDot"
    | "sysDashDot"
    | "sysDashDotDot";
  /** Line cap style */
  cap?: "flat" | "sq" | "rnd";
  /** Line join style */
  join?: "round" | "bevel" | "miter";
  /** Compound line type */
  compound?: "sng" | "dbl" | "thickThin" | "thinThick" | "tri";
  /** No line */
  noFill?: boolean;
}

/**
 * Fill properties.
 */
export interface ChartFill {
  /** Solid color fill */
  solid?: ChartColor;
  /** No fill */
  noFill?: boolean;
  /** Pattern fill */
  pattern?: {
    preset: string;
    foreground?: ChartColor;
    background?: ChartColor;
  };
  /** Gradient fill */
  gradient?: {
    stops: Array<{ position: number; color: ChartColor }>;
    angle?: number;
    /** Linear, circle, rect, shape */
    type?: "linear" | "circle" | "rect" | "shape";
  };
}

/**
 * Shape properties (c:spPr).
 */
export interface ShapeProperties {
  fill?: ChartFill;
  line?: ChartLine;
  /** Structured effect list (shadow, glow, soft edge, reflection, blur). */
  effectList?: EffectList;
  /** 3D scene (camera + light rig) */
  scene3d?: Scene3D;
  /** 3D shape properties (bevel / extrusion / material) */
  sp3d?: ShapeProperties3D;
  /** @internal Raw XML string for perfect round-trip fidelity */
  _rawXml?: string;
}

/**
 * Effect list (a:effectLst) — drop shadow, outer/inner glow, soft edge,
 * reflection, blur. Multiple effects can be combined.
 */
export interface EffectList {
  /** Blur effect (a:blur) */
  blur?: {
    /** Blur radius in EMU */
    radius?: number;
    /** If true, blur grows shape bounds */
    grow?: boolean;
  };
  /** Outer shadow (a:outerShdw) */
  outerShadow?: Shadow;
  /** Inner shadow (a:innerShdw) */
  innerShadow?: Shadow;
  /** Preset shadow (a:prstShdw val="shdw1".."shdw20") */
  presetShadow?: {
    preset: string;
    distance?: number;
    direction?: number;
    color?: ChartColor;
  };
  /** Outer glow (a:glow) */
  glow?: {
    radius: number;
    color: ChartColor;
  };
  /** Soft edge (a:softEdge) */
  softEdge?: {
    radius: number;
  };
  /** Reflection (a:reflection) */
  reflection?: {
    /** Blur radius in EMU */
    blurRadius?: number;
    /** Start opacity 0-100000 */
    startOpacity?: number;
    /** Start position 0-100000 */
    startPosition?: number;
    /** End opacity */
    endOpacity?: number;
    /** End position */
    endPosition?: number;
    /** Distance from shape */
    distance?: number;
    /** Direction in 60000ths of a degree */
    direction?: number;
    /** Fade direction */
    fadeDirection?: number;
    /** Horizontal scale 0-100000 */
    scaleHorizontal?: number;
    /** Vertical scale */
    scaleVertical?: number;
    /** Horizontal skew angle */
    skewHorizontal?: number;
    /** Vertical skew angle */
    skewVertical?: number;
    /** Alignment: bl/br/ctr/l/r/t/tl/tr/b */
    alignment?: "b" | "bl" | "br" | "ctr" | "l" | "r" | "t" | "tl" | "tr";
    /** Rotate with shape */
    rotateWithShape?: boolean;
  };
}

/**
 * Shadow effect (used for both innerShdw and outerShdw).
 */
export interface Shadow {
  /** Blur radius in EMU */
  blurRadius?: number;
  /** Distance from shape in EMU */
  distance?: number;
  /** Direction in 60000ths of a degree (outer only) */
  direction?: number;
  /** Alignment (outer only) */
  alignment?: "b" | "bl" | "br" | "ctr" | "l" | "r" | "t" | "tl" | "tr";
  /** Shadow color */
  color: ChartColor;
  /** Rotate with shape (outer only) */
  rotateWithShape?: boolean;
  /** Horizontal scale (outer only) */
  scaleHorizontal?: number;
  /** Vertical scale (outer only) */
  scaleVertical?: number;
  /** Horizontal skew (outer only) */
  skewHorizontal?: number;
  /** Vertical skew (outer only) */
  skewVertical?: number;
}

/**
 * 3D scene (a:scene3d) — camera and light rig.
 */
export interface Scene3D {
  camera?: {
    preset: string;
    fov?: number;
    zoom?: number;
    rotation?: { lat: number; lon: number; rev: number };
  };
  lightRig?: {
    rig: string;
    direction: string;
    rotation?: { lat: number; lon: number; rev: number };
  };
  backdrop?: unknown;
}

/**
 * 3D shape properties (a:sp3d) — bevels, extrusion, material.
 */
export interface ShapeProperties3D {
  /** Z-axis height (a:sp3d/@z) */
  z?: number;
  /** Extrusion height (a:sp3d/@extrusionH) */
  extrusionHeight?: number;
  /** Contour width (a:sp3d/@contourW) */
  contourWidth?: number;
  /** Preset material (legacy, matte, plastic, metal, etc.) */
  material?: string;
  /** Top bevel */
  bevelTop?: Bevel;
  /** Bottom bevel */
  bevelBottom?: Bevel;
  /** Extrusion color (a:extrusionClr) */
  extrusionColor?: ChartColor;
  /** Contour color (a:contourClr) */
  contourColor?: ChartColor;
}

export interface Bevel {
  /** Width in EMU */
  width?: number;
  /** Height in EMU */
  height?: number;
  /** Preset: angle, artDeco, circle, etc. */
  preset?: string;
}

/**
 * Underline style (OOXML a:u attribute values).
 * See §20.1.10.82 of ECMA-376.
 */
export type UnderlineStyle =
  | "none"
  | "words"
  | "sng"
  | "dbl"
  | "heavy"
  | "dotted"
  | "dottedHeavy"
  | "dash"
  | "dashHeavy"
  | "dashLong"
  | "dashLongHeavy"
  | "dotDash"
  | "dotDashHeavy"
  | "dotDotDash"
  | "dotDotDashHeavy"
  | "wavy"
  | "wavyHeavy"
  | "wavyDbl";

/** Strike-through style (OOXML a:strike). */
export type StrikeStyle = "noStrike" | "sngStrike" | "dblStrike";

/** Capitalization (OOXML a:cap). */
export type CapStyle = "none" | "small" | "all";

/** Paragraph alignment (OOXML a:algn). */
export type ParagraphAlignment = "l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist";

/**
 * Text properties for chart text elements.
 *
 * Maps to OOXML `<a:rPr>` / `<a:defRPr>` run properties. All fields are optional
 * and only those set are serialised; raw XML fallback (`_rawXml`) is used when
 * preserving unparseable round-trip content.
 */
export interface ChartTextProperties {
  /** Font size in hundredths of a point (e.g. 1000 = 10pt) */
  size?: number;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /**
   * Underline. `true` is a shorthand for `"sng"`; `false`/omitted means none.
   * For full OOXML control use the string variant.
   */
  underline?: boolean | UnderlineStyle;
  /** Strike-through */
  strike?: StrikeStyle;
  /** Font color */
  color?: ChartColor;
  /** Font family (Latin typeface — maps to a:latin/@typeface) */
  fontFamily?: string;
  /** East Asian typeface (a:ea/@typeface) */
  eastAsianFamily?: string;
  /** Complex-script typeface (a:cs/@typeface) */
  complexScriptFamily?: string;
  /** Rotation in 60000ths of a degree (applied at bodyPr level when used on paragraph text) */
  rotation?: number;
  /** Baseline offset (percentage * 1000 — positive=superscript, negative=subscript) */
  baseline?: number;
  /** Character kerning cut-off, hundredths of a point (a:rPr/@kern) */
  kern?: number;
  /** Character spacing in hundredths of a point (a:rPr/@spc) */
  spacing?: number;
  /** Capitalisation */
  cap?: CapStyle;
  /** Language (a:rPr/@lang, e.g. "en-US") */
  lang?: string;
  /** @internal Raw XML string for perfect round-trip fidelity */
  _rawXml?: string;
}

/**
 * Rich text for chart labels and titles.
 */
export interface ChartRichText {
  paragraphs: ChartParagraph[];
  /** Body properties (wrapping, anchoring, rotation) */
  bodyProperties?: ChartBodyProperties;
}

/**
 * Body-level text properties (OOXML a:bodyPr).
 */
export interface ChartBodyProperties {
  /** Rotation in 60000ths of a degree */
  rotation?: number;
  /** Horizontal overflow: overflow/clip */
  horizontalOverflow?: "overflow" | "clip";
  /** Vertical anchor */
  anchor?: "t" | "ctr" | "b" | "just" | "dist";
  /** Text wrapping */
  wrap?: "none" | "square";
  /** Vertical text (stacked) */
  vertical?:
    | "horz"
    | "vert"
    | "vert270"
    | "wordArtVert"
    | "eaVert"
    | "mongolianVert"
    | "wordArtVertRtl";
}

/**
 * Paragraph-level properties (OOXML a:pPr).
 */
export interface ChartParagraphProperties {
  /** Alignment */
  alignment?: ParagraphAlignment;
  /** Indent in EMU (914400 = 1 inch) */
  indent?: number;
  /** Left margin in EMU */
  marginLeft?: number;
  /** Right margin in EMU */
  marginRight?: number;
  /** Bullet: character / auto-number / none */
  bullet?: ChartBullet;
  /** Line spacing (percentage * 1000 for a:spcPct, or points * 100 for a:spcPts) */
  lineSpacing?: ChartLineSpacing;
  /** Space before paragraph */
  spaceBefore?: ChartLineSpacing;
  /** Space after paragraph */
  spaceAfter?: ChartLineSpacing;
  /** Level (0-8) for nested lists */
  level?: number;
  /** Default run properties applied to runs without their own properties */
  defaultRunProperties?: ChartTextProperties;
}

export type ChartBullet =
  | { type: "none" }
  | { type: "char"; character: string }
  | { type: "autoNum"; scheme: string; startAt?: number };

export type ChartLineSpacing =
  | { type: "percentage"; value: number }
  | { type: "points"; value: number };

export interface ChartParagraph {
  /** Paragraph properties (a:pPr) */
  properties?: ChartParagraphProperties;
  /** Legacy alias for default run properties (kept for backward compat). */
  runProperties?: ChartTextProperties;
  /** Text runs */
  runs?: ChartTextRun[];
  /** End-paragraph run properties (a:endParaRPr) */
  endParaRunProperties?: ChartTextProperties;
}

export interface ChartTextRun {
  text: string;
  /** Run properties (a:rPr) */
  properties?: ChartTextProperties;
  /** Hyperlink — a:hlinkClick */
  hyperlink?: {
    /** Rel ID pointing to the target (for external hyperlinks) */
    relationshipId?: string;
    /** Tooltip text */
    tooltip?: string;
  };
}

// ============================================================================
// Data References
// ============================================================================

/**
 * Number reference (c:numRef).
 */
export interface NumberReference {
  /** Formula reference (e.g. "Sheet1!$B$2:$B$5") */
  formula: string;
  /** Cached numeric values */
  cache?: NumberCache;
}

export interface NumberCache {
  formatCode?: string;
  pointCount?: number;
  points: Array<{ index: number; value: number | null; formatCode?: string }>;
}

/**
 * String reference (c:strRef).
 */
export interface StringReference {
  /** Formula reference */
  formula: string;
  /** Cached string values */
  cache?: StringCache;
}

export interface StringCache {
  pointCount?: number;
  points: Array<{ index: number; value: string }>;
}

/**
 * Number literal (c:numLit).
 */
export interface NumberLiteral {
  formatCode?: string;
  pointCount?: number;
  points: Array<{ index: number; value: number | null }>;
}

/**
 * String literal (c:strLit).
 */
export interface StringLiteral {
  pointCount?: number;
  points: Array<{ index: number; value: string }>;
}

/**
 * Multi-level string reference (c:multiLvlStrRef).
 */
export interface MultiLevelStringReference {
  formula: string;
  cache?: MultiLevelStringCache;
}

export interface MultiLevelStringCache {
  pointCount?: number;
  levels: StringCache[];
}

/**
 * Axis data source — category axis data.
 */
export interface AxisDataSource {
  numRef?: NumberReference;
  numLit?: NumberLiteral;
  strRef?: StringReference;
  strLit?: StringLiteral;
  multiLvlStrRef?: MultiLevelStringReference;
}

/**
 * Number data source — value axis data.
 */
export interface NumberDataSource {
  numRef?: NumberReference;
  numLit?: NumberLiteral;
}

// ============================================================================
// Marker
// ============================================================================

/**
 * Series marker.
 */
export interface ChartMarker {
  symbol?:
    | "circle"
    | "dash"
    | "diamond"
    | "dot"
    | "none"
    | "picture"
    | "plus"
    | "square"
    | "star"
    | "triangle"
    | "x"
    | "auto";
  size?: number;
  spPr?: ShapeProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

// ============================================================================
// Data Labels
// ============================================================================

/**
 * Data labels configuration.
 */
export interface DataLabels {
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLeaderLines?: boolean;
  separator?: string;
  position?: DataLabelPosition;
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** Individual data label overrides */
  entries?: DataLabelEntry[];
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface DataLabelEntry {
  index: number;
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  position?: DataLabelPosition;
  layout?: ChartLayout;
  text?: ChartRichText;
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** Delete this specific label */
  delete?: boolean;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
  /** @internal Raw c:tx XML for round-trip of rich text overrides */
  rawTx?: string;
}

// ============================================================================
// Trendline and Error Bars
// ============================================================================

export interface Trendline {
  type: TrendlineType;
  name?: string;
  order?: number; // for polynomial
  period?: number; // for moving average
  forward?: number;
  backward?: number;
  intercept?: number;
  displayRSqr?: boolean;
  displayEq?: boolean;
  spPr?: ShapeProperties;
  trendlineLbl?: TrendlineLabel;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface TrendlineLabel {
  layout?: ChartLayout;
  text?: ChartRichText;
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** @internal Raw c:tx XML for round-trip */
  rawTx?: string;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface ErrorBars {
  errDir?: ErrorBarDirection;
  barDir: ErrorBarType;
  errValType: ErrorBarValueType;
  noEndCap?: boolean;
  val?: number;
  plus?: NumberDataSource;
  minus?: NumberDataSource;
  spPr?: ShapeProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

// ============================================================================
// Series Types
// ============================================================================

/**
 * Base properties shared by all series types.
 */
export interface SeriesBase {
  /** Series index (0-based order) */
  index: number;
  /** Series plot order */
  order: number;
  /** Series name */
  tx?: { strRef?: StringReference; value?: string };
  /** Shape properties */
  spPr?: ShapeProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * Bar/column chart series.
 */
export interface BarSeries extends SeriesBase {
  invertIfNegative?: boolean;
  pictureOptions?: PictureOptions;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars;
  cat?: AxisDataSource;
  val?: NumberDataSource;
  shape?: BarShape;
}

/**
 * Line chart series.
 */
export interface LineSeries extends SeriesBase {
  marker?: ChartMarker;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars;
  cat?: AxisDataSource;
  val?: NumberDataSource;
  smooth?: boolean;
}

/**
 * Pie/doughnut chart series.
 */
export interface PieSeries extends SeriesBase {
  explosion?: number;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  cat?: AxisDataSource;
  val?: NumberDataSource;
}

/**
 * Area chart series.
 */
export interface AreaSeries extends SeriesBase {
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars;
  cat?: AxisDataSource;
  val?: NumberDataSource;
}

/**
 * Scatter chart series.
 */
export interface ScatterSeries extends SeriesBase {
  marker?: ChartMarker;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars[];
  xVal?: AxisDataSource;
  yVal?: NumberDataSource;
  smooth?: boolean;
}

/**
 * Bubble chart series.
 */
export interface BubbleSeries extends SeriesBase {
  invertIfNegative?: boolean;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  trendlines?: Trendline[];
  errorBars?: ErrorBars[];
  xVal?: AxisDataSource;
  yVal?: NumberDataSource;
  bubbleSize?: NumberDataSource;
  bubble3D?: boolean;
}

/**
 * Radar chart series.
 */
export interface RadarSeries extends SeriesBase {
  marker?: ChartMarker;
  dataPoints?: DataPoint[];
  dataLabels?: DataLabels;
  cat?: AxisDataSource;
  val?: NumberDataSource;
}

/**
 * Surface chart series.
 */
export interface SurfaceSeries extends SeriesBase {
  cat?: AxisDataSource;
  val?: NumberDataSource;
}

/**
 * Stock chart series - same as line series (HLC or OHLC).
 */
export type StockSeries = LineSeries;

/**
 * Individual data point override.
 */
export interface DataPoint {
  index: number;
  invertIfNegative?: boolean;
  marker?: ChartMarker;
  bubble3D?: boolean;
  explosion?: number;
  spPr?: ShapeProperties;
  pictureOptions?: PictureOptions;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface PictureOptions {
  applyToFront?: boolean;
  applyToSides?: boolean;
  applyToEnd?: boolean;
  pictureFormat?: PictureFormat;
  pictureStackUnit?: number;
}

// ============================================================================
// Chart Type Groups
// ============================================================================

export interface BarChartGroup {
  type: "bar" | "bar3D";
  barDir: BarDirection;
  grouping: BarGrouping;
  varyColors?: boolean;
  series: BarSeries[];
  dataLabels?: DataLabels;
  gapWidth?: number;
  overlap?: number;
  serLines?: ShapeProperties;
  axisIds: number[];
  /** 3D bar shape */
  shape?: BarShape;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface LineChartGroup {
  type: "line" | "line3D";
  grouping: LineGrouping;
  varyColors?: boolean;
  series: LineSeries[];
  dataLabels?: DataLabels;
  marker?: boolean;
  smooth?: boolean;
  hiLowLines?: ShapeProperties;
  upDownBars?: UpDownBars;
  dropLines?: ShapeProperties;
  axisIds: number[];
  extLst?: string;
}

export interface PieChartGroup {
  type: "pie" | "pie3D";
  varyColors?: boolean;
  series: PieSeries[];
  dataLabels?: DataLabels;
  firstSliceAng?: number;
  axisIds?: number[]; // pie charts typically don't use axes
  extLst?: string;
}

export interface DoughnutChartGroup {
  type: "doughnut";
  varyColors?: boolean;
  series: PieSeries[];
  dataLabels?: DataLabels;
  firstSliceAng?: number;
  holeSize?: number;
  axisIds?: number[];
  extLst?: string;
}

export interface AreaChartGroup {
  type: "area" | "area3D";
  grouping: LineGrouping;
  varyColors?: boolean;
  series: AreaSeries[];
  dataLabels?: DataLabels;
  dropLines?: ShapeProperties;
  axisIds: number[];
  extLst?: string;
}

export interface ScatterChartGroup {
  type: "scatter";
  scatterStyle: ScatterStyle;
  varyColors?: boolean;
  series: ScatterSeries[];
  dataLabels?: DataLabels;
  axisIds: number[];
  extLst?: string;
}

export interface BubbleChartGroup {
  type: "bubble";
  varyColors?: boolean;
  series: BubbleSeries[];
  dataLabels?: DataLabels;
  bubbleScale?: number;
  showNegBubbles?: boolean;
  sizeRepresents?: "area" | "w";
  axisIds: number[];
  extLst?: string;
}

export interface RadarChartGroup {
  type: "radar";
  radarStyle: RadarStyle;
  varyColors?: boolean;
  series: RadarSeries[];
  dataLabels?: DataLabels;
  axisIds: number[];
  extLst?: string;
}

export interface StockChartGroup {
  type: "stock";
  varyColors?: boolean;
  series: StockSeries[];
  dataLabels?: DataLabels;
  hiLowLines?: ShapeProperties;
  upDownBars?: UpDownBars;
  dropLines?: ShapeProperties;
  axisIds: number[];
  extLst?: string;
}

export interface SurfaceChartGroup {
  type: "surface" | "surface3D";
  wireframe?: boolean;
  series: SurfaceSeries[];
  dataLabels?: DataLabels;
  bandFormats?: BandFormat[];
  axisIds: number[];
  extLst?: string;
}

export interface OfPieChartGroup {
  type: "ofPie";
  ofPieType: OfPieType;
  varyColors?: boolean;
  series: PieSeries[];
  dataLabels?: DataLabels;
  gapWidth?: number;
  splitType?: SplitType;
  splitPos?: number;
  custSplit?: number[];
  secondPieSize?: number;
  serLines?: ShapeProperties;
  axisIds?: number[];
  extLst?: string;
}

export type ChartTypeGroup =
  | BarChartGroup
  | LineChartGroup
  | PieChartGroup
  | DoughnutChartGroup
  | AreaChartGroup
  | ScatterChartGroup
  | BubbleChartGroup
  | RadarChartGroup
  | StockChartGroup
  | SurfaceChartGroup
  | OfPieChartGroup;

export interface UpDownBars {
  gapWidth?: number;
  upBars?: ShapeProperties;
  downBars?: ShapeProperties;
}

export interface BandFormat {
  index: number;
  spPr?: ShapeProperties;
}

// ============================================================================
// Axes
// ============================================================================

/**
 * Base axis properties shared by all axis types.
 */
export interface AxisBase {
  /** Unique axis ID */
  axId: number;
  /** Scaling */
  scaling?: {
    orientation?: AxisOrientation;
    max?: number;
    min?: number;
    logBase?: number;
  };
  /** Delete the axis (hide it) */
  delete?: boolean;
  /** Axis position */
  axPos: AxisPosition;
  /** Major gridlines */
  majorGridlines?: ShapeProperties;
  /** Minor gridlines */
  minorGridlines?: ShapeProperties;
  /** Axis title */
  title?: ChartTitle;
  /** Number format */
  numFmt?: { formatCode: string; sourceLinked?: boolean };
  /** Major tick mark type */
  majorTickMark?: TickMark;
  /** Minor tick mark type */
  minorTickMark?: TickMark;
  /** Tick label position */
  tickLblPos?: TickLabelPosition;
  /** Shape properties for the axis line */
  spPr?: ShapeProperties;
  /** Text properties for tick labels */
  txPr?: ChartTextProperties;
  /** Cross axis ID (the axis this one crosses) */
  crossAx: number;
  /** Where this axis crosses the other */
  crosses?: AxisCrosses;
  /** Explicit cross value */
  crossesAt?: number;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * Category axis (c:catAx).
 */
export interface CategoryAxis extends AxisBase {
  axisType: "cat";
  /** Auto label ordering */
  auto?: boolean;
  /** Label alignment */
  lblAlgn?: LabelAlignment;
  /** Label offset (percentage, 0-1000) */
  lblOffset?: number;
  /** Tick label skip */
  tickLblSkip?: number;
  /** Tick mark skip */
  tickMarkSkip?: number;
  /** No multi-level labels */
  noMultiLvlLbl?: boolean;
}

/**
 * Value axis (c:valAx).
 */
export interface ValueAxis extends AxisBase {
  axisType: "val";
  /** Cross between categories or midpoints */
  crossBetween?: "between" | "midCat";
  /** Major unit */
  majorUnit?: number;
  /** Minor unit */
  minorUnit?: number;
  /** Display units (hundreds, thousands, millions, etc.) */
  dispUnits?: DisplayUnits;
}

/**
 * Date axis (c:dateAx).
 */
export interface DateAxis extends AxisBase {
  axisType: "date";
  /** Auto date detection */
  auto?: boolean;
  /** Label offset */
  lblOffset?: number;
  /** Base time unit */
  baseTimeUnit?: TimeUnit;
  /** Major unit */
  majorUnit?: number;
  /** Major time unit */
  majorTimeUnit?: TimeUnit;
  /** Minor unit */
  minorUnit?: number;
  /** Minor time unit */
  minorTimeUnit?: TimeUnit;
}

/**
 * Series axis (c:serAx) — used in 3D charts.
 */
export interface SeriesAxis extends AxisBase {
  axisType: "ser";
  /** Tick label skip */
  tickLblSkip?: number;
  /** Tick mark skip */
  tickMarkSkip?: number;
}

export type ChartAxis = CategoryAxis | ValueAxis | DateAxis | SeriesAxis;

export interface DisplayUnits {
  builtInUnit?:
    | "hundreds"
    | "thousands"
    | "tenThousands"
    | "hundredThousands"
    | "millions"
    | "tenMillions"
    | "hundredMillions"
    | "billions"
    | "trillions";
  custUnit?: number;
  label?: ChartTitle;
}

// ============================================================================
// Layout
// ============================================================================

/**
 * Manual layout positioning.
 */
export interface ChartLayout {
  manualLayout?: ManualLayout;
}

export interface ManualLayout {
  layoutTarget?: "inner" | "outer";
  xMode?: "edge" | "factor";
  yMode?: "edge" | "factor";
  wMode?: "edge" | "factor";
  hMode?: "edge" | "factor";
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

// ============================================================================
// Title and Legend
// ============================================================================

/**
 * Chart title (c:title).
 */
export interface ChartTitle {
  text?: ChartRichText;
  /** Title from data reference */
  strRef?: StringReference;
  /** Raw c:tx XML for round-trip fidelity (used instead of text when present) */
  rawTx?: string;
  layout?: ChartLayout;
  overlay?: boolean;
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * Chart legend (c:legend).
 */
export interface ChartLegend {
  legendPos?: LegendPosition;
  legendEntries?: LegendEntry[];
  layout?: ChartLayout;
  overlay?: boolean;
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

export interface LegendEntry {
  index: number;
  delete?: boolean;
  txPr?: ChartTextProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

// ============================================================================
// 3D View
// ============================================================================

/**
 * 3D view settings (c:view3D).
 */
export interface View3D {
  rotX?: number;
  rotY?: number;
  depthPercent?: number;
  rAngAx?: boolean;
  hPercent?: number;
  perspective?: number;
}

// ============================================================================
// Plot Area and Chart
// ============================================================================

/**
 * Plot area (c:plotArea).
 */
export interface PlotArea {
  layout?: ChartLayout;
  /** One or more chart type groups (combo charts have multiple) */
  chartTypes: ChartTypeGroup[];
  /** Axes */
  axes: ChartAxis[];
  /** Data table */
  dataTable?: DataTable;
  /** Shape properties for the plot area background */
  spPr?: ShapeProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * Data table display options (c:dTable).
 */
export interface DataTable {
  showHorzBorder?: boolean;
  showVertBorder?: boolean;
  showOutline?: boolean;
  showKeys?: boolean;
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * The main chart element (c:chart).
 */
export interface ChartData {
  title?: ChartTitle;
  autoTitleDeleted?: boolean;
  pivotFormats?: PivotFormat[];
  view3D?: View3D;
  floor?: ShapeProperties;
  sideWall?: ShapeProperties;
  backWall?: ShapeProperties;
  plotArea: PlotArea;
  legend?: ChartLegend;
  plotVisOnly?: boolean;
  dispBlanksAs?: DisplayBlanksAs;
  showDLblsOverMax?: boolean;
  /** Extension list at c:chart level (raw XML for round-trip) */
  extLst?: string;
}

export interface PivotFormat {
  index: number;
  spPr?: ShapeProperties;
  txPr?: ChartTextProperties;
  marker?: ChartMarker;
  dataLabels?: DataLabels;
  /** Raw XML for a single c:dLbl inside pivotFmt (not wrapped in c:dLbls) */
  rawDLbl?: string;
}

// ============================================================================
// Chart Space (top-level)
// ============================================================================

/**
 * The top-level chart container (c:chartSpace).
 * This is the root element of a chart XML file.
 */
export interface ChartModel {
  /** Chart data */
  chart: ChartData;
  /** Chart style (numeric style index) — legacy c:style */
  style?: number;
  /** Modern chart style via mc:AlternateContent/c14:style (raw XML for round-trip) */
  alternateContentStyle?: string;
  /** Chart-level shape properties */
  spPr?: ShapeProperties;
  /** Chart-level text properties */
  txPr?: ChartTextProperties;
  /** Print settings */
  printSettings?: PrintSettings;
  /** External data reference */
  externalData?: { id: string; autoUpdate?: boolean };
  /** Rounding on file load (preserve round-trip) */
  roundedCorners?: boolean;
  /** Language */
  lang?: string;
  /** Whether the chart uses the 1904 date system */
  date1904?: boolean;
  /** Pivot source information (raw XML for round-trip) */
  pivotSource?: string;
  /** Color map override (raw XML for round-trip) */
  clrMapOvr?: string;
  /** Chart protection settings (raw XML for round-trip) */
  protection?: string;
  /** Extension list at c:chartSpace level (raw XML for round-trip) */
  extLst?: string;
  /** Extra namespace declarations for round-trip fidelity */
  extraNamespaces?: Record<string, string>;
}

export interface PrintSettings {
  /** Header/footer content (raw XML string for round-trip) */
  headerFooter?: string;
  pageMargins?: {
    b: number;
    l: number;
    r: number;
    t: number;
    header: number;
    footer: number;
  };
  pageSetup?: {
    orientation?: "portrait" | "landscape";
    paperSize?: number;
  };
}

// ============================================================================
// Chart Style and Colors
// ============================================================================

/**
 * Chart style model (stored in styleN.xml alongside the chart).
 */
export interface ChartStyleModel {
  /** Raw XML preserved for round-trip */
  rawXml?: string;
  /** @internal Structured style ID (from cs:chartStyle/@id) */
  id?: number;
}

/**
 * Chart color model (stored in colorsN.xml alongside the chart).
 */
export interface ChartColorsModel {
  /** Raw XML preserved for round-trip */
  rawXml?: string;
  /** Color style method attribute (cycle/withinLinear/acrossLinear) */
  method?: string;
  /** @internal Sequence ID */
  id?: number;
  /** Color palette — each entry is either a theme reference or sRGB color */
  colors?: ChartColorsEntry[];
}

/**
 * A single color in a chart's colors.xml palette.
 */
export interface ChartColorsEntry {
  /** Theme color reference (a:schemeClr@val) */
  theme?: string;
  /** Straight sRGB hex (without #) */
  srgb?: string;
  /** Luminance modulation */
  lumMod?: number;
  /** Luminance offset */
  lumOff?: number;
  /** Tint */
  tint?: number;
  /** Shade */
  shade?: number;
  /** Saturation modulation */
  satMod?: number;
  /** Alpha (0-100000) */
  alpha?: number;
}

// ============================================================================
// Chart Anchor (Drawing Integration)
// ============================================================================

/**
 * Chart range for worksheet placement.
 */
export interface ChartRange {
  tl: AnchorModel;
  br: AnchorModel;
  editAs?: "oneCell" | "twoCell" | "absolute";
}

// ============================================================================
// High-Level API Input Types
// ============================================================================

/**
 * Simplified input for creating a chart programmatically.
 *
 * Usage:
 * ```ts
 * worksheet.addChart(
 *   { type: "bar", series: [{ values: "Sheet1!$B$1:$B$5" }], title: "Sales" },
 *   "A1:H15"
 * );
 * ```
 */
export interface AddChartOptions {
  /** Chart type */
  type: ChartType;
  /** Series definitions */
  series?: AddChartSeriesOptions[];
  /** Chart title text, formula reference, or structured rich text */
  title?: string | { formula: string } | ChartRichText;
  /** Show legend */
  showLegend?: boolean;
  /** Legend position */
  legendPosition?: LegendPosition;
  /** Grouping for bar/line/area */
  grouping?: BarGrouping | LineGrouping;
  /** Direction for bar charts */
  barDir?: BarDirection;
  /** Scatter style */
  scatterStyle?: ScatterStyle;
  /** Radar style */
  radarStyle?: RadarStyle;
  /** Of-pie type */
  ofPieType?: OfPieType;
  /** Vary colors by point */
  varyColors?: boolean;
  /** 3D view */
  view3D?: View3D;
  /** Display blanks as */
  displayBlanksAs?: DisplayBlanksAs;
  /** Hole size for doughnut (0-90) */
  holeSize?: number;
  /** Style index */
  style?: number;
  /** Wireframe mode for surface charts */
  wireframe?: boolean;
  /** Surface band formats (per level colouring) */
  bandFormats?: Array<{
    /** 0-based band index */
    index: number;
    /** Shape properties for this band */
    spPr?: ShapeProperties | AddShapeFillOptions;
  }>;
  /** Bubble scale (percent) for bubble charts */
  bubbleScale?: number;
  /** Show negative bubbles for bubble charts */
  showNegBubbles?: boolean;
  /** Size represents for bubble charts */
  sizeRepresents?: "area" | "w";
  /** Split type for ofPie charts */
  splitType?: SplitType;
  /** Split position for ofPie charts */
  splitPos?: number;
  /** Second pie size for ofPie charts (percent) */
  secondPieSize?: number;
  /** 3D bar shape */
  shape?: "box" | "cone" | "coneToMax" | "cylinder" | "pyramid" | "pyramidToMax";
  /** Category (X) axis configuration */
  categoryAxis?: AddAxisOptions;
  /** Value (Y) axis configuration */
  valueAxis?: AddAxisOptions;
  /** Data labels for the entire chart type group */
  dataLabels?: AddDataLabelsOptions;
  /** Gap width percentage (bar/column charts, default 150) */
  gapWidth?: number;
  /** Overlap percentage (bar/column charts, -100 to 100) */
  overlap?: number;
  /** Show markers on line/radar (default true for line, false for radar "filled") */
  showMarker?: boolean;
  /** Smooth lines by default for all series (line/scatter) */
  smooth?: boolean;
  /** First slice angle for pie/doughnut charts (0-360, default 0) */
  firstSliceAng?: number;
  /** Show hi-low lines for line/stock charts */
  hiLowLines?: boolean;
  /** Show up-down bars for line/stock charts */
  upDownBars?:
    | boolean
    | {
        gapWidth?: number;
        upBars?: ShapeProperties | AddShapeFillOptions;
        downBars?: ShapeProperties | AddShapeFillOptions;
      };
  /** Show drop lines for line/area charts */
  dropLines?: boolean;
  /** Show series lines for bar/ofPie charts */
  serLines?: boolean;
  /** Show data table below chart */
  dataTable?:
    | boolean
    | {
        showHorzBorder?: boolean;
        showVertBorder?: boolean;
        showOutline?: boolean;
        showKeys?: boolean;
      };
  /** Show data labels over max value */
  showDLblsOverMax?: boolean;
  /** Plot visible cells only (default true) */
  plotVisOnly?: boolean;

  /** Title layout options (position, overlay, spPr/txPr) */
  titleOptions?: AddTitleOptions;
  /** Legend options (layout, entries, spPr/txPr) */
  legendOptions?: AddLegendOptions;
  /** Plot area layout and background */
  plotAreaOptions?: AddPlotAreaOptions;
  /** 3D chart: floor background */
  floor?: ShapeProperties | AddShapeFillOptions;
  /** 3D chart: side wall background */
  sideWall?: ShapeProperties | AddShapeFillOptions;
  /** 3D chart: back wall background */
  backWall?: ShapeProperties | AddShapeFillOptions;
}

/**
 * Simplified fill / line options for chart shapes.
 */
export interface AddShapeFillOptions {
  /** Solid fill color (hex) */
  fill?: string;
  /** Border color (hex) */
  border?: string;
  /** Border width in points */
  borderWidth?: number;
  /** Gradient fill */
  gradient?: ChartFill["gradient"];
  /** Pattern fill */
  pattern?: ChartFill["pattern"];
  /** No fill */
  noFill?: boolean;
}

/**
 * Title layout / formatting options.
 */
export interface AddTitleOptions {
  /** Manual layout (relative or absolute positioning) */
  layout?: ChartLayout;
  /** Whether the title overlays the plot area */
  overlay?: boolean;
  /** Shape properties for the title frame */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties for the title */
  txPr?: ChartTextProperties;
}

/**
 * Legend layout / entry overrides.
 */
export interface AddLegendOptions {
  /** Manual layout */
  layout?: ChartLayout;
  /** Whether the legend overlays the plot area */
  overlay?: boolean;
  /** Per-entry customisations (delete or restyle an entry) */
  entries?: Array<{
    /** 0-based legend entry index */
    index: number;
    /** Hide this entry (leaves the series plotted) */
    hidden?: boolean;
    /** Text properties for this entry */
    txPr?: ChartTextProperties;
  }>;
  /** Shape properties for the legend frame */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties for the legend */
  txPr?: ChartTextProperties;
}

/**
 * Plot area layout / styling.
 */
export interface AddPlotAreaOptions {
  /** Manual layout */
  layout?: ChartLayout;
  /** Shape properties for the plot area background */
  spPr?: ShapeProperties | AddShapeFillOptions;
}

/**
 * Options for a single chart type group within a combo chart.
 * Each group can use its own chart type and optionally bind to a secondary axis.
 */
export interface ComboChartGroupOptions extends AddChartOptions {
  /**
   * When true, the series in this group are plotted against a secondary
   * value axis (right side). The builder will automatically create
   * a secondary catAx/valAx pair and wire them up.
   */
  useSecondaryAxis?: boolean;
}

/**
 * Options for creating a combo chart with multiple overlaid chart type groups.
 *
 * Example: bar + line combo with secondary axis on the line:
 * ```ts
 * worksheet.addComboChart(
 *   {
 *     groups: [
 *       { type: "bar", series: [{ values: "Sheet1!$B$1:$B$5" }] },
 *       { type: "line", series: [{ values: "Sheet1!$C$1:$C$5" }], useSecondaryAxis: true }
 *     ],
 *     title: "Sales vs Growth"
 *   },
 *   "A1:H15"
 * );
 * ```
 */
export interface AddComboChartOptions {
  /** Chart type groups — at least 2 for a combo chart */
  groups: ComboChartGroupOptions[];
  /** Chart title text, formula reference, or structured rich text */
  title?: string | { formula: string } | ChartRichText;
  /** Show legend */
  showLegend?: boolean;
  /** Legend position */
  legendPosition?: LegendPosition;
  /** Display blanks as */
  displayBlanksAs?: DisplayBlanksAs;
  /** Style index */
  style?: number;
  /** Plot visible cells only */
  plotVisOnly?: boolean;
  /** Show data labels over max */
  showDLblsOverMax?: boolean;
  /** Data table */
  dataTable?:
    | boolean
    | {
        showHorzBorder?: boolean;
        showVertBorder?: boolean;
        showOutline?: boolean;
        showKeys?: boolean;
      };
  /** Title layout options */
  titleOptions?: AddTitleOptions;
  /** Legend layout / entry options */
  legendOptions?: AddLegendOptions;
  /** Plot area layout and background */
  plotAreaOptions?: AddPlotAreaOptions;
  /** 3D view */
  view3D?: View3D;
  /** 3D chart floor background */
  floor?: ShapeProperties | AddShapeFillOptions;
  /** 3D chart side wall background */
  sideWall?: ShapeProperties | AddShapeFillOptions;
  /** 3D chart back wall background */
  backWall?: ShapeProperties | AddShapeFillOptions;
}

export interface AddChartSeriesOptions {
  /** Series name or reference */
  name?: string | { formula: string };
  /** Category values reference */
  categories?: string;
  /** Values reference */
  values: string;
  /** X values for scatter/bubble */
  xValues?: string;
  /** Bubble size for bubble charts */
  bubbleSize?: string;
  /** Fill color (hex, e.g. "#FF0000" or "FF0000") */
  fill?: string;
  /** Line color (hex) */
  line?: string;
  /** Line width in points */
  lineWidth?: number;
  /** Dash style for the line */
  lineDash?: ChartLine["dash"];
  /** Marker configuration */
  marker?: AddChartMarkerOptions;
  /** Smooth lines (line/scatter) */
  smooth?: boolean;
  /** Data labels for this series */
  dataLabels?: AddDataLabelsOptions;
  /** Trendline configuration (single or multiple) */
  trendline?: AddTrendlineOptions | AddTrendlineOptions[];
  /** Error bars configuration */
  errorBars?: AddErrorBarsOptions | AddErrorBarsOptions[];
  /** Data point overrides */
  dataPoints?: AddDataPointOptions[];
  /** Invert if negative (bar/bubble) */
  invertIfNegative?: boolean;
  /** Explosion percentage for pie/doughnut */
  explosion?: number;
  /** Bubble 3D effect */
  bubble3D?: boolean;
  /** Picture fill options (bar charts only) */
  pictureFill?: {
    /** Rel ID for a picture blip */
    relationshipId?: string;
    /** How to stretch: stretch (whole bar) or stack (per unit) */
    fillMode?: "stretch" | "stack" | "stackScale";
    /** Picture scale (for stackScale) */
    scale?: number;
    /** Apply to data points / bar sides / front */
    applyToFront?: boolean;
    applyToSides?: boolean;
    applyToEnd?: boolean;
  };
  /** Advanced shape properties (overrides fill/line when set) */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties for this series */
  txPr?: ChartTextProperties;
}

/**
 * Options for a chart marker on a series.
 */
export interface AddChartMarkerOptions {
  /** Marker symbol */
  symbol?: ChartMarker["symbol"];
  /** Marker size (2-72) */
  size?: number;
  /** Fill color (hex) */
  fill?: string;
  /** Border/outline color (hex) */
  border?: string;
}

/**
 * Options for data labels on a series or chart type group.
 */
export interface AddDataLabelsOptions {
  /** Show legend key */
  showLegendKey?: boolean;
  /** Show category name */
  showCatName?: boolean;
  /** Show series name */
  showSerName?: boolean;
  /** Show value */
  showVal?: boolean;
  /** Show percentage (pie/doughnut) */
  showPercent?: boolean;
  /** Show bubble size */
  showBubbleSize?: boolean;
  /** Show leader lines */
  showLeaderLines?: boolean;
  /** Label position */
  position?: DataLabelPosition;
  /** Separator between label parts */
  separator?: string;
  /** Number format */
  numFmt?: string;
  /** Whether number format is linked to source */
  numFmtLinked?: boolean;
  /** Shape properties for label frame */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties for label */
  txPr?: ChartTextProperties;
  /** Per-entry overrides (keyed by 0-based point index) */
  entries?: AddDataLabelEntryOptions[];
}

/**
 * Override for a single data label entry.
 */
export interface AddDataLabelEntryOptions {
  /** 0-based point index */
  index: number;
  /** Hide this entry entirely */
  delete?: boolean;
  /** Custom label text (plain string or rich text) */
  text?: string | ChartRichText;
  /** Position (overrides group-level) */
  position?: DataLabelPosition;
  /** Number format */
  numFmt?: string;
  numFmtLinked?: boolean;
  /** Shape properties */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties */
  txPr?: ChartTextProperties;
  /** Show flags (overrides group-level) */
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLegendKey?: boolean;
}

/**
 * Options for a trendline on a series.
 */
export interface AddTrendlineOptions {
  /** Trendline type */
  type: TrendlineType;
  /** Trendline name (displayed in legend) */
  name?: string;
  /** Polynomial order (2-6, for type "poly") */
  order?: number;
  /** Moving average period (for type "movingAvg") */
  period?: number;
  /** Forward forecast periods */
  forward?: number;
  /** Backward forecast periods */
  backward?: number;
  /** Y-intercept value */
  intercept?: number;
  /** Display R-squared value on chart */
  displayRSqr?: boolean;
  /** Display equation on chart */
  displayEq?: boolean;
  /** Line color (hex) */
  line?: string;
  /** Line width in points */
  lineWidth?: number;
  /** Dash style */
  lineDash?: ChartLine["dash"];
  /** Trendline label (text, layout, style) */
  label?: AddTrendlineLabelOptions;
}

/**
 * Trendline label styling.
 */
export interface AddTrendlineLabelOptions {
  /** Custom label text (structured rich text) */
  text?: ChartRichText;
  /** Number format */
  numFmt?: string;
  numFmtLinked?: boolean;
  /** Layout */
  layout?: ChartLayout;
  /** Shape properties */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties */
  txPr?: ChartTextProperties;
}

/**
 * Options for error bars on a series.
 */
export interface AddErrorBarsOptions {
  /** Error bar direction */
  direction?: ErrorBarDirection;
  /** Which sides to show */
  barDir?: ErrorBarType;
  /** Value type */
  type: ErrorBarValueType;
  /** Fixed value (for type "fixedVal" or "percentage") */
  value?: number;
  /** No end cap */
  noEndCap?: boolean;
  /** Custom plus values formula (for type "cust") */
  plus?: string;
  /** Custom minus values formula (for type "cust") */
  minus?: string;
  /** Line color (hex) */
  line?: string;
  /** Line width in points */
  lineWidth?: number;
  /** Dash style */
  lineDash?: ChartLine["dash"];
  /** Shape properties (advanced — overrides line/lineWidth/lineDash) */
  spPr?: ShapeProperties | AddShapeFillOptions;
}

/**
 * Options for a data point override.
 */
export interface AddDataPointOptions {
  /** 0-based point index */
  index: number;
  /** Fill color (hex) */
  fill?: string;
  /** Border color (hex) */
  border?: string;
  /** Explosion (pie/doughnut) */
  explosion?: number;
  /** Bubble 3D */
  bubble3D?: boolean;
  /** Marker override for this data point */
  marker?: AddChartMarkerOptions;
  /** Invert if negative */
  invertIfNegative?: boolean;
}

/**
 * Axis configuration options for the builder.
 */
export interface AddAxisOptions {
  /** Axis title text */
  title?: string;
  /** Number format (e.g. "#,##0", "0.00%") */
  numFmt?: string;
  /** Whether number format is linked to source */
  numFmtLinked?: boolean;
  /** Minimum value */
  min?: number;
  /** Maximum value */
  max?: number;
  /** Major unit */
  majorUnit?: number;
  /** Minor unit */
  minorUnit?: number;
  /** Major tick mark type */
  majorTickMark?: TickMark;
  /** Minor tick mark type */
  minorTickMark?: TickMark;
  /** Tick label position */
  tickLblPos?: TickLabelPosition;
  /** Show major gridlines */
  majorGridlines?: boolean;
  /** Show minor gridlines */
  minorGridlines?: boolean;
  /** Axis orientation */
  orientation?: AxisOrientation;
  /** Cross between categories or midpoints (for value axis) */
  crossBetween?: "between" | "midCat";
  /** Axis label rotation in degrees (-90 to 90) */
  textRotation?: number;
  /** Delete/hide the axis */
  hidden?: boolean;
  /** Logarithmic base (e.g. 10) */
  logBase?: number;
  /** Label alignment (category axis only) */
  lblAlgn?: "ctr" | "l" | "r";
  /** Label offset percentage (category/date axis, default 100) */
  lblOffset?: number;
  /** Skip every N tick labels */
  tickLblSkip?: number;
  /** Skip every N tick marks */
  tickMarkSkip?: number;
  /** Cross axis at this value ("autoZero", "min", "max") */
  crosses?: "autoZero" | "min" | "max";
  /** Cross axis at a specific numeric value */
  crossesAt?: number;
  /** Display units for value axis */
  displayUnits?: DisplayUnits["builtInUnit"];
  /** Axis line color (hex) */
  lineColor?: string;
  /** Axis line width in points */
  lineWidth?: number;
  /** Line dash style */
  lineDash?: ChartLine["dash"];
  /** Custom display unit value (for value axis). */
  customUnit?: number;
  /** Display unit label (shown near axis when displayUnits is set) */
  displayUnitsLabel?: string | ChartRichText;
  /** Base time unit (for date axis: days/months/years) */
  baseTimeUnit?: "days" | "months" | "years";
  /** Major time unit (for date axis) */
  majorTimeUnit?: "days" | "months" | "years";
  /** Minor time unit (for date axis) */
  minorTimeUnit?: "days" | "months" | "years";
  /** Shape properties for the axis line and tick marks (advanced) */
  spPr?: ShapeProperties | AddShapeFillOptions;
  /** Text properties for tick labels */
  txPr?: ChartTextProperties;
  /** Rich styling for the axis title (overrides `title` string form) */
  titleOptions?: AddTitleOptions;
  /** Major gridlines styling (pass shape properties to style gridlines) */
  majorGridlinesStyle?: ShapeProperties | AddShapeFillOptions;
  /** Minor gridlines styling */
  minorGridlinesStyle?: ShapeProperties | AddShapeFillOptions;
}

/**
 * Anchor range input for addChart.
 *
 * Supported forms:
 * - **String**: `"A1:H15"` (two-cell anchor) or `"A1"` (defaults to 10×15 cells).
 * - **Two-cell**: `{ tl, br, editAs? }` — top-left and bottom-right cells.
 * - **One-cell**: `{ tl, ext }` — top-left cell plus absolute extent (EMU).
 * - **Absolute**: `{ pos, ext }` — absolute position and extent in EMU.
 *
 * For EMU conversions: 914400 EMU = 1 inch. Typical chart 4×3 inches =
 * `{ cx: 3657600, cy: 2743200 }`.
 */
export type AddChartRange =
  | string
  | {
      /** Two-cell anchor: top-left and bottom-right */
      tl: { col: number; row: number } | string;
      br: { col: number; row: number } | string;
      editAs?: "oneCell" | "twoCell" | "absolute";
    }
  | {
      /** One-cell anchor: top-left + fixed extent */
      tl: { col: number; row: number } | string;
      ext: { cx: number; cy: number };
      editAs?: "oneCell";
    }
  | {
      /** Absolute anchor: position + extent, both in EMU */
      pos: { x: number; y: number };
      ext: { cx: number; cy: number };
      editAs?: "absolute";
    };
