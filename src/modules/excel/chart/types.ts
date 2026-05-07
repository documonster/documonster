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
  /**
   * DrawingML scheme colour name that could not be mapped onto a theme
   * index (e.g. `"phClr"` — the DrawingML "placeholder colour" token —
   * or a future / vendor addition). When set, the writer emits
   * `<a:schemeClr val="…">` (NOT `<a:sysClr>`); this preserves the
   * original element identity on round-trip. Prefer {@link theme} when
   * the name is one of the 12 canonical scheme slots.
   */
  schemeName?: string;
  /** System color (e.g. "windowText", "window") */
  sysClr?: string;
  /** Preset color name (e.g. "black", "white") */
  prstClr?: string;
  /**
   * Tint modifier. Stored as a **fraction in 0..1** (e.g. `0.5` for
   * 50% tint); the writer multiplies by `100000` to produce the OOXML
   * `<a:tint val="…"/>` integer.
   *
   * ⚠ This convention differs from {@link ChartColorsEntry.tint} and
   * {@link ChartColorVariation.tint}, which store the **raw OOXML
   * integer** (0..100000) because those palette types pass the DrawingML
   * modifier value through unchanged. Copying a value between the two
   * shapes without conversion will land the wrong colour — use
   *
   *     chartColorsEntry.tint = Math.round(chartColor.tint * 100000)
   *     chartColor.tint       = chartColorsEntry.tint / 100000
   *
   * at the boundary.
   */
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
    /**
     * Gradient colour stops. `position` is a **fraction** in the range
     * `[0, 1]` (0 = start, 1 = end), regardless of OOXML's on-disk
     * encoding which uses hundredths-of-a-percent (`0`–`100000`). The
     * renderer converts to/from the wire format.
     */
    stops: Array<{ position: number; color: ChartColor }>;
    /**
     * Linear-gradient angle in **degrees** (0 = left-to-right, 90 =
     * top-to-bottom). Ignored for `circle` / `rect` / `shape` paths.
     * OOXML stores this as 60000ths of a degree; the renderer converts.
     */
    angle?: number;
    /**
     * DrawingML `<a:lin/@scaled>` — when `true` (the default Excel
     * emits), the angle scales with the shape's aspect ratio; when
     * `false`, the angle is independent of shape geometry. The
     * difference is visible: a 45° gradient on a 2:1 rectangle aims
     * at a different corner depending on this flag. The writer
     * defaults to `true` when omitted (matching Excel), so only
     * `false` changes the on-disk bytes; but preserving the authored
     * value is needed to avoid silently drifting the flag on
     * round-trip.
     */
    scaled?: boolean;
    /** Linear, circle, rect, shape */
    type?: "linear" | "circle" | "rect" | "shape";
    /**
     * Focal rectangle for non-linear gradients. Components are
     * fractions in `[0, 1]` (left/top/right/bottom insets from the
     * bounding box). Absent or all-zero means the focal centre is the
     * shape centre — OOXML's `<a:fillToRect l="50000" t="50000"
     * r="50000" b="50000"/>` equivalent. The renderer converts to
     * hundredths-of-a-percent on write.
     */
    fillToRect?: { left?: number; top?: number; right?: number; bottom?: number };
  };
  /**
   * Picture (blip) fill — `<a:blipFill>/<a:blip r:embed="rIdN"/>`.
   *
   * {@link ChartBlipFill.relationshipId} is the final wire value; when
   * the caller has not yet registered the image with the chart part,
   * `_pendingImage` carries the raw data until the worksheet-side
   * `_registerChart` path wires up the media entry and chart rel.
   */
  blip?: ChartBlipFill;
}

/**
 * Picture fill for chart shapes — serialises as `<a:blipFill>`.
 *
 * `relationshipId` is the authoritative output: once present, the
 * writer emits `<a:blip r:embed="rIdN"/>`. `_pendingImage` is a
 * construction-time staging slot consumed by the worksheet registration
 * path — after `_registerChart` runs, it is removed and
 * `relationshipId` is set.
 */
export interface ChartBlipFill {
  /** Chart-part relationship id referring to the image. */
  relationshipId?: string;
  /**
   * Stretch rectangle (`<a:srcRect l t r b>`) with fractional
   * (per-mille, 0-100000) insets. Absent for a full-bleed fill.
   */
  sourceRectangle?: { left?: number; top?: number; right?: number; bottom?: number };
  /**
   * Tile or stretch mode. `stretch` maps to `<a:stretch><a:fillRect/></a:stretch>`;
   * `tile` maps to `<a:tile/>` with the given tiling parameters; `none`
   * emits neither child (which Excel treats as stretch).
   */
  fillMode?: "stretch" | "tile" | "none";
  /** Tile options, honoured only when {@link fillMode} is `"tile"`. */
  tile?: {
    tx?: number;
    ty?: number;
    sx?: number;
    sy?: number;
    flip?: "none" | "x" | "y" | "xy";
    alignment?: string;
  };
  /**
   * @internal Staging payload consumed by the worksheet-side chart
   * registration path. Never serialised directly; see
   * {@link AddChartSeriesOptions.pictureFill.image}.
   */
  _pendingImage?: AddChartPictureFillImage;
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
  /**
   * `a:xfrm` — position / size / rotation / flip. Structured version
   * of the transform Excel emits on shapes that have been manually
   * positioned on the chart. When absent the shape uses its parent's
   * automatic layout.
   */
  transform?: ShapeTransform;
  /**
   * `a:prstGeom` — the preset shape geometry (rectangle, roundRect,
   * ellipse, arrow, callout, etc.) used by non-rectangular chart
   * shapes. Present on `c:shapeGroupSprite`-style chart decorations;
   * charts themselves default to `rect` and omit this element.
   */
  presetGeometry?: PresetGeometry;
  /**
   * `a:custGeom` — freeform geometry for fully custom shapes. Only
   * meaningful for shapes the user drew with the Excel freeform tool.
   * Structured access covers path segments; the `paths` array mirrors
   * the OOXML `<a:pathLst>` container.
   */
  customGeometry?: CustomGeometry;
  /** @internal Raw XML string for perfect round-trip fidelity */
  _rawXml?: string;
}

/**
 * `a:xfrm` — shape transform: origin, extent, rotation, flip.
 *
 * OOXML coordinates are in EMUs (1/914400 inch) for position/extent and
 * in 1/60000 of a degree for `rotation`. excelts exposes the raw units
 * so round-trip is lossless; helpers in `shape-properties.ts` convert
 * to / from points and degrees when a caller prefers friendlier units.
 */
export interface ShapeTransform {
  /** `<a:off x y>` — top-left origin in EMU. */
  offsetX?: number;
  offsetY?: number;
  /** `<a:ext cx cy>` — width / height in EMU. */
  width?: number;
  height?: number;
  /**
   * `@rot` on `a:xfrm` — rotation in 1/60000 of a degree (positive
   * values rotate clockwise). Absent means "not rotated".
   */
  rotation?: number;
  /** `@flipH="1"` — horizontal flip. */
  flipHorizontal?: boolean;
  /** `@flipV="1"` — vertical flip. */
  flipVertical?: boolean;
}

/**
 * `a:prstGeom` — preset shape geometry.
 *
 * `preset` is one of the ~180 DrawingML preset shape names
 * (`"rect"`, `"roundRect"`, `"ellipse"`, `"rightArrow"`, `"cloud"`…).
 * `adjustments` are the shape-specific `gd` parameters Excel exposes
 * as the little yellow diamond handles on interactive shapes; each
 * entry has the symbolic name (e.g. `"adj1"`) and a per-mille
 * integer value.
 */
export interface PresetGeometry {
  preset: string;
  adjustments?: Array<{ name: string; fmla: string }>;
}

/**
 * `a:custGeom` — custom freeform geometry.
 *
 * Represented at path granularity: each path carries its own bounding
 * box (`w` / `h`) and the ordered list of drawing commands. This
 * matches Excel's serialisation closely — a preview renderer can walk
 * `commands` linearly and emit SVG `<path d="…">` data.
 */
export interface CustomGeometry {
  /**
   * `a:pathLst/a:path` — one or more subpaths. Each subpath declares
   * its own logical coordinate space via {@link CustomGeometryPath.w}
   * and {@link CustomGeometryPath.h}; command coordinates are
   * relative to that space.
   */
  paths?: CustomGeometryPath[];
  /**
   * `a:avLst` — adjustment values, as on {@link PresetGeometry}. Very
   * rare on custom geometry but retained for byte-preserving round
   * trip.
   */
  adjustments?: Array<{ name: string; fmla: string }>;
}

export interface CustomGeometryPath {
  /** Width of the path's local coordinate space in EMU. */
  w?: number;
  /** Height of the path's local coordinate space in EMU. */
  h?: number;
  /** Fill mode — maps to `a:path/@fill`. */
  fill?: "none" | "norm" | "lighten" | "darken" | "lightenLess" | "darkenLess";
  /** Stroke flag — `a:path/@stroke="1"`. */
  stroke?: boolean;
  /**
   * Drawing commands in order: `moveTo`, `lnTo`, `arcTo`, `cubicBezTo`,
   * `quadBezTo`, `close`. Each command's `points` array carries the
   * 0-2 control points in local path coordinates.
   */
  commands: CustomGeometryCommand[];
}

export interface CustomGeometryCommand {
  type: "moveTo" | "lnTo" | "arcTo" | "cubicBezTo" | "quadBezTo" | "close";
  /** Points consumed by the command (empty for `close`). */
  points?: Array<{ x: number; y: number }>;
  /**
   * For `arcTo`: `wR`, `hR`, `stAng`, `swAng`. OOXML keeps arc
   * parameters as explicit named attributes rather than as points, so
   * they ride here to preserve the original shape.
   */
  arcParams?: { wR: number; hR: number; stAng: number; swAng: number };
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
  /**
   * Suppress this whole `c:dLbls` block. OOXML expresses this as the
   * `<c:delete val="1"/>` choice-left branch of `CT_DLbls`; when set,
   * only `c:delete` (and `c:extLst`) is emitted. Typical usage: turn
   * off the default data labels inherited from a chart style / theme
   * while keeping per-point overrides in {@link entries} visible.
   */
  delete?: boolean;
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
  /**
   * Excel 2013+ "Value From Cells" — corresponds to the
   * `c15:datalabelsRange` extension emitted inside
   * `c:dLbls/c:extLst/c:ext[uri={CE6537A1-…}]`.
   *
   * When set, each data label displays the text from the corresponding
   * cell in {@link DataLabelsRange.formula} instead of (or in addition
   * to) the series/category/value drawn from the point itself. This is
   * the feature Excel surfaces as the "Value From Cells" checkbox in
   * the Format Data Labels pane.
   *
   * The matching per-point `<c:dLbl>` entries would traditionally
   * hold `<c:tx><c:rich>…</c:rich></c:tx>` with the cached label
   * string so viewers that don't understand the extension still
   * render something sensible. This library currently emits the
   * extension (`<c15:datalabelsRange>` inside `c:dLbls/c:extLst`)
   * with its cache, but does NOT auto-generate placeholder
   * per-point `<c:dLbl>` entries — Excel and every modern viewer
   * reads the `c15:` extension directly, so the fallback placeholders
   * are optional. Callers that need the fallback for a legacy
   * consumer can populate {@link DataLabels.entries} manually with
   * the same cached values.
   */
  dataLabelsRange?: DataLabelsRange;
  /** Extension list (raw XML for round-trip) */
  extLst?: string;
}

/**
 * "Value From Cells" range for a data-label series. Points at a worksheet
 * range whose values are used as the text of each data label, matching
 * the `c15:datalabelsRange` extension (Office 2013+).
 */
export interface DataLabelsRange {
  /** Excel formula referring to the cells whose values become labels. */
  formula: string;
  /**
   * Optional cached values — one entry per source point, parallel to the
   * series' data. Writers that want byte-preserving round-trip should
   * populate this from the worksheet before serialising so readers
   * without a formula engine still see the right labels; the
   * cache-populator auto-fills it when left empty.
   */
  cache?: StringCache;
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
  /**
   * Which error-bar cap to show (x / y axis): `"x"`, `"y"` or both.
   * Maps to OOXML `c:errDir`.
   */
  errDir?: ErrorBarDirection;
  /**
   * Whether the bar extends above (`"plus"`), below (`"minus"`) or
   * both (`"both"`) the data point. The field is named `barDir` for
   * historical reasons; the matching OOXML element is `c:errBarType`
   * (see `ST_ErrBarType`). It is NOT the same kind of value as the
   * chart-level {@link BarChartGroup.barDir} (`"col"` / `"bar"`).
   */
  barDir: ErrorBarType;
  /**
   * How `val` / `plus` / `minus` are interpreted — fixed value,
   * percentage, standard deviation, standard error or custom range.
   */
  errValType: ErrorBarValueType;
  /** Suppress the end-cap (`<c:noEndCap val="1"/>`). */
  noEndCap?: boolean;
  /**
   * Error magnitude for `fixedVal` / `percentage` / `stdDev`. For
   * `"stdErr"` this is ignored (Excel computes it from the series);
   * for `"cust"` {@link plus} and {@link minus} are required instead.
   */
  val?: number;
  /** Positive-direction reference (used with `errValType === "cust"`). */
  plus?: NumberDataSource;
  /** Negative-direction reference (used with `errValType === "cust"`). */
  minus?: NumberDataSource;
  /** Error-bar shape properties (colour, line width, end-cap style). */
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
  /**
   * Picture fill options — `CT_AreaSer/c:pictureOptions`. Controls
   * how a texture / blip fill applies across the area
   * (`pictureFormat`: `stretch` | `stack` | `stackScale`, etc.).
   * Parsed from source XML by `_processSeries`; round-tripped to
   * preserve texture-filled areas.
   */
  pictureOptions?: PictureOptions;
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
  /**
   * Depth-direction gap for 3D bar charts (`c:gapDepth`). Expressed as
   * a percentage (0-500) — 150 is Excel's default. Ignored for 2D
   * bar charts.
   */
  gapDepth?: number;
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
  /** Depth-direction gap for 3D line charts (`c:gapDepth`, 0-500). */
  gapDepth?: number;
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
  /** Depth-direction gap for 3D area charts (`c:gapDepth`, 0-500). */
  gapDepth?: number;
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
  // `CT_StockChart` in the Chart2014 schema does NOT carry a
  // `varyColors` attribute — the lines are always single-coloured per
  // series. Previous versions of this library accepted the option and
  // emitted `<c:varyColors>`, which LibreOffice's strict validator
  // rejects. The field is intentionally absent.
  series: StockSeries[];
  dataLabels?: DataLabels;
  /**
   * Per ECMA-376 `CT_StockChart` sequence: `dropLines` appears **before**
   * `hiLowLines`. The previous model had no notion of ordering
   * because it used object field declaration order — the writer took
   * advantage of that to emit the correct sequence, so only the
   * serialiser cares. Keep the TypeScript declaration order matching
   * the schema for clarity.
   */
  dropLines?: ShapeProperties;
  hiLowLines?: ShapeProperties;
  upDownBars?: UpDownBars;
  axisIds: number[];
  extLst?: string;
}

export interface SurfaceChartGroup {
  type: "surface" | "surface3D";
  wireframe?: boolean;
  series: SurfaceSeries[];
  // NOTE: `CT_SurfaceChart` has no `dLbls` child per ECMA-376 §21.2.2.204
  // (it only allows `wireframe, ser*, bandFmts?, axId{2,3}`). The
  // builder's validator explicitly rejects `opts.dataLabels` for
  // surface charts, and the parser now drops any parsed `c:dLbls`
  // encountered under `c:surfaceChart` rather than letting it escape
  // into the model. Keeping the field off the type ensures programmatic
  // callers can't route a data-labels mutation here and have the
  // writer silently emit invalid OOXML.
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
  /** Extension list (raw XML for round-trip). */
  extLst?: string;
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
  /** Extension list (raw XML for round-trip). */
  extLst?: string;
}

// ============================================================================
// Layout
// ============================================================================

/**
 * Manual layout positioning.
 */
export interface ChartLayout {
  manualLayout?: ManualLayout;
  /**
   * @internal Raw XML preserved for round-trip of layout variants the
   * structured model does not (yet) cover — used by the ChartEx parser
   * on `cx:layout` where the DrawingML-like layout syntax differs from
   * `c:layout / c:manualLayout`. Not intended for public consumption.
   */
  _rawXml?: string;
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
  /**
   * Horizontal / vertical alignment within the position slot. Chart2014
   * `CT_Legend/@align` admits `"ctr"` (centred — the default Excel
   * always emits), `"l"` / `"r"` (left / right within horizontal
   * positions), or `"t"` / `"b"` (top / bottom within vertical
   * positions). Absent means inherit the default (`"ctr"`).
   */
  align?: "ctr" | "l" | "r" | "t" | "b";
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
  /** Extension list (raw XML for round-trip). */
  extLst?: string;
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
  /**
   * Structured representation of the single `c:dLbl` that pivot-chart
   * samples from Excel emit directly inside `c:pivotFmt` (without the
   * enclosing `c:dLbls` wrapper). Most callers will leave this
   * undefined — the field exists so editors can mutate pivot-series
   * label styling without dropping into raw XML.
   */
  dLbl?: DataLabelEntry;
  /**
   * @deprecated Use {@link dLbl} for structured access. Retained so
   * files parsed before the structured slot landed continue to
   * round-trip. Writers prefer `dLbl` when both are present.
   */
  rawDLbl?: string;
  /** Extension list (raw XML for round-trip). */
  extLst?: string;
}

/**
 * Structured metadata for a pivot chart, corresponding to MS Office 2010+
 * `c14:pivotOptions` extension (ECMA-376 MS-XLSX §2.3.11, namespace
 * `http://schemas.microsoft.com/office/drawing/2007/8/2/chart`).
 *
 * Written as `c:chartSpace/c:extLst/c:ext/c14:pivotOptions`, so Excel
 * recognises the metadata on load; prior versions of excelts wrote this
 * data under a private `excelts:` namespace that Excel silently discarded.
 *
 * All `dropZone*` fields are boolean flags controlling whether the
 * corresponding family of PivotTable fields gets drop-zone controls on
 * the chart when {@link dropZonesVisible} is enabled. Granularity is
 * per-axis-category, not per-field — this matches the OOXML schema. The
 * earlier `fieldButtons[]` / `filters[]` array shape never matched any
 * part of the OOXML grammar and has been removed in this release.
 */
export interface PivotChartOptions {
  /**
   * Whether any drop-zone controls can appear on the pivot chart. When
   * `false`, none of the other `dropZone*` flags have any visual effect.
   * Absent (`undefined`) means Excel uses its default (typically `true`).
   */
  dropZonesVisible?: boolean;
  /**
   * Whether a control for each PivotTable field on the Filter (page) axis
   * of the source PivotTable appears on the chart when
   * {@link dropZonesVisible} is `true`.
   */
  dropZoneFilter?: boolean;
  /**
   * Whether a control for each PivotTable field on the Row axis of the
   * source PivotTable appears on the chart when {@link dropZonesVisible}
   * is `true`.
   */
  dropZoneCategories?: boolean;
  /**
   * Whether a control for each PivotTable field on the Data (values) axis
   * of the source PivotTable appears on the chart when
   * {@link dropZonesVisible} is `true`.
   */
  dropZoneData?: boolean;
  /**
   * Whether a control for each PivotTable field on the Column axis of the
   * source PivotTable appears on the chart when {@link dropZonesVisible}
   * is `true`.
   */
  dropZoneSeries?: boolean;
  /**
   * Whether Excel should refresh the linked pivot cache when opening the
   * workbook. Mapped onto `pivotCacheDefinition/@refreshOnLoad` in the
   * sheet parts, not into `c14:pivotOptions`.
   */
  refreshOnOpen?: boolean;
  /**
   * Whether the pivot chart shows the expand/collapse field buttons
   * introduced in Office 2014+. Serialised as the
   * `c16:showExpandCollapseFieldButtons` child inside the
   * `c:chartSpace/c:extLst/c:ext[uri={E28EC0CA-…}]/c16:pivotOptions16`
   * extension — a separate extension from the 2010 `c14:pivotOptions`.
   *
   * Both extensions can coexist on the same chart: the c14 block
   * controls the drop-zone visibility, the c16 block controls the
   * expand/collapse affordances on field buttons.
   */
  showExpandCollapseFieldButtons?: boolean;
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
  /**
   * `c:userShapes` — user-drawn annotation shapes overlaid on the chart.
   *
   * OOXML models these as a reference to a separate drawing part
   * (`drawings/drawingN.xml`). Full structural support would require
   * reproducing the entire DrawingML shape subsystem inside the chart
   * pipeline; for now we keep the `r:id` so the reference survives
   * round-trip, and the referenced part is preserved alongside the
   * chart rels.
   *
   * The target drawing part itself is carried through untouched via
   * the standard chart rels mechanism — this field only captures the
   * relationship id so we know it's there.
   */
  userShapesRelId?: string;
  /** Pivot source information (raw XML for round-trip) */
  pivotSource?: string;
  /**
   * Structured pivot chart options corresponding to the MS Office 2010+
   * `c14:pivotOptions` extension. Parsed from and serialised into
   * `c:chartSpace/c:extLst/c:ext[uri=…chart]/c14:pivotOptions` — the only
   * location Excel recognises for these settings.
   *
   * Exists only on pivot charts (i.e. when {@link pivotSource} is set);
   * writers emit nothing when `pivotSource` is absent.
   */
  pivotOptions?: PivotChartOptions;
  /** Color map override (raw XML for round-trip) */
  clrMapOvr?: string;
  /** Chart protection settings (raw XML for round-trip) */
  protection?: string;
  /** Extension list at c:chartSpace level (raw XML for round-trip) */
  extLst?: string;
  /** Extra namespace declarations for round-trip fidelity */
  extraNamespaces?: Record<string, string>;
  /**
   * Vendor-extension elements observed while parsing that are not part of
   * the `c:` / `a:` / `r:` OOXML namespaces and are not already captured
   * via `extLst` / raw XML targets. Mirrors
   * {@link ChartExModel.unknownElements} so `strictTemplateMode` can warn
   * when a structural rebuild would drop them. Purely informational in
   * the default `preserve` mode.
   */
  unknownElements?: ChartUnknownElement[];
}

/**
 * Describes one unstructured child element discovered while parsing a
 * classic chart part. `path` uses `/` as the separator relative to the
 * nearest `c:` ancestor that was already recognised, matching the format
 * of {@link ChartExUnknownElement}.
 */
export interface ChartUnknownElement {
  /** Fully-qualified element name (e.g. `c15:customTag`). */
  name: string;
  /** Slash-separated breadcrumb, e.g. `c:chartSpace/c15:customTag`. */
  path: string;
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
  /**
   * Structured per-element style definitions (Office 2013+).
   *
   * Each key is one of the ~25 well-known `cs:*` children of
   * `cs:chartStyle` — `categoryAxis`, `chartArea`, `dataLabel`,
   * `dataPoint`, `dataPointLine`, `dataPointMarker`, `gridlineMajor`,
   * `legend`, `plotArea`, `title`, `trendline` and so on (see
   * ECMA-376 Part 1 §21.2). Each definition carries structured slots
   * for the `cs:*Ref` indices that Excel looks up from the chart's
   * theme, plus verbatim raw XML for `spPr` / `defRPr` / `bodyPr` /
   * `fontRef` so the DrawingML sub-tree survives round-trip without
   * this module needing to model every DrawingML child.
   *
   * When absent, the writer falls back to `rawXml` (for round-trip
   * fidelity of files loaded with an unrecognised style) or to the
   * legacy id-only form (`<cs:chartStyle id="N"/>`).
   */
  elements?: Record<string, ChartStyleElement>;
}

/**
 * Per-element entry inside `<cs:chartStyle>`. The four `*RefIdx`
 * fields mirror Excel's theme indices for lines (`cs:lnRef`), fills
 * (`cs:fillRef`), effects (`cs:effectRef`), and the meta slot that
 * governs font (`cs:fontRef`). Everything else is kept as raw XML so
 * callers can round-trip unfamiliar DrawingML without losing bytes.
 */
export interface ChartStyleElement {
  /** `cs:lnRef/@idx` — theme line index. */
  lnRefIdx?: number;
  /** `cs:fillRef/@idx` — theme fill index. */
  fillRefIdx?: number;
  /** `cs:effectRef/@idx` — theme effect index. */
  effectRefIdx?: number;
  /**
   * `cs:fontRef/@idx` — either `minor`, `major`, or `none` for the
   * theme font slot.
   */
  fontRefIdx?: "minor" | "major" | "none";
  /** Verbatim `cs:fontRef` inner XML (colour reference + modifiers). */
  fontRefBody?: string;
  /** Verbatim `cs:spPr` XML (lines/fill/effects on the element). */
  spPrXml?: string;
  /** Verbatim `cs:defRPr` XML (default run properties for text). */
  defRPrXml?: string;
  /** Verbatim `cs:bodyPr` XML (text-body properties). */
  bodyPrXml?: string;
  /**
   * Attributes on the element itself other than the structured ones
   * above, preserved as a key/value map (e.g. `mods` on `chartArea`).
   */
  attributes?: Record<string, string>;
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
  /**
   * `cs:variation` blocks — structured per-index colour modifiers that
   * Excel uses to generate a gradient of related colours from the main
   * palette. Each variation carries a list of DrawingML modifiers
   * (tint / shade / satMod / lumMod / alpha …) that apply on top of
   * the `colors` palette at the matching index.
   *
   * Variations are emitted after the palette entries inside
   * `<cs:colorStyle>`. When this array is absent the writer falls back
   * to preserving anything present in {@link rawXml} so legacy round
   * trips are not disturbed.
   */
  variations?: ChartColorVariation[];
}

/**
 * A single `<cs:variation>` block — a list of colour modifiers applied
 * at the matching palette index. The modifier names mirror the DrawingML
 * schemeClr / srgbClr children and are stored as raw integer values in
 * the OOXML per-mille scale (0-100000).
 */
export interface ChartColorVariation {
  /** `<a:lumMod val="…"/>` — luminance modulation. */
  lumMod?: number;
  /** `<a:lumOff val="…"/>` — luminance offset. */
  lumOff?: number;
  /**
   * `<a:tint val="…"/>` as the **raw OOXML integer** (0..100000). ⚠
   * {@link ChartColor.tint} uses a 0..1 fraction; do not assign between
   * the two types without scaling (×100000 / ÷100000).
   */
  tint?: number;
  /** `<a:shade val="…"/>`. */
  shade?: number;
  /** `<a:satMod val="…"/>`. */
  satMod?: number;
  /** `<a:alpha val="…"/>`. */
  alpha?: number;
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
  /**
   * Tint as the **raw OOXML integer** (0..100000). ⚠ {@link ChartColor.tint}
   * uses a 0..1 fraction; do not assign between the two types without
   * scaling (×100000 / ÷100000).
   */
  tint?: number;
  /** Shade */
  shade?: number;
  /** Saturation modulation */
  satMod?: number;
  /** Alpha (0-100000) */
  alpha?: number;
}

/** Pivot source information written as c:pivotSource for classic pivot charts. */
export type PivotChartSource =
  | string
  | {
      /** Pivot table name, e.g. `PivotTable1` or `[Book.xlsx]Pivot!PivotTable1`. */
      name: string;
      /** Pivot chart format id. Defaults to 0. */
      fmtId?: number;
      /** Structured pivot chart metadata emitted as a chart-space extension. */
      options?: PivotChartOptions;
    };

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
  /**
   * Chart title. Accepts:
   *   - `string` — literal title text
   *   - `{ formula: string }` — formula reference resolved at read time
   *   - {@link ChartRichText} — structured rich text for per-run formatting
   *   - `null` — explicitly suppress the title (Excel will NOT
   *     auto-generate one; `autoTitleDeleted="1"` is emitted)
   *
   * Omit the option entirely to let Excel auto-title the chart per its
   * default behaviour.
   */
  title?: string | { formula: string } | ChartRichText | null;
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
  /** Modern chart style sidecar written to xl/charts/styleN.xml */
  chartStyle?: ChartStyleModel;
  /** Modern chart colors sidecar written to xl/charts/colorsN.xml */
  chartColors?: ChartColorsModel;
  /** Pivot table source for creating a classic pivot chart. */
  pivotSource?: PivotChartSource;
  /** Pivot chart field buttons, filters, and refresh metadata. */
  pivotChartOptions?: PivotChartOptions;
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
  /**
   * Gap depth percentage for 3-D charts (`c:gapDepth`, 0-500). Only
   * valid on `bar3D` / `line3D` / `area3D`; rejected on every other
   * type (including `pie3D`, which has no `gapDepth` child in its
   * `CT_Pie3DChart` definition despite the name suggesting it). Sets
   * the depth-direction spacing between series in the z-axis
   * extrusion.
   */
  gapDepth?: number;
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

type CommonAddChartOptions = Omit<
  AddChartOptions,
  | "type"
  | "series"
  | "barDir"
  | "grouping"
  | "scatterStyle"
  | "radarStyle"
  | "ofPieType"
  | "holeSize"
  | "wireframe"
  | "bandFormats"
  | "bubbleScale"
  | "showNegBubbles"
  | "sizeRepresents"
  | "splitType"
  | "splitPos"
  | "secondPieSize"
  | "shape"
  | "firstSliceAng"
  | "gapWidth"
  | "overlap"
  | "dataLabels"
  | "showMarker"
  | "smooth"
>;

export type AddBarChartSeriesOptions = Omit<
  AddChartSeriesOptions,
  "xValues" | "bubbleSize" | "bubble3D" | "explosion"
>;

export interface AddBarChartOptions extends CommonAddChartOptions {
  type?: "bar" | "bar3D";
  series?: AddBarChartSeriesOptions[];
  grouping?: BarGrouping;
  barDir?: BarDirection;
  dataLabels?: AddDataLabelsOptions;
  gapWidth?: number;
  overlap?: number;
  shape?: BarShape;
}

export type AddPieChartSeriesOptions = Omit<
  AddChartSeriesOptions,
  "xValues" | "bubbleSize" | "bubble3D" | "trendline" | "errorBars" | "pictureFill"
>;

export interface AddPieChartOptions extends CommonAddChartOptions {
  type?: "pie" | "pie3D" | "doughnut" | "ofPie";
  series?: AddPieChartSeriesOptions[];
  holeSize?: number;
  firstSliceAng?: number;
  dataLabels?: AddDataLabelsOptions;
  gapWidth?: number;
  ofPieType?: OfPieType;
  splitType?: SplitType;
  splitPos?: number;
  secondPieSize?: number;
}

export type AddScatterChartSeriesOptions = Omit<
  AddChartSeriesOptions,
  "categories" | "bubbleSize" | "bubble3D" | "pictureFill" | "explosion"
>;

export interface AddScatterChartOptions extends CommonAddChartOptions {
  type?: "scatter";
  series?: AddScatterChartSeriesOptions[];
  scatterStyle?: ScatterStyle;
  dataLabels?: AddDataLabelsOptions;
  showMarker?: boolean;
  smooth?: boolean;
}

export type AddSurfaceChartSeriesOptions = Omit<
  AddChartSeriesOptions,
  "dataLabels" | "trendline" | "errorBars" | "marker" | "bubbleSize" | "bubble3D"
>;

export interface AddSurfaceChartOptions extends CommonAddChartOptions {
  type?: "surface" | "surface3D";
  series?: AddSurfaceChartSeriesOptions[];
  wireframe?: boolean;
  bandFormats?: AddChartOptions["bandFormats"];
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
export interface AddComboChartOptions extends Pick<
  AddChartOptions,
  | "title"
  | "showLegend"
  | "legendPosition"
  | "displayBlanksAs"
  | "style"
  | "chartStyle"
  | "chartColors"
  | "pivotSource"
  | "pivotChartOptions"
  | "plotVisOnly"
  | "showDLblsOverMax"
  | "dataTable"
  | "titleOptions"
  | "legendOptions"
  | "plotAreaOptions"
  | "view3D"
  | "floor"
  | "sideWall"
  | "backWall"
> {
  /** Chart type groups — at least 2 for a combo chart */
  groups: ComboChartGroupOptions[];
}

/**
 * High-level image source for {@link AddChartSeriesOptions.pictureFill}'s
 * `image` field.
 *
 * The various shapes exist so callers don't have to marshal their native
 * data just to fill a bar:
 *
 *   - `Uint8Array` — raw binary; extension is sniffed from magic bytes
 *     (PNG / JPEG / GIF) with PNG as fallback.
 *   - `string` — either a `data:image/<type>;base64,…` URL or a bare
 *     base64 payload (extension inferred from the data URL prefix when
 *     present, else PNG).
 *   - `ChartPictureFillImageData` — structured object with explicit
 *     `extension` + either buffer or base64.
 *   - `{ workbookImageId: number }` — points at an image previously
 *     registered via {@link Workbook.addImage}; no new media entry is
 *     allocated, just a new chart rel.
 */
export type AddChartPictureFillImage =
  | Uint8Array
  | string
  | ChartPictureFillImageData
  | { workbookImageId: number };

/**
 * Structured variant of {@link AddChartPictureFillImage} that mirrors
 * the worksheet `ImageData` shape but keeps this types module independent
 * of the worksheet image types. The worksheet-side registration path
 * accepts both shapes interchangeably.
 */
export interface ChartPictureFillImageData {
  /** Image extension — used for media filename and content type. */
  extension: "png" | "jpeg" | "gif";
  /** Raw binary payload (preferred). */
  buffer?: Uint8Array;
  /** Bare base64 (no data: URL prefix). */
  base64?: string;
}

export interface AddChartSeriesOptions {
  /** Series name or reference */
  name?: string | { formula: string };
  /** Category values reference or structured category data source. */
  categories?: string | AxisDataSource;
  /** Values reference */
  values: string;
  /** X values for scatter/bubble. */
  xValues?: string | AxisDataSource;
  /**
   * Semantic type of {@link xValues} — controls whether a `string`
   * reference is wrapped as a `numRef` (default, matches OOXML scatter
   * regular usage) or a `strRef` (text-categorical x axis, produces a
   * cat axis even on a scatter chart).
   *
   * - `"number"` (default): `xValues` is a numeric range; Excel plots
   *   `(x, y)` pairs.
   * - `"text"`: `xValues` points at text cells; Excel treats them as
   *   evenly-spaced categorical labels — useful when the natural x
   *   dimension is names/dates-as-strings rather than measurements.
   *
   * Ignored when {@link xValues} is already a structured
   * {@link AxisDataSource} (the caller has already picked strRef vs
   * numRef explicitly).
   */
  xValueType?: "number" | "text";
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
    /**
     * Rel ID for a pre-registered chart-part image relationship.
     * Callers that already manage the image rel by hand can pass the
     * existing `rId…` string here. Most callers should prefer
     * {@link image} — the builder wires up the media entry, chart rel
     * and correct r:id automatically.
     */
    relationshipId?: string;
    /**
     * High-level image source: either a raw payload ({@link Buffer} /
     * {@link Uint8Array} / base64 string), a structured
     * `ImageData`-shaped object, or a `{ workbookImageId }` pointing at
     * a previously registered {@link Workbook.addImage} result.
     *
     * When set, the worksheet-side `_registerChart` path stores the
     * image in the workbook's media collection, allocates a new
     * relationship on the chart part (`rIdN` — non-conflicting with
     * style/colors rels), and populates {@link relationshipId}
     * automatically. Subsequent writes emit the correct `<a:blipFill>`
     * with the matching `r:embed` reference.
     *
     * `extension` is inferred from the buffer's magic bytes when the
     * caller passes a raw `Uint8Array`: PNG / JPEG / GIF are
     * recognised (every `<a:blipFill>`-supported format) and
     * anything else — WebP / AVIF / TIFF / BMP / SVG — is dropped
     * entirely rather than relabelled as `png`. Emitting an image
     * with a wrong extension would produce a broken picture in
     * Excel, so "drop the blip" is safer than "guess".
     *
     * For string inputs the rule differs: a `data:image/<type>;base64,…`
     * URL requires `<type>` to be `png` / `jpeg` / `jpg` / `gif` (other
     * content-types are dropped), but a **bare base64 payload** (no
     * `data:` prefix) has no embedded content-type and so is
     * assumed to be PNG — the most common encode target when
     * callers pass an already-stripped payload. Authors who need a
     * different format for bare base64 should pass a structured
     * {@link ChartPictureFillImageData} with an explicit
     * `extension`, or prefix with the matching data URL.
     */
    image?: AddChartPictureFillImage;
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
  // NOTE: classic chart **series** have no `txPr` slot in OOXML —
  // `CT_BarSer`, `CT_LineSer`, `CT_PieSer`, `CT_ScatterSer`,
  // `CT_BubbleSer`, `CT_AreaSer`, `CT_RadarSer` and `CT_SurfaceSer`
  // all omit the element (only the chart-space, title, legend,
  // axis, and data-labels nodes carry `txPr`). The previous
  // `txPr?: ChartTextProperties` field on this options bag was
  // accepted by the builder and stored on the series object, but no
  // writer ever emitted it — every programmatic caller was
  // silently losing their styling on save. For per-run / per-label
  // text styling use `dataLabels.txPr`; for axis / title text use
  // `valueAxis.txPr` / `titleOptions.txPr`.
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
  /**
   * Excel 2013+ "Value From Cells". When set, each data label's text is
   * read from the given worksheet range instead of the series/value.
   *
   * Accepts either a plain formula string (most common) or a structured
   * {@link DataLabelsRange} for callers who want to pre-populate the
   * cache. The builder wires the value up as the MS `c15:datalabelsRange`
   * extension and generates placeholder per-point `<c:dLbl>` entries
   * carrying the cached strings so viewers that don't understand the
   * extension still show the right labels.
   *
   * Typical usage: `valueFromCells: "Sheet1!$C$2:$C$10"`.
   */
  valueFromCells?: string | DataLabelsRange;
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
  /** Major unit (must be > 0; when both are set, minorUnit must be \u2264 majorUnit) */
  majorUnit?: number;
  /** Minor unit (must be > 0; when both are set, must be \u2264 majorUnit) */
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
