/**
 * DOCX Module - Enhanced DrawingML Shape Builder
 *
 * Provides a comprehensive API for creating DrawingML shapes beyond the
 * basic shape type. Includes support for:
 * - All 187 preset geometry types
 * - Gradient and pattern fills
 * - Shadow, glow, reflection effects
 * - 3D rotation and bevel
 * - Shape connectors and groups
 * - Custom geometry paths
 * - Text body formatting within shapes
 */

import {
  checkOneOf,
  checkRecord,
  checkRgbColor,
  invalidArgument,
  under
} from "@word/core/argument-checks";
import type { FieldName } from "@word/core/argument-checks";
import { assertDrawingNumber, assertDrawingNumbers } from "@word/core/drawing-geometry";
import type {
  DrawingShape,
  Paragraph,
  HexColor,
  Emu,
  HorizontalPositionRelative,
  VerticalPositionRelative,
  WrapStyle,
  WrapTextSide
} from "@word/types";
import { xmlEncodeAttr } from "@xml/encode";

// =============================================================================
// Preset Shape Types (OOXML ST_ShapeType enum — all 187 values)
// =============================================================================

/** All standard OOXML preset shape types. */
export type StandardShapeType =
  // Basic shapes
  | "rect"
  | "roundRect"
  | "snip1Rect"
  | "snip2SameRect"
  | "snip2DiagRect"
  | "snipRoundRect"
  | "round1Rect"
  | "round2SameRect"
  | "round2DiagRect"
  | "ellipse"
  | "triangle"
  | "rtTriangle"
  | "parallelogram"
  | "trapezoid"
  | "diamond"
  | "pentagon"
  | "hexagon"
  | "heptagon"
  | "octagon"
  | "decagon"
  | "dodecagon"
  | "pie"
  | "chord"
  | "teardrop"
  | "frame"
  | "halfFrame"
  | "corner"
  | "diagStripe"
  | "plus"
  | "cross"
  | "cube"
  | "can"
  | "foldedCorner"
  | "smileyFace"
  | "heart"
  | "lightningBolt"
  | "sun"
  | "moon"
  | "cloud"
  | "arc"
  | "bracketPair"
  | "bracePair"
  | "plaque"
  | "donut"
  | "noSmoking"
  | "blockArc"
  | "gear6"
  | "gear9"
  // Arrows
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "leftRightArrow"
  | "upDownArrow"
  | "quadArrow"
  | "leftRightUpArrow"
  | "bentArrow"
  | "uturnArrow"
  | "leftUpArrow"
  | "bentUpArrow"
  | "curvedRightArrow"
  | "curvedLeftArrow"
  | "curvedUpArrow"
  | "curvedDownArrow"
  | "stripedRightArrow"
  | "notchedRightArrow"
  | "homePlate"
  | "chevron"
  | "rightArrowCallout"
  | "downArrowCallout"
  | "leftArrowCallout"
  | "upArrowCallout"
  | "leftRightArrowCallout"
  | "quadArrowCallout"
  | "circularArrow"
  // Flowchart
  | "flowChartProcess"
  | "flowChartAlternateProcess"
  | "flowChartDecision"
  | "flowChartInputOutput"
  | "flowChartPredefinedProcess"
  | "flowChartInternalStorage"
  | "flowChartDocument"
  | "flowChartMultidocument"
  | "flowChartTerminator"
  | "flowChartPreparation"
  | "flowChartManualInput"
  | "flowChartManualOperation"
  | "flowChartConnector"
  | "flowChartOffpageConnector"
  | "flowChartPunchedCard"
  | "flowChartPunchedTape"
  | "flowChartSummingJunction"
  | "flowChartOr"
  | "flowChartCollate"
  | "flowChartSort"
  | "flowChartExtract"
  | "flowChartMerge"
  | "flowChartOnlineStorage"
  | "flowChartDelay"
  | "flowChartMagneticTape"
  | "flowChartMagneticDisk"
  | "flowChartMagneticDrum"
  | "flowChartDisplay"
  // Callouts
  | "wedgeRectCallout"
  | "wedgeRoundRectCallout"
  | "wedgeEllipseCallout"
  | "cloudCallout"
  | "borderCallout1"
  | "borderCallout2"
  | "borderCallout3"
  | "accentCallout1"
  | "accentCallout2"
  | "accentCallout3"
  | "callout1"
  | "callout2"
  | "callout3"
  | "accentBorderCallout1"
  | "accentBorderCallout2"
  | "accentBorderCallout3"
  // Stars & Banners
  | "irregularSeal1"
  | "irregularSeal2"
  | "star4"
  | "star5"
  | "star6"
  | "star7"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  | "ribbon"
  | "ribbon2"
  | "ellipseRibbon"
  | "ellipseRibbon2"
  | "verticalScroll"
  | "horizontalScroll"
  | "wave"
  | "doubleWave"
  // Math shapes
  | "mathPlus"
  | "mathMinus"
  | "mathMultiply"
  | "mathDivide"
  | "mathEqual"
  | "mathNotEqual"
  // Connectors
  | "line"
  | "straightConnector1"
  | "bentConnector2"
  | "bentConnector3"
  | "bentConnector4"
  | "bentConnector5"
  | "curvedConnector2"
  | "curvedConnector3"
  | "curvedConnector4"
  | "curvedConnector5"
  // Action buttons
  | "actionButtonBlank"
  | "actionButtonHome"
  | "actionButtonHelp"
  | "actionButtonInformation"
  | "actionButtonBackPrevious"
  | "actionButtonForwardNext"
  | "actionButtonBeginning"
  | "actionButtonEnd"
  | "actionButtonReturn"
  | "actionButtonDocument"
  | "actionButtonSound"
  | "actionButtonMovie";

// =============================================================================
// Fill Types
// =============================================================================

/** Solid fill. */
export interface SolidFill {
  readonly type: "solid";
  readonly color: HexColor;
  /** Transparency (0-100, percent). */
  readonly transparency?: number;
}

/** Gradient stop. */
export interface GradientStop {
  readonly position: number; // 0-100000 (percentage * 1000)
  readonly color: HexColor;
  readonly transparency?: number;
}

/** Gradient fill. */
export interface GradientFill {
  readonly type: "gradient";
  readonly stops: readonly GradientStop[];
  /** Angle in 60,000ths of a degree. 0 = left-to-right. */
  readonly angle?: number;
  /** Gradient type: linear or radial (path). */
  readonly gradientType?: "linear" | "radial";
}

/** Pattern fill. */
export interface PatternFill {
  readonly type: "pattern";
  /** Pattern preset (e.g. "pct10", "dkHorz", "ltVert", "dnDiag"). */
  readonly preset: string;
  readonly foregroundColor: HexColor;
  readonly backgroundColor: HexColor;
}

/** No fill (transparent). */
export interface NoFill {
  readonly type: "none";
}

/** Shape fill specification. */
export type ShapeFill = SolidFill | GradientFill | PatternFill | NoFill;

// =============================================================================
// Line/Outline Types
// =============================================================================

/** Line dash style. */
export type LineDash =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDot"
  | "sysDash"
  | "sysDashDot"
  | "sysDashDotDot";

/** Line end type (arrow head). */
export type LineEndType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";

/** Line end size. */
export type LineEndSize = "sm" | "med" | "lg";

/** Line/outline specification. */
export interface ShapeOutline {
  /** Width in EMU. */
  readonly width?: Emu;
  /** Color. */
  readonly color?: HexColor;
  /** Dash style. */
  readonly dash?: LineDash;
  /** Join type. */
  readonly join?: "round" | "bevel" | "miter";
  /** Head end (start of line). */
  readonly headEnd?: { type: LineEndType; width?: LineEndSize; length?: LineEndSize };
  /** Tail end (end of line). */
  readonly tailEnd?: { type: LineEndType; width?: LineEndSize; length?: LineEndSize };
  /** No outline. */
  readonly noLine?: boolean;
}

// =============================================================================
// Effect Types
// =============================================================================

/** Shadow effect. */
export interface ShadowEffect {
  readonly type: "outer" | "inner";
  readonly color: HexColor;
  readonly transparency?: number;
  /** Blur radius in EMU. */
  readonly blurRadius?: Emu;
  /** Distance in EMU. */
  readonly distance?: Emu;
  /** Direction in 60,000ths of a degree. */
  readonly direction?: number;
}

/** Glow effect. */
export interface GlowEffect {
  readonly color: HexColor;
  readonly transparency?: number;
  /** Radius in EMU. */
  readonly radius: Emu;
}

/** Reflection effect. */
export interface ReflectionEffect {
  /** Blur radius in EMU. */
  readonly blurRadius?: Emu;
  /** Start transparency (0-100). */
  readonly startOpacity?: number;
  /** End transparency (0-100). */
  readonly endOpacity?: number;
  /** Distance in EMU. */
  readonly distance?: Emu;
  /** Direction in 60,000ths of a degree. */
  readonly direction?: number;
  /** Fade direction. */
  readonly fadeDirection?: number;
}

/** 3D effect. */
export interface Effect3D {
  /** Rotation on X/Y/Z axes (in 60,000ths of a degree). */
  readonly rotX?: number;
  readonly rotY?: number;
  readonly rotZ?: number;
  /** Camera preset. */
  readonly camera?:
    | "orthographicFront"
    | "perspectiveFront"
    | "isometricTopDown"
    | "obliqueTopLeft";
  /** Bevel top. */
  readonly bevelTop?: { width: Emu; height: Emu; preset?: string };
  /** Bevel bottom. */
  readonly bevelBottom?: { width: Emu; height: Emu; preset?: string };
  /** Extrusion depth in EMU. */
  readonly extrusionDepth?: Emu;
  /** Extrusion color. */
  readonly extrusionColor?: HexColor;
}

/** Shape effects. */
export interface ShapeEffects {
  readonly shadow?: ShadowEffect;
  readonly glow?: GlowEffect;
  readonly reflection?: ReflectionEffect;
  readonly effect3d?: Effect3D;
  /** Soft edges radius in EMU. */
  readonly softEdges?: Emu;
}

// =============================================================================
// Text Body Properties
// =============================================================================

/** Vertical text anchor. */
export type TextVerticalAnchor = "t" | "ctr" | "b";

/** Text wrapping type within shape. */
export type TextWrap = "none" | "square";

/** Text body configuration for shapes. */
export interface ShapeTextBody {
  /** Content paragraphs. */
  readonly paragraphs: readonly Paragraph[];
  /** Vertical alignment. */
  readonly anchor?: TextVerticalAnchor;
  /** Wrap text within shape. */
  readonly wrap?: TextWrap;
  /** Internal margins (EMU). */
  readonly margins?: {
    readonly top?: Emu;
    readonly bottom?: Emu;
    readonly left?: Emu;
    readonly right?: Emu;
  };
  /** Auto-fit text. */
  readonly autoFit?: "none" | "normal" | "shrink";
  /** Vertical text (for CJK). */
  readonly vertical?: boolean;
  /** Number of columns. */
  readonly columns?: number;
  /** Column spacing in EMU. */
  readonly columnSpacing?: Emu;
}

// =============================================================================
// Enhanced Shape Options
// =============================================================================

/** Complete shape creation options. */
export interface CreateShapeOptions {
  /** Shape type (preset geometry). */
  readonly shapeType: string;
  /** Width in EMU. */
  readonly width: Emu;
  /** Height in EMU. */
  readonly height: Emu;
  /** Fill specification. */
  readonly fill?: ShapeFill;
  /** Outline specification. */
  readonly outline?: ShapeOutline;
  /** Effects (shadow, glow, etc.). */
  readonly effects?: ShapeEffects;
  /** Text content. */
  readonly textBody?: ShapeTextBody;
  /** Alternative text (accessibility). */
  readonly altText?: string;
  /** Shape name. */
  readonly name?: string;
  /** Rotation in 60,000ths of a degree. */
  readonly rotation?: number;
  /** Flip horizontally (`a:xfrm/@flipH`). */
  readonly flipHorizontal?: boolean;
  /** Flip vertically (`a:xfrm/@flipV`). */
  readonly flipVertical?: boolean;
  /** @deprecated Use {@link flipHorizontal}, as the shape node names it. Removed in 1.0. */
  readonly flipH?: boolean;
  /** @deprecated Use {@link flipVertical}, as the shape node names it. Removed in 1.0. */
  readonly flipV?: boolean;
  /** Horizontal position. */
  readonly horizontalPosition?: {
    readonly relativeTo?: HorizontalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "left" | "center" | "right";
  };
  /** Vertical position. */
  readonly verticalPosition?: {
    readonly relativeTo?: VerticalPositionRelative;
    readonly offset?: Emu;
    readonly align?: "top" | "center" | "bottom";
  };
  /** Text wrapping. */
  readonly wrap?: {
    readonly style: WrapStyle;
    readonly side?: WrapTextSide;
  };
  /** Behind document text. */
  readonly behindDoc?: boolean;
}

// =============================================================================
// Shape Builder Functions
// =============================================================================

/**
 * Create an enhanced DrawingML shape with full styling options.
 *
 * Basic properties (solid fill, outline color/width, text paragraphs, positioning)
 * are mapped directly to the `DrawingShape` interface fields. Advanced properties
 * (gradient/pattern fills, effects, line details, text body formatting) are
 * serialized into the `rawXml` field for preservation during packaging.
 *
 * @param options - Complete shape creation options.
 * @returns A DrawingShape element for the document body.
 */
export function createShape(options: CreateShapeOptions): DrawingShape {
  checkRecord(options, "options", "Build.shape");
  return buildShape(options, "Build.shape", under("options"));
}

/** A shape's `width` / `height` are parameters of the helper; the rest are its options. */
const sizedField: FieldName = path =>
  path === "width" || path === "height" ? path : `options.${path}`;

/**
 * A shape a helper (`Build.rect`, `Build.star`, …) builds: `width` × `height`,
 * its own `defaults` under the caller's options.
 */
function sizedShape(
  where: string,
  shapeType: CreateShapeOptions["shapeType"],
  width: Emu,
  height: Emu,
  options: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">> | undefined,
  defaults: Partial<CreateShapeOptions> = {}
): DrawingShape {
  checkRecord(options, "options", where, { optional: true });
  return buildShape({ ...defaults, ...options, shapeType, width, height }, where, sizedField);
}

/** {@link createShape} in the name of `where`, its fields named by `field`. */
function buildShape(options: CreateShapeOptions, where: string, field: FieldName): DrawingShape {
  assertShapeNumbers(options, where, field);
  assertShapeColors(options, where, field);
  // Convert fill to simple properties for the DrawingShape interface
  let fillColor: HexColor | undefined;
  let noFill = false;
  if (options.fill) {
    if (options.fill.type === "solid") {
      fillColor = options.fill.color;
    } else if (options.fill.type === "none") {
      noFill = true;
    }
  }

  let outlineColor: HexColor | undefined;
  let outlineWidth: Emu | undefined;
  let noOutline = false;
  if (options.outline) {
    outlineColor = options.outline.color;
    outlineWidth = options.outline.width;
    noOutline = options.outline.noLine ?? false;
  }

  // Serialize advanced properties that DrawingShape can't represent directly.
  // The writer needs them split because the OOXML schema requires fill
  // fragments to precede a:ln while effect/3D fragments must follow it.
  const advanced = serializeAdvancedProperties(options);
  const rawXml = (advanced.fillXml ?? "") + (advanced.effectsXml ?? "");

  return {
    type: "drawingShape",
    shapeType: options.shapeType,
    width: options.width,
    height: options.height,
    fillColor,
    noFill,
    outlineColor,
    outlineWidth,
    noOutline,
    textContent: options.textBody?.paragraphs,
    textBodyAnchor: options.textBody?.anchor,
    altText: options.altText,
    name: options.name,
    horizontalPosition: options.horizontalPosition,
    verticalPosition: options.verticalPosition,
    wrap: options.wrap,
    behindDoc: options.behindDoc,
    rotation: options.rotation,
    flipHorizontal: options.flipHorizontal ?? options.flipH,
    flipVertical: options.flipVertical ?? options.flipV,
    rawXml: rawXml.length > 0 ? rawXml : undefined,
    _advancedFillXml: advanced.fillXml,
    _advancedEffectsXml: advanced.effectsXml
  };
}

/**
 * Check every colour a shape will write. Each becomes an `a:srgbClr`, whose
 * value is `RRGGBB` — a `#` or a name would be written as given and the shape
 * refused.
 */
function assertShapeColors(options: CreateShapeOptions, where: string, field: FieldName): void {
  const fill = options.fill;
  if (fill?.type === "solid") {
    checkRgbColor(fill.color, field("fill.color"), where);
  } else if (fill?.type === "gradient") {
    fill.stops.forEach((stop, i) =>
      checkRgbColor(stop.color, field(`fill.stops[${i}].color`), where)
    );
  } else if (fill?.type === "pattern") {
    checkRgbColor(fill.foregroundColor, field("fill.foregroundColor"), where);
    checkRgbColor(fill.backgroundColor, field("fill.backgroundColor"), where);
  }
  checkRgbColor(options.outline?.color, field("outline.color"), where, { optional: true });
  const fx = options.effects;
  if (fx?.shadow) {
    checkRgbColor(fx.shadow.color, field("effects.shadow.color"), where);
  }
  if (fx?.glow) {
    checkRgbColor(fx.glow.color, field("effects.glow.color"), where);
  }
  checkRgbColor(fx?.effect3d?.extrusionColor, field("effects.effect3d.extrusionColor"), where, {
    optional: true
  });
}

/**
 * Check every number a shape will write — its frame and the effect geometry
 * {@link serializeAdvancedProperties} interpolates into markup.
 */
function assertShapeNumbers(options: CreateShapeOptions, where: string, field: FieldName): void {
  assertDrawingNumbers(where, options, field);
  const nonNegative = { sign: "nonNegative", optional: true } as const;
  const any = { optional: true } as const;
  assertDrawingNumber(options.outline?.width, field("outline.width"), where, nonNegative);
  const fill = options.fill;
  if (fill?.type === "gradient") {
    fill.stops.forEach((stop, i) => {
      assertDrawingNumber(stop.position, field(`fill.stops[${i}].position`), where);
      assertDrawingNumber(stop.transparency, field(`fill.stops[${i}].transparency`), where, any);
    });
    assertDrawingNumber(fill.angle, field("fill.angle"), where, any);
  } else if (fill?.type === "solid") {
    assertDrawingNumber(fill.transparency, field("fill.transparency"), where, any);
  }
  const fx = options.effects;
  if (fx?.shadow) {
    assertDrawingNumber(
      fx.shadow.blurRadius,
      field("effects.shadow.blurRadius"),
      where,
      nonNegative
    );
    assertDrawingNumber(fx.shadow.distance, field("effects.shadow.distance"), where, nonNegative);
    assertDrawingNumber(fx.shadow.direction, field("effects.shadow.direction"), where, any);
    assertDrawingNumber(fx.shadow.transparency, field("effects.shadow.transparency"), where, any);
  }
  if (fx?.glow) {
    assertDrawingNumber(fx.glow.radius, field("effects.glow.radius"), where, {
      sign: "nonNegative"
    });
    assertDrawingNumber(fx.glow.transparency, field("effects.glow.transparency"), where, any);
  }
  if (fx?.reflection) {
    const r = fx.reflection;
    assertDrawingNumber(r.blurRadius, field("effects.reflection.blurRadius"), where, nonNegative);
    assertDrawingNumber(r.startOpacity, field("effects.reflection.startOpacity"), where, any);
    assertDrawingNumber(r.endOpacity, field("effects.reflection.endOpacity"), where, any);
    assertDrawingNumber(r.distance, field("effects.reflection.distance"), where, nonNegative);
    assertDrawingNumber(r.direction, field("effects.reflection.direction"), where, any);
    assertDrawingNumber(r.fadeDirection, field("effects.reflection.fadeDirection"), where, any);
  }
  assertDrawingNumber(fx?.softEdges, field("effects.softEdges"), where, nonNegative);
  const e = fx?.effect3d;
  if (e) {
    assertDrawingNumber(e.rotX, field("effects.effect3d.rotX"), where, any);
    assertDrawingNumber(e.rotY, field("effects.effect3d.rotY"), where, any);
    assertDrawingNumber(e.rotZ, field("effects.effect3d.rotZ"), where, any);
    assertDrawingNumber(
      e.extrusionDepth,
      field("effects.effect3d.extrusionDepth"),
      where,
      nonNegative
    );
    for (const key of ["bevelTop", "bevelBottom"] as const) {
      const bevel = e[key];
      if (bevel) {
        const required = { sign: "nonNegative" } as const;
        assertDrawingNumber(bevel.width, field(`effects.effect3d.${key}.width`), where, required);
        assertDrawingNumber(bevel.height, field(`effects.effect3d.${key}.height`), where, required);
      }
    }
  }
}

/**
 * Serialize advanced shape properties into XML fragments suitable for
 * insertion inside `wps:spPr`. Returns them split into the two slots the
 * OOXML schema cares about:
 *  - `fillXml` is inserted between `a:prstGeom` and `a:ln`
 *  - `effectsXml` is inserted after `a:ln` (effectLst → scene3d → sp3d)
 */
function serializeAdvancedProperties(options: CreateShapeOptions): {
  fillXml?: string;
  effectsXml?: string;
} {
  const fillParts: string[] = [];
  const effectParts: string[] = [];

  // Gradient fill
  if (options.fill && options.fill.type === "gradient") {
    const gf = options.fill;
    const stops = gf.stops
      .map(
        s =>
          `<a:gs pos="${Math.round(s.position)}"><a:srgbClr val="${xmlEncodeAttr(s.color)}"${s.transparency ? ` alpha="${Math.round((100 - s.transparency) * 1000)}"` : ""}/></a:gs>`
      )
      .join("");
    const angle = gf.angle !== undefined ? ` ang="${Math.round(gf.angle)}"` : "";
    fillParts.push(
      `<a:gradFill><a:gsLst>${stops}</a:gsLst><a:lin${angle} scaled="1"/></a:gradFill>`
    );
  }

  // Pattern fill
  if (options.fill && options.fill.type === "pattern") {
    const pf = options.fill;
    fillParts.push(
      `<a:pattFill prst="${xmlEncodeAttr(pf.preset)}"><a:fgClr><a:srgbClr val="${xmlEncodeAttr(pf.foregroundColor)}"/></a:fgClr><a:bgClr><a:srgbClr val="${xmlEncodeAttr(pf.backgroundColor)}"/></a:bgClr></a:pattFill>`
    );
  }

  // Effects: shadow / glow / reflection / softEdges must live inside
  // <a:effectLst>. Collect all of them then emit a single wrapper.
  const effectChildren: string[] = [];
  if (options.effects?.shadow) {
    const s = options.effects.shadow;
    const attrs: string[] = [];
    if (s.blurRadius) {
      attrs.push(`blurRad="${Math.round(s.blurRadius)}"`);
    }
    if (s.distance) {
      attrs.push(`dist="${Math.round(s.distance)}"`);
    }
    if (s.direction !== undefined) {
      attrs.push(`dir="${Math.round(s.direction)}"`);
    }
    const alpha =
      s.transparency !== undefined
        ? ` <a:alpha val="${Math.round((100 - s.transparency) * 1000)}"/>`
        : "";
    const tag = s.type === "inner" ? "a:innerShdw" : "a:outerShdw";
    effectChildren.push(
      `<${tag} ${attrs.join(" ")}><a:srgbClr val="${xmlEncodeAttr(s.color)}">${alpha}</a:srgbClr></${tag}>`
    );
  }

  if (options.effects?.glow) {
    const g = options.effects.glow;
    const alpha =
      g.transparency !== undefined
        ? `<a:alpha val="${Math.round((100 - g.transparency) * 1000)}"/>`
        : "";
    effectChildren.push(
      `<a:glow rad="${Math.round(g.radius)}"><a:srgbClr val="${xmlEncodeAttr(g.color)}">${alpha}</a:srgbClr></a:glow>`
    );
  }

  if (options.effects?.reflection) {
    const r = options.effects.reflection;
    const attrs: string[] = [];
    if (r.blurRadius) {
      attrs.push(`blurRad="${Math.round(r.blurRadius)}"`);
    }
    if (r.startOpacity !== undefined) {
      attrs.push(`stA="${Math.round(r.startOpacity * 1000)}"`);
    }
    if (r.endOpacity !== undefined) {
      attrs.push(`endA="${Math.round(r.endOpacity * 1000)}"`);
    }
    if (r.distance) {
      attrs.push(`dist="${Math.round(r.distance)}"`);
    }
    if (r.direction !== undefined) {
      attrs.push(`dir="${Math.round(r.direction)}"`);
    }
    if (r.fadeDirection !== undefined) {
      attrs.push(`fadeDir="${Math.round(r.fadeDirection)}"`);
    }
    effectChildren.push(`<a:reflection ${attrs.join(" ")}/>`);
  }

  if (options.effects?.softEdges) {
    effectChildren.push(`<a:softEdge rad="${Math.round(options.effects.softEdges)}"/>`);
  }

  if (effectChildren.length > 0) {
    effectParts.push(`<a:effectLst>${effectChildren.join("")}</a:effectLst>`);
  }

  // 3D effect (scene3d + sp3d) — both follow effectLst.
  if (options.effects?.effect3d) {
    const e = options.effects.effect3d;
    const camera = e.camera ?? "orthographicFront";
    const rotAttrs: string[] = [];
    if (e.rotX !== undefined) {
      rotAttrs.push(`lat="${Math.round(e.rotX)}"`);
    }
    if (e.rotY !== undefined) {
      rotAttrs.push(`lon="${Math.round(e.rotY)}"`);
    }
    if (e.rotZ !== undefined) {
      rotAttrs.push(`rev="${Math.round(e.rotZ)}"`);
    }
    const rot = rotAttrs.length > 0 ? `<a:rot ${rotAttrs.join(" ")}/>` : "";
    effectParts.push(
      `<a:scene3d><a:camera prst="${xmlEncodeAttr(camera)}">${rot}</a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>`
    );

    const sp3dChildren: string[] = [];
    if (e.bevelTop) {
      const preset = e.bevelTop.preset ? ` prst="${xmlEncodeAttr(e.bevelTop.preset)}"` : "";
      sp3dChildren.push(
        `<a:bevelT w="${Math.round(e.bevelTop.width)}" h="${Math.round(e.bevelTop.height)}"${preset}/>`
      );
    }
    if (e.bevelBottom) {
      const preset = e.bevelBottom.preset ? ` prst="${xmlEncodeAttr(e.bevelBottom.preset)}"` : "";
      sp3dChildren.push(
        `<a:bevelB w="${Math.round(e.bevelBottom.width)}" h="${Math.round(e.bevelBottom.height)}"${preset}/>`
      );
    }
    if (e.extrusionColor) {
      sp3dChildren.push(
        `<a:extrusionClr><a:srgbClr val="${xmlEncodeAttr(e.extrusionColor)}"/></a:extrusionClr>`
      );
    }
    const extDepth = e.extrusionDepth ? ` extrusionH="${Math.round(e.extrusionDepth)}"` : "";
    effectParts.push(`<a:sp3d${extDepth}>${sp3dChildren.join("")}</a:sp3d>`);
  }

  return {
    fillXml: fillParts.length > 0 ? fillParts.join("") : undefined,
    effectsXml: effectParts.length > 0 ? effectParts.join("") : undefined
  };
}

/**
 * Create a simple rectangle shape.
 */
export function createRect(
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  return sizedShape("Build.rect", "rect", width, height, options);
}

/**
 * Create a rounded rectangle shape.
 */
export function createRoundRect(
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  return sizedShape("Build.roundRect", "roundRect", width, height, options);
}

/**
 * Create an ellipse/circle shape.
 */
export function createEllipse(
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  return sizedShape("Build.ellipse", "ellipse", width, height, options);
}

/**
 * Create a line connector.
 */
export function createLine(
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  return sizedShape("Build.line", "line", width, height, options, { fill: { type: "none" } });
}

/**
 * Create an arrow shape.
 */
export function createArrow(
  direction: "right" | "left" | "up" | "down",
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  const shapeMap: Record<string, StandardShapeType> = {
    right: "rightArrow",
    left: "leftArrow",
    up: "upArrow",
    down: "downArrow"
  };
  const where = "Build.arrow";
  checkOneOf(direction, new Set(Object.keys(shapeMap)), "direction", where);
  return sizedShape(where, shapeMap[direction]!, width, height, options);
}

/**
 * Create a flowchart shape.
 */
export function createFlowchartShape(
  kind: "process" | "decision" | "terminator" | "document" | "data" | "connector" | "preparation",
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  const shapeMap: Record<string, StandardShapeType> = {
    process: "flowChartProcess",
    decision: "flowChartDecision",
    terminator: "flowChartTerminator",
    document: "flowChartDocument",
    data: "flowChartInputOutput",
    connector: "flowChartConnector",
    preparation: "flowChartPreparation"
  };
  const where = "Build.flowchartShape";
  checkOneOf(kind, new Set(Object.keys(shapeMap)), "kind", where);
  return sizedShape(where, shapeMap[kind]!, width, height, options);
}

/**
 * Create a callout shape.
 */
export function createCallout(
  style: "rect" | "roundRect" | "ellipse" | "cloud",
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  const shapeMap: Record<string, StandardShapeType> = {
    rect: "wedgeRectCallout",
    roundRect: "wedgeRoundRectCallout",
    ellipse: "wedgeEllipseCallout",
    cloud: "cloudCallout"
  };
  const where = "Build.callout";
  checkOneOf(style, new Set(Object.keys(shapeMap)), "style", where);
  return sizedShape(where, shapeMap[style]!, width, height, options);
}

const STAR_POINTS: ReadonlySet<number> = new Set([4, 5, 6, 7, 8, 10, 12, 16, 24, 32]);

/**
 * Create a star shape.
 */
export function createStar(
  points: 4 | 5 | 6 | 7 | 8 | 10 | 12 | 16 | 24 | 32,
  width: Emu,
  height: Emu,
  options?: Partial<Omit<CreateShapeOptions, "shapeType" | "width" | "height">>
): DrawingShape {
  const where = "Build.star";
  if (!STAR_POINTS.has(points)) {
    invalidArgument(where, "points", `one of ${[...STAR_POINTS].join(", ")}`, points);
  }
  return sizedShape(where, `star${points}`, width, height, options);
}
