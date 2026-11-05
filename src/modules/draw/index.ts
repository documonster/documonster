/**
 * `draw` — the shared drawing engine.
 *
 * One structured display list, one walker, many backends. Producers (charts,
 * sparklines, diagram engines) build a {@link DrawList}; backends implement
 * {@link DrawSurface}. Nothing passes an SVG string between renderers any more,
 * which is what used to let each backend drift into its own interpretation of the
 * same picture.
 *
 * @example
 * ```ts
 * import { Draw, type DrawList } from "documonster/draw";
 *
 * const list: DrawList = {
 *   width: 100,
 *   height: 50,
 *   children: [
 *     { kind: "rect", x: 10, y: 10, width: 80, height: 30,
 *       paint: { fill: { r: 0.2, g: 0.4, b: 0.8, a: 1 } } }
 *   ]
 * };
 * const svg = Draw.toSvg(list);
 * ```
 *
 * ## Backends
 *
 * Two ship here: `Draw.toSvg` serialises to markup, and `Draw.toRgba`
 * paints pixels. A PDF page is the third and lives in `documonster/pdf`
 * (`Pdf.createDrawSurface`), because it draws onto a page builder rather than
 * producing a standalone artefact. Any of them can be driven from the same list —
 * that is the whole point of having one.
 *
 * Pixels rather than a PNG: encoding one needs DEFLATE, which sits a layer above this
 * module. Pair `Draw.toRgba` with any encoder — `Archive.encodePng` from
 * `documonster/archive` is the one this repository uses.
 */

import * as DrawNs from "@draw/api/draw";
import * as MatrixNs from "@draw/api/matrix";
import { checkPathCommands, checkRasterCanvas } from "@draw/arguments";
import type {
  RasterPoint as DrawRasterPoint,
  DrawRasterLineOptions,
  DrawRasterStrokeOptions,
  DrawRasterTextOptions,
  DrawRasterTextRotation
} from "@draw/raster/canvas";
import { parseRasterFont as parseRasterFontLenient } from "@draw/raster/glyph-outline";
import type {
  GlyphOutline as DrawGlyphOutline,
  GlyphPoint as DrawGlyphPoint,
  RasterFont as DrawRasterFont,
  RasterFontSource as DrawRasterFontSource
} from "@draw/raster/glyph-outline";
import {
  createScaledRasterSurface,
  rasterizeToRgba as rasterizeToRgbaLenient
} from "@draw/raster/surface";
import type {
  RasterSurfaceOptions as DrawRasterSurfaceOptions,
  RasterizeOptions as DrawRasterizeOptions,
  RasterizedImage as DrawRasterizedImage,
  RgbaImage as DrawRgbaImage
} from "@draw/raster/surface";
import type { DrawSurface } from "@draw/surface";
import type { ToSvgOptions as DrawToSvgOptions } from "@draw/svg";
import { toSvg as toSvgLenient } from "@draw/svg";
import { flattenPath as flattenPathImpl } from "@draw/types";
import type {
  DrawPathCommand,
  DrawSubpath,
  RasterTextStyle as DrawRasterTextStyle
} from "@draw/types";
import type { Rgba01 as DrawRgba01 } from "@utils/svg-lex";

// =============================================================================
// Public API — the `Draw` and `Matrix` namespaces (tree-shaken via `export * as`)
// =============================================================================

export * as Draw from "@draw/api/draw";
export * as Matrix from "@draw/api/matrix";

// =============================================================================
// Constants
// =============================================================================

export { POINTS_PER_PIXEL } from "@draw/text";
export { DEFAULT_TEXT_FAMILY } from "@draw/types";

// =============================================================================
// Errors
// =============================================================================

export { DrawError, DrawArgumentError, isDrawError } from "@draw/errors";

// =============================================================================
// Types
// =============================================================================

export type { DrawSurface } from "@draw/surface";
export type {
  DrawToSvgOptions,
  DrawRasterSurfaceOptions,
  DrawRasterizeOptions,
  DrawRasterizedImage,
  DrawRgbaImage
};
/**
 * Reachable from `Draw.Backend.BasicRasterCanvas`'s own signatures — its point lists and the
 * font it rasterises glyphs from. Without the names a consumer can call the methods but
 * cannot declare a variable to hold what they take.
 */
export type {
  DrawRasterPoint,
  DrawRasterLineOptions,
  DrawRasterStrokeOptions,
  DrawRasterTextOptions,
  DrawRasterTextRotation,
  DrawGlyphOutline,
  DrawGlyphPoint,
  DrawRasterFont
};
/**
 * How a font may be supplied: bytes, a parsed face (`Draw.parseRasterFont`), or a face
 * inside a collection.
 *
 * Named because it appears in `DrawRasterizeOptions.fonts` and
 * `BasicRasterCanvas.setFonts`.
 *
 * There is deliberately **no process-wide font registry**, published or otherwise.
 * Fonts belong to a render — `DrawRasterizeOptions.fonts` — or to a canvas —
 * `BasicRasterCanvas.setFonts`. A module-level setter would let one caller change what
 * another caller's canvas draws with; the one that existed was used by nothing but the
 * tests, and removing it took two mutable globals with it.
 */
export type { DrawRasterFontSource };
/**
 * Re-exported because it is reachable from this module's own signatures —
 * `DrawPaint.fill`, `DrawTextStyle.fill` and the return of `Draw.parseColor` are
 * all `DrawRgba01`. Without the name a consumer can assign one but cannot declare one.
 */
export type { DrawRgba01 };
export type {
  DrawBox,
  DrawClip,
  DrawFlattenPathOptions,
  DrawSubpath,
  DrawList,
  DrawMatrix,
  DrawNode,
  DrawPaint,
  DrawPathCommand,
  DrawPoint,
  DrawTextAnchor,
  DrawTextLine,
  DrawTextStyle
} from "@draw/types";
export type { DrawRasterTextStyle };

// =============================================================================
// Deprecated flat names — forwards to the namespaces above
// =============================================================================

/** @deprecated Use {@link DrawNs.parseColor | Draw.parseColor}. Removed in 1.0. */
export const cssColour = DrawNs.parseColor;
/** @deprecated Use {@link DrawNs.getRelativeLuminance | Draw.getRelativeLuminance}. Removed in 1.0. */
export const relativeLuminance = DrawNs.getRelativeLuminance;
/** @deprecated Use {@link DrawNs.withAlpha | Draw.withAlpha}. Removed in 1.0. */
export const translucent = DrawNs.withAlpha;
/** @deprecated Use {@link DrawNs.measureText | Draw.measureText}. Removed in 1.0. */
export const measureText = DrawNs.measureText;
/** @deprecated Use {@link DrawNs.measureWidest | Draw.measureWidest}. Removed in 1.0. */
export const widestText = DrawNs.measureWidest;
/** @deprecated Use {@link DrawNs.wrapText | Draw.wrapText}. Removed in 1.0. */
export const wrapText = DrawNs.wrapText;
/** @deprecated Use {@link DrawNs.render | Draw.render}. Removed in 1.0. */
export const renderDrawList = DrawNs.render;
/** @deprecated Use {@link DrawNs.Backend.renderNode | Draw.Backend.renderNode}. Removed in 1.0. */
export const renderNode = DrawNs.Backend.renderNode;
/** @deprecated Use {@link DrawNs.Backend.SvgSurface | Draw.Backend.SvgSurface}. Removed in 1.0. */
export const SvgSurface = DrawNs.Backend.SvgSurface;
/** @deprecated Use {@link DrawNs.Backend.SvgSurface | Draw.Backend.SvgSurface}. Removed in 1.0. */
export type SvgSurface = DrawNs.Backend.SvgSurface;
/**
 * @deprecated Use {@link DrawNs.toSvg | Draw.toSvg}, which checks the list and its options.
 * This form keeps writing what it is given, as it always did. Removed in 1.0.
 */
export const toSvg = toSvgLenient;
/**
 * @deprecated Use {@link DrawNs.Backend.createRasterSurface | Draw.Backend.createRasterSurface}, which takes
 * `{ scale }`. Removed in 1.0.
 */
export function createRasterSurface(
  canvas: DrawNs.Backend.BasicRasterCanvas,
  scale?: number
): DrawSurface {
  checkRasterCanvas("createRasterSurface", canvas);
  return createScaledRasterSurface(canvas, scale);
}
/** @deprecated Use {@link DrawNs.Backend.downsample | Draw.Backend.downsample}. Removed in 1.0. */
export const downsample = DrawNs.Backend.downsample;
/** @deprecated Use {@link DrawNs.Backend.normalizeSamples | Draw.Backend.normalizeSamples}. Removed in 1.0. */
export const normalizeSamples = DrawNs.Backend.normalizeSamples;
/**
 * @deprecated Use {@link DrawNs.toRgba | Draw.toRgba}, which checks the list and its options
 * and refuses a list with no area. This form keeps its old reading — a 0×0 list rasterises
 * to a 1×1 transparent image. Removed in 1.0.
 */
export const rasterizeToRgba = rasterizeToRgbaLenient;
/** @deprecated Use {@link DrawNs.isSimpleText | Draw.isSimpleText}. Removed in 1.0. */
export const isSimpleText = DrawNs.isSimpleText;
/** @deprecated Use {@link DrawNs.getTextFeatures | Draw.getTextFeatures}. Removed in 1.0. */
export const textFeaturesOf = DrawNs.getTextFeatures;
/** @deprecated Use {@link DrawNs.Backend.BasicRasterCanvas | Draw.Backend.BasicRasterCanvas}. Removed in 1.0. */
export const BasicRasterCanvas = DrawNs.Backend.BasicRasterCanvas;
/** @deprecated Use {@link DrawNs.Backend.BasicRasterCanvas | Draw.Backend.BasicRasterCanvas}. Removed in 1.0. */
export type BasicRasterCanvas = DrawNs.Backend.BasicRasterCanvas;
/** @deprecated Use {@link DrawNs.Backend.fontHasGlyph | Draw.Backend.fontHasGlyph}. Removed in 1.0. */
export const fontHasGlyph = DrawNs.Backend.fontHasGlyph;
/**
 * @deprecated Use {@link DrawNs.parseRasterFont | Draw.parseRasterFont}, which takes the bytes or
 * `{ data, collectionIndex }` and throws for bytes it cannot read (this form returns an empty
 * face, and face 0 for an index past the end of a collection). Removed in 1.0.
 */
export const parseRasterFont = parseRasterFontLenient;
/** @deprecated Use {@link DrawNs.Geometry.arcToCubics | Draw.Geometry.arcToCubics}. Removed in 1.0. */
export const arcToCubics = DrawNs.Geometry.arcToCubics;
/**
 * @deprecated Use {@link DrawNs.Geometry.flattenPath | Draw.Geometry.flattenPath}, which takes
 * `{ scale }` and checks it. This form multiplies by whatever `scale` it is given, as it always
 * did. Removed in 1.0.
 */
export function flattenPath(commands: readonly DrawPathCommand[], scale?: number): DrawSubpath[] {
  checkPathCommands("flattenPath", "commands", commands);
  return flattenPathImpl(commands, scale);
}
/** @deprecated Use {@link DrawNs.Geometry.createRectNode | Draw.Geometry.createRectNode}. Removed in 1.0. */
export const rectNode = DrawNs.Geometry.createRectNode;
/** @deprecated Use {@link DrawNs.Geometry.roundedRectToPath | Draw.Geometry.roundedRectToPath}. Removed in 1.0. */
export const roundedRectToPath = DrawNs.Geometry.roundedRectToPath;
/** @deprecated Use {@link DrawNs.Geometry.sectorToPath | Draw.Geometry.sectorToPath}. Removed in 1.0. */
export const sectorToPath = DrawNs.Geometry.sectorToPath;
/** @deprecated Use {@link MatrixNs.IDENTITY | Matrix.IDENTITY}. Removed in 1.0. */
export const IDENTITY = MatrixNs.IDENTITY;
/** @deprecated Use {@link MatrixNs.apply | Matrix.apply}. Removed in 1.0. */
export const apply = MatrixNs.apply;
/** @deprecated Use {@link MatrixNs.multiply | Matrix.multiply}. Removed in 1.0. */
export const multiply = MatrixNs.multiply;
/** @deprecated Use {@link MatrixNs.rotate | Matrix.rotate}. Removed in 1.0. */
export const rotate = MatrixNs.rotate;
/** @deprecated Use {@link MatrixNs.scale | Matrix.scale}. Removed in 1.0. */
export const scale = MatrixNs.scale;
/** @deprecated Use {@link MatrixNs.translate | Matrix.translate}. Removed in 1.0. */
export const translate = MatrixNs.translate;
/** @deprecated Use {@link MatrixNs.getRotation | Matrix.getRotation}. Removed in 1.0. */
export const rotationOf = MatrixNs.getRotation;
/** @deprecated Use {@link MatrixNs.getUniformScale | Matrix.getUniformScale}. Removed in 1.0. */
export const uniformScale = MatrixNs.getUniformScale;

// =============================================================================
// Deprecated unprefixed type names — every flat type of this entry carries `Draw`
// =============================================================================

/** @deprecated Use {@link DrawToSvgOptions}. Removed in 1.0. */
export type ToSvgOptions = DrawToSvgOptions;
/** @deprecated Use {@link DrawRasterSurfaceOptions}. Removed in 1.0. */
export type RasterSurfaceOptions = DrawRasterSurfaceOptions;
/** @deprecated Use {@link DrawRasterizeOptions}. Removed in 1.0. */
export type RasterizeOptions = DrawRasterizeOptions;
/** @deprecated Use {@link DrawRasterizedImage}. Removed in 1.0. */
export type RasterizedImage = DrawRasterizedImage;
/** @deprecated Use {@link DrawRgbaImage}. Removed in 1.0. */
export type RgbaImage = DrawRgbaImage;
/** @deprecated Use {@link DrawRasterPoint}. Removed in 1.0. */
export type RasterPoint = DrawRasterPoint;
/** @deprecated Use {@link DrawGlyphOutline}. Removed in 1.0. */
export type GlyphOutline = DrawGlyphOutline;
/** @deprecated Use {@link DrawRasterFont}. Removed in 1.0. */
export type RasterFont = DrawRasterFont;
/** @deprecated Use {@link DrawRasterFontSource}. Removed in 1.0. */
export type RasterFontSource = DrawRasterFontSource;
/** @deprecated Use {@link DrawRgba01}. Removed in 1.0. */
export type Rgba01 = DrawRgba01;
/** @deprecated Use {@link DrawRasterTextStyle}. Removed in 1.0. */
export type RasterTextStyle = DrawRasterTextStyle;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
