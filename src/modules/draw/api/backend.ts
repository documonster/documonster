/**
 * `Draw.Backend` — the plumbing beneath `Draw.toSvg`, `Draw.toRgba` and `Draw.render`,
 * for a caller driving a surface by hand: the SVG and raster surfaces themselves, the
 * canvas the raster surface paints, walking one node rather than a whole list, the
 * supersampling helpers, and the font coverage check the rasteriser uses.
 *
 * Most callers want `Draw.toSvg(list)` or `Draw.toRgba(list)`, which do all of this in
 * one call.
 */

import {
  checkMatrix,
  checkNode,
  checkOptionalObject,
  checkRasterCanvas,
  checkSurface,
  fail
} from "@draw/arguments";
import type { BasicRasterCanvas } from "@draw/raster/canvas";
import { fontHasGlyph as fontHasGlyphImpl } from "@draw/raster/glyph-outline";
import type { RasterFont } from "@draw/raster/glyph-outline";
import {
  createRasterSurface as createRasterSurfaceImpl,
  downsample as downsampleImpl,
  normalizeSamples as normalizeSamplesImpl
} from "@draw/raster/surface";
import type { RasterSurfaceOptions } from "@draw/raster/surface";
import { renderNode as renderNodeImpl } from "@draw/render";
import type { DrawSurface } from "@draw/surface";
import type { DrawMatrix, DrawNode } from "@draw/types";

export { SvgSurface } from "@draw/svg";
export { BasicRasterCanvas } from "@draw/raster/canvas";

/**
 * Whether `font` can draw `codePoint`, for a font that may not implement `hasGlyph`: its
 * `hasGlyph` when it has one, else whether `getOutline` returns an outline.
 *
 * @throws {DrawArgumentError} `font` is not a `RasterFont` (no `getOutline` function).
 */
export function fontHasGlyph(font: RasterFont, codePoint: number): boolean {
  if (
    typeof font !== "object" ||
    font === null ||
    typeof (font as { getOutline?: unknown }).getOutline !== "function"
  ) {
    fail("Draw.Backend.fontHasGlyph", "font", "a RasterFont (an object with getOutline)", font);
  }
  return fontHasGlyphImpl(font, codePoint);
}

function checkPositiveInteger(where: string, field: string, value: unknown): void {
  if (typeof value !== "number" || !Number.isInteger(value) || value <= 0) {
    fail(where, field, "a positive integer", value);
  }
}

/**
 * Build a raster surface over an existing canvas. `options.scale` maps list user units to
 * canvas pixels (default 1).
 */
export function createRasterSurface(
  canvas: BasicRasterCanvas,
  options?: RasterSurfaceOptions
): DrawSurface {
  checkRasterCanvas("Draw.Backend.createRasterSurface", canvas);
  checkOptionalObject("Draw.Backend.createRasterSurface", "options", options);
  const scale = options?.scale;
  if (scale !== undefined && (typeof scale !== "number" || !Number.isFinite(scale) || scale <= 0)) {
    fail("Draw.Backend.createRasterSurface", "options.scale", "a finite number > 0", scale);
  }
  return createRasterSurfaceImpl(canvas, options);
}

/**
 * Box-filter a straight-alpha RGBA buffer of `width·factor × height·factor` pixels down by
 * `factor`, weighting the colour channels by alpha.
 *
 * @throws {DrawArgumentError} a size is not a positive integer, or `source` is not
 *   `width · factor · height · factor · 4` bytes.
 */
export function downsample(
  source: Uint8Array,
  width: number,
  height: number,
  factor: number
): Uint8Array {
  checkPositiveInteger("Draw.Backend.downsample", "width", width);
  checkPositiveInteger("Draw.Backend.downsample", "height", height);
  checkPositiveInteger("Draw.Backend.downsample", "factor", factor);
  const expected = width * factor * height * factor * 4;
  const isBytes =
    ArrayBuffer.isView(source) &&
    (source as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1;
  if (!isBytes || source.length !== expected) {
    fail(
      "Draw.Backend.downsample",
      "source",
      `a Uint8Array of ${expected} bytes (width·factor × height·factor RGBA)`,
      isBytes ? `${source.length} bytes` : source
    );
  }
  return downsampleImpl(source, width, height, factor);
}

/**
 * The supersampling factor a render of `pixelWidth × pixelHeight` will use for a requested
 * `samples` (default 3): floored, clamped to 1–4, and lowered until the internal canvas
 * fits the memory limit.
 */
export function normalizeSamples(
  samples: number | undefined,
  pixelWidth: number,
  pixelHeight: number
): number {
  checkPositiveInteger("Draw.Backend.normalizeSamples", "pixelWidth", pixelWidth);
  checkPositiveInteger("Draw.Backend.normalizeSamples", "pixelHeight", pixelHeight);
  return normalizeSamplesImpl(samples, pixelWidth, pixelHeight);
}

/** Draw one node onto a surface, folding `ctm` into its geometry. */
export function renderNode(node: DrawNode, ctm: DrawMatrix, surface: DrawSurface): void {
  checkNode("Draw.Backend.renderNode", "node", node);
  checkMatrix("Draw.Backend.renderNode", "ctm", ctm);
  checkSurface("Draw.Backend.renderNode", surface);
  renderNodeImpl(node, ctm, surface);
}
