/**
 * `Draw` namespace surface — display-list rendering, colour, text measurement
 * and path construction.
 *
 * `import { Draw } from "documonster/draw"` →
 *   `Draw.toSvg(list)`, `Draw.toPng(list)`, `Draw.toRgba(list)`, `Draw.render(list, surface)`,
 *   `Draw.parseColor("#3366cc")`, `Draw.measureText(text, style)`, …
 *
 * The matrix helpers live in their own `Matrix` namespace: `scale`, `rotate` and
 * `apply` are generic enough that, beside the rendering functions, a bare `scale` reads
 * as "scale the drawing" rather than "build a scaling matrix".
 *
 * Two labelled sub-namespaces hold the low-level pieces, so this one stays about the
 * main tasks: `Draw.Backend` (surfaces, the raster canvas, single-node walking,
 * supersampling, glyph coverage) and `Draw.Geometry` (arcs, sectors, rounded
 * rectangles, path flattening).
 *
 * The directory is `api/` rather than the usual `surface/` because `draw/surface.ts`
 * already names the backend interface, and `@draw/surface` cannot mean both.
 */

import { background, checkedRasterize, positive } from "@draw/api/rasterize";
import {
  at,
  checkBackground,
  checkColor,
  checkDrawList,
  checkString,
  checkSurface,
  checkTextStyle,
  fail
} from "@draw/arguments";
import { cssColor, relativeLuminance, translucent } from "@draw/colour";
import { isByteSource, readRasterFont } from "@draw/raster/glyph-outline";
import type { RasterFont } from "@draw/raster/glyph-outline";
import type { RasterizedImage, RasterizeOptions } from "@draw/raster/surface";
import { renderDrawList } from "@draw/render";
import type { DrawSurface } from "@draw/surface";
import { toSvg as toSvgImpl } from "@draw/svg";
import type { ToSvgOptions } from "@draw/svg";
import { measureText as measureTextImpl, widestText, wrapText as wrapTextImpl } from "@draw/text";
import type { DrawList, DrawTextStyle } from "@draw/types";
import { isSimpleText as isSimpleTextImpl, textFeaturesOf } from "@utils/complex-text";
import { isRecordArgument } from "@utils/object";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import type { Rgba01 } from "@utils/svg-lex";

const toSvgOptionsSchema = s.object(
  {
    width: positive().optional(),
    height: positive().optional(),
    background: background(),
    comment: s.string().optional()
  },
  { expected: "an object", readonly: true }
);
const _toSvgOptionsMatch: SchemaMatches<typeof toSvgOptionsSchema, ToSvgOptions> = true;
void _toSvgOptionsMatch;

// --- Backends ---

/**
 * Serialise a display list to an SVG document.
 *
 * @throws {DrawArgumentError} the list or an option has no valid reading — a `NaN`
 *   coordinate, a colour that is not `{ r, g, b, a }` in 0–1, an unknown node `kind`.
 */
export function toSvg(list: DrawList, options?: ToSvgOptions): string {
  checkDrawList("Draw.toSvg", list);
  check(toSvgOptionsSchema.optional(), options, at("Draw.toSvg", "options"));
  if (options?.background !== undefined) {
    checkBackground("Draw.toSvg", "options.background", options.background);
  }
  return toSvgImpl(list, options);
}

/**
 * Render a display list to pixels: an RGBA image plus what could not be drawn.
 *
 * Pixels rather than a file: for a PNG use `Draw.toPng`; for another format, pair these pixels with its
 * encoder.
 *
 * @throws {DrawArgumentError} the list or an option has no valid reading; a list with no
 *   area (`width` or `height` of 0) has nothing to rasterise.
 * @throws {DrawError} the output would exceed the 64-megapixel limit.
 */
export function toRgba(list: DrawList, options?: RasterizeOptions): RasterizedImage {
  return checkedRasterize("Draw.toRgba", list, options);
}

export { toPng } from "@draw/api/png";

// --- Walking ---

/**
 * Draw a whole list onto a surface — the one walker every backend goes through.
 *
 * @throws {DrawArgumentError} the list has no valid reading, or `surface` lacks a drawing method.
 */
export function render(list: DrawList, surface: DrawSurface): void {
  checkDrawList("Draw.render", list);
  checkSurface("Draw.render", surface);
  renderDrawList(list, surface);
}

// --- Color ---

function checkToken(where: string, token: unknown): void {
  if (token !== undefined && typeof token !== "string") {
    fail(where, "token", "a CSS color string", token);
  }
}

/**
 * An authored CSS colour — `#rgb`, `#rrggbb`, `#rrggbbaa`, `rgb()`, `hsl()` or a named
 * colour — as an IR colour. `none` and `transparent` are fully transparent; a token that
 * cannot be read, or `undefined`, is black (a visible wrong colour is easier to notice than
 * an invisible shape).
 */
export function parseColor(token: string | undefined): Rgba01 {
  checkToken("Draw.parseColor", token);
  return cssColor(token);
}

/**
 * An authored colour at the given alpha, which replaces the token's own. `none` and
 * `transparent` stay fully transparent.
 *
 * @throws {DrawArgumentError} `alpha` is not a number from 0 to 1.
 */
export function withAlpha(token: string | undefined, alpha: number): Rgba01 {
  checkToken("Draw.withAlpha", token);
  if (typeof alpha !== "number" || !(alpha >= 0 && alpha <= 1)) {
    fail("Draw.withAlpha", "alpha", "a number from 0 to 1", alpha);
  }
  return translucent(token, alpha);
}

/**
 * Perceived brightness of a colour on 0…1, by the BT.709 coefficients over the stored
 * channels — for choosing ink that stays readable on a generated fill.
 */
export function getRelativeLuminance(color: Rgba01): number {
  checkColor("Draw.getRelativeLuminance", "color", color);
  return relativeLuminance(color);
}

// --- Text ---

/**
 * The width of `text` in the display list's own units; multi-line input reports the
 * widest line.
 */
export function measureText(text: string, style: DrawTextStyle): number {
  checkString("Draw.measureText", "text", text);
  checkTextStyle("Draw.measureText", "style", style);
  return measureTextImpl(text, style);
}

/** The width of the widest of several strings; an empty list measures 0. */
export function measureWidest(texts: readonly string[], style: DrawTextStyle): number {
  if (!Array.isArray(texts)) {
    fail("Draw.measureWidest", "texts", "an array of strings", texts);
  }
  texts.forEach((text, i) => checkString("Draw.measureWidest", `texts[${i}]`, text));
  checkTextStyle("Draw.measureWidest", "style", style);
  return widestText(texts, style);
}

/**
 * Break text into lines no wider than `maxWidth`: Latin at spaces and hyphens, East Asian
 * text between characters with kinsoku. Existing newlines are kept. A `maxWidth` of 0 or
 * less has no wrap that satisfies it, and `Infinity` needs none: both return the paragraphs
 * unchanged.
 *
 * @throws {DrawArgumentError} `maxWidth` is not a number, or is `NaN`.
 */
export function wrapText(text: string, style: DrawTextStyle, maxWidth: number): string[] {
  checkString("Draw.wrapText", "text", text);
  checkTextStyle("Draw.wrapText", "style", style);
  if (typeof maxWidth !== "number" || Number.isNaN(maxWidth)) {
    fail("Draw.wrapText", "maxWidth", "a number", maxWidth);
  }
  return wrapTextImpl(text, style, maxWidth);
}

/** Whether `text` can be drawn one glyph per code point, in the order given. */
export function isSimpleText(text: string): boolean {
  checkString("Draw.isSimpleText", "text", text);
  return isSimpleTextImpl(text);
}

/**
 * Features of `text` that a one-glyph-per-code-point renderer cannot honour: the scripts
 * that need OpenType shaping, and whether it needs bidi reordering. Simple text reports
 * none.
 */
export function getTextFeatures(text: string): { scripts: string[]; rtl: boolean } {
  checkString("Draw.getTextFeatures", "text", text);
  return textFeaturesOf(text);
}

// --- Fonts ---

/**
 * Parse a `.ttf` / `.ttc` file into a face the rasteriser can draw with — parse once and
 * pass it in `DrawRasterizeOptions.fonts` for every render.
 *
 * `source` is the file's bytes, or `{ data, collectionIndex }` to select a face inside a
 * `.ttc` — which matters: macOS keeps `Songti SC` at Black in face 0 of `Songti.ttc`, so
 * always reading face 0 draws a heavier weight.
 *
 * @throws {DrawArgumentError} `source` is not bytes or `{ data, collectionIndex }`, or
 *   `source.collectionIndex` names no face.
 * @throws {DrawError} the bytes are not a TrueType font with glyph outlines.
 */
export function parseRasterFont(
  source: Uint8Array | { readonly data: Uint8Array; readonly collectionIndex?: number }
): RasterFont {
  if (isByteSource(source)) {
    return readRasterFont(source, 0, { data: "source", collectionIndex: "source" });
  }
  if (!isRecordArgument(source)) {
    fail("Draw.parseRasterFont", "source", "font bytes or { data, collectionIndex }", source);
  }
  return readRasterFont(source.data, source.collectionIndex ?? 0, {
    data: "source.data",
    collectionIndex: "source.collectionIndex"
  });
}

// --- Low level ---
export * as Backend from "@draw/api/backend";
export * as Geometry from "@draw/api/geometry";
