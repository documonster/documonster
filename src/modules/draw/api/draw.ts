/**
 * `Draw` namespace surface — display-list rendering, colour, text measurement
 * and path construction.
 *
 * `import { Draw } from "documonster/draw"` →
 *   `Draw.toSvg(list)`, `Draw.toRgba(list)`, `Draw.render(list, surface)`,
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

import {
  checkBackground,
  checkColor,
  checkDrawList,
  checkFinite,
  checkOptionalBoolean,
  checkOptionalObject,
  checkOptionalPositive,
  checkOptionalString,
  checkPositive,
  checkString,
  checkSurface,
  checkTextStyle,
  fail
} from "@draw/arguments";
import { cssColor, relativeLuminance, translucent } from "@draw/colour";
import { readRasterFont } from "@draw/raster/glyph-outline";
import type { RasterFont } from "@draw/raster/glyph-outline";
import { rasterizeToRgba } from "@draw/raster/surface";
import type { RasterizedImage, RasterizeOptions } from "@draw/raster/surface";
import { renderDrawList } from "@draw/render";
import type { DrawSurface } from "@draw/surface";
import { toSvg as toSvgImpl } from "@draw/svg";
import type { ToSvgOptions } from "@draw/svg";
import { measureText as measureTextImpl, widestText, wrapText as wrapTextImpl } from "@draw/text";
import type { DrawList, DrawTextStyle } from "@draw/types";
import { isSimpleText as isSimpleTextImpl, textFeaturesOf } from "@utils/complex-text";
import type { Rgba01 } from "@utils/svg-lex";

// --- Backends ---

/**
 * Serialise a display list to an SVG document.
 *
 * @throws {DrawArgumentError} the list or an option has no valid reading — a `NaN`
 *   coordinate, a colour that is not `{ r, g, b, a }` in 0–1, an unknown node `kind`.
 */
export function toSvg(list: DrawList, options?: ToSvgOptions): string {
  checkDrawList("Draw.toSvg", list);
  checkOptionalObject("Draw.toSvg", "options", options);
  checkOptionalPositive("Draw.toSvg", "options.width", options?.width);
  checkOptionalPositive("Draw.toSvg", "options.height", options?.height);
  if (options?.background !== undefined) {
    checkBackground("Draw.toSvg", "options.background", options.background);
  }
  checkOptionalString("Draw.toSvg", "options.comment", options?.comment);
  return toSvgImpl(list, options);
}

/**
 * Render a display list to pixels: an RGBA image plus what could not be drawn.
 *
 * Pixels rather than a PNG — pair this with any encoder (`Archive.encodePng` from
 * `documonster/archive`).
 *
 * @throws {DrawArgumentError} the list or an option has no valid reading; a list with no
 *   area (`width` or `height` of 0) has nothing to rasterise.
 * @throws {DrawError} the output would exceed the 64-megapixel limit.
 */
export function toRgba(list: DrawList, options?: RasterizeOptions): RasterizedImage {
  checkDrawList("Draw.toRgba", list);
  checkPositive("Draw.toRgba", "list.width", list.width);
  checkPositive("Draw.toRgba", "list.height", list.height);
  checkOptionalObject("Draw.toRgba", "options", options);
  if (options !== undefined) {
    checkOptionalPositive("Draw.toRgba", "options.width", options.width);
    checkOptionalPositive("Draw.toRgba", "options.height", options.height);
    checkOptionalPositive("Draw.toRgba", "options.scale", options.scale);
    if (options.samples !== undefined) {
      checkFinite("Draw.toRgba", "options.samples", options.samples);
      if (options.samples < 1) {
        fail("Draw.toRgba", "options.samples", "a finite number >= 1", options.samples);
      }
    }
    if (options.background !== undefined) {
      checkBackground("Draw.toRgba", "options.background", options.background);
    }
    if (options.fonts !== undefined) {
      if (!Array.isArray(options.fonts)) {
        fail("Draw.toRgba", "options.fonts", "an array of fonts", options.fonts);
      }
      options.fonts.forEach((font: unknown, i) => {
        if (typeof font !== "object" || font === null) {
          fail(
            "Draw.toRgba",
            `options.fonts[${i}]`,
            "font bytes, a RasterFont or { data, collectionIndex }",
            font
          );
        }
      });
    }
    checkOptionalBoolean("Draw.toRgba", "options.useSystemFonts", options.useSystemFonts);
  }
  return rasterizeToRgba(list, options);
}

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
 * pass it in `RasterizeOptions.fonts` for every render.
 *
 * `collectionIndex` selects a face inside a `.ttc`, and matters: macOS keeps `Songti SC`
 * at Black in face 0 of `Songti.ttc`, so always reading face 0 draws a heavier weight.
 *
 * @throws {DrawArgumentError} `data` is not bytes, or `collectionIndex` names no face.
 * @throws {DrawError} the bytes are not a TrueType font with glyph outlines.
 */
export function parseRasterFont(data: Uint8Array, collectionIndex?: number): RasterFont {
  return readRasterFont(data, collectionIndex);
}

// --- Low level ---
export * as Backend from "@draw/api/backend";
export * as Geometry from "@draw/api/geometry";
