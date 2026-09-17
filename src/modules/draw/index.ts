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
 * import { toSvg, type DrawList } from "documonster/draw";
 *
 * const list: DrawList = {
 *   width: 100,
 *   height: 50,
 *   children: [
 *     { kind: "rect", x: 10, y: 10, width: 80, height: 30,
 *       paint: { fill: { r: 0.2, g: 0.4, b: 0.8, a: 1 } } }
 *   ]
 * };
 * const svg = toSvg(list);
 * ```
 *
 * ## Backends
 *
 * Two ship here: {@link toSvg} serialises to markup, and {@link rasterizeToRgba}
 * paints pixels. A PDF page is the third and lives in `documonster/pdf`
 * (`createPdfDrawSurface`), because it draws onto a page builder rather than
 * producing a standalone artefact. Any of them can be driven from the same list —
 * that is the whole point of having one.
 *
 * Pixels rather than a PNG: encoding one needs DEFLATE, which sits a layer above this
 * module. Pair {@link rasterizeToRgba} with any encoder. The pairing this repository
 * uses internally is `excel/chart/render/draw-raster-png.ts`, which is that and nothing
 * else — it is not currently part of the published surface.
 */

export { cssColour, relativeLuminance, translucent } from "@draw/colour";
export { POINTS_PER_PIXEL, measureText, widestText, wrapText } from "@draw/text";
export type { DrawSurface } from "@draw/surface";
export { renderDrawList, renderNode } from "@draw/render";
export { SvgSurface, toSvg } from "@draw/svg";
export type { ToSvgOptions } from "@draw/svg";
export {
  createRasterSurface,
  downsample,
  normalizeSamples,
  rasterizeToRgba
} from "@draw/raster/surface";
export type { RasterizeOptions, RasterizedImage, RgbaImage } from "@draw/raster/surface";
/**
 * Ask, before rendering, whether text can be laid out one glyph per code point.
 *
 * `RasterizedImage.textWarnings` reports this after the fact; these answer beforehand,
 * so a caller can route Arabic or Devanagari to `toSvg` — which is correct for them —
 * instead of producing pixels it will have to discard.
 */
export { isSimpleText, textFeaturesOf } from "@utils/complex-text";
export { BasicRasterCanvas } from "@draw/raster/canvas";
/**
 * Reachable from {@link BasicRasterCanvas}'s own signatures — its point lists and the
 * font it rasterises glyphs from. Without the names a consumer can call the methods but
 * cannot declare a variable to hold what they take.
 */
export type { RasterPoint } from "@draw/raster/canvas";
export type { GlyphOutline, RasterFont } from "@draw/raster/glyph-outline";
/**
 * Parse font bytes once, to hand the result to several renders.
 *
 * `RasterizeOptions.fonts` accepts bytes directly, but re-parses them on every call —
 * for a CJK face that means rebuilding a 43,000-entry `cmap` each time, and a fresh
 * object also misses the glyph cache, which is keyed on outline identity. It is also
 * the only way to choose a face inside a `.ttc`.
 *
 * `fontHasGlyph` answers coverage for any `RasterFont`, including one a consumer
 * implemented themselves without the optional `hasGlyph` member.
 */
export { fontHasGlyph, parseRasterFont } from "@draw/raster/glyph-outline";
/**
 * How a font may be supplied: bytes, a parsed face, or a face inside a collection.
 *
 * Named because it appears in `RasterizeOptions.fonts` and
 * `BasicRasterCanvas.setFonts`.
 *
 * There is deliberately **no process-wide font registry**, published or otherwise.
 * Fonts belong to a render — `RasterizeOptions.fonts` — or to a canvas —
 * `BasicRasterCanvas.setFonts`. A module-level setter would let one caller change what
 * another caller's canvas draws with; the one that existed was used by nothing but the
 * tests, and removing it took two mutable globals with it.
 */
export type { RasterFontSource } from "@draw/raster/glyph-outline";
export {
  DEFAULT_TEXT_FAMILY,
  IDENTITY,
  apply,
  arcToCubics,
  multiply,
  rotate,
  flattenPath,
  rectNode,
  roundedRectToPath,
  rotationOf,
  scale,
  sectorToPath,
  translate,
  uniformScale
} from "@draw/types";
/**
 * Re-exported because it is reachable from this module's own signatures —
 * `DrawPaint.fill`, `DrawTextStyle.fill` and the return of {@link cssColour} are
 * all `Rgba01`. Without the name a consumer can assign one but cannot declare one,
 * and would have to resort to `ReturnType<typeof cssColour>`.
 */
export type { Rgba01 } from "@utils/svg-lex";
export type {
  DrawBox,
  DrawClip,
  DrawSubpath,
  DrawList,
  DrawMatrix,
  DrawNode,
  DrawPaint,
  DrawPathCommand,
  DrawPoint,
  DrawTextAnchor,
  DrawTextLine,
  DrawTextStyle,
  RasterTextStyle
} from "@draw/types";
