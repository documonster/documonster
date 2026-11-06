/**
 * `Draw.toPng` — a display list straight to a PNG file.
 *
 * The one file in `draw/` that may import `archive/`: a PNG is a DEFLATE stream plus CRC-32-checked chunks, and
 * the encoder for it lives with those primitives (`@archive/png`). A second encoder here would put a second
 * DEFLATE in the tree; lowering the first to `utils` would put a compression engine at Layer 0 for every
 * consumer. The exception is registered in `scripts/verify-layers.ts`, and `scripts/treeshake-verify.ts` pins
 * that only a caller of `Draw.toPng` pays for DEFLATE — `Draw.toSvg` and `Draw.toRgba` stay archive-free.
 */

import { encodePng } from "@archive/png";
import { checkedRasterize, positive } from "@draw/api/rasterize";
import { at } from "@draw/arguments";
import type { RasterizedImage, RasterizeOptions } from "@draw/raster/surface";
import type { DrawList } from "@draw/types";
import { check } from "@utils/schema";

/** What {@link toPng} takes: the rasteriser's options plus the resolution written into the file. */
export interface ToPngOptions extends RasterizeOptions {
  /**
   * Physical resolution in dots per inch, written as a `pHYs` chunk so Word, Excel and print pipelines size
   * the image as intended rather than assuming 96 DPI. Omitted, the file states no resolution.
   */
  readonly dpi?: number;
}

/**
 * What {@link toPng} returns: the PNG file and what the rasteriser could not draw — the report `Draw.toRgba`
 * gives, because a PNG with holes where a script's glyphs should be is otherwise indistinguishable from a
 * correct one.
 */
export interface PngImage extends Omit<RasterizedImage, "data"> {
  /** The PNG file. */
  readonly png: Uint8Array;
}

const dpiSchema = positive().optional();

/**
 * Render a display list and encode it as PNG, for the public call `where`.
 *
 * @internal Shared with `Mermaid.toPng`; the published form is {@link toPng}.
 */
export function rasterizePng(
  where: string,
  list: DrawList,
  options: ToPngOptions | undefined
): PngImage {
  const dpi = options?.dpi;
  check(dpiSchema, dpi, at(where, "options.dpi"));
  const { data, ...image } = checkedRasterize(where, list, options);
  return { ...image, png: encodePng(data, image.width, image.height, { dpi }) };
}

/**
 * Render a display list to a PNG file — `Draw.toRgba` and PNG encoding in one call.
 *
 * Check `uncoveredCodePoints` and `textWarnings` on the result: they say what the picture is missing, exactly as
 * `Draw.toRgba` reports it.
 *
 * @throws {DrawArgumentError} the list or an option has no valid reading — what `Draw.toRgba` refuses, and a
 *   `dpi` that is not a positive finite number.
 * @throws {DrawError} the output would exceed the 64-megapixel limit.
 */
export function toPng(list: DrawList, options?: ToPngOptions): PngImage {
  return rasterizePng("Draw.toPng", list, options);
}
