/**
 * The checked entry to the rasteriser, shared by `Draw.toRgba`, `Draw.toPng` and `Mermaid.toPng` so each names
 * itself in an argument error. Not part of the `Draw` namespace: `api/draw.ts` is published whole.
 */

import { at, checkBackground, checkDrawList, checkPositive } from "@draw/arguments";
import type { RasterFontSource } from "@draw/raster/glyph-outline";
import { rasterizeToRgba } from "@draw/raster/surface";
import type { RasterizedImage, RasterizeOptions } from "@draw/raster/surface";
import type { DrawList } from "@draw/types";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import type { Rgba01 } from "@utils/svg-lex";

/** A finite number above zero: a scale, a size, a resolution. */
export const positive = () =>
  s
    .number({ finite: true, expected: "a finite number > 0" })
    .refine(v => v > 0, "a finite number > 0");
/** Checked by `checkBackground`, which knows the CSS colour grammar. */
export const background = () => s.unknown<string | Rgba01>().optional();

/**
 * The rasteriser's own settings, shared with `Mermaid.toPng` so the two cannot drift; a caller that sizes the picture
 * itself (`width`, `height`) or paints a background adds those.
 */
export const rasterShape = {
  scale: positive().optional(),
  samples: s
    .number({ finite: true })
    .refine(v => v >= 1, "a finite number >= 1")
    .optional(),
  fonts: s
    .readonlyArray(
      s.custom<RasterFontSource>(
        v => typeof v === "object" && v !== null,
        "font bytes, a RasterFont or { data, collectionIndex }"
      ),
      { expected: "an array of fonts" }
    )
    .optional(),
  useSystemFonts: s.boolean().optional()
};

const rasterizeOptionsSchema = s.object(
  {
    width: positive().optional(),
    height: positive().optional(),
    background: background(),
    ...rasterShape
  },
  { expected: "an object", readonly: true }
);
const _rasterizeOptionsMatch: SchemaMatches<typeof rasterizeOptionsSchema, RasterizeOptions> = true;
void _rasterizeOptionsMatch;

/**
 * `Draw.toRgba` with its argument errors named for the public call `where` — `Draw.toPng` and `Mermaid.toPng`
 * rasterise through it too.
 */
export function checkedRasterize(
  where: string,
  list: DrawList,
  options: RasterizeOptions | undefined
): RasterizedImage {
  checkDrawList(where, list);
  checkPositive(where, "list.width", list.width);
  checkPositive(where, "list.height", list.height);
  check(rasterizeOptionsSchema.optional(), options, at(where, "options"));
  if (options?.background !== undefined) {
    checkBackground(where, "options.background", options.background);
  }
  return rasterizeToRgba(list, options);
}
