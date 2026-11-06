/**
 * The checks every chart renderer runs on its `ChartRenderOptions` before it draws.
 *
 * `toSvg` used to check nothing and `toPng` only `scale` and `dpi`, so a `NaN` width became
 * `width="NaN"` (or `"0"`) in the markup, a negative one a negative `viewBox`, a PNG one pixel
 * wide, or a `DrawArgumentError` from an internal canvas the caller never called; a number
 * where the title or the background belonged failed with a `TypeError`. One check, shared by
 * the classic and the Office 2016 renderers and by every public name that reaches them.
 */

import type { RasterFontSource } from "@draw/raster/glyph-outline";
import type {
  ChartRenderOptions,
  RegionMapDataOptions,
  RegionMapMatchRule
} from "@excel/chart/render/chart-renderer";
import { checkTopology } from "@excel/chart/render/topojson";
import type { ChartThemeColors } from "@excel/chart/shared/chart-utils";
import { ExcelChartOptionsError } from "@excel/errors";
import type {} from "@excel/utils/arguments";
import { checkBackgroundWith } from "@utils/background-check";
import { describeValue as describe } from "@utils/describe-value";
import { check, s } from "@utils/schema";
import type { Schema, SchemaMatches } from "@utils/schema";
import { isNoPaintColor, parseCssColor } from "@utils/svg-lex";
import type { Rgba01 } from "@utils/svg-lex";
import type { OoxmlThemeColorName } from "@utils/theme-colors";

const THEME_SLOTS: readonly OoxmlThemeColorName[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink"
];

function fail(where: string, field: string, expected: string, value: unknown): never {
  throw new ExcelChartOptionsError(
    field,
    `${where}: ${field} must be ${expected}, got ${describe(value)}`
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function checkOptionalPositive(where: string, field: string, value: unknown): void {
  if (value !== undefined && (typeof value !== "number" || !Number.isFinite(value) || value <= 0)) {
    fail(where, field, "a finite number > 0", value);
  }
}

/** A background as every renderer takes it — the one check (`checkBackgroundWith`, `@utils/background-check`). */
function checkOptionalBackground(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkBackgroundWith(where, field, value, fail);
  }
}

/**
 * The fill a caller asked for, as a CSS colour: `background`, then the deprecated
 * `backgroundColor`; `undefined` when neither is given, and the chart area's own fill applies.
 * An `{ r, g, b, a }` is written in percentages, which carry its channels without rounding them
 * to 8 bits; `"transparent"` paints nothing.
 */
export function chartBackground(options: {
  readonly background?: string | Rgba01;
  readonly backgroundColor?: string;
}): string | undefined {
  const background: string | Rgba01 | undefined = options.background ?? options.backgroundColor;
  if (background === undefined) {
    return undefined;
  }
  if (typeof background === "string") {
    return background;
  }
  const pct = (channel: number): string => `${Number((channel * 100).toFixed(4))}%`;
  return `rgba(${pct(background.r)}, ${pct(background.g)}, ${pct(background.b)}, ${Number(
    background.a.toFixed(6)
  )})`;
}

/**
 * Whether a caller chose the fonts a PNG is drawn with — `fonts`, or `useSystemFonts: false`. Only
 * the built-in rasteriser can honour that; a browser's own SVG engine draws with the page's fonts,
 * so `toPng` uses the browser engine only when the caller left the fonts to the platform.
 */
export function requestsFontControl(options: {
  readonly fonts?: readonly unknown[];
  readonly useSystemFonts?: boolean;
}): boolean {
  return options.fonts !== undefined || options.useSystemFonts === false;
}

function chartOptionsError(field: string, message: string): ExcelChartOptionsError {
  return new ExcelChartOptionsError(field, message);
}

const positive = () =>
  s.number({ finite: true, min: Number.MIN_VALUE, expected: "a finite number > 0" }).optional();

const cssColor = () =>
  s
    .custom<string>(
      v => typeof v === "string" && (isNoPaintColor(v) || parseCssColor(v) !== undefined),
      "a CSS colour (#rgb, #rrggbb, #rrggbbaa, rgb(), hsl(), a colour name, transparent or none)"
    )
    .optional();

const themeColorsShape = Object.fromEntries(
  THEME_SLOTS.map(slot => [
    slot,
    s
      .custom<string>(
        v => typeof v === "string" && /^[0-9A-Fa-f]{6}$/.test(v),
        'six hex digits "RRGGBB"'
      )
      .optional()
  ])
) as Record<OoxmlThemeColorName, Schema<string, true>>;
const themeColorsSchema = s.object(themeColorsShape, {
  expected: 'an object of theme colours { accent1: "RRGGBB", … }',
  unknownKeys: "reject",
  readonly: true
});
const _themeColorsMatch: SchemaMatches<typeof themeColorsSchema, ChartThemeColors> = true;
void _themeColorsMatch;

const matchRule = s.custom<RegionMapMatchRule>(
  v => typeof v === "string" && (v === "id" || /^property:.+/.test(v)),
  '"id" or "property:<key>"'
);
const regionMapSchema = s.object(
  {
    // Checked by `checkTopology`, which names the path inside the topology.
    topology: s.unknown(),
    objectName: s.unknown<string>(),
    match: s.union([matchRule, s.array(matchRule)]).optional(),
    projection: s.literal("mercator", "miller", "albers", "robinson").optional(),
    strokeColor: cssColor()
  },
  { expected: "an object { topology, objectName, … }" }
);
const _regionMapMatch: SchemaMatches<typeof regionMapSchema, RegionMapDataOptions> = true;
void _regionMapMatch;

const renderOptionsSchema = s.object(
  {
    width: positive(),
    height: positive(),
    title: s.string().optional(),
    // Checked by `checkBackgroundWith`, the one background check every renderer shares.
    background: s.unknown<string | Rgba01>().optional(),
    backgroundColor: cssColor(),
    fonts: s
      .readonlyArray(
        s.custom<RasterFontSource>(
          v => typeof v === "object" && v !== null,
          "font bytes, a RasterFont or { data, collectionIndex }"
        ),
        { expected: "an array of fonts" }
      )
      .optional(),
    useSystemFonts: s.boolean().optional(),
    themeColors: themeColorsSchema.optional(),
    scale: positive(),
    dpi: positive(),
    regionMap: regionMapSchema.optional()
  },
  { expected: "an options object" }
);
const _renderOptionsMatch: SchemaMatches<typeof renderOptionsSchema, ChartRenderOptions> = true;
void _renderOptionsMatch;

/**
 * Check a renderer's options, naming `where` — the public call (`ChartClassic.toSvg`,
 * `Chart.toPng`, …) — and each field by its path from `options`.
 *
 * @throws {ExcelChartOptionsError} an option has no valid reading.
 */
export function checkChartRenderOptions(where: string, options: unknown): void {
  check(renderOptionsSchema.optional(), options, {
    where,
    field: "options",
    error: chartOptionsError
  });
  if (options === undefined) {
    return;
  }
  checkOptionalBackground(where, "options.background", options.background);
  const regionMap = options.regionMap;
  if (regionMap !== undefined) {
    checkTopology(regionMap.topology, regionMap.objectName, {
      where,
      topology: "options.regionMap.topology",
      objectName: "options.regionMap.objectName"
    });
  }
}

/**
 * The surface and destination box the deprecated `drawChartExPdf` draws into: a
 * surface with the marks a chart puts down, and a box of finite numbers with an area.
 *
 * @throws {ExcelChartOptionsError} either has no valid reading.
 */
export function checkDrawPdfArguments(where: string, surface: unknown, rect: unknown): void {
  for (const method of ["drawRect", "drawLine", "drawText"] as const) {
    if (typeof (surface as Record<string, unknown> | null)?.[method] !== "function") {
      throw new ExcelChartOptionsError(
        "surface",
        `${where}: surface must be a PDF page or a surface like one — it has no ${method}()`
      );
    }
  }
  if (!isRecord(rect)) {
    fail(where, "rect", "a box { x, y, width, height }", rect);
  }
  for (const key of ["x", "y"] as const) {
    if (typeof rect[key] !== "number" || !Number.isFinite(rect[key])) {
      fail(where, `rect.${key}`, "a finite number", rect[key]);
    }
  }
  checkOptionalPositive(where, "rect.width", rect.width ?? Number.NaN);
  checkOptionalPositive(where, "rect.height", rect.height ?? Number.NaN);
}
