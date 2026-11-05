/**
 * The checks every chart renderer runs on its `ChartRenderOptions` before it draws.
 *
 * `toSvg` used to check nothing and `toPng` only `scale` and `dpi`, so a `NaN` width became
 * `width="NaN"` (or `"0"`) in the markup, a negative one a negative `viewBox`, a PNG one pixel
 * wide, or a `DrawArgumentError` from an internal canvas the caller never called; a number
 * where the title or the background belonged failed with a `TypeError`. One check, shared by
 * the classic and the Office 2016 renderers and by every public name that reaches them.
 */

import { checkTopology } from "@excel/chart/render/topojson";
import { ExcelChartOptionsError } from "@excel/errors";
import { checkBackgroundWith } from "@utils/background-check";
import { describeValue as describe } from "@utils/describe-value";
import { isNoPaintColor, parseCssColor } from "@utils/svg-lex";
import type { Rgba01 } from "@utils/svg-lex";

const PROJECTIONS: readonly string[] = ["mercator", "miller", "albers", "robinson"];
const THEME_SLOTS: readonly string[] = [
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

function checkOptionalString(where: string, field: string, value: unknown): void {
  if (value !== undefined && typeof value !== "string") {
    fail(where, field, "a string", value);
  }
}

function checkOptionalCssColor(where: string, field: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "string" || (!isNoPaintColor(value) && parseCssColor(value) === undefined)) {
    fail(
      where,
      field,
      "a CSS colour (#rgb, #rrggbb, #rrggbbaa, rgb(), hsl(), a colour name, transparent or none)",
      value
    );
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

/**
 * Check a renderer's options, naming `where` — the public call (`ChartClassic.toSvg`,
 * `Chart.toPng`, …) — and each field by its path from `options`.
 *
 * @throws {ExcelChartOptionsError} an option has no valid reading.
 */
export function checkChartRenderOptions(where: string, options: unknown): void {
  if (options === undefined) {
    return;
  }
  if (!isRecord(options)) {
    fail(where, "options", "an options object", options);
  }
  checkOptionalPositive(where, "options.width", options.width);
  checkOptionalPositive(where, "options.height", options.height);
  checkOptionalPositive(where, "options.scale", options.scale);
  checkOptionalPositive(where, "options.dpi", options.dpi);
  checkOptionalString(where, "options.title", options.title);
  checkOptionalBackground(where, "options.background", options.background);
  checkOptionalCssColor(where, "options.backgroundColor", options.backgroundColor);
  const fonts = options.fonts;
  if (fonts !== undefined) {
    if (!Array.isArray(fonts)) {
      fail(where, "options.fonts", "an array of fonts", fonts);
    }
    fonts.forEach((font: unknown, i) => {
      if (typeof font !== "object" || font === null) {
        fail(
          where,
          `options.fonts[${i}]`,
          "font bytes, a RasterFont or { data, collectionIndex }",
          font
        );
      }
    });
  }
  if (options.useSystemFonts !== undefined && typeof options.useSystemFonts !== "boolean") {
    fail(where, "options.useSystemFonts", "a boolean", options.useSystemFonts);
  }
  const themeColors = options.themeColors;
  if (themeColors !== undefined) {
    if (!isRecord(themeColors)) {
      fail(
        where,
        "options.themeColors",
        'an object of theme colours { accent1: "RRGGBB", … }',
        themeColors
      );
    }
    for (const [slot, value] of Object.entries(themeColors)) {
      if (!THEME_SLOTS.includes(slot)) {
        fail(
          where,
          "options.themeColors",
          `an object whose keys are theme slots (${THEME_SLOTS.join(", ")})`,
          slot
        );
      }
      if (value !== undefined && (typeof value !== "string" || !/^[0-9A-Fa-f]{6}$/.test(value))) {
        fail(where, `options.themeColors.${slot}`, 'six hex digits "RRGGBB"', value);
      }
    }
  }
  const regionMap = options.regionMap;
  if (regionMap === undefined) {
    return;
  }
  if (!isRecord(regionMap)) {
    fail(where, "options.regionMap", "an object { topology, objectName, … }", regionMap);
  }
  checkTopology(regionMap.topology, regionMap.objectName, {
    where,
    topology: "options.regionMap.topology",
    objectName: "options.regionMap.objectName"
  });
  const rules =
    regionMap.match === undefined
      ? []
      : Array.isArray(regionMap.match)
        ? regionMap.match
        : [regionMap.match];
  rules.forEach((rule: unknown, i) => {
    if (typeof rule !== "string" || (rule !== "id" && !/^property:.+/.test(rule))) {
      fail(
        where,
        Array.isArray(regionMap.match)
          ? `options.regionMap.match[${i}]`
          : "options.regionMap.match",
        '"id" or "property:<key>"',
        rule
      );
    }
  });
  if (regionMap.projection !== undefined && !PROJECTIONS.includes(regionMap.projection as string)) {
    fail(
      where,
      "options.regionMap.projection",
      `one of ${PROJECTIONS.map(p => JSON.stringify(p)).join(", ")}`,
      regionMap.projection
    );
  }
  checkOptionalCssColor(where, "options.regionMap.strokeColor", regionMap.strokeColor);
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
