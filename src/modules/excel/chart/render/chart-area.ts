/**
 * The chart area: the chart's own fill and border, from the model.
 *
 * A chart is drawn inside its chart area, which Office paints from `c:chartSpace/c:spPr`
 * (classic) or `cx:chartSpace/cx:spPr` (Office 2016). It is part of the picture, so it is a
 * node of the display list and every backend — SVG, PNG, PDF — shows the same one. The PDF
 * bridges used to draw a white rectangle with a fixed grey border instead, whatever the
 * chart said, while SVG and PNG drew no border at all.
 *
 * What the model leaves out is what Office draws for a chart area on its automatic format.
 * Excel writes that format out in full when it saves a chart (the `c:spPr` of every
 * Excel-saved chart under `excel/__tests__/data`), it is the `cs:chartArea` of Excel's default
 * chart style (`chart-sidecar.ts`), and it is what LibreOffice's MSO-compatible importer
 * applies to a chart space with no `spPr` (`oox/source/drawingml/chart/objectformatter.cxx`):
 *
 * - fill: `bg1` for chart styles 1–32 (the default, `c:style` 2), `lt1` for 33–40 and `dk1`
 *   for 41–48;
 * - border: a solid 0.75 pt line in `tx1` at 15 % luminance plus 85 % — `#D9D9D9` in the
 *   default Office theme.
 *
 * Each of the two is defaulted separately: an `spPr` that sets only a fill keeps the
 * automatic border, and a line that names a width but no colour keeps the automatic colour.
 * `noFill` turns either off.
 *
 * `c:roundedCorners` absent means rounded corners in Excel (which is why every Office writer,
 * this library included, writes `val="0"`); Office 2016 charts have no such element. The radius
 * is not specified anywhere, so {@link ROUNDED_CORNER_RADIUS} approximates Excel's.
 */

import { cssColor } from "@draw/colour";
import { rectNode } from "@draw/types";
import type { DrawNode, DrawPaint } from "@draw/types";
import type { ChartColor, ChartFill, ShapeProperties } from "@excel/chart/model/types";
import { resolveChartColorRgba } from "@excel/chart/shared/chart-utils";
import type { ChartThemeColors } from "@excel/chart/shared/chart-utils";
import { getSpPrFill, getSpPrLine } from "@excel/chart/shared/shape-properties";
import type { Rgba01 } from "@utils/svg-lex";
import { emuToPt } from "@utils/units";

/** Office's automatic chart-area border: `tx1` at 15 % luminance plus 85 % (`#D9D9D9` by default). */
const AUTOMATIC_BORDER: ChartColor = { theme: 0, lumMod: 15000, lumOff: 85000 };
/** Office's automatic chart-area border width: 9525 EMU. */
const AUTOMATIC_BORDER_WIDTH = 0.75;
/** Corner radius of a rounded chart area, in chart units. Not specified; approximates Excel. */
export const ROUNDED_CORNER_RADIUS = 8;

/** What a renderer needs to know about a model's chart area. */
export interface ChartAreaSource {
  /** `c:chartSpace/c:spPr` or `cx:chartSpace/cx:spPr`. */
  readonly spPr?: ShapeProperties | undefined;
  /** `c:style` (1–48), which picks the automatic fill. Default 2. */
  readonly style?: number | undefined;
  /** Whether the corners are rounded (`c:roundedCorners`). */
  readonly rounded: boolean;
  /** The workbook's theme colours, for a scheme colour (`ChartRenderOptions.themeColors`). */
  readonly themeColors?: ChartThemeColors | undefined;
}

/** The chart area's fill, border and corners, resolved. */
export interface ChartArea {
  readonly fill: Rgba01 | undefined;
  readonly stroke: Rgba01 | undefined;
  readonly strokeWidth: number;
  readonly rx: number;
}

/** Resolve a model's chart area. */
export function resolveChartArea(source: ChartAreaSource): ChartArea {
  const fill = source.spPr ? getSpPrFill(source.spPr) : undefined;
  const line = source.spPr ? getSpPrLine(source.spPr) : undefined;
  const resolvedFill = areaFill(fill, source.style, source.themeColors);
  const stroke = line?.noFill
    ? undefined
    : (resolveColor(line?.color, source.themeColors) ??
      resolveColor(AUTOMATIC_BORDER, source.themeColors));
  return {
    fill: resolvedFill,
    stroke,
    strokeWidth:
      stroke === undefined
        ? 0
        : typeof line?.width === "number"
          ? emuToPt(line.width)
          : AUTOMATIC_BORDER_WIDTH,
    rx: source.rounded ? ROUNDED_CORNER_RADIUS : 0
  };
}

/**
 * The chart-area node of a `width` × `height` chart, or `undefined` when it paints nothing.
 * `background`, when given, is the caller's fill (`ChartRenderOptions.background`) and
 * replaces the model's — `"transparent"` or `"none"` paints none — while the border stays the
 * model's. The border is drawn inside the box, as Office draws it: centred on the edge, half
 * of it would fall outside the picture.
 */
export function chartAreaNode(
  model: ChartArea,
  width: number,
  height: number,
  background?: string
): DrawNode | undefined {
  const area: ChartArea =
    background === undefined ? model : { ...model, fill: backgroundFill(background) };
  if (area.fill === undefined && area.stroke === undefined) {
    return undefined;
  }
  const inset = area.stroke === undefined ? 0 : area.strokeWidth / 2;
  const paint: DrawPaint = {
    ...(area.fill === undefined ? {} : { fill: area.fill }),
    ...(area.stroke === undefined ? {} : { stroke: area.stroke, strokeWidth: area.strokeWidth })
  };
  return rectNode(
    { x: inset, y: inset, width: width - 2 * inset, height: height - 2 * inset },
    paint,
    area.rx > 0 ? area.rx : undefined
  );
}

/**
 * The `c:style` a classic model carries: its own `style`, or the one inside the
 * `mc:AlternateContent` Excel writes it in (`c14:style` 101–148 is `c:style` 1–48).
 */
export function classicChartStyle(model: {
  readonly style?: number;
  readonly alternateContentStyle?: string;
}): number | undefined {
  if (typeof model.style === "number") {
    return model.style;
  }
  const raw = model.alternateContentStyle;
  if (raw === undefined) {
    return undefined;
  }
  const fallback = /<c:style\s+val="(\d+)"/.exec(raw);
  if (fallback) {
    return Number(fallback[1]);
  }
  const c14 = /<c14:style\s+val="(\d+)"/.exec(raw);
  return c14 ? Number(c14[1]) - 100 : undefined;
}

function backgroundFill(background: string): Rgba01 | undefined {
  const colour = cssColor(background);
  return colour.a <= 0 ? undefined : colour;
}

function areaFill(
  fill: ChartFill | undefined,
  style: number | undefined,
  themeColors: ChartThemeColors | undefined
): Rgba01 | undefined {
  if (fill?.noFill) {
    return undefined;
  }
  const explicit =
    resolveColor(fill?.solid, themeColors) ??
    // A gradient or pattern has no single colour; its first stop or its foreground stands
    // for it, as for every other filled element of the preview.
    resolveColor(fill?.gradient?.stops?.[0]?.color, themeColors) ??
    resolveColor(fill?.pattern?.foreground, themeColors);
  if (explicit) {
    return explicit;
  }
  const index = style ?? 2;
  // `dk1` for the dark styles, `bg1` / `lt1` (one slot under the chart colour map) otherwise.
  return resolveColor({ theme: index >= 41 && index <= 48 ? 0 : 1 }, themeColors);
}

function resolveColor(
  color: ChartColor | undefined,
  themeColors: ChartThemeColors | undefined
): Rgba01 | undefined {
  return resolveChartColorRgba(color, themeColors);
}
