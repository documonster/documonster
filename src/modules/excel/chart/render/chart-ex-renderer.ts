/**
 * ChartEx renderer — renders a ChartExModel to SVG / PNG / vector PDF.
 *
 * This file owns the "render to image" responsibility (SVG string
 * emission, PNG rasterisation, and vector PDF drawing). The companion
 * `chart-ex-serialize.ts` owns the orthogonal "emit cx: OOXML XML"
 * responsibility. The two files share no private helpers — only
 * neutral utilities from `chart-utils` — so neither imports the other.
 */

import type {
  ChartExAxis,
  ChartExDataEntry,
  ChartExModel,
  ChartExSeries
} from "@excel/chart/model/chart-ex-types";
import type { ChartTitle, ShapeProperties } from "@excel/chart/model/types";
import type {
  ChartPdfDrawingSurface,
  ChartPdfPathOp,
  ChartRenderOptions,
  RegionMapDataOptions,
  RegionMapMatchRule
} from "@excel/chart/render/chart-renderer";
import { renderSvgToPng } from "@excel/chart/render/chart-renderer";
import type { ResolvedRing, TopologyLike } from "@excel/chart/render/topojson";
import { resolveTopologyObject } from "@excel/chart/render/topojson";
import type { ChartRect, PdfColor } from "@excel/chart/shared/chart-utils";
import {
  COLORS,
  DEFAULT_HEIGHT,
  DEFAULT_WIDTH,
  clamp01,
  escapeXml,
  escapeXmlAttr,
  fmt,
  formatNumber,
  hexToPdfColor,
  insetRect,
  interpolateColor,
  polar,
  previewShapeFillColor,
  valueToY,
  withAlpha
} from "@excel/chart/shared/chart-utils";
import { getSpPrFill } from "@excel/chart/shared/shape-properties";
import { ChartOptionsError } from "@excel/errors";

/**
 * Local alias kept for file-level readability. ChartEx code has always
 * called these `SvgRect`; the underlying shape matches the shared
 * {@link ChartRect} exactly.
 */
type SvgRect = ChartRect;

/**
 * Return the maximum absolute value of `values`, floored at `min`.
 * Folds via a loop rather than `Math.max(min, ...values.map(...))`
 * — the spread form blows the JS call stack past ~100k elements.
 * Matches the defensive style used elsewhere in this file
 * (`valueRange`, `boxStats`, `buildBubbles`).
 */
function maxAbsValue(values: readonly number[], min: number): number {
  let max = min;
  for (const v of values) {
    const abs = Math.abs(v);
    if (Number.isFinite(abs) && abs > max) {
      max = abs;
    }
  }
  return max;
}

export function renderChartExSvg(model: ChartExModel, options: ChartRenderOptions = {}): string {
  const width = options.width ?? DEFAULT_WIDTH;
  const height = options.height ?? DEFAULT_HEIGHT;
  const title = options.title ?? chartTitleText(model.chartSpace.chart.title);
  // Count newlines in the title so `getPlotRect` can expand the top
  // margin for multi-paragraph titles; a 3-line title with the default
  // 52px top margin would otherwise overflow into the plot area.
  const titleLineCount = title ? title.split(/\r?\n/).length : 0;
  const plot = getPlotRect(width, height, !!title, titleLineCount);
  const backgroundColor = options.backgroundColor ?? "#fff";
  const series =
    model.chartSpace.chart.plotArea.plotAreaRegion?.series ??
    model.chartSpace.chart.plotArea.series ??
    [];
  const parts: string[] = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`
  );
  if (backgroundColor !== "transparent") {
    parts.push(`<rect width="100%" height="100%" fill="${escapeXmlAttr(backgroundColor)}"/>`);
  }
  if (title) {
    // Multi-paragraph titles arrive as a `\n`-joined string from
    // `chartTitleText`. SVG `<text>` normalises whitespace (including
    // newlines) to spaces, so we must emit each paragraph as its own
    // `<tspan>` with an explicit baseline offset to stack them
    // vertically. Single-line titles stay on the fast path and emit a
    // bare `<text>` node — matches the previous byte-for-byte output
    // for the common case.
    const lines = title.split(/\r?\n/);
    if (lines.length === 1) {
      parts.push(
        `<text x="${fmt(width / 2)}" y="26" text-anchor="middle" font-family="Arial" font-size="18" fill="#222">${escapeXml(title)}</text>`
      );
    } else {
      const lineHeightEm = 1.2;
      const tspans = lines
        .map(
          (line, i) =>
            `<tspan x="${fmt(width / 2)}"${i === 0 ? "" : ` dy="${lineHeightEm}em"`}>${escapeXml(line)}</tspan>`
        )
        .join("");
      parts.push(
        `<text x="${fmt(width / 2)}" y="26" text-anchor="middle" font-family="Arial" font-size="18" fill="#222">${tspans}</text>`
      );
    }
  }
  for (let i = 0; i < series.length; i++) {
    renderChartExSeriesSvg(parts, model, series[i], plot, i, options);
  }
  // Only draw a legend when the model carries one. Previously this
  // call was unconditional, producing a synthetic legend on charts that
  // explicitly hid the legend (e.g. builder used `showLegend: false`,
  // which stores `legend: undefined`). The classic renderer honours
  // the legend model — ChartEx now matches.
  if (model.chartSpace.chart.legend) {
    renderChartExLegend(
      parts,
      series,
      width,
      height,
      !!title,
      model.chartSpace.chart.legend.legendPos
    );
  }
  parts.push("</svg>");
  return parts.join("");
}

export async function renderChartExPng(
  model: ChartExModel,
  options: ChartRenderOptions = {}
): Promise<Uint8Array> {
  const width = options.width ?? DEFAULT_WIDTH;
  const height = options.height ?? DEFAULT_HEIGHT;
  return renderSvgToPng(renderChartExSvg(model, { ...options, width, height }), {
    width,
    height,
    scale: options.scale,
    dpi: options.dpi
  });
}

/**
 * Layout IDs whose PDF geometry is expressed in `drawRect` /
 * `drawPath` primitives rather than SVG-specific filters or raster
 * fallbacks. {@link drawChartExPdf} consults this set to decide
 * whether a vector path is available; any layout not listed falls
 * back to the raster pipeline in `chartToPdf`.
 *
 * As of the regionMap port this set covers every ChartEx layout the
 * library currently emits, so vector is the default and raster is
 * only reached when the caller passes `chartToPdf(chart, {
 * forceRaster: true })`.
 *
 * Exported so the `@pdf/excel-bridge` `chartToPdf` helper can make
 * the same decision without reimplementing it.
 */
export const VECTOR_PDF_CHART_EX_LAYOUT_IDS: readonly string[] = [
  "sunburst",
  "treemap",
  "funnel",
  "waterfall",
  "boxWhisker",
  // Histogram shares its `clusteredColumn` rendering path with plain
  // Pareto (minus the cumulative line) — the parser/builder normalises
  // both to `layoutId: "clusteredColumn"`, so this one entry covers
  // both types. The Pareto branch is detected at draw time via
  // `series.layoutPr?.paretoLine`.
  "clusteredColumn",
  // Standalone `paretoLine` layoutId — Excel emits this for the
  // cumulative-percent line component of a paired Pareto chart. The
  // SVG and PDF paths render it as a line-with-points curve rather
  // than falling through to the column default.
  "paretoLine",
  // regionMap now has a full vector path (topology polygons, centroid
  // preview, hex-tile fallback — all three modes ported from the SVG
  // emitter). Callers who need raster output can still opt in via
  // `chartToPdf(chart, { forceRaster: true })`.
  "regionMap"
];

/**
 * True when every series of a ChartEx model has a layoutId that
 * `drawChartExPdf` can render as vector PDF. Used by `chartToPdf` to
 * route the chart to the vector or raster path automatically.
 */
export function canRenderChartExAsVectorPdf(model: ChartExModel): boolean {
  const plotArea = model.chartSpace.chart.plotArea;
  const seriesList = plotArea.plotAreaRegion?.series ?? plotArea.series ?? [];
  if (seriesList.length === 0) {
    return false;
  }
  const supported = new Set(VECTOR_PDF_CHART_EX_LAYOUT_IDS);
  return seriesList.every(s => supported.has(s.layoutId));
}

/**
 * Draw a ChartEx chart as vector content onto a
 * {@link ChartPdfDrawingSurface}. All layout IDs listed in
 * {@link VECTOR_PDF_CHART_EX_LAYOUT_IDS} are supported — the
 * surrounding geometry lives in the "collect" functions shared with
 * the SVG renderer, so the SVG and PDF paths stay equivalent by
 * construction.
 *
 * For any layout ID not in the set the function throws; callers
 * (notably `@pdf/excel-bridge.chartToPdf`) must pre-filter via
 * {@link canRenderChartExAsVectorPdf}. Filtering ahead of time rather
 * than silently skipping preserves "fail loud on unsupported" — a
 * silent no-op would produce an empty page and hide the mistake.
 */
export function drawChartExPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  rect: { x: number; y: number; width: number; height: number },
  options: { title?: string; regionMap?: RegionMapDataOptions } = {}
): ChartPdfDrawingSurface {
  const plotArea = model.chartSpace.chart.plotArea;
  const seriesList = plotArea.plotAreaRegion?.series ?? plotArea.series ?? [];
  const titleText = options.title ?? chartTitleText(model.chartSpace.chart.title);
  // `titleHeight` must scale with the number of lines so multi-paragraph
  // titles don't overflow the plot area. Line-height of 19.2 matches
  // the 16pt font × 1.2 line-height applied below in the drawText loop.
  const titleLines = titleText ? titleText.split(/\r?\n/).length : 0;
  const titleHeight = titleText ? Math.max(28, 20 + titleLines * 19.2) : 0;

  // Internal drawing uses SVG coordinates (y=0 at top, increasing downward).
  // The `rect` parameter is in PDF coordinates (y=0 at page bottom).
  // We compute the plot in SVG-local space (origin 0,0 at top-left of chart)
  // then use a flipping surface to emit correct PDF coordinates.
  const plot: SvgRect = {
    x: 12,
    y: titleHeight + 12,
    width: Math.max(10, rect.width - 24),
    height: Math.max(10, rect.height - titleHeight - 24)
  };

  // Y-flipping surface: converts SVG-local y to PDF-page y.
  // Formula: pdfY = rect.y + rect.height - localY
  // For rects: PDF rect(x, pdfBottomY, w, h) where pdfBottomY = rect.y + rect.height - (localY + h)
  const flipY = (localY: number): number => rect.y + rect.height - localY;
  const flipped: ChartPdfDrawingSurface = {
    drawRect(o) {
      surface.drawRect({
        ...o,
        x: rect.x + o.x,
        y: flipY(o.y + o.height),
        width: o.width,
        height: o.height
      });
      return flipped;
    },
    drawLine(o) {
      surface.drawLine({
        ...o,
        x1: rect.x + o.x1,
        y1: flipY(o.y1),
        x2: rect.x + o.x2,
        y2: flipY(o.y2)
      });
      return flipped;
    },
    drawText(text, o) {
      surface.drawText(text, {
        ...o,
        x: rect.x + o.x,
        y: flipY(o.y)
      });
      return flipped;
    },
    drawCircle: surface.drawCircle
      ? o => {
          surface.drawCircle!({
            ...o,
            cx: rect.x + o.cx,
            cy: flipY(o.cy)
          });
          return flipped;
        }
      : undefined,
    drawPath: surface.drawPath
      ? (ops, pathOpts) => {
          const flippedOps: ChartPdfPathOp[] = ops.map(op => {
            if (op.op === "close") {
              return op;
            }
            if (op.op === "curve") {
              return {
                ...op,
                x1: rect.x + op.x1,
                y1: flipY(op.y1),
                x2: rect.x + op.x2,
                y2: flipY(op.y2),
                x3: rect.x + op.x3,
                y3: flipY(op.y3)
              };
            }
            return { ...op, x: rect.x + op.x, y: flipY(op.y) };
          });
          surface.drawPath!(flippedOps, pathOpts);
          return flipped;
        }
      : undefined
  } as ChartPdfDrawingSurface;

  // Page background: a light frame matching the SVG background rect.
  surface.drawRect({
    x: rect.x,
    y: rect.y,
    width: rect.width,
    height: rect.height,
    fill: { r: 1, g: 1, b: 1 },
    stroke: { r: 0.85, g: 0.85, b: 0.85 }
  });

  if (titleText) {
    // Multi-paragraph titles arrive as a `\n`-joined string from
    // `chartTitleText`. `drawText` is a single-line primitive on every
    // surface we support, so previously a two-paragraph title rendered
    // as one line containing a literal `\n`. Split explicitly and
    // stack paragraphs vertically with a `fontSize * 1.2` line-height
    // (same convention as the SVG path below).
    const fontSize = 16;
    const lineHeight = Math.round(fontSize * 1.2);
    const lines = titleText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      flipped.drawText(lines[i], {
        x: rect.width / 2,
        y: 20 + i * lineHeight,
        fontSize,
        anchor: "middle",
        color: hexToPdfColor("#222222")
      });
    }
  }

  for (const series of seriesList) {
    if (series.layoutId === "treemap") {
      drawTreemapPdf(flipped, model, series, plot);
    } else if (series.layoutId === "sunburst") {
      drawSunburstPdf(flipped, model, series, plot);
    } else if (series.layoutId === "funnel") {
      drawFunnelPdf(flipped, model, series, plot);
    } else if (series.layoutId === "waterfall") {
      drawWaterfallPdf(flipped, model, series, plot);
    } else if (series.layoutId === "boxWhisker") {
      drawBoxWhiskerPdf(flipped, model, series, plot);
    } else if (series.layoutId === "regionMap") {
      drawRegionMapPdf(flipped, model, series, plot, options.regionMap);
    } else if (series.layoutId === "clusteredColumn") {
      // Both `histogram` and `pareto` live under clusteredColumn after
      // builder normalisation; distinguishing them is a single runtime
      // flag (`layoutPr.paretoLine`).
      if (series.layoutPr?.paretoLine) {
        drawParetoPdf(flipped, model, series, plot, { drawColumns: true });
      } else {
        drawHistogramPdf(flipped, model, series, plot);
      }
    } else if (series.layoutId === "paretoLine") {
      // Standalone paretoLine (distinct from `clusteredColumn` with
      // `paretoLine` flag) — Excel emits this for the cumulative-
      // percent line overlay when the author builds a paired Pareto
      // chart via the UI. The bars come from a sibling `clusteredColumn`
      // series; this series draws ONLY the overlay curve. PDF dispatch
      // was missing this case and threw `layoutId 'paretoLine' is not
      // supported`; then when added, unconditionally redrew the columns
      // on top of the companion series' bars.
      drawParetoPdf(flipped, model, series, plot, { drawColumns: false });
    } else {
      throw new Error(
        `drawChartExPdf: layoutId '${series.layoutId}' is not supported by the vector path. ` +
          `Gate on canRenderChartExAsVectorPdf(model) before calling.`
      );
    }
  }
  return surface;
}

function drawTreemapPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const root = buildHierarchy(refs.hierarchy, categories, values);
  for (const cell of collectTreemapCells(root, plot)) {
    surface.drawRect({
      x: cell.rect.x,
      y: cell.rect.y,
      width: cell.rect.width,
      height: cell.rect.height,
      fill: hexToPdfColor(cell.color),
      stroke: { r: 1, g: 1, b: 1 }
    });
    if (cell.label) {
      surface.drawText(cell.label, {
        x: cell.rect.x + 4,
        y: cell.rect.y + 14,
        fontSize: 10,
        color: { r: 1, g: 1, b: 1 },
        anchor: "start"
      });
    }
  }
}

function drawSunburstPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const root = buildHierarchy(refs.hierarchy, categories, values);
  if (!surface.drawPath) {
    // Without drawPath we can only approximate slice fills with a
    // flood-filled wedge — not worth the complexity. Fall back to a
    // neutral rect so the chart placeholder is still visible rather
    // than silently empty, matching how classic charts degrade when
    // `drawPath` is missing.
    surface.drawRect({
      x: plot.x,
      y: plot.y,
      width: plot.width,
      height: plot.height,
      stroke: { r: 0.7, g: 0.7, b: 0.7 }
    });
    return;
  }
  for (const slice of collectSunburstSlices(root, plot)) {
    const ops = ringSliceToPdfPath(slice);
    if (ops.length === 0) {
      continue;
    }
    surface.drawPath(ops, {
      fill: hexToPdfColor(slice.color),
      stroke: { r: 1, g: 1, b: 1 }
    });
  }
}

/**
 * Shared PDF axis renderer for the ChartEx layouts that borrow
 * classic-chart axis furniture — histogram, pareto, waterfall,
 * boxWhisker. Mirrors the SVG {@link renderAxes} helper: four evenly
 * spaced horizontal gridlines, a darker baseline + left edge, and
 * numeric min/max labels on the y axis. Keeping the two functions
 * visually aligned means the PDF matches the SVG preview pixel-for-
 * pixel modulo rasterisation.
 */
function drawAxesPdf(
  surface: ChartPdfDrawingSurface,
  plot: SvgRect,
  range: { min: number; max: number }
): void {
  const gridColor = hexToPdfColor("#D9D9D9");
  const frameColor = hexToPdfColor("#444444");
  const textColor = hexToPdfColor("#555555");
  // Four interior gridlines plus the top and bottom baselines give six
  // data-aligned labels (min, 20%, 40%, 60%, 80%, max). Emitting them
  // alongside the gridlines — same loop, same `i` — keeps the labels
  // mathematically in sync with the lines so a user always sees a
  // tick number directly to the left of each grid mark.
  for (let i = 1; i < 5; i++) {
    const t = i / 5;
    const y = plot.y + plot.height * t;
    surface.drawLine({
      x1: plot.x,
      y1: y,
      x2: plot.x + plot.width,
      y2: y,
      color: gridColor
    });
    // `t` counts from the top; the numeric value is therefore the
    // interpolation between `max` (at t = 0, top) and `min` (at
    // t = 1, bottom). Mirrors the SVG emitter in `renderAxes`.
    const value = range.max + (range.min - range.max) * t;
    surface.drawText(formatNumber(value), {
      x: plot.x - 8,
      y: y + 3,
      fontSize: 10,
      color: textColor,
      anchor: "end"
    });
  }
  surface.drawLine({
    x1: plot.x,
    y1: plot.y + plot.height,
    x2: plot.x + plot.width,
    y2: plot.y + plot.height,
    color: frameColor
  });
  surface.drawLine({
    x1: plot.x,
    y1: plot.y,
    x2: plot.x,
    y2: plot.y + plot.height,
    color: frameColor
  });
  surface.drawText(formatNumber(range.max), {
    x: plot.x - 8,
    y: plot.y + 3,
    fontSize: 10,
    color: textColor,
    anchor: "end"
  });
  surface.drawText(formatNumber(range.min), {
    x: plot.x - 8,
    y: plot.y + plot.height + 3,
    fontSize: 10,
    color: textColor,
    anchor: "end"
  });
}

/**
 * Shared PDF column renderer. Draws the same `rect + category label`
 * pair the SVG {@link renderColumnSvg} helper emits, plus the axes
 * via {@link drawAxesPdf}. Used directly by the histogram PDF path
 * and reused by waterfall / pareto (which compute their own column
 * geometry but share the axis framing).
 */
function drawColumnsPdf(
  surface: ChartPdfDrawingSurface,
  values: number[],
  categories: string[],
  plot: SvgRect,
  color: string,
  axis?: ChartExAxis
): void {
  const range = valueRange(values, axis);
  drawAxesPdf(surface, plot, range);
  const count = Math.max(1, values.length);
  const groupWidth = plot.width / count;
  const zero = valueToY(0, range.min, range.max, plot);
  const fill = hexToPdfColor(color);
  const labelColor = hexToPdfColor("#555555");
  values.forEach((value, i) => {
    const y = valueToY(value, range.min, range.max, plot);
    surface.drawRect({
      x: plot.x + i * groupWidth + groupWidth * 0.18,
      y: Math.min(y, zero),
      width: groupWidth * 0.64,
      height: Math.abs(zero - y),
      fill
    });
    surface.drawText(categories[i] ?? String(i + 1), {
      x: plot.x + i * groupWidth + groupWidth / 2,
      y: plot.y + plot.height + 18,
      fontSize: 10,
      color: labelColor,
      anchor: "middle"
    });
  });
}

function drawHistogramPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const bins = buildHistogramBins(refs.values, series.layoutPr?.binning);
  drawColumnsPdf(
    surface,
    bins.map(bin => bin.count),
    bins.map(bin => bin.label),
    plot,
    COLORS[0],
    findValueAxis(model)
  );
}

function drawParetoPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect,
  options: { drawColumns?: boolean } = {}
): void {
  const { drawColumns = true } = options;
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  // Filter non-finite values before sorting. `Array.prototype.sort`
  // is undefined on `NaN` comparator outputs (V8 TimSort keeps them
  // in their discovery position, WebKit may shuffle them arbitrarily),
  // so a blank / `#N/A` source cell produces bars at a random X
  // position and desynces the cumulative curve from the column heights.
  // Pareto convention also drops missing rows entirely.
  const sorted = values
    .map((value, i) => ({ value, category: categories[i] ?? String(i + 1) }))
    .filter(item => Number.isFinite(item.value))
    .sort((a, b) => b.value - a.value);
  const sortedValues = sorted.map(item => item.value);
  if (drawColumns) {
    drawColumnsPdf(
      surface,
      sortedValues,
      sorted.map(item => item.category),
      plot,
      COLORS[0],
      findValueAxis(model)
    );
  }

  // Cumulative polyline. Rendered as connected `drawLine` segments so
  // the path stays visible on surfaces without `drawPath`, matching
  // the SVG polyline behaviour. `drawCircle` is used for the dots
  // when available; otherwise small rects serve as markers.
  //
  // Non-finite values (NaN from blank / `#N/A` source cells) would
  // otherwise poison `Math.max(0, …)` and zero the visible `> 0` check
  // — mirrors the SVG pareto guard.
  const positiveSum = sortedValues.reduce(
    (sum, v) => sum + (Number.isFinite(v) ? Math.max(0, v) : 0),
    0
  );
  // When the dataset has no positive values (all-zero or all-negative
  // Pareto input), the cumulative-percent curve is undefined — every
  // point would collapse to the baseline and imply a flat 0 %
  // cumulative trace. Suppress the overlay entirely so the author
  // notices the input is out of range, rather than silently emitting
  // a flat line that reads as valid data.
  if (positiveSum > 0) {
    const total = positiveSum;
    let cumulative = 0;
    const count = Math.max(1, sortedValues.length);
    const step = plot.width / count;
    const points = sortedValues.map((value, i) => {
      cumulative += Number.isFinite(value) ? Math.max(0, value) : 0;
      return {
        x: plot.x + i * step + step / 2,
        y: plot.y + plot.height - (cumulative / total) * plot.height
      };
    });
    const lineColor = hexToPdfColor(COLORS[1]);
    for (let i = 1; i < points.length; i++) {
      surface.drawLine({
        x1: points[i - 1].x,
        y1: points[i - 1].y,
        x2: points[i].x,
        y2: points[i].y,
        color: lineColor,
        lineWidth: 2
      });
    }
    for (const p of points) {
      if (surface.drawCircle) {
        surface.drawCircle({ cx: p.x, cy: p.y, r: 3, fill: lineColor });
      } else {
        surface.drawRect({
          x: p.x - 3,
          y: p.y - 3,
          width: 6,
          height: 6,
          fill: lineColor
        });
      }
    }
    surface.drawText("Cumulative %", {
      x: plot.x + plot.width - 4,
      y: plot.y + 12,
      fontSize: 10,
      color: lineColor,
      anchor: "end"
    });
  }
}

function drawWaterfallPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const subtotalIdx = new Set(series.layoutPr?.subtotals?.map(s => s.idx) ?? []);
  let running = 0;
  const spans = values.map((value, i) => {
    // Subtotal convention: span `0 → running sum` (not the scalar
    // value stored at the row). See `renderWaterfallSvg` for the
    // matching rationale.
    if (subtotalIdx.has(i)) {
      const end = running;
      return { start: 0, end, value: end, total: true, gap: false };
    }
    // NaN guard mirrors the SVG path — `collectChartExNumbers` emits
    // `NaN` for sparse slots; adding that into `running` poisons every
    // subsequent bar. Emit a zero-height placeholder span and leave
    // `running` untouched. Flag as `gap` so the colour picker below
    // routes to a neutral grey rather than "increase" green
    // (`value: 0 >= 0` would otherwise paint the gap row the same as a
    // zero-increase row, making missing data indistinguishable from a
    // real zero delta).
    if (!Number.isFinite(value)) {
      return { start: running, end: running, value: 0, total: false, gap: true };
    }
    const start = running;
    const end = running + value;
    running = end;
    return { start, end, value, total: false, gap: false };
  });
  const range = valueRange(
    spans.flatMap(s => [s.start, s.end]),
    findValueAxis(model)
  );
  drawAxesPdf(surface, plot, range);
  const count = Math.max(1, spans.length);
  const groupWidth = plot.width / count;
  const labelColor = hexToPdfColor("#555555");
  const connectorColor = hexToPdfColor("#888888");
  const centers: Array<{ x: number; y: number }> = [];
  spans.forEach((span, i) => {
    const y1 = valueToY(span.start, range.min, range.max, plot);
    const y2 = valueToY(span.end, range.min, range.max, plot);
    const x = plot.x + i * groupWidth + groupWidth * 0.18;
    const colorHex = span.gap
      ? "#BFBFBF"
      : span.total
        ? shapeFillColor(series.layoutPr?.totalSpPr, COLORS[2])
        : span.value >= 0
          ? shapeFillColor(series.layoutPr?.increaseSpPr, "#70AD47")
          : shapeFillColor(series.layoutPr?.decreaseSpPr, "#C00000");
    surface.drawRect({
      x,
      y: Math.min(y1, y2),
      width: groupWidth * 0.64,
      // Gap rows should have zero visible height — `Math.max(1, 0)`
      // produced a 1-pixel green sliver that looked like a tiny
      // positive delta. Skip the `Math.max` clamp when the span is
      // intentionally flat.
      height: span.gap ? 0 : Math.max(1, Math.abs(y1 - y2)),
      fill: hexToPdfColor(colorHex)
    });
    surface.drawText(categories[i] ?? String(i + 1), {
      x: plot.x + i * groupWidth + groupWidth / 2,
      y: plot.y + plot.height + 18,
      fontSize: 10,
      color: labelColor,
      anchor: "middle"
    });
    centers.push({ x: x + groupWidth * 0.64, y: y2 });
  });
  if (series.layoutPr?.connectorLines !== false) {
    for (let i = 0; i < centers.length - 1; i++) {
      surface.drawLine({
        x1: centers[i].x,
        y1: centers[i].y,
        x2: centers[i + 1].x - groupWidth * 0.64,
        y2: centers[i].y,
        color: connectorColor,
        dashPattern: [3, 3]
      });
    }
  }
}

function drawFunnelPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  // Use the true maximum magnitude — flooring at 1 collapses charts
  // whose data is entirely sub-1 (conversion rates, probabilities,
  // proportions) into a tiny funnel because every stage width then
  // scales as `value / 1` rather than `value / max`. Only fall back
  // to `1` when there are no positive, finite values at all so
  // downstream arithmetic never divides by zero.
  const trueMax = maxAbsValue(values, 0);
  const max = trueMax > 0 ? trueMax : 1;
  const count = Math.max(1, values.length);
  const h = plot.height / count;
  const labelColor: PdfColor = { r: 1, g: 1, b: 1 };
  const whiteStroke: PdfColor = { r: 1, g: 1, b: 1 };
  // Pre-resolve per-point `dataPt/@idx` overrides into a map, mirroring
  // `renderFunnelSvg`. Previously the PDF path used the default
  // `COLORS` palette regardless of `dataPt` overrides, so a
  // round-tripped funnel with custom stage colours rendered correctly
  // in SVG but reverted to the default palette in PDF output — a
  // silent divergence between the two backends.
  const pointFills = new Map<number, string>();
  if (series.dataPt) {
    for (const dp of series.dataPt) {
      if (dp.spPr) {
        pointFills.set(dp.idx, shapeFillColor(dp.spPr, COLORS[dp.idx % COLORS.length]));
      }
    }
  }
  values.forEach((value, i) => {
    // Non-finite values must not propagate into the polygon vertices —
    // see `renderFunnelSvg` for the rationale.
    const absValue = Number.isFinite(value) ? Math.abs(value) : 0;
    const rawNext = values[i + 1];
    const nextAbs = Number.isFinite(rawNext) ? Math.abs(rawNext) : absValue;
    const topW = (absValue / max) * plot.width;
    const bottomW = (nextAbs / max) * plot.width;
    const y = plot.y + i * h;
    const cx = plot.x + plot.width / 2;
    const fill = hexToPdfColor(pointFills.get(i) ?? COLORS[i % COLORS.length]);
    // Trapezoid = quadrilateral polygon: `drawPath` with move + 3×line
    // + close. When the surface lacks drawPath we still keep a filled
    // rectangle (retains the colour signal) and overlay four stroke-
    // only lines that trace the trapezoid's real outline — so the
    // funnel silhouette remains recognisable even without vector
    // polygon filling.
    if (surface.drawPath) {
      surface.drawPath(
        [
          { op: "move", x: cx - topW / 2, y },
          { op: "line", x: cx + topW / 2, y },
          { op: "line", x: cx + bottomW / 2, y: y + h * 0.88 },
          { op: "line", x: cx - bottomW / 2, y: y + h * 0.88 },
          { op: "close" }
        ],
        { fill, stroke: whiteStroke }
      );
    } else {
      surface.drawRect({
        x: cx - Math.max(topW, bottomW) / 2,
        y,
        width: Math.max(topW, bottomW),
        height: h * 0.88,
        fill
      });
      // Four white strokes reproducing the trapezoid outline. Top and
      // bottom are horizontal; the two sides angle inwards when
      // bottomW < topW (narrowing funnel) or outward for a growing
      // layer. The rect behind still carries the colour; these lines
      // carve the trapezoid silhouette on top.
      const yBottom = y + h * 0.88;
      surface.drawLine({
        x1: cx - topW / 2,
        y1: y,
        x2: cx + topW / 2,
        y2: y,
        color: whiteStroke
      });
      surface.drawLine({
        x1: cx + topW / 2,
        y1: y,
        x2: cx + bottomW / 2,
        y2: yBottom,
        color: whiteStroke
      });
      surface.drawLine({
        x1: cx + bottomW / 2,
        y1: yBottom,
        x2: cx - bottomW / 2,
        y2: yBottom,
        color: whiteStroke
      });
      surface.drawLine({
        x1: cx - bottomW / 2,
        y1: yBottom,
        x2: cx - topW / 2,
        y2: y,
        color: whiteStroke
      });
    }
    surface.drawText(categories[i] ?? String(i + 1), {
      x: cx,
      y: y + h * 0.55,
      fontSize: 11,
      color: labelColor,
      anchor: "middle"
    });
  });
}

function drawBoxWhiskerPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const groups =
    categories.length > 0
      ? groupValuesByCategory(values, categories)
      : new Map([["Values", values]]);
  const allValues = Array.from(groups.values()).flat();
  const range = valueRange(allValues, findValueAxis(model));
  drawAxesPdf(surface, plot, range);
  const keys = Array.from(groups.keys());
  const groupWidth = plot.width / Math.max(1, keys.length);
  const whiskerColor = hexToPdfColor("#555555");
  const medianColor = hexToPdfColor("#333333");
  const outlierStroke = hexToPdfColor("#333333");
  const meanColor = hexToPdfColor("#333333");
  const labelColor = hexToPdfColor("#555555");
  keys.forEach((key, i) => {
    const stats = boxStats(groups.get(key) ?? [], series.layoutPr?.quartileMethod ?? "exclusive");
    const cx = plot.x + i * groupWidth + groupWidth / 2;
    const w = groupWidth * 0.38;
    const q1 = valueToY(stats.q1, range.min, range.max, plot);
    const q3 = valueToY(stats.q3, range.min, range.max, plot);
    const med = valueToY(stats.median, range.min, range.max, plot);
    const low = valueToY(stats.low, range.min, range.max, plot);
    const high = valueToY(stats.high, range.min, range.max, plot);
    const seriesColorHex = COLORS[i % COLORS.length];
    const seriesColor = hexToPdfColor(seriesColorHex);
    // Whisker (vertical line spanning low..high).
    surface.drawLine({ x1: cx, y1: high, x2: cx, y2: low, color: whiskerColor });
    // IQR box with 0.35 alpha fill, matching the SVG `withAlpha`.
    surface.drawRect({
      x: cx - w / 2,
      y: Math.min(q1, q3),
      width: w,
      height: Math.abs(q3 - q1),
      fill: { ...seriesColor, a: 0.35 },
      stroke: seriesColor
    });
    // Median line.
    surface.drawLine({
      x1: cx - w / 2,
      y1: med,
      x2: cx + w / 2,
      y2: med,
      color: medianColor
    });
    // Mean marker (circle).
    if (series.layoutPr?.showMeanMarker !== false) {
      const meanY = valueToY(stats.mean, range.min, range.max, plot);
      if (surface.drawCircle) {
        surface.drawCircle({ cx, cy: meanY, r: 3, fill: meanColor });
      } else {
        surface.drawRect({
          x: cx - 3,
          y: meanY - 3,
          width: 6,
          height: 6,
          fill: meanColor
        });
      }
    }
    // Mean line — dashed horizontal inside the IQR box.
    if (series.layoutPr?.showMeanLine) {
      const meanY = valueToY(stats.mean, range.min, range.max, plot);
      surface.drawLine({
        x1: cx - w / 2,
        y1: meanY,
        x2: cx + w / 2,
        y2: meanY,
        color: meanColor,
        dashPattern: [3, 2]
      });
    }
    // Inner points — small translucent dots, one per non-outlier
    // sample. Matches the SVG path; iterate `stats.nonOutliers` (not
    // the full raw group) so outlier values don't get double-plotted
    // as both an inner-point dot AND a hollow outlier ring when both
    // flags are enabled.
    if (series.layoutPr?.showInnerPoints) {
      const innerFill = { ...seriesColor, a: 0.55 };
      for (const value of stats.nonOutliers) {
        const y = valueToY(value, range.min, range.max, plot);
        if (surface.drawCircle) {
          surface.drawCircle({ cx: cx - w * 0.62, cy: y, r: 1.6, fill: innerFill });
        } else {
          surface.drawRect({
            x: cx - w * 0.62 - 1.6,
            y: y - 1.6,
            width: 3.2,
            height: 3.2,
            fill: innerFill
          });
        }
      }
    }
    // Outliers — hollow circles outside the IQR whiskers.
    if (series.layoutPr?.showOutlierPoints !== false) {
      for (const outlier of stats.outliers) {
        const y = valueToY(outlier, range.min, range.max, plot);
        if (surface.drawCircle) {
          surface.drawCircle({ cx: cx + w * 0.62, cy: y, r: 2, stroke: outlierStroke });
        } else {
          // No drawCircle → approximate the hollow ring with a stroke-
          // only rect. Matches SVG's `fill="none" stroke="#333"` circle
          // more faithfully than an opaque rect would; keeps the bbox
          // size identical so overlays with inner points don't shift.
          surface.drawRect({
            x: cx + w * 0.62 - 2,
            y: y - 2,
            width: 4,
            height: 4,
            stroke: outlierStroke
          });
        }
      }
    }
    surface.drawText(key, {
      x: cx,
      y: plot.y + plot.height + 18,
      fontSize: 10,
      color: labelColor,
      anchor: "middle"
    });
  });
}

/**
 * Vector PDF path for ChartEx `regionMap`. Mirrors `renderRegionMapSvg`
 * with the same three-mode dispatch — TopoJSON first, centroid
 * preview second, hex-tile fallback last — so the decision logic
 * stays in a single place instead of drifting between backends. All
 * the math (projections, extent normalisation, centroid lookup,
 * label matching) is imported verbatim from the SVG helpers; only
 * the "emit" step switches from string concatenation to `drawRect` /
 * `drawPath` / `drawCircle` / `drawLine` / `drawText` calls.
 *
 * Inside the function the three branches are:
 *
 *   1. Topology supplied and at least one feature matches → draw
 *      filled polygons per feature via `drawPath` (or a `drawLine`
 *      outline on surfaces without `drawPath`), plus matched-feature
 *      labels at each ring centroid.
 *   2. Known centroids cover every category → draw a muted frame
 *      rectangle + graticule + scaled circle per region + optional
 *      text labels; missing centroids spill into the hex-tile grid.
 *   3. No centroids hit → the whole plot becomes the hex-tile grid.
 */
function drawRegionMapPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect,
  mapOptions?: RegionMapDataOptions
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const range = valueRange(values, findValueAxis(model));

  // 1. TopoJSON branch.
  if (mapOptions?.topology) {
    const drawn = tryDrawRegionMapWithTopologyPdf(
      surface,
      values,
      categories,
      series,
      plot,
      range,
      mapOptions
    );
    if (drawn) {
      return;
    }
    // Fall through to the centroid preview.
  }

  // 2. Centroid preview.
  const records = values.map((value, i) => ({
    value,
    label: categories[i] ?? String(i + 1),
    coord: lookupRegionCoordinate(categories[i] ?? String(i + 1))
  }));
  const known = records.filter(
    (record): record is { value: number; label: string; coord: RegionCoordinate } => !!record.coord
  );
  if (known.length === 0) {
    // 3. No centroid hits: hex tiles over the full plot.
    drawRegionMapTileFallbackPdf(surface, records, range, plot);
    return;
  }

  // Frame + graticule.
  surface.drawRect({
    x: plot.x,
    y: plot.y,
    width: plot.width,
    height: plot.height,
    fill: hexToPdfColor("#F7FBFF"),
    stroke: hexToPdfColor("#C7DFF2")
  });
  drawRegionMapGraticulePdf(surface, plot);

  // Scaled circle per region.
  const labelMode = series.layoutPr?.regionLabels ?? "bestFit";
  const whiteStroke: PdfColor = { r: 1, g: 1, b: 1 };
  const regionLabelColor = hexToPdfColor("#1F3B53");
  for (const record of known) {
    const projected = projectRegionCoordinate(
      record.coord,
      series.layoutPr?.projection ?? "miller",
      plot
    );
    // Skip non-finite values and clamp `t` to `[0, 1]` — same guard
    // as the SVG path; without it `radius` / `interpolateColor`
    // receive `NaN` or out-of-range input and draw invisible or
    // garbage-coloured dots. Matches the TOPO path's clamp.
    if (!Number.isFinite(record.value)) {
      continue;
    }
    const rawT = (record.value - range.min) / Math.max(1e-9, range.max - range.min);
    const t = Math.max(0, Math.min(1, rawT));
    const radius = 6 + Math.sqrt(t) * 14;
    const fillColor: PdfColor = {
      ...hexToPdfColor(interpolateColor("#D9EAF7", "#2F75B5", t)),
      // Matches the SVG `opacity="0.92"`; `/ExtGState` will materialise
      // it on capable surfaces, others render opaque which is
      // acceptable for a preview-grade dot.
      a: 0.92
    };
    if (surface.drawCircle) {
      surface.drawCircle({
        cx: projected.x,
        cy: projected.y,
        r: radius,
        fill: fillColor,
        stroke: whiteStroke
      });
    } else {
      surface.drawRect({
        x: projected.x - radius,
        y: projected.y - radius,
        width: radius * 2,
        height: radius * 2,
        fill: fillColor,
        stroke: whiteStroke
      });
    }
    if (labelMode === "showAll" || (labelMode === "bestFit" && radius >= 9)) {
      surface.drawText(record.label, {
        x: projected.x,
        y: projected.y + 3,
        fontSize: 9,
        color: regionLabelColor,
        anchor: "middle"
      });
    }
  }

  // Unknown regions drift into a hex-tile strip at the bottom.
  const unknown = records.filter(record => !record.coord);
  if (unknown.length > 0) {
    const fallbackHeight = Math.min(52, plot.height * 0.25);
    drawRegionMapTileFallbackPdf(surface, unknown, range, {
      x: plot.x + 8,
      y: plot.y + plot.height - fallbackHeight - 8,
      width: plot.width - 16,
      height: fallbackHeight
    });
  }
}

/**
 * PDF equivalent of `tryRenderRegionMapWithTopology`. Returns `true`
 * when at least one feature matched and was drawn; otherwise the
 * caller falls back to the centroid preview. Reuses every piece of
 * business logic from the SVG version (match rule ordering,
 * projection, extent computation, label reverse lookup) so the
 * decision boundaries stay identical across backends.
 */
function tryDrawRegionMapWithTopologyPdf(
  surface: ChartPdfDrawingSurface,
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect,
  range: { min: number; max: number },
  mapOptions: RegionMapDataOptions
): boolean {
  let features: ReturnType<typeof resolveTopologyObject>;
  try {
    features = resolveTopologyObject(mapOptions.topology as TopologyLike, mapOptions.objectName);
  } catch {
    return false;
  }
  if (features.length === 0) {
    return false;
  }

  const matchRules: RegionMapMatchRule[] = (() => {
    const raw = mapOptions.match ?? "id";
    return Array.isArray(raw) ? (raw.length > 0 ? raw : ["id"]) : [raw];
  })();
  const candidateKeys = (f: (typeof features)[number]): string[] => {
    const keys: string[] = [];
    for (const rule of matchRules) {
      const raw = rule.startsWith("property:")
        ? (f.properties?.[rule.slice(9)] as string | number | undefined)
        : (f.id as string | number | undefined);
      if (raw !== undefined && raw !== null) {
        keys.push(String(raw).trim().toLowerCase());
      }
    }
    return keys;
  };

  const valueByLabel = new Map<string, number>();
  categories.forEach((label, i) => {
    const norm = normaliseLabel(label);
    if (norm) {
      valueByLabel.set(norm, values[i]);
    }
  });
  if (valueByLabel.size === 0) {
    return false;
  }

  const projection = mapOptions.projection ?? series.layoutPr?.projection ?? "miller";
  const extent = computeProjectionExtent(features, projection);
  if (!extent) {
    return false;
  }

  // Pre-pass: resolve matches and count hits BEFORE drawing anything.
  // PDF surface calls can't be rolled back, so if we committed the
  // frame + outlines to the surface and then returned `false`, the
  // caller would layer its centroid fallback on top of our partial
  // topo-map — the SVG path had the same bug and buffers fragments
  // instead. Here we just do the decision first and only touch the
  // surface on the success path.
  const resolvedMatch = new Map<
    (typeof features)[number],
    { key: string; value: number } | undefined
  >();
  let matchedCount = 0;
  for (const feature of features) {
    const keys = candidateKeys(feature);
    let hit: { key: string; value: number } | undefined;
    for (const key of keys) {
      const value = valueByLabel.get(key);
      if (value !== undefined) {
        hit = { key, value };
        matchedCount += 1;
        break;
      }
    }
    resolvedMatch.set(feature, hit);
  }
  if (matchedCount === 0) {
    return false;
  }

  // Frame. PDF surface `drawRect` does not expose `borderRadius` via
  // the chart surface interface, so the SVG's `rx="14"` rounded
  // corners become sharp corners here — the only intentional visual
  // divergence between the two backends for regionMap.
  surface.drawRect({
    x: plot.x,
    y: plot.y,
    width: plot.width,
    height: plot.height,
    fill: hexToPdfColor("#F7FBFF"),
    stroke: hexToPdfColor("#C7DFF2")
  });

  const strokeHex = mapOptions.strokeColor ?? "#FFFFFF";
  const strokeColor = hexToPdfColor(strokeHex);
  for (const feature of features) {
    const match = resolvedMatch.get(feature);
    const fillHex = match
      ? (() => {
          const t = (match.value - range.min) / Math.max(1e-9, range.max - range.min);
          return interpolateColor("#D9EAF7", "#2F75B5", Math.max(0, Math.min(1, t)));
        })()
      : "#E9EEF3";
    const fill = hexToPdfColor(fillHex);
    const ops = featureToPdfPathOps(feature.rings, projection, plot, extent);
    if (ops.length === 0) {
      continue;
    }
    if (surface.drawPath) {
      surface.drawPath(ops, { fill, stroke: strokeColor });
    } else {
      // Minimal-surface fallback: outline the feature by walking the
      // move/line ops as a series of drawLine calls. Fill is lost
      // (no way to flood-fill a polygon without drawPath) but the
      // silhouette remains.
      drawPathOpsAsLines(surface, ops, strokeColor);
    }
  }

  // Label reverse lookup (original category casing).
  const labelByKey = new Map<string, string>();
  categories.forEach(label => {
    const norm = normaliseLabel(label);
    if (norm) {
      labelByKey.set(norm, label);
    }
  });

  const labelMode = series.layoutPr?.regionLabels ?? "bestFit";
  const topoLabelColor = hexToPdfColor("#1F3B53");
  if (labelMode === "showAll" || labelMode === "bestFit") {
    for (const feature of features) {
      const match = resolvedMatch.get(feature);
      if (!match || feature.rings.length === 0) {
        continue;
      }
      const centroidLonLat = ringCentroid(feature.rings[0]);
      if (!centroidLonLat) {
        continue;
      }
      const projected = projectLonLatToPlot(
        centroidLonLat[0],
        centroidLonLat[1],
        projection,
        plot,
        extent
      );
      const originalLabel =
        labelByKey.get(match.key) ??
        (typeof feature.id === "string" ? feature.id : String(feature.id ?? ""));
      surface.drawText(originalLabel, {
        x: projected.x,
        y: projected.y + 3,
        fontSize: 9,
        color: topoLabelColor,
        anchor: "middle"
      });
    }
  }

  return true;
}

/**
 * Convert a feature's resolved lon/lat rings into
 * `ChartPdfPathOp[]`. Each ring becomes `move` + `line×n` + `close`,
 * which is the exact one-to-one mapping of the SVG `M x y L x y … Z`
 * form `featureToSvgPath` emits — no curve approximation, because
 * TopoJSON arcs are already polylines once `resolveTopologyObject`
 * has dequantised them.
 */
function featureToPdfPathOps(
  rings: ResolvedRing[],
  projection: NonNullable<RegionMapDataOptions["projection"]>,
  plot: SvgRect,
  extent: { minX: number; maxX: number; minY: number; maxY: number }
): ChartPdfPathOp[] {
  const ops: ChartPdfPathOp[] = [];
  for (const ring of rings) {
    if (ring.length < 2) {
      continue;
    }
    const pts = ring.map(([lon, lat]) => projectLonLatToPlot(lon, lat, projection, plot, extent));
    ops.push({ op: "move", x: pts[0].x, y: pts[0].y });
    for (let i = 1; i < pts.length; i++) {
      ops.push({ op: "line", x: pts[i].x, y: pts[i].y });
    }
    ops.push({ op: "close" });
  }
  return ops;
}

/**
 * Last-resort degradation for surfaces without `drawPath`: trace a
 * polygon's move/line/close ops as individual `drawLine` segments.
 * Produces a stroke-only outline (fill is unreachable without a
 * path-filling primitive). Only reachable on custom surfaces; both
 * `PdfPageBuilder` and `PdfEditorPage` implement `drawPath`.
 */
function drawPathOpsAsLines(
  surface: ChartPdfDrawingSurface,
  ops: ChartPdfPathOp[],
  color: PdfColor
): void {
  let subpathStart: { x: number; y: number } | undefined;
  let prev: { x: number; y: number } | undefined;
  for (const op of ops) {
    if (op.op === "move") {
      subpathStart = { x: op.x, y: op.y };
      prev = subpathStart;
    } else if (op.op === "line" && prev) {
      surface.drawLine({ x1: prev.x, y1: prev.y, x2: op.x, y2: op.y, color });
      prev = { x: op.x, y: op.y };
    } else if (op.op === "close" && prev && subpathStart) {
      surface.drawLine({
        x1: prev.x,
        y1: prev.y,
        x2: subpathStart.x,
        y2: subpathStart.y,
        color
      });
      prev = subpathStart;
    }
  }
}

/** PDF equivalent of `renderRegionMapGraticule` — three vertical and two horizontal reference lines. */
function drawRegionMapGraticulePdf(surface: ChartPdfDrawingSurface, plot: SvgRect): void {
  const color = hexToPdfColor("#E5F0FA");
  for (let i = 1; i < 4; i++) {
    const x = plot.x + (plot.width * i) / 4;
    surface.drawLine({ x1: x, y1: plot.y + 8, x2: x, y2: plot.y + plot.height - 8, color });
  }
  for (let i = 1; i < 3; i++) {
    const y = plot.y + (plot.height * i) / 3;
    surface.drawLine({ x1: plot.x + 8, y1: y, x2: plot.x + plot.width - 8, y2: y, color });
  }
}

/**
 * Hex-tile fallback (PDF). Used when the category list contains no
 * regions the centroid table recognises, or as a bottom strip next
 * to the centroid preview for unmatched regions. Mirrors
 * `renderRegionMapTileFallback` exactly — a grid of hexagons whose
 * fill tracks the value scale, with optional labels when the cell
 * is large enough.
 */
function drawRegionMapTileFallbackPdf(
  surface: ChartPdfDrawingSurface,
  records: Array<{ value: number; label: string }>,
  range: { min: number; max: number },
  plot: SvgRect
): void {
  const count = Math.max(1, records.length);
  const cols = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / cols);
  const cellW = plot.width / cols;
  const cellH = plot.height / rows;
  const radius = Math.min(cellW, cellH) * 0.38;
  const whiteStroke: PdfColor = { r: 1, g: 1, b: 1 };
  const labelColor = hexToPdfColor("#1F3B53");
  records.forEach((record, i) => {
    if (!Number.isFinite(record.value)) {
      return;
    }
    const cx = plot.x + (i % cols) * cellW + cellW / 2;
    const cy = plot.y + Math.floor(i / cols) * cellH + cellH / 2;
    const t = clamp01((record.value - range.min) / Math.max(1e-9, range.max - range.min));
    const fill = hexToPdfColor(interpolateColor("#D9EAF7", "#2F75B5", t));
    const hexOps = hexagonPathOps(cx, cy, radius);
    if (surface.drawPath) {
      surface.drawPath(hexOps, { fill, stroke: whiteStroke });
    } else {
      // Approximate the hexagon with a centred rect + outline — the
      // colour remains, the exact hex shape becomes a square but the
      // grid layout is preserved.
      surface.drawRect({
        x: cx - radius,
        y: cy - radius,
        width: radius * 2,
        height: radius * 2,
        fill,
        stroke: whiteStroke
      });
    }
    if (cellW > 34 && cellH > 22) {
      surface.drawText(record.label, {
        x: cx,
        y: cy + 3,
        fontSize: 9,
        color: labelColor,
        anchor: "middle"
      });
    }
  });
}

function hexagonPathOps(cx: number, cy: number, radius: number): ChartPdfPathOp[] {
  const ops: ChartPdfPathOp[] = [];
  for (let i = 0; i < 6; i++) {
    const angle = Math.PI / 6 + (Math.PI * 2 * i) / 6;
    const x = cx + Math.cos(angle) * radius;
    const y = cy + Math.sin(angle) * radius;
    ops.push(i === 0 ? { op: "move", x, y } : { op: "line", x, y });
  }
  ops.push({ op: "close" });
  return ops;
}

/**
 * Convert a sunburst ring slice into `ChartPdfPathOp[]`, approximating
 * each arc with a cubic Bézier curve. PDF's path grammar has no
 * native arc primitive, but a single ≤ 90° arc can be drawn to
 * within ~0.027 % max error using the standard "kappa" control
 * coefficient `(4/3) * tan((end-start)/4)`; longer arcs are split
 * into quadrants.
 *
 * The resulting path walks outer-arc forwards, then inner-arc
 * backwards, closing the ring slice — the same topology
 * `renderRingSlice` produces for SVG, just expressed as Beziers.
 */
function ringSliceToPdfPath(slice: SunburstSlice): ChartPdfPathOp[] {
  const ops: ChartPdfPathOp[] = [];
  const outerStart = polar(slice.cx, slice.cy, slice.outer, slice.start);
  ops.push({ op: "move", x: outerStart.x, y: outerStart.y });
  appendArcAsBeziers(ops, slice.cx, slice.cy, slice.outer, slice.start, slice.end);
  const innerEnd = polar(slice.cx, slice.cy, slice.inner, slice.end);
  ops.push({ op: "line", x: innerEnd.x, y: innerEnd.y });
  appendArcAsBeziers(ops, slice.cx, slice.cy, slice.inner, slice.end, slice.start);
  ops.push({ op: "close" });
  return ops;
}

function appendArcAsBeziers(
  ops: ChartPdfPathOp[],
  cx: number,
  cy: number,
  radius: number,
  start: number,
  end: number
): void {
  if (radius <= 0) {
    return;
  }
  const totalSweep = end - start;
  if (totalSweep === 0) {
    return;
  }
  // Split the sweep so each sub-arc is ≤ 90° (π/2 radians). Keeping
  // sub-arcs small is what bounds the cubic Bézier approximation
  // error below ~0.03 %; larger sweeps visibly distort at the
  // midpoint.
  const steps = Math.max(1, Math.ceil(Math.abs(totalSweep) / (Math.PI / 2)));
  const stepSweep = totalSweep / steps;
  let theta = start;
  for (let i = 0; i < steps; i++) {
    const next = theta + stepSweep;
    const kappa = (4 / 3) * Math.tan(stepSweep / 4);
    const p0 = polar(cx, cy, radius, theta);
    const p1 = {
      x: p0.x - kappa * radius * Math.sin(theta),
      y: p0.y + kappa * radius * Math.cos(theta)
    };
    const p3 = polar(cx, cy, radius, next);
    const p2 = {
      x: p3.x + kappa * radius * Math.sin(next),
      y: p3.y - kappa * radius * Math.cos(next)
    };
    ops.push({ op: "curve", x1: p1.x, y1: p1.y, x2: p2.x, y2: p2.y, x3: p3.x, y3: p3.y });
    theta = next;
  }
}

interface HierarchyNode {
  name: string;
  value: number;
  children: HierarchyNode[];
}

function renderChartExSeriesSvg(
  parts: string[],
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect,
  seriesIndex: number,
  renderOptions: ChartRenderOptions = {}
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  // Locate the value axis once per series so `valScaling.min/max` can be
  // honoured. Non-valued layouts (treemap / sunburst / funnel / regionMap)
  // ignore the axis; columnar layouts pass it into `valueRange` below.
  const valueAxis = findValueAxis(model);
  switch (series.layoutId) {
    case "funnel":
      renderFunnelSvg(parts, values, categories, series, plot);
      return;
    case "waterfall":
      renderWaterfallSvg(parts, values, categories, series, plot, valueAxis);
      return;
    case "clusteredColumn":
      if (series.layoutPr?.paretoLine) {
        renderParetoSvg(parts, values, categories, series, plot, valueAxis, {
          drawColumns: true
        });
      } else {
        renderHistogramSvg(parts, values, series, plot, valueAxis);
      }
      return;
    case "paretoLine":
      // `paretoLine` is a valid layoutId distinct from `clusteredColumn`
      // + `layoutPr.paretoLine`. Excel stores the paired Pareto chart
      // as two sibling series — a `clusteredColumn` for the bars and a
      // `paretoLine` for the cumulative curve — so the standalone
      // variant must emit ONLY the overlay line (the columns come from
      // the companion series, if any). Previously this case fell
      // through to `renderParetoSvg` which unconditionally redrew the
      // columns in sorted order on top of the companion series' bars.
      renderParetoSvg(parts, values, categories, series, plot, valueAxis, {
        drawColumns: false
      });
      return;
    case "boxWhisker":
      renderBoxWhiskerSvg(parts, values, categories, series, plot, valueAxis);
      return;
    case "treemap":
      renderTreemapSvg(parts, buildHierarchy(refs.hierarchy, categories, values), plot);
      return;
    case "sunburst":
      renderSunburstSvg(parts, buildHierarchy(refs.hierarchy, categories, values), plot);
      return;
    case "regionMap":
      renderRegionMapSvg(
        parts,
        values,
        categories,
        series,
        plot,
        renderOptions.regionMap,
        valueAxis
      );
      return;
    default:
      renderColumnSvg(
        parts,
        values,
        categories,
        plot,
        COLORS[seriesIndex % COLORS.length],
        valueAxis
      );
  }
}

function resolveChartExRefs(
  model: ChartExModel,
  series: ChartExSeries
): { values: number[]; categories: string[]; hierarchy: string[][] } {
  const dataById = new Map(model.chartSpace.chartData.data.map(entry => [entry.id, entry]));
  const entries = (series.dataRefs ?? [])
    .map(ref => (ref.dataId === undefined ? undefined : dataById.get(ref.dataId)))
    .filter((entry): entry is ChartExDataEntry => !!entry);

  // Per ECMA-376 Chart2014 `ST_DimType`, the dimension's `@type`
  // attribute carries its semantic role:
  //   - `"val"` / `"y"` / `"size"` → primary numeric axis
  //   - `"cat"` → categorical (string) axis
  //   - `"x"` → value-axis category (histogram / pareto bin)
  //   - `"from"` / `"to"` → waterfall transitions
  //   - `"classification"` → hierarchical classifier
  //
  // Previously the renderer picked "the first strDim / numDim in
  // declaration order" which happened to work for simple sunburst /
  // treemap layouts but silently mis-routed dimensions whenever
  // Excel authored the data in a different order (common in
  // waterfall and pareto files, where the `type="x"` numDim often
  // precedes the primary `type="val"` one). Honour the dimension
  // type so the numeric payload lands on `values` regardless of
  // declaration order, and keep a "declaration order" fallback for
  // ChartEx files whose dims carry the permissive `"val"` default.
  const pickNumDim = (): number[] => {
    const valDim = entries.find(
      entry => entry.numDim && (entry.numDim.type === "val" || entry.numDim.type === "y")
    );
    if (valDim) {
      return collectChartExNumbers(valDim);
    }
    // No `val`/`y` — fall back to `size` (bubble / ChartEx size
    // dimension) and then to the first numDim in order.
    const sizeDim = entries.find(entry => entry.numDim?.type === "size");
    if (sizeDim) {
      return collectChartExNumbers(sizeDim);
    }
    const first = entries.find(entry => entry.numDim);
    return first ? collectChartExNumbers(first) : [];
  };
  const pickPrimaryCategoryDim = (): string[] => {
    const catDim = entries.find(entry => entry.strDim?.type === "cat");
    if (catDim) {
      return collectChartExStrings(catDim);
    }
    const first = entries.find(entry => entry.strDim);
    return first ? collectChartExStrings(first) : [];
  };
  // Hierarchy = every `strDim` that isn't the primary category axis.
  // Preserves their declaration order so sunburst/treemap rings still
  // stack outward in the author's intent.
  const primaryCatEntry =
    entries.find(entry => entry.strDim?.type === "cat") ?? entries.find(entry => entry.strDim);
  const hierarchy = entries
    .filter(entry => entry.strDim && entry !== primaryCatEntry)
    .map(entry => collectChartExStrings(entry));

  return {
    values: pickNumDim(),
    categories: pickPrimaryCategoryDim(),
    hierarchy
  };
}

function collectChartExStrings(entry: ChartExDataEntry): string[] {
  const levels = entry.strDim?.levels ?? [];
  const first = levels[0];
  if (!first) {
    return [];
  }
  // Hard ceiling on densification — prevents malicious / malformed XML
  // from allocating a gigabyte-scale array via a bogus `ptCount`.
  // Matches the classic `collectNumberValues` guard (`SPARSE_ARRAY_CEILING`).
  const SPARSE_ARRAY_CEILING = 2_097_152;
  const declared = first.ptCount ?? first.points.length;
  const count = Math.min(Math.max(first.points.length, declared), SPARSE_ARRAY_CEILING);
  const values = Array.from({ length: count }, (_, i) => String(i + 1));
  for (const point of first.points) {
    if (point.index >= 0 && point.index < count) {
      values[point.index] = point.value;
    }
  }
  return values;
}

function collectChartExNumbers(entry: ChartExDataEntry): number[] {
  const levels = entry.numDim?.levels ?? [];
  const first = levels[0];
  if (!first) {
    return [];
  }
  const SPARSE_ARRAY_CEILING = 2_097_152;
  const declared = first.ptCount ?? first.points.length;
  const count = Math.min(Math.max(first.points.length, declared), SPARSE_ARRAY_CEILING);
  // Mirror the classic `collectNumberValues` semantics: sparse slots
  // are gaps, not zeros. Excel omits `<cx:pt>` entries for blank /
  // `#N/A` source cells; the classic renderer encodes those as `NaN`
  // so `valueToY` / bar builders skip them. Previously the ChartEx
  // path filled gaps with `0`, producing phantom zero-value entries
  // that poisoned the data range (waterfall spanning to 0, histogram
  // bars showing at "empty" categories) and diverged from the classic
  // rendering of identical data.
  const values = Array.from({ length: count }, () => NaN);
  for (const point of first.points) {
    if (point.index >= 0 && point.index < count) {
      values[point.index] = point.value;
    }
  }
  return values;
}

function renderColumnSvg(
  parts: string[],
  values: number[],
  categories: string[],
  plot: SvgRect,
  color: string,
  axis?: ChartExAxis
): void {
  const range = valueRange(values, axis);
  renderAxes(parts, plot, range);
  const count = Math.max(1, values.length);
  const groupWidth = plot.width / count;
  const zero = valueToY(0, range.min, range.max, plot);
  values.forEach((value, i) => {
    const y = valueToY(value, range.min, range.max, plot);
    parts.push(
      `<rect x="${fmt(plot.x + i * groupWidth + groupWidth * 0.18)}" y="${fmt(Math.min(y, zero))}" width="${fmt(groupWidth * 0.64)}" height="${fmt(Math.abs(zero - y))}" fill="${color}"/>`
    );
    parts.push(
      svgText(
        plot.x + i * groupWidth + groupWidth / 2,
        plot.y + plot.height + 18,
        categories[i] ?? String(i + 1),
        10,
        "#555",
        "middle"
      )
    );
  });
}

function renderHistogramSvg(
  parts: string[],
  values: number[],
  series: ChartExSeries,
  plot: SvgRect,
  axis?: ChartExAxis
): void {
  const bins = buildHistogramBins(values, series.layoutPr?.binning);
  renderColumnSvg(
    parts,
    bins.map(bin => bin.count),
    bins.map(bin => bin.label),
    plot,
    COLORS[0],
    axis
  );
}

function renderParetoSvg(
  parts: string[],
  values: number[],
  categories: string[],
  _series: ChartExSeries,
  plot: SvgRect,
  axis?: ChartExAxis,
  options: { drawColumns?: boolean } = {}
): void {
  const { drawColumns = true } = options;
  // Filter non-finite values before sorting — see `drawParetoPdf` for
  // the full explanation. `NaN` comparator outputs make the sort order
  // implementation-defined, which desyncs bars from the cumulative curve.
  const sorted = values
    .map((value, i) => ({ value, category: categories[i] ?? String(i + 1) }))
    .filter(item => Number.isFinite(item.value))
    .sort((a, b) => b.value - a.value);
  const sortedValues = sorted.map(item => item.value);
  if (drawColumns) {
    renderColumnSvg(
      parts,
      sortedValues,
      sorted.map(item => item.category),
      plot,
      COLORS[0],
      axis
    );
  }
  // Suppress the cumulative overlay when the dataset has no positive
  // contribution. A `|| 1` fallback previously produced a flat line
  // at the baseline for all-zero / all-negative data, visually
  // indistinguishable from legitimate 0 % cumulative. See the PDF
  // path (`drawParetoPdf`) for the matching rationale.
  //
  // `sortedValues` may carry `NaN` for blank / `#N/A` source cells
  // (`collectChartExNumbers` preserves slot identity). `Math.max(0, NaN)`
  // is `NaN`, and `NaN > 0` is `false`, so a single missing value
  // previously suppressed the entire cumulative line. Coerce non-finite
  // values to zero contribution so the line tracks the real positive
  // data.
  const positiveSum = sortedValues.reduce(
    (sum, v) => sum + (Number.isFinite(v) ? Math.max(0, v) : 0),
    0
  );
  if (positiveSum > 0) {
    const total = positiveSum;
    let cumulative = 0;
    const count = Math.max(1, sortedValues.length);
    const step = plot.width / count;
    const points = sortedValues.map(value => {
      cumulative += Number.isFinite(value) ? Math.max(0, value) : 0;
      return {
        y: plot.y + plot.height - (cumulative / total) * plot.height
      };
    });
    // Re-derive X positions so the type stays simple (no array-index
    // closure capture) — keeps the polyline emission below straight.
    const plotted = points.map((p, i) => ({ x: plot.x + i * step + step / 2, y: p.y }));
    parts.push(
      `<polyline points="${plotted.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="none" stroke="${COLORS[1]}" stroke-width="2"/>`
    );
    for (const p of plotted) {
      parts.push(`<circle cx="${fmt(p.x)}" cy="${fmt(p.y)}" r="3" fill="${COLORS[1]}"/>`);
    }
    // Emit the "Cumulative %" caption whenever the cumulative line is
    // drawn, matching the PDF path. The previous
    // `series.layoutPr?.paretoLine` guard only fired for the paired
    // `clusteredColumn + paretoLine` variant; the standalone
    // `layoutId: "paretoLine"` case (added at line 1527) reached this
    // function without the flag set and silently dropped the caption
    // — producing SVG output that disagreed with the PDF.
    // `_series` is kept as a parameter in case future heuristics want
    // to tweak placement, but the caption is no longer gated by the
    // layoutPr flag.
    parts.push(svgText(plot.x + plot.width - 4, plot.y + 12, "Cumulative %", 10, COLORS[1], "end"));
  }
}

function renderWaterfallSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect,
  axis?: ChartExAxis
): void {
  const subtotalIdx = new Set(series.layoutPr?.subtotals?.map(s => s.idx) ?? []);
  let running = 0;
  const spans = values.map((value, i) => {
    // Excel's waterfall convention: a subtotal column spans `0 → running`
    // — it visualises the cumulative sum up to that point, not the
    // scalar value stored at the row. Author convention leaves the
    // subtotal row's numeric value at `0` (the subtotal is derived),
    // so the old `end = value` read `0` and the subtotal bar
    // collapsed to zero height. Worse, `running = end` then reset
    // the running sum, corrupting every subsequent bar.
    if (subtotalIdx.has(i)) {
      const end = running;
      // Keep `running` unchanged — the next bar should start from the
      // same cumulative sum the subtotal displays.
      return { start: 0, end, value: end, total: true, gap: false };
    }
    // `collectChartExNumbers` emits `NaN` for sparse `<cx:pt>` slots
    // (blanks or `#N/A` source cells). Adding NaN into `running`
    // permanently poisons it, collapsing every subsequent bar's height
    // to zero (via `fmt(NaN) → "0"`). Treat a blank slot as a
    // zero-height span at the current running total and flag it as a
    // `gap` so the colour picker routes to neutral grey rather than
    // "increase" green (`value: 0 >= 0` would otherwise paint the gap
    // row identically to a zero-increase row — visually
    // indistinguishable from real data). Leave `running` advancing as
    // if the missing value were `0`; this matches Excel's own behaviour
    // for blank waterfall rows.
    if (!Number.isFinite(value)) {
      return { start: running, end: running, value: 0, total: false, gap: true };
    }
    const start = running;
    const end = running + value;
    running = end;
    return { start, end, value, total: false, gap: false };
  });
  const range = valueRange(
    spans.flatMap(s => [s.start, s.end]),
    axis
  );
  renderAxes(parts, plot, range);
  const count = Math.max(1, spans.length);
  const groupWidth = plot.width / count;
  const centers: Array<{ x: number; y: number }> = [];
  spans.forEach((span, i) => {
    const y1 = valueToY(span.start, range.min, range.max, plot);
    const y2 = valueToY(span.end, range.min, range.max, plot);
    const x = plot.x + i * groupWidth + groupWidth * 0.18;
    const color = span.gap
      ? "#BFBFBF"
      : span.total
        ? shapeFillColor(series.layoutPr?.totalSpPr, COLORS[2])
        : span.value >= 0
          ? shapeFillColor(series.layoutPr?.increaseSpPr, "#70AD47")
          : shapeFillColor(series.layoutPr?.decreaseSpPr, "#C00000");
    // Gap rows should render invisibly. `Math.max(1, …)` floored
    // zero-height spans to a 1-pixel sliver that reads as a tiny
    // positive delta; skip the floor when the span is flagged flat.
    const height = span.gap ? 0 : Math.max(1, Math.abs(y1 - y2));
    parts.push(
      `<rect x="${fmt(x)}" y="${fmt(Math.min(y1, y2))}" width="${fmt(groupWidth * 0.64)}" height="${fmt(height)}" fill="${color}"/>`
    );
    parts.push(
      svgText(
        plot.x + i * groupWidth + groupWidth / 2,
        plot.y + plot.height + 18,
        categories[i] ?? String(i + 1),
        10,
        "#555",
        "middle"
      )
    );
    centers.push({ x: x + groupWidth * 0.64, y: y2 });
  });
  if (series.layoutPr?.connectorLines !== false) {
    for (let i = 0; i < centers.length - 1; i++) {
      parts.push(
        `<line x1="${fmt(centers[i].x)}" y1="${fmt(centers[i].y)}" x2="${fmt(centers[i + 1].x - groupWidth * 0.64)}" y2="${fmt(centers[i].y)}" stroke="#888" stroke-dasharray="3 3"/>`
      );
    }
  }
}

function renderFunnelSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect
): void {
  // See `drawFunnelPdf` for why we use the true max (not floored at
  // 1): fractional-magnitude data would otherwise render as a tiny
  // off-centre funnel.
  const trueMax = maxAbsValue(values, 0);
  const max = trueMax > 0 ? trueMax : 1;
  const count = Math.max(1, values.length);
  const h = plot.height / count;
  // Pre-resolve per-point `dataPt/@idx` overrides into a map so the
  // hot loop stays O(n). Previously funnel charts ignored
  // `<cx:dataPt>` entirely and forced the preview palette, so an
  // Excel-authored funnel with individually-coloured stages rendered
  // with the wrong colours after a round-trip — even though the
  // authored XML was preserved byte-for-byte on write.
  const pointFills = new Map<number, string>();
  if (series.dataPt) {
    for (const dp of series.dataPt) {
      if (dp.spPr) {
        pointFills.set(dp.idx, shapeFillColor(dp.spPr, COLORS[dp.idx % COLORS.length]));
      }
    }
  }
  values.forEach((value, i) => {
    // Non-finite values (blank / `#N/A` source cells emit NaN via
    // `collectChartExNumbers`) must not propagate into the polygon
    // vertices — `Math.abs(NaN) = NaN`, and `fmt(NaN)` returns `"0"`,
    // collapsing the stage's edges to the SVG page origin at (0, 0).
    // Coerce the width contribution to zero so a gap stage renders
    // as a degenerate zero-width wedge rather than a triangle pointing
    // at the page corner. Same for `values[i+1]`, where `??` would
    // only have coalesced `null`/`undefined`, not NaN.
    const absValue = Number.isFinite(value) ? Math.abs(value) : 0;
    const rawNext = values[i + 1];
    const nextAbs = Number.isFinite(rawNext) ? Math.abs(rawNext) : absValue;
    const topW = (absValue / max) * plot.width;
    const bottomW = (nextAbs / max) * plot.width;
    const y = plot.y + i * h;
    const cx = plot.x + plot.width / 2;
    const fill = pointFills.get(i) ?? COLORS[i % COLORS.length];
    parts.push(
      `<polygon points="${fmt(cx - topW / 2)},${fmt(y)} ${fmt(cx + topW / 2)},${fmt(y)} ${fmt(cx + bottomW / 2)},${fmt(y + h * 0.88)} ${fmt(cx - bottomW / 2)},${fmt(y + h * 0.88)}" fill="${fill}" stroke="#fff"/>`
    );
    parts.push(svgText(cx, y + h * 0.55, categories[i] ?? String(i + 1), 11, "#fff", "middle"));
  });
}

function renderBoxWhiskerSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect,
  axis?: ChartExAxis
): void {
  const groups =
    categories.length > 0
      ? groupValuesByCategory(values, categories)
      : new Map([["Values", values]]);
  const allValues = Array.from(groups.values()).flat();
  const range = valueRange(allValues, axis);
  renderAxes(parts, plot, range);
  const keys = Array.from(groups.keys());
  const groupWidth = plot.width / Math.max(1, keys.length);
  keys.forEach((key, i) => {
    const stats = boxStats(groups.get(key) ?? [], series.layoutPr?.quartileMethod ?? "exclusive");
    const cx = plot.x + i * groupWidth + groupWidth / 2;
    const w = groupWidth * 0.38;
    const q1 = valueToY(stats.q1, range.min, range.max, plot);
    const q3 = valueToY(stats.q3, range.min, range.max, plot);
    const med = valueToY(stats.median, range.min, range.max, plot);
    const low = valueToY(stats.low, range.min, range.max, plot);
    const high = valueToY(stats.high, range.min, range.max, plot);
    parts.push(
      `<line x1="${fmt(cx)}" y1="${fmt(high)}" x2="${fmt(cx)}" y2="${fmt(low)}" stroke="#555"/>`
    );
    parts.push(
      `<rect x="${fmt(cx - w / 2)}" y="${fmt(Math.min(q1, q3))}" width="${fmt(w)}" height="${fmt(Math.abs(q3 - q1))}" fill="${withAlpha(COLORS[i % COLORS.length], 0.35)}" stroke="${COLORS[i % COLORS.length]}"/>`
    );
    parts.push(
      `<line x1="${fmt(cx - w / 2)}" y1="${fmt(med)}" x2="${fmt(cx + w / 2)}" y2="${fmt(med)}" stroke="#333"/>`
    );
    if (series.layoutPr?.showMeanMarker !== false) {
      parts.push(
        `<circle cx="${fmt(cx)}" cy="${fmt(valueToY(stats.mean, range.min, range.max, plot))}" r="3" fill="#333"/>`
      );
    }
    if (series.layoutPr?.showMeanLine) {
      const mean = valueToY(stats.mean, range.min, range.max, plot);
      parts.push(
        `<line x1="${fmt(cx - w / 2)}" y1="${fmt(mean)}" x2="${fmt(cx + w / 2)}" y2="${fmt(mean)}" stroke="#333" stroke-dasharray="3 2"/>`
      );
    }
    if (series.layoutPr?.showInnerPoints) {
      // Iterate `nonOutliers` (NOT the full raw group) so outlier
      // samples don't get double-painted as both a filled inner-point
      // dot AND a hollow outlier ring when both flags are enabled.
      // Matches Excel's semantics: "inner points" are individual
      // non-outlier observations.
      for (const value of stats.nonOutliers) {
        parts.push(
          `<circle cx="${fmt(cx - w * 0.62)}" cy="${fmt(valueToY(value, range.min, range.max, plot))}" r="1.6" fill="${COLORS[i % COLORS.length]}" opacity="0.55"/>`
        );
      }
    }
    if (series.layoutPr?.showOutlierPoints !== false) {
      for (const outlier of stats.outliers) {
        parts.push(
          `<circle cx="${fmt(cx + w * 0.62)}" cy="${fmt(valueToY(outlier, range.min, range.max, plot))}" r="2" fill="none" stroke="#333"/>`
        );
      }
    }
    parts.push(svgText(cx, plot.y + plot.height + 18, key, 10, "#555", "middle"));
  });
}

function renderTreemapSvg(parts: string[], root: HierarchyNode, plot: SvgRect): void {
  for (const cell of collectTreemapCells(root, plot)) {
    parts.push(
      `<rect x="${fmt(cell.rect.x)}" y="${fmt(cell.rect.y)}" width="${fmt(cell.rect.width)}" height="${fmt(cell.rect.height)}" fill="${cell.color}" stroke="#fff"/>`
    );
    if (cell.label) {
      parts.push(svgText(cell.rect.x + 4, cell.rect.y + 14, cell.label, 10, "#fff", "start"));
    }
  }
}

function renderSunburstSvg(parts: string[], root: HierarchyNode, plot: SvgRect): void {
  for (const slice of collectSunburstSlices(root, plot)) {
    parts.push(
      renderRingSlice(
        slice.cx,
        slice.cy,
        slice.outer,
        slice.inner,
        slice.start,
        slice.end,
        slice.color
      )
    );
  }
}

/**
 * Treemap geometry collector. Shared between the SVG emitter (which
 * prints `<rect>` elements) and the PDF emitter (which calls
 * `surface.drawRect` / `surface.drawText`). Returning plain data —
 * rect, colour, optional label — keeps both backends honest: any
 * future visual regression in the geometry shows up on the test that
 * exercises the primary (SVG) backend, not only on the secondary
 * (PDF) one.
 *
 * The label threshold (> 40 × 18 pixels) mirrors the historical SVG
 * code; both backends honour it identically so small treemap cells
 * degrade to colour-only the same way everywhere.
 */
export interface TreemapCell {
  rect: SvgRect;
  color: string;
  label: string | undefined;
}

export function collectTreemapCells(root: HierarchyNode, plot: SvgRect): TreemapCell[] {
  const nodes = root.children.length > 0 ? root.children : [{ ...root, name: "Values" }];
  const entries = sliceDice(nodes, plot, true);
  // Drop degenerate cells (zero-value nodes produce zero-width or
  // zero-height rects). Without the filter, those rects are still
  // emitted with `stroke="#fff"` — browsers render the stroke on a
  // collapsed rect as a visible 1-pixel line, producing parasitic
  // white seams between otherwise-adjacent coloured tiles. Color
  // palette indices stay aligned with the remaining nodes so
  // neighbouring tiles keep their authored colour mapping.
  return entries
    .filter(entry => entry.rect.width > 0.5 && entry.rect.height > 0.5)
    .map((entry, i) => ({
      rect: entry.rect,
      color: COLORS[i % COLORS.length],
      label: entry.rect.width > 40 && entry.rect.height > 18 ? entry.node.name : undefined
    }));
}

/**
 * Sunburst geometry collector. Recursively walks the hierarchy,
 * emitting one {@link SunburstSlice} per non-root node in the same
 * angular/ring order the SVG renderer has always used, so the two
 * backends stay pixel-equivalent modulo rasterisation.
 *
 * Colour handling follows the original `renderSunburstNode`
 * "increment on every visit" rule so sibling slices and their
 * descendants walk through `COLORS` in the same sequence regardless of
 * which backend is driving.
 */
export interface SunburstSlice {
  cx: number;
  cy: number;
  /** Outer ring radius, always >= `inner`. */
  outer: number;
  /** Inner ring radius; 0 for the innermost ring. */
  inner: number;
  /** Start angle in radians, 0 = +X axis. */
  start: number;
  /** End angle in radians. */
  end: number;
  color: string;
}

export function collectSunburstSlices(root: HierarchyNode, plot: SvgRect): SunburstSlice[] {
  const slices: SunburstSlice[] = [];
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  // `hierarchyDepth(root)` counts the invisible root, but the recursive
  // emitter skips `depth === 0` (the root doesn't draw). The number of
  // *visible* rings is therefore `depth - 1`. Dividing `radius` by
  // `depth` directly left the outermost `1 / depth` of the plot radius
  // blank (leaf slices stopped at `ring * (depth - 1)` instead of
  // reaching `radius`), wasting one full ring of visual space —
  // progressively worse for shallow hierarchies (a single-level tree
  // halved the rendered radius).
  const maxDepth = Math.max(1, hierarchyDepth(root) - 1);
  const radius = Math.min(plot.width, plot.height) / 2.25;
  // Seed `colorIndex = -1` so the root's "consumed" palette slot (which
  // the recursive emitter reserves via `colorIndex + 1` pre-increment)
  // lands on index `0`. The old `colorIndex = 0` made the root eat
  // `COLORS[0]`, then its first visible child drew at `COLORS[1]` —
  // every sunburst started with orange instead of the accent-1 blue
  // that every other ChartEx type uses.
  collectSunburstSlicesRecursive(slices, root, cx, cy, 0, Math.PI * 2, 0, maxDepth, radius, -1);
  return slices;
}

function collectSunburstSlicesRecursive(
  out: SunburstSlice[],
  node: HierarchyNode,
  cx: number,
  cy: number,
  startAngle: number,
  endAngle: number,
  depth: number,
  maxDepth: number,
  radius: number,
  colorIndex: number
): number {
  const ring = radius / Math.max(1, maxDepth);
  let nextColorIndex = colorIndex;
  if (depth > 0) {
    out.push({
      cx,
      cy,
      outer: ring * depth,
      inner: ring * (depth - 1),
      start: startAngle,
      end: endAngle,
      color: COLORS[colorIndex % COLORS.length]
    });
  }
  const total = node.children.reduce((sum, c) => sum + Math.max(0, c.value), 0) || 1;
  let angle = startAngle;
  for (let i = 0; i < node.children.length; i++) {
    const child = node.children[i];
    // Snap the last child to `endAngle` exactly. Summing N floating
    // fractions of a range almost never reproduces the range bound
    // (IEEE-754 drift on the order of 1e-13 per step), and on deep
    // hierarchies with many siblings that drift accumulates until the
    // outermost ring's sweep slips below the `2π - 1e-9` full-circle
    // guard downstream — the ring then renders as a degenerate
    // near-invisible arc instead of the closing slice.
    const next =
      i === node.children.length - 1
        ? endAngle
        : angle + (Math.max(0, child.value) / total) * (endAngle - startAngle);
    nextColorIndex = collectSunburstSlicesRecursive(
      out,
      child,
      cx,
      cy,
      angle,
      next,
      depth + 1,
      maxDepth,
      radius,
      nextColorIndex + 1
    );
    angle = next;
  }
  return nextColorIndex;
}

function renderRegionMapSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect,
  mapOptions?: RegionMapDataOptions,
  axis?: ChartExAxis
): void {
  const range = valueRange(values, axis);

  // Path A: caller supplied a TopoJSON dataset. Try to draw real
  // region polygons; fall through to the centroid preview if anything
  // goes wrong (invalid topology, zero matches, unsupported geometry).
  if (mapOptions?.topology) {
    const drawn = tryRenderRegionMapWithTopology(
      parts,
      values,
      categories,
      series,
      plot,
      range,
      mapOptions
    );
    if (drawn) {
      return;
    }
    // Fall through to the centroid preview — this gives the caller
    // something to look at even when their topology fails to match
    // any labels, rather than producing an empty chart.
  }

  const records = values.map((value, i) => ({
    value,
    label: categories[i] ?? String(i + 1),
    coord: lookupRegionCoordinate(categories[i] ?? String(i + 1))
  }));
  const known = records.filter(
    (record): record is { value: number; label: string; coord: RegionCoordinate } => !!record.coord
  );
  if (known.length === 0) {
    renderRegionMapTileFallback(parts, records, range, plot);
    return;
  }

  parts.push(
    `<rect x="${fmt(plot.x)}" y="${fmt(plot.y)}" width="${fmt(plot.width)}" height="${fmt(plot.height)}" rx="14" fill="#F7FBFF" stroke="#C7DFF2" data-region-map-mode="geographic-preview"/>`
  );
  renderRegionMapGraticule(parts, plot);
  const labelMode = series.layoutPr?.regionLabels ?? "bestFit";
  for (const record of known) {
    const projected = projectRegionCoordinate(
      record.coord,
      series.layoutPr?.projection ?? "miller",
      plot
    );
    // Skip non-finite values — `collectChartExNumbers` preserves `NaN`
    // for blank / `#N/A` cells, and `(NaN - range.min) / …` propagates
    // `NaN` into both `radius` (a `NaN` passed to `fmt` emits `"0"`,
    // drawing an invisible 0-radius dot) and `interpolateColor`
    // (produces a garbage colour off the clamp path). Also clamp `t`
    // to `[0, 1]` so out-of-range values still receive a defined
    // colour — mirrors the TOPO branch's clamp at `tryRender…WithTopology`.
    if (!Number.isFinite(record.value)) {
      continue;
    }
    const rawT = (record.value - range.min) / Math.max(1e-9, range.max - range.min);
    const t = Math.max(0, Math.min(1, rawT));
    const radius = 6 + Math.sqrt(t) * 14;
    parts.push(
      `<circle cx="${fmt(projected.x)}" cy="${fmt(projected.y)}" r="${fmt(radius)}" fill="${interpolateColor("#D9EAF7", "#2F75B5", t)}" stroke="#fff" stroke-width="1.5" opacity="0.92"/>`
    );
    if (labelMode === "showAll" || (labelMode === "bestFit" && radius >= 9)) {
      parts.push(svgText(projected.x, projected.y + 3, record.label, 9, "#1F3B53", "middle"));
    }
  }

  const unknown = records.filter(record => !record.coord);
  if (unknown.length > 0) {
    const fallbackHeight = Math.min(52, plot.height * 0.25);
    renderRegionMapTileFallback(
      parts,
      unknown,
      range,
      {
        x: plot.x + 8,
        y: plot.y + plot.height - fallbackHeight - 8,
        width: plot.width - 16,
        height: fallbackHeight
      },
      "unmatched"
    );
  }
}

/**
 * Render the region map using a user-supplied TopoJSON topology.
 *
 * Returns `true` when at least one feature was drawn so the caller
 * skips the centroid-dot fallback. Returns `false` on any failure
 * (topology invalid, geometry not found, zero matches) so the caller
 * can degrade gracefully to the built-in preview.
 *
 * Matching policy: case-insensitive + trimmed comparison, consistent
 * with `normaliseLabel` used by the centroid table. Matches every
 * feature once — unmatched features are drawn as a neutral-fill
 * outline so the world map provides context even for the regions
 * the author didn't supply data for.
 */
function tryRenderRegionMapWithTopology(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect,
  range: { min: number; max: number },
  mapOptions: RegionMapDataOptions
): boolean {
  let features: ReturnType<typeof resolveTopologyObject>;
  try {
    features = resolveTopologyObject(mapOptions.topology as TopologyLike, mapOptions.objectName);
  } catch {
    return false;
  }
  if (features.length === 0) {
    return false;
  }

  // Normalise the `match` option into an ordered rule list. A single
  // rule stays functionally identical to the pre-matchers behaviour;
  // an array lets callers express locale-aware fall-backs such as
  // `["property:name_zh", "property:name", "id"]` without writing
  // custom code per workbook.
  const matchRules: RegionMapMatchRule[] = (() => {
    const raw = mapOptions.match ?? "id";
    return Array.isArray(raw) ? (raw.length > 0 ? raw : ["id"]) : [raw];
  })();
  const candidateKeys = (f: (typeof features)[number]): string[] => {
    const keys: string[] = [];
    for (const rule of matchRules) {
      const raw = rule.startsWith("property:")
        ? (f.properties?.[rule.slice(9)] as string | number | undefined)
        : (f.id as string | number | undefined);
      if (raw !== undefined && raw !== null) {
        keys.push(String(raw).trim().toLowerCase());
      }
    }
    return keys;
  };

  // Build value map: label → value.
  const valueByLabel = new Map<string, number>();
  categories.forEach((label, i) => {
    const norm = normaliseLabel(label);
    if (norm) {
      valueByLabel.set(norm, values[i]);
    }
  });
  if (valueByLabel.size === 0) {
    return false;
  }

  // Determine data extents in the projected unit square so we can
  // scale to the plot rectangle. Compute once over all feature
  // coordinates the user supplied — this keeps the map centred even
  // when the TopoJSON covers a bounding box larger than the matched
  // countries.
  const projection = mapOptions.projection ?? series.layoutPr?.projection ?? "miller";
  const extent = computeProjectionExtent(features, projection);
  if (!extent) {
    return false;
  }

  // Buffer the SVG fragments into a local array instead of pushing
  // directly to `parts`. If we ultimately return `false` (no feature
  // matched a category), the caller falls through to the centroid
  // preview — we MUST NOT leave a half-drawn world outline underneath
  // that preview, or the composite image shows both layers. Flush the
  // buffer only once we've decided to claim the chart area.
  const buffer: string[] = [];
  buffer.push(
    `<rect x="${fmt(plot.x)}" y="${fmt(plot.y)}" width="${fmt(plot.width)}" height="${fmt(plot.height)}" rx="14" fill="#F7FBFF" stroke="#C7DFF2" data-region-map-mode="topojson"/>`
  );

  // Resolve each feature against the rule list once and cache both the
  // winning key and its value, so the fill loop and the label loop
  // below don't re-scan the rules (which would be O(rules × features × 2)
  // for world-atlas scale topologies).
  const resolvedMatch = new Map<
    (typeof features)[number],
    { key: string; value: number } | undefined
  >();
  for (const feature of features) {
    const keys = candidateKeys(feature);
    let hit: { key: string; value: number } | undefined;
    for (const key of keys) {
      const value = valueByLabel.get(key);
      if (value !== undefined) {
        hit = { key, value };
        break;
      }
    }
    resolvedMatch.set(feature, hit);
  }

  // Compute per-feature fill: matched features get the value-scaled
  // colour, unmatched get a neutral base so the world still appears.
  const stroke = mapOptions.strokeColor ?? "#FFFFFF";
  let matchedCount = 0;
  for (const feature of features) {
    const match = resolvedMatch.get(feature);
    const fill = match
      ? (() => {
          matchedCount++;
          const t = (match.value - range.min) / Math.max(1e-9, range.max - range.min);
          return interpolateColor("#D9EAF7", "#2F75B5", Math.max(0, Math.min(1, t)));
        })()
      : "#E9EEF3";
    const path = featureToSvgPath(feature.rings, projection, plot, extent);
    if (path) {
      buffer.push(
        `<path d="${path}" fill="${fill}" stroke="${escapeXmlAttr(stroke)}" stroke-width="0.5" stroke-linejoin="round"/>`
      );
    }
  }

  // No feature matched any author-supplied category. Abandon the
  // buffered world outline — the caller will draw the centroid
  // preview on a clean plot area instead of layering it over our
  // partial output.
  if (matchedCount === 0) {
    return false;
  }

  // Build a reverse lookup (normalised category → original label) once
  // so label emission can find the author-supplied spelling regardless
  // of which matcher rule produced the hit.
  const labelByKey = new Map<string, string>();
  categories.forEach(label => {
    const norm = normaliseLabel(label);
    if (norm) {
      labelByKey.set(norm, label);
    }
  });

  // Labels for matched features — place at the first ring's centroid.
  const labelMode = series.layoutPr?.regionLabels ?? "bestFit";
  if (labelMode === "showAll" || labelMode === "bestFit") {
    for (const feature of features) {
      const match = resolvedMatch.get(feature);
      if (!match || feature.rings.length === 0) {
        continue;
      }
      const centroidLonLat = ringCentroid(feature.rings[0]);
      if (!centroidLonLat) {
        continue;
      }
      const projected = projectLonLatToPlot(
        centroidLonLat[0],
        centroidLonLat[1],
        projection,
        plot,
        extent
      );
      const originalLabel =
        labelByKey.get(match.key) ??
        (typeof feature.id === "string" ? feature.id : String(feature.id ?? ""));
      buffer.push(svgText(projected.x, projected.y + 3, originalLabel, 9, "#1F3B53", "middle"));
    }
  }

  // Commit the buffered fragments now that we know at least one
  // feature matched — the caller will see a complete topo-map
  // rendering and skip the centroid preview.
  for (const frag of buffer) {
    parts.push(frag);
  }
  return true;
}

/**
 * Unit-square → plot-rect projection shared by the topology
 * renderer. Uses the same projection table as `projectRegionCoordinate`
 * but feeds the result through the per-dataset extent so the bundle
 * of features is centred in the plot regardless of the topology's
 * coordinate extent.
 */
function projectLonLatToPlot(
  lon: number,
  lat: number,
  projection: NonNullable<RegionMapDataOptions["projection"]>,
  plot: SvgRect,
  extent: { minX: number; maxX: number; minY: number; maxY: number }
): { x: number; y: number } {
  const raw = projectLonLatRaw(lon, lat, projection);
  const nx =
    extent.maxX === extent.minX ? 0.5 : (raw.x - extent.minX) / (extent.maxX - extent.minX);
  const ny =
    extent.maxY === extent.minY ? 0.5 : (raw.y - extent.minY) / (extent.maxY - extent.minY);
  return {
    x: plot.x + 14 + nx * Math.max(1, plot.width - 28),
    y: plot.y + 14 + ny * Math.max(1, plot.height - 28)
  };
}

/**
 * Project a single (lon, lat) pair into the projection's raw
 * coordinate space (not yet plot-normalised). The four projections
 * mirror `projectRegionCoordinate`'s formulas — duplicated here
 * because the existing function already applies plot padding, which
 * we need to defer until after the per-dataset extent is known.
 */
function projectLonLatRaw(
  lon: number,
  lat: number,
  projection: NonNullable<RegionMapDataOptions["projection"]>
): { x: number; y: number } {
  const clampedLon = Math.max(-180, Math.min(180, lon));
  // The ±85° clamp only applies to projections with a Mercator-style
  // singularity at the poles (`log(tan(π/4 ± π/2)) → ∞`). Albers,
  // Robinson, and plain equirectangular are all finite at ±90°, so
  // clamping their input destroys 5° of polar data for no reason —
  // Antarctica and high-latitude research stations silently flatten
  // onto the `±85°` parallel when rendered under those projections.
  const clampedLat =
    projection === "mercator" || projection === "miller"
      ? Math.max(-MERCATOR_LAT_CLAMP_DEG, Math.min(MERCATOR_LAT_CLAMP_DEG, lat))
      : Math.max(-90, Math.min(90, lat));
  switch (projection) {
    case "mercator": {
      const rad = (clampedLat * Math.PI) / 180;
      return {
        x: clampedLon / 360,
        y: -Math.log(Math.tan(Math.PI / 4 + rad / 2)) / (2 * Math.PI)
      };
    }
    case "miller": {
      const rad = (clampedLat * Math.PI) / 180;
      return {
        x: clampedLon / 360,
        y: -(1.25 * Math.log(Math.tan(Math.PI / 4 + 0.4 * rad))) / (2 * Math.PI)
      };
    }
    case "albers": {
      const { rawX, rawY } = rawAlbers(clampedLon, clampedLat);
      return { x: rawX, y: rawY };
    }
    case "robinson": {
      const { nx, ny } = projectRobinson(clampedLon, clampedLat);
      // projectRobinson returns 0..1; subtract 0.5 to centre.
      return { x: nx - 0.5, y: ny - 0.5 };
    }
    default: {
      return { x: clampedLon / 360, y: -clampedLat / 180 };
    }
  }
}

/**
 * Derive the min/max projected extent across every ring in the
 * provided feature list. Used to normalise the topology's coordinate
 * space into the plot rectangle without distortion.
 */
function computeProjectionExtent(
  features: ReturnType<typeof resolveTopologyObject>,
  projection: NonNullable<RegionMapDataOptions["projection"]>
): { minX: number; maxX: number; minY: number; maxY: number } | undefined {
  let minX = Infinity;
  let maxX = -Infinity;
  let minY = Infinity;
  let maxY = -Infinity;
  for (const feature of features) {
    for (const ring of feature.rings) {
      for (const [lon, lat] of ring) {
        const { x, y } = projectLonLatRaw(lon, lat, projection);
        if (x < minX) {
          minX = x;
        }
        if (x > maxX) {
          maxX = x;
        }
        if (y < minY) {
          minY = y;
        }
        if (y > maxY) {
          maxY = y;
        }
      }
    }
  }
  if (!Number.isFinite(minX) || !Number.isFinite(maxX)) {
    return undefined;
  }
  return { minX, maxX, minY, maxY };
}

/**
 * Convert a list of rings (outer + holes) into an SVG `path` `d`
 * attribute. Holes are represented by alternating fill-rule within the
 * same `<path>` — the default nonzero rule handles inner polygons
 * drawn clockwise in TopoJSON output from d3.
 */
function featureToSvgPath(
  rings: ResolvedRing[],
  projection: NonNullable<RegionMapDataOptions["projection"]>,
  plot: SvgRect,
  extent: { minX: number; maxX: number; minY: number; maxY: number }
): string | undefined {
  const segments: string[] = [];
  for (const ring of rings) {
    if (ring.length < 2) {
      continue;
    }
    const pts = ring.map(([lon, lat]) => projectLonLatToPlot(lon, lat, projection, plot, extent));
    segments.push(`M${fmt(pts[0].x)} ${fmt(pts[0].y)}`);
    for (let i = 1; i < pts.length; i++) {
      segments.push(`L${fmt(pts[i].x)} ${fmt(pts[i].y)}`);
    }
    segments.push("Z");
  }
  return segments.length > 0 ? segments.join("") : undefined;
}

/**
 * Geometric (area-weighted) centroid of a polygon ring in lon/lat
 * space. Uses the shoelace formula (Bourke 1988) so vertex density
 * does not bias the result — a country with a densely sampled
 * coastline and a sparse inland border still has its label centred
 * on the polygon's visual mass. The previous implementation returned
 * the vertex mean, which for long-coastline countries (Norway,
 * Chile, Indonesia) sat visibly off-centre.
 *
 * Falls back to the vertex mean when the ring's signed area rounds
 * to zero (degenerate / self-intersecting polygons where the
 * shoelace formula is undefined).
 */
function ringCentroid(ring: ResolvedRing): [number, number] | undefined {
  const n = ring.length;
  if (n === 0) {
    return undefined;
  }
  if (n < 3) {
    // A 1- or 2-point ring has no interior. Return the vertex mean;
    // callers use this centroid purely for label placement.
    let sumX = 0;
    let sumY = 0;
    for (const [x, y] of ring) {
      sumX += x;
      sumY += y;
    }
    return [sumX / n, sumY / n];
  }
  let signedArea = 0;
  let cx = 0;
  let cy = 0;
  for (let i = 0; i < n; i++) {
    const [x0, y0] = ring[i];
    const [x1, y1] = ring[(i + 1) % n];
    const cross = x0 * y1 - x1 * y0;
    signedArea += cross;
    cx += (x0 + x1) * cross;
    cy += (y0 + y1) * cross;
  }
  signedArea *= 0.5;
  if (!Number.isFinite(signedArea) || Math.abs(signedArea) < 1e-12) {
    // Degenerate ring (zero area / numerical collapse) — fall back to
    // the vertex mean so we still return *something* for the label.
    let sumX = 0;
    let sumY = 0;
    for (const [x, y] of ring) {
      sumX += x;
      sumY += y;
    }
    return [sumX / n, sumY / n];
  }
  const scale = 1 / (6 * signedArea);
  return [cx * scale, cy * scale];
}

/** Lowercase + trim label to match the case-insensitive lookup policy. */
function normaliseLabel(label: string): string {
  return label.trim().toLowerCase();
}

interface RegionCoordinate {
  lon: number;
  lat: number;
}

const REGION_COORDINATES: Record<string, RegionCoordinate> = {
  // Approximate country centroids used for regionMap preview rendering.
  // Coordinates are degrees longitude (east positive) / latitude (north
  // positive) of the country's geographic centre. The exact centroid
  // method doesn't matter for a preview-grade renderer — a dot anywhere
  // inside the country is visually acceptable.
  //
  // Coverage targets the ISO-3166-1 short names and common aliases
  // Excel users feed regionMap; synonyms (uk → united kingdom,
  // usa → united states, etc.) are added so the lookup is case- and
  // form-insensitive.
  afghanistan: { lon: 66, lat: 34 },
  albania: { lon: 20, lat: 41 },
  algeria: { lon: 3, lat: 28 },
  angola: { lon: 17, lat: -12 },
  argentina: { lon: -64, lat: -34 },
  armenia: { lon: 45, lat: 40 },
  australia: { lon: 134, lat: -25 },
  austria: { lon: 14, lat: 47 },
  azerbaijan: { lon: 47, lat: 40 },
  bahamas: { lon: -78, lat: 24 },
  bahrain: { lon: 50, lat: 26 },
  bangladesh: { lon: 90, lat: 24 },
  belarus: { lon: 28, lat: 53 },
  belgium: { lon: 4, lat: 50 },
  belize: { lon: -88, lat: 17 },
  benin: { lon: 2, lat: 9 },
  bhutan: { lon: 90, lat: 27 },
  bolivia: { lon: -64, lat: -16 },
  "bosnia and herzegovina": { lon: 18, lat: 44 },
  botswana: { lon: 24, lat: -22 },
  brazil: { lon: -52, lat: -10 },
  brunei: { lon: 114, lat: 4 },
  bulgaria: { lon: 25, lat: 43 },
  "burkina faso": { lon: -2, lat: 12 },
  burundi: { lon: 30, lat: -3 },
  cambodia: { lon: 105, lat: 13 },
  cameroon: { lon: 12, lat: 5 },
  canada: { lon: -106, lat: 56 },
  "central african republic": { lon: 21, lat: 7 },
  chad: { lon: 19, lat: 15 },
  chile: { lon: -71, lat: -30 },
  china: { lon: 104, lat: 35 },
  colombia: { lon: -72, lat: 4 },
  "costa rica": { lon: -84, lat: 10 },
  croatia: { lon: 16, lat: 45 },
  cuba: { lon: -78, lat: 22 },
  cyprus: { lon: 33, lat: 35 },
  "czech republic": { lon: 15, lat: 50 },
  czechia: { lon: 15, lat: 50 },
  "democratic republic of the congo": { lon: 23, lat: -2 },
  denmark: { lon: 10, lat: 56 },
  djibouti: { lon: 43, lat: 12 },
  "dominican republic": { lon: -70, lat: 19 },
  ecuador: { lon: -79, lat: -2 },
  egypt: { lon: 30, lat: 26 },
  "el salvador": { lon: -89, lat: 14 },
  "equatorial guinea": { lon: 10, lat: 2 },
  eritrea: { lon: 39, lat: 15 },
  estonia: { lon: 26, lat: 59 },
  ethiopia: { lon: 40, lat: 9 },
  fiji: { lon: 178, lat: -18 },
  finland: { lon: 26, lat: 64 },
  france: { lon: 2, lat: 46 },
  gabon: { lon: 11, lat: -1 },
  gambia: { lon: -15, lat: 13 },
  georgia: { lon: 43, lat: 42 },
  germany: { lon: 10, lat: 51 },
  ghana: { lon: -2, lat: 8 },
  greece: { lon: 22, lat: 39 },
  greenland: { lon: -42, lat: 72 },
  guatemala: { lon: -90, lat: 15 },
  guinea: { lon: -10, lat: 11 },
  "guinea-bissau": { lon: -15, lat: 12 },
  guyana: { lon: -58, lat: 5 },
  haiti: { lon: -72, lat: 19 },
  honduras: { lon: -86, lat: 15 },
  "hong kong": { lon: 114, lat: 22 },
  hungary: { lon: 20, lat: 47 },
  iceland: { lon: -19, lat: 65 },
  india: { lon: 78, lat: 22 },
  indonesia: { lon: 118, lat: -2 },
  iran: { lon: 53, lat: 32 },
  iraq: { lon: 44, lat: 33 },
  ireland: { lon: -8, lat: 53 },
  israel: { lon: 35, lat: 31 },
  italy: { lon: 12, lat: 43 },
  "ivory coast": { lon: -5, lat: 7 },
  "côte d’ivoire": { lon: -5, lat: 7 },
  jamaica: { lon: -77, lat: 18 },
  japan: { lon: 138, lat: 37 },
  jordan: { lon: 36, lat: 31 },
  kazakhstan: { lon: 68, lat: 48 },
  kenya: { lon: 38, lat: 0 },
  kuwait: { lon: 47, lat: 29 },
  kyrgyzstan: { lon: 75, lat: 41 },
  laos: { lon: 104, lat: 18 },
  latvia: { lon: 25, lat: 56 },
  lebanon: { lon: 36, lat: 33 },
  lesotho: { lon: 28, lat: -29 },
  liberia: { lon: -9, lat: 6 },
  libya: { lon: 17, lat: 27 },
  liechtenstein: { lon: 9, lat: 47 },
  lithuania: { lon: 24, lat: 56 },
  luxembourg: { lon: 6, lat: 49 },
  macedonia: { lon: 21, lat: 41 },
  "north macedonia": { lon: 21, lat: 41 },
  madagascar: { lon: 47, lat: -20 },
  malawi: { lon: 34, lat: -13 },
  malaysia: { lon: 102, lat: 4 },
  mali: { lon: -4, lat: 17 },
  malta: { lon: 14, lat: 35 },
  mauritania: { lon: -11, lat: 20 },
  mauritius: { lon: 57, lat: -20 },
  mexico: { lon: -102, lat: 23 },
  moldova: { lon: 28, lat: 47 },
  monaco: { lon: 7, lat: 43 },
  mongolia: { lon: 104, lat: 46 },
  montenegro: { lon: 19, lat: 42 },
  morocco: { lon: -7, lat: 32 },
  mozambique: { lon: 35, lat: -18 },
  myanmar: { lon: 96, lat: 21 },
  burma: { lon: 96, lat: 21 },
  namibia: { lon: 18, lat: -22 },
  nepal: { lon: 84, lat: 28 },
  netherlands: { lon: 5, lat: 52 },
  "new zealand": { lon: 174, lat: -41 },
  nicaragua: { lon: -85, lat: 13 },
  niger: { lon: 8, lat: 17 },
  nigeria: { lon: 8, lat: 10 },
  "north korea": { lon: 127, lat: 40 },
  norway: { lon: 10, lat: 62 },
  oman: { lon: 57, lat: 21 },
  pakistan: { lon: 70, lat: 30 },
  palestine: { lon: 35, lat: 32 },
  panama: { lon: -80, lat: 9 },
  "papua new guinea": { lon: 144, lat: -6 },
  paraguay: { lon: -58, lat: -23 },
  peru: { lon: -75, lat: -10 },
  philippines: { lon: 122, lat: 13 },
  poland: { lon: 20, lat: 52 },
  portugal: { lon: -8, lat: 40 },
  qatar: { lon: 51, lat: 25 },
  "republic of the congo": { lon: 15, lat: -1 },
  romania: { lon: 25, lat: 46 },
  russia: { lon: 100, lat: 60 },
  rwanda: { lon: 30, lat: -2 },
  "saudi arabia": { lon: 45, lat: 24 },
  senegal: { lon: -15, lat: 14 },
  serbia: { lon: 21, lat: 44 },
  "sierra leone": { lon: -12, lat: 8 },
  singapore: { lon: 104, lat: 1 },
  slovakia: { lon: 19, lat: 49 },
  slovenia: { lon: 15, lat: 46 },
  somalia: { lon: 46, lat: 5 },
  "south africa": { lon: 25, lat: -29 },
  "south korea": { lon: 128, lat: 36 },
  "south sudan": { lon: 30, lat: 7 },
  spain: { lon: -4, lat: 40 },
  "sri lanka": { lon: 81, lat: 7 },
  sudan: { lon: 30, lat: 15 },
  suriname: { lon: -56, lat: 4 },
  sweden: { lon: 15, lat: 62 },
  switzerland: { lon: 8, lat: 47 },
  syria: { lon: 38, lat: 35 },
  taiwan: { lon: 121, lat: 24 },
  tajikistan: { lon: 71, lat: 39 },
  tanzania: { lon: 35, lat: -6 },
  thailand: { lon: 101, lat: 15 },
  togo: { lon: 1, lat: 8 },
  "trinidad and tobago": { lon: -61, lat: 11 },
  tunisia: { lon: 9, lat: 34 },
  turkey: { lon: 35, lat: 39 },
  turkmenistan: { lon: 59, lat: 40 },
  uganda: { lon: 32, lat: 1 },
  ukraine: { lon: 32, lat: 49 },
  "united arab emirates": { lon: 54, lat: 24 },
  uae: { lon: 54, lat: 24 },
  "united kingdom": { lon: -2, lat: 54 },
  uk: { lon: -2, lat: 54 },
  britain: { lon: -2, lat: 54 },
  "great britain": { lon: -2, lat: 54 },
  "united states": { lon: -98, lat: 39 },
  "united states of america": { lon: -98, lat: 39 },
  usa: { lon: -98, lat: 39 },
  us: { lon: -98, lat: 39 },
  uruguay: { lon: -56, lat: -33 },
  uzbekistan: { lon: 64, lat: 41 },
  venezuela: { lon: -66, lat: 8 },
  vietnam: { lon: 108, lat: 16 },
  yemen: { lon: 48, lat: 15 },
  zambia: { lon: 28, lat: -14 },
  zimbabwe: { lon: 30, lat: -19 }
};

function lookupRegionCoordinate(label: string): RegionCoordinate | undefined {
  // Consult the pre-normalised lookup table so queries and keys go
  // through the identical `normalizeRegionLabel` transform. The old
  // `REGION_COORDINATES[normalizeRegionLabel(label)]` path applied the
  // strip-the/republic-of regex to the *query* only, leaving keys like
  // `"democratic republic of the congo"` unreachable via their own
  // canonical name — `lookup("Democratic Republic of the Congo")`
  // normalised to `"democratic congo"` and missed the key.
  return NORMALISED_REGION_COORDINATES.get(normalizeRegionLabel(label));
}

function normalizeRegionLabel(label: string): string {
  return label
    .trim()
    .toLowerCase()
    .replace(/\b(the|republic of)\b/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/**
 * Pre-normalised view of {@link REGION_COORDINATES}. Built once at module
 * load so runtime lookups don't re-normalise keys on every query.
 */
const NORMALISED_REGION_COORDINATES: Map<string, RegionCoordinate> = new Map(
  Object.entries(REGION_COORDINATES).map(([k, v]) => [normalizeRegionLabel(k), v])
);

function renderRegionMapGraticule(parts: string[], plot: SvgRect): void {
  for (let i = 1; i < 4; i++) {
    const x = plot.x + (plot.width * i) / 4;
    parts.push(
      `<line x1="${fmt(x)}" y1="${fmt(plot.y + 8)}" x2="${fmt(x)}" y2="${fmt(plot.y + plot.height - 8)}" stroke="#E5F0FA"/>`
    );
  }
  for (let i = 1; i < 3; i++) {
    const y = plot.y + (plot.height * i) / 3;
    parts.push(
      `<line x1="${fmt(plot.x + 8)}" y1="${fmt(y)}" x2="${fmt(plot.x + plot.width - 8)}" y2="${fmt(y)}" stroke="#E5F0FA"/>`
    );
  }
}

// Mercator projection clamps latitude to `±MERCATOR_LAT_CLAMP_DEG`
// to avoid the singularity at the poles (where `log(tan(π/4 + π/2)) → ∞`).
// 85° is the de facto standard used by Web Mercator / EPSG:3857 tile
// services; clamping to a smaller value would visibly shear circumpolar
// regions. Two separate code paths in this file (`projectLonLatRaw`
// and `projectRegionCoordinate`) previously used different values (85
// vs 84), producing slightly different output for the same topology —
// the centroid path and the polygon path rendered at different zoom
// levels. Unify here.
const MERCATOR_LAT_CLAMP_DEG = 85;

function projectRegionCoordinate(
  coord: RegionCoordinate,
  projection: NonNullable<NonNullable<ChartExSeries["layoutPr"]>["projection"]>,
  plot: SvgRect
): { x: number; y: number } {
  const lon = Math.max(-180, Math.min(180, coord.lon));
  // Only mercator/miller need the ±85° clamp (log-singularity at the
  // poles). Albers / Robinson / equirectangular are finite at ±90°;
  // clamping their input silently drops 5° of polar data.
  const lat =
    projection === "mercator" || projection === "miller"
      ? Math.max(-MERCATOR_LAT_CLAMP_DEG, Math.min(MERCATOR_LAT_CLAMP_DEG, coord.lat))
      : Math.max(-90, Math.min(90, coord.lat));
  let nx: number;
  let ny: number;
  if (projection === "mercator") {
    nx = (lon + 180) / 360;
    const rad = (lat * Math.PI) / 180;
    ny = 0.5 - Math.log(Math.tan(Math.PI / 4 + rad / 2)) / (2 * Math.PI);
  } else if (projection === "miller") {
    nx = (lon + 180) / 360;
    const rad = (lat * Math.PI) / 180;
    ny = 0.5 - (1.25 * Math.log(Math.tan(Math.PI / 4 + 0.4 * rad))) / (2 * Math.PI);
  } else if (projection === "albers") {
    ({ nx, ny } = projectAlbers(lon, lat));
  } else if (projection === "robinson") {
    ({ nx, ny } = projectRobinson(lon, lat));
  } else {
    // Fallback "plain" equirectangular projection: latitude in degrees
    // spans the range `[-90, 90]`, so the normalised Y is `0.5 - lat/180`.
    // The previous constant `190` was a typo — it compressed the
    // vertical axis by ~5.3%, producing visible drift for high-latitude
    // regions (89° mapped to `0.032` instead of `0.006`). Caught by the
    // deep audit, not covered by any existing snapshot.
    nx = (lon + 180) / 360;
    ny = 0.5 - lat / 180;
  }
  return {
    x: plot.x + 14 + clamp01(nx) * Math.max(1, plot.width - 28),
    y: plot.y + 14 + clamp01(ny) * Math.max(1, plot.height - 28)
  };
}

/**
 * Albers Equal-Area Conic projection, normalised to the 0..1 unit square.
 *
 * Uses world-wide standard parallels (φ1=20°, φ2=60°) centred on the
 * equator and prime meridian — the same defaults `proj4` and `d3-geo`
 * apply when no explicit parallels are supplied. The returned `nx`/`ny`
 * are normalised by the world extent so the 19-country centroid table
 * lands inside the plot rectangle for every canvas size.
 *
 * Formulas follow Snyder, "Map Projections — A Working Manual", USGS
 * Professional Paper 1395 §14. Equal-area means the cone coefficient
 * `n` is the same for every latitude, so we compute it once at module
 * load.
 */
const ALBERS_PHI1 = (20 * Math.PI) / 180;
const ALBERS_PHI2 = (60 * Math.PI) / 180;
const ALBERS_N = (Math.sin(ALBERS_PHI1) + Math.sin(ALBERS_PHI2)) / 2;
const ALBERS_C =
  Math.cos(ALBERS_PHI1) * Math.cos(ALBERS_PHI1) + 2 * ALBERS_N * Math.sin(ALBERS_PHI1);
const ALBERS_RHO0 = Math.sqrt(ALBERS_C) / ALBERS_N;
// Pre-compute the world extent in raw Albers units so the projected
// coordinates can be normalised to 0..1.
const ALBERS_RANGE = (() => {
  let minX = Infinity;
  let maxX = -Infinity;
  let minY = Infinity;
  let maxY = -Infinity;
  for (let lon = -180; lon <= 180; lon += 10) {
    for (let lat = -84; lat <= 84; lat += 4) {
      const { rawX, rawY } = rawAlbers(lon, lat);
      if (rawX < minX) {
        minX = rawX;
      }
      if (rawX > maxX) {
        maxX = rawX;
      }
      if (rawY < minY) {
        minY = rawY;
      }
      if (rawY > maxY) {
        maxY = rawY;
      }
    }
  }
  return { minX, maxX, minY, maxY };
})();

function rawAlbers(lon: number, lat: number): { rawX: number; rawY: number } {
  const phi = (lat * Math.PI) / 180;
  const lambda = (lon * Math.PI) / 180;
  const rho = Math.sqrt(Math.max(0, ALBERS_C - 2 * ALBERS_N * Math.sin(phi))) / ALBERS_N;
  const theta = ALBERS_N * lambda;
  return {
    rawX: rho * Math.sin(theta),
    rawY: ALBERS_RHO0 - rho * Math.cos(theta)
  };
}

function projectAlbers(lon: number, lat: number): { nx: number; ny: number } {
  const { rawX, rawY } = rawAlbers(lon, lat);
  const { minX, maxX, minY, maxY } = ALBERS_RANGE;
  // `rawAlbers` follows the Snyder §14 convention where `rawY` grows
  // northward (rawY at +90° ≈ +1.39, at −90° ≈ −0.74). Normalising
  // directly to [0..1] would put the North Pole at ny=1 — but SVG / PDF
  // y grows downward, so ny=1 renders at the BOTTOM of the plot. The
  // other projections in this file (`mercator`, `miller`,
  // equirectangular, `robinson`) all flip the sign explicitly (see
  // `projectRobinson` at line 2888: `sign = lat >= 0 ? -1 : 1`).
  // Without the flip, every Albers-projected map rendered upside down
  // relative to its siblings — the USA sat across the southern
  // hemisphere on every regionMap chart. Mirror the normalised output
  // so north = top for this projection too.
  return {
    nx: (rawX - minX) / (maxX - minX),
    ny: 1 - (rawY - minY) / (maxY - minY)
  };
}

/**
 * Robinson projection — a pseudocylindrical compromise projection whose
 * X and Y scaling factors are defined by a look-up table at 5° latitude
 * intervals with linear interpolation between entries.
 *
 * The factor tables below come from Robinson's original 1974 publication
 * (as reproduced in Snyder & Voxland, "An Album of Map Projections",
 * USGS Professional Paper 1453, Table 34). `PLEN` is the length of a
 * parallel relative to the equator; `PDFE` is the perpendicular distance
 * from the equator. Values are symmetric about the equator, so the
 * implementation looks up `|φ|` and mirrors the Y sign.
 */
const ROBINSON_PLEN = [
  1, 0.9986, 0.9954, 0.99, 0.9822, 0.973, 0.96, 0.9427, 0.9216, 0.8962, 0.8679, 0.835, 0.7986,
  0.7597, 0.7186, 0.6732, 0.6213, 0.5722, 0.5322
];
const ROBINSON_PDFE = [
  0, 0.062, 0.124, 0.186, 0.248, 0.31, 0.372, 0.434, 0.4958, 0.5571, 0.6176, 0.6769, 0.7346, 0.7903,
  0.8435, 0.8936, 0.9394, 0.9761, 1
];

function robinsonFactors(absLatDeg: number): { plen: number; pdfe: number } {
  // Table is indexed by lat / 5°. Use linear interpolation between the
  // two surrounding rows for non-multiple latitudes.
  const idxFloat = Math.min(ROBINSON_PLEN.length - 1, absLatDeg / 5);
  const lo = Math.floor(idxFloat);
  const hi = Math.min(ROBINSON_PLEN.length - 1, lo + 1);
  const t = idxFloat - lo;
  return {
    plen: ROBINSON_PLEN[lo] + (ROBINSON_PLEN[hi] - ROBINSON_PLEN[lo]) * t,
    pdfe: ROBINSON_PDFE[lo] + (ROBINSON_PDFE[hi] - ROBINSON_PDFE[lo]) * t
  };
}

function projectRobinson(lon: number, lat: number): { nx: number; ny: number } {
  const { plen, pdfe } = robinsonFactors(Math.abs(lat));
  const sign = lat >= 0 ? -1 : 1; // SVG y grows downward
  // Robinson's conventional aspect ratio is plen_equator : 2 * pdfe_pole
  //   = 1 : 0.5072, so the normalised Y uses that half-height.
  const nx = 0.5 + (lon / 180) * plen * 0.5;
  const ny = 0.5 + sign * pdfe * 0.5072 * 0.5;
  return { nx, ny };
}

function renderRegionMapTileFallback(
  parts: string[],
  records: Array<{ value: number; label: string }>,
  range: { min: number; max: number },
  plot: SvgRect,
  mode = "tile-fallback"
): void {
  const count = Math.max(1, records.length);
  const cols = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / cols);
  const cellW = plot.width / cols;
  const cellH = plot.height / rows;
  const radius = Math.min(cellW, cellH) * 0.38;
  records.forEach((record, i) => {
    if (!Number.isFinite(record.value)) {
      return;
    }
    const cx = plot.x + (i % cols) * cellW + cellW / 2;
    const cy = plot.y + Math.floor(i / cols) * cellH + cellH / 2;
    const t = clamp01((record.value - range.min) / Math.max(1e-9, range.max - range.min));
    const points = hexagonPoints(cx, cy, radius);
    parts.push(
      `<polygon points="${points}" fill="${interpolateColor("#D9EAF7", "#2F75B5", t)}" stroke="#fff" data-region-map-mode="${mode}"/>`
    );
    if (cellW > 34 && cellH > 22) {
      parts.push(svgText(cx, cy + 3, record.label, 9, "#1F3B53", "middle"));
    }
  });
}

function hexagonPoints(cx: number, cy: number, radius: number): string {
  return Array.from({ length: 6 }, (_, i) => {
    const angle = Math.PI / 6 + (Math.PI * 2 * i) / 6;
    return `${fmt(cx + Math.cos(angle) * radius)},${fmt(cy + Math.sin(angle) * radius)}`;
  }).join(" ");
}

function buildHistogramBins(
  values: number[],
  binning: NonNullable<ChartExSeries["layoutPr"]>["binning"] | undefined
): Array<{ label: string; count: number }> {
  // `collectChartExNumbers` encodes blank / `#N/A` source cells as
  // `NaN` (matching the classic renderer's gap semantics). NaN
  // propagates through `sort((a,b) => a-b)` (the comparator returns
  // `NaN`, which the sort treats as "no swap", leaving the NaNs
  // wherever they landed) — so `sorted[0]` or `sorted[N-1]` can be
  // `NaN`, producing `rawSize = NaN` and throwing downstream. Strip
  // non-finite values up front: a histogram of "blanks mixed with
  // numbers" means "bin the numbers, ignore the blanks", matching
  // Excel's own behaviour.
  const finite = values.filter(v => Number.isFinite(v));
  if (finite.length === 0) {
    return [];
  }
  const sorted = finite.slice().sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  if (binning?.binType === "categories") {
    const counts = new Map<string, number>();
    for (const value of finite) {
      const key = String(value);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return Array.from(counts, ([label, count]) => ({ label, count }));
  }
  const binCount =
    binning?.binCount ??
    (binning?.binType === "binCount" ? 10 : Math.ceil(Math.sqrt(finite.length)));
  // Auto bin width. The previous `Math.max(1, …)` floor collapsed
  // fractional datasets (percentages / probabilities / 0-1 ranges)
  // into a single bin — e.g. `[0.1..0.9]` with `binCount=3` produced
  // `rawSize=1` and a single `[0.1, 1.1]` bin. Scale the bin width to
  // the data range instead, but fall back to 1 when the data collapses
  // to a single point (`max === min`) so the `> 0` guard below doesn't
  // fire for a legitimate all-identical dataset.
  const span = max - min;
  const rawSize = binning?.binSize ?? (span > 0 ? span / Math.max(1, binCount) : 1);
  // Guard against a caller-supplied or computation-produced
  // non-positive bin width. Previously the `for` loop below never
  // advanced when `rawSize <= 0`, degenerating into a 1000-iteration
  // spin (stopped only by the `bins.length > 1000` safety valve) that
  // wasted CPU and produced a chart full of zero-width "bins". Fail
  // loud with a descriptive error instead — this catches user mistakes
  // (e.g. `binning: { binSize: 0 }`) and edge cases where `max === min`
  // with `binCount === 0`.
  if (!(rawSize > 0) || !Number.isFinite(rawSize)) {
    throw new ChartOptionsError(
      `Histogram bin size must be a positive finite number; got ${rawSize}.`
    );
  }
  const start = binning?.underflow ?? min;
  const end = binning?.overflow ?? max;
  const bins: Array<{ low: number; high: number; label: string; count: number }> = [];
  if (binning?.underflow !== undefined) {
    bins.push({
      low: -Infinity,
      high: binning.underflow,
      label: `<=${formatNumber(binning.underflow)}`,
      count: 0
    });
  }
  // Cap the bin count at a sane limit. Excel's own histogram UI
  // accepts up to ~1000 bins; going higher produces unreadable output
  // anyway. `HISTOGRAM_BIN_CAP` used to `break` silently, truncating
  // the last bin's upper bound and mis-positioning the overflow bin
  // — users with a tiny `rawSize` relative to `end - start` got a
  // chart with no visual clue that data was cut off. Fail loud with
  // `ChartOptionsError` instead; callers with legitimately wide
  // ranges should reduce their bin count or widen `rawSize`.
  const HISTOGRAM_BIN_CAP = 1000;
  const expectedBinCount = Math.ceil((end - start) / rawSize);
  if (expectedBinCount > HISTOGRAM_BIN_CAP) {
    throw new ChartOptionsError(
      `Histogram would produce ${expectedBinCount} bins with binSize=${rawSize} over [${start}, ${end}]; ` +
        `the renderer caps at ${HISTOGRAM_BIN_CAP}. Widen the bin size or narrow the data range.`
    );
  }
  // Emit exactly enough bins to cover `[start, end]`. Compute the bin
  // count up front and iterate by index so repeated `low += rawSize`
  // IEEE-754 drift doesn't produce a spurious extra bin — e.g.
  // `start=0.05, end=0.95, rawSize=0.1` previously drifted past `end`
  // because the `0.1`s accumulate (summing ten of them lands at
  // `0.9999999999999999`, not `1.0`), emitting an 11th empty bin that
  // trailed the chart at the right edge. `Math.round` the count so
  // we stop at the last bin whose UPPER edge is still `<= end`;
  // anything beyond `end` is captured by `bins[bins.length - 1]`'s
  // upper fence or the overflow bin.
  const normalSpan = end - start;
  // `Math.round(normalSpan / rawSize)` handles the exact-multiple case
  // (e.g. `span=10, rawSize=2` → 5 bins, not 4 from `Math.floor`
  // after `9.999…`). Adding `8*EPSILON*|span|` is a drift cushion that
  // pulls a `9.999…99` computation up to 10 before rounding.
  const expectedNormalBins = Math.max(
    1,
    Math.round((normalSpan + Number.EPSILON * 8 * normalSpan) / rawSize)
  );
  for (let i = 0; i < expectedNormalBins; i++) {
    const low = start + i * rawSize;
    const high = i === expectedNormalBins - 1 ? end : start + (i + 1) * rawSize;
    bins.push({ low, high, label: `${formatNumber(low)}-${formatNumber(high)}`, count: 0 });
  }
  if (binning?.overflow !== undefined) {
    bins.push({
      low: binning.overflow,
      high: Infinity,
      label: `>${formatNumber(binning.overflow)}`,
      count: 0
    });
  }
  // Degenerate case: `start === end` (all input values identical, or
  // caller set `underflow === overflow`) means the generator loop
  // emitted zero "normal" bins. Without a fallback, the counting loop
  // below would try to write `bins[-1].count++` and throw. Guarantee
  // at least one bin exists by synthesising a unit-width bucket at
  // `start`; this matches Excel's own single-bin output for all-
  // identical data.
  if (bins.length === 0) {
    bins.push({
      low: start,
      high: start + rawSize,
      label: `${formatNumber(start)}-${formatNumber(start + rawSize)}`,
      count: 0
    });
  }
  const closedLeft = binning?.intervalClosed === "l";
  // Index of the lowest "normal" bin (first bin after an optional
  // underflow sentinel). Values equal to the axis minimum need to
  // land here for right-closed intervals — `value > b.low` would
  // otherwise drop them into the fallback. Mirrors Excel's own
  // "values less than or equal to this bin" semantics where the
  // lowest bin has no effective lower bound.
  const firstNormalBinIdx = binning?.underflow !== undefined ? 1 : 0;
  // Similarly, the highest "normal" bin must accept values equal to
  // the axis maximum under left-closed intervals (`value < b.high`
  // otherwise excludes them).
  const lastNormalBinIdx = bins.length - (binning?.overflow !== undefined ? 2 : 1);
  // Iterate over the NaN-filtered array so sparse blanks don't leak
  // into the `bins[bins.length - 1]` fallback at the loop tail.
  for (const value of finite) {
    const bin =
      bins.find((b, idx) => {
        const lowHit = closedLeft
          ? value >= b.low
          : idx === firstNormalBinIdx
            ? value >= b.low
            : value > b.low;
        const highHit = closedLeft
          ? idx === lastNormalBinIdx
            ? value <= b.high
            : value < b.high
          : value <= b.high;
        return lowHit && highHit;
      }) ?? bins[bins.length - 1];
    bin.count++;
  }
  return bins.map(({ label, count }) => ({ label, count }));
}

function groupValuesByCategory(values: number[], categories: string[]): Map<string, number[]> {
  const groups = new Map<string, number[]>();
  values.forEach((value, i) => {
    const key = categories[i] ?? "Values";
    const list = groups.get(key) ?? [];
    list.push(value);
    groups.set(key, list);
  });
  return groups;
}

function boxStats(values: number[], method: "inclusive" | "exclusive") {
  const sorted = values
    .slice()
    .filter(Number.isFinite)
    .sort((a, b) => a - b);
  const safe = sorted.length > 0 ? sorted : [0];
  const q1 = percentile(safe, 0.25, method);
  const median = percentile(safe, 0.5, method);
  const q3 = percentile(safe, 0.75, method);
  const iqr = q3 - q1;
  const lowFence = q1 - 1.5 * iqr;
  const highFence = q3 + 1.5 * iqr;
  const nonOutliers = safe.filter(v => v >= lowFence && v <= highFence);
  const outliers = safe.filter(v => v < lowFence || v > highFence);
  // Whisker bounds are the smallest and largest *non-outlier* values.
  // Use `reduce` instead of `Math.min(...arr)` / `Math.max(...arr)` so the
  // implementation is safe for large samples (the spread form blows the
  // JS call stack past ~100k elements).
  const whiskerSource = nonOutliers.length ? nonOutliers : safe;
  const low = whiskerSource.reduce((acc, v) => (v < acc ? v : acc), whiskerSource[0]);
  const high = whiskerSource.reduce((acc, v) => (v > acc ? v : acc), whiskerSource[0]);
  return {
    q1,
    median,
    q3,
    low,
    high,
    mean: safe.reduce((sum, v) => sum + v, 0) / safe.length,
    outliers,
    // `nonOutliers` — every finite sample within `[lowFence, highFence]`.
    // Exposed so renderers that draw "inner points" do not double-plot
    // outliers as both a filled inner-point dot AND a hollow outlier
    // ring. Excel's convention: inner points are individual non-outlier
    // observations; outlier points are the `|v - median| > 1.5·IQR`
    // samples, and the two overlays must be disjoint.
    nonOutliers
  };
}

function percentile(values: number[], p: number, method: "inclusive" | "exclusive"): number {
  if (values.length === 0) {
    return NaN;
  }
  if (values.length === 1) {
    return values[0];
  }
  // `rank` is 1-indexed per the convention used by NIST / Excel's
  // `PERCENTILE.INC` / `PERCENTILE.EXC`. Clamp to `[1, N]` so the
  // interpolation `fraction = rank - lower` stays in `[0, 1]`.
  // Previously `rank` could fall below 1 for the exclusive method with
  // small `p` (e.g. N=2, p=0.25, exclusive: rank = 3*0.25 = 0.75 → lower
  // clamped to 1 but fraction = −0.25, producing a point lower than
  // both source values — wrong sign). The clamp here matches Excel's
  // behaviour of returning the minimum / maximum for out-of-band `p`.
  const rawRank = method === "inclusive" ? 1 + (values.length - 1) * p : (values.length + 1) * p;
  const rank = Math.max(1, Math.min(values.length, rawRank));
  const lower = Math.max(1, Math.floor(rank));
  const upper = Math.min(values.length, Math.ceil(rank));
  const fraction = rank - lower;
  return values[lower - 1] + (values[upper - 1] - values[lower - 1]) * fraction;
}

function buildHierarchy(levels: string[][], categories: string[], values: number[]): HierarchyNode {
  const root: HierarchyNode = { name: "root", value: 0, children: [] };
  values.forEach((value, i) => {
    let node = root;
    // Preserve explicit empty-string labels (`""` is a legitimate
    // node name in Excel's hierarchy data — "Unassigned" / "Blank"
    // category rolls up under a visible empty slice). The previous
    // `filter(Boolean)` dropped every level where the user had
    // intentionally left the label empty, collapsing those points into
    // the wrong parent. Also filter `null`/`undefined` which do mean
    // "no hierarchy level at this depth".
    const path = [
      ...levels.map(level => level[i]).filter(v => v !== undefined && v !== null),
      categories[i] ?? String(i + 1)
    ];
    // Clamp negative contributions to zero at insert time. Sunburst /
    // treemap layouts use angular / areal sweep proportional to
    // `node.value`; a mix of positive and negative values would
    // otherwise net out to a small (or zero) parent total, producing
    // a ring with zero angular span. NaN / Infinity likewise
    // degrade to zero so they don't poison the sum.
    const safeValue = Number.isFinite(value) && value > 0 ? value : 0;
    for (const name of path) {
      let child = node.children.find(c => c.name === name);
      if (!child) {
        child = { name, value: 0, children: [] };
        node.children.push(child);
      }
      child.value += safeValue;
      node = child;
    }
    root.value += safeValue;
  });
  return root;
}

function sliceDice(
  nodes: HierarchyNode[],
  rect: SvgRect,
  horizontal: boolean
): Array<{ node: HierarchyNode; rect: SvgRect }> {
  const total = nodes.reduce((sum, n) => sum + Math.max(0, n.value), 0) || 1;
  let offset = 0;
  const result: Array<{ node: HierarchyNode; rect: SvgRect }> = [];
  for (const node of nodes) {
    const share = Math.max(0, node.value) / total;
    const r = horizontal
      ? { x: rect.x + offset, y: rect.y, width: rect.width * share, height: rect.height }
      : { x: rect.x, y: rect.y + offset, width: rect.width, height: rect.height * share };
    result.push({ node, rect: r });
    if (node.children.length > 0 && r.width > 18 && r.height > 18) {
      // Push entries one at a time instead of `result.push(...recursive)`.
      // `Function.prototype.apply`-style spread passes each element as a
      // separate function argument, which V8 / JSC throw `RangeError:
      // Maximum call stack size exceeded` on past ~100k entries. The
      // file already documents this defensive pattern in
      // `maxAbsValue` / `valueRange` / `boxStats`; this loop was the
      // last hold-out.
      const children = sliceDice(node.children, insetRect(r, 4), !horizontal);
      for (const child of children) {
        result.push(child);
      }
    }
    offset += horizontal ? r.width : r.height;
  }
  return result;
}

function renderRingSlice(
  cx: number,
  cy: number,
  outer: number,
  inner: number,
  start: number,
  end: number,
  color: string
): string {
  const sweep = end - start;
  // Full ring: SVG can't describe a 360° arc with a single `A` — start
  // and end points coincide so the renderer would emit an empty path.
  // Build a full ring from two 180° arcs on each radius instead. The
  // epsilon guards against float drift (sunburst with one leaf at a
  // given depth has `sweep === endAngle - startAngle === 2π`).
  if (sweep >= Math.PI * 2 - 1e-9) {
    return (
      `<path d="M ${fmt(cx - outer)} ${fmt(cy)} ` +
      `A ${fmt(outer)} ${fmt(outer)} 0 1 1 ${fmt(cx + outer)} ${fmt(cy)} ` +
      `A ${fmt(outer)} ${fmt(outer)} 0 1 1 ${fmt(cx - outer)} ${fmt(cy)} ` +
      `M ${fmt(cx - inner)} ${fmt(cy)} ` +
      `A ${fmt(inner)} ${fmt(inner)} 0 1 0 ${fmt(cx + inner)} ${fmt(cy)} ` +
      `A ${fmt(inner)} ${fmt(inner)} 0 1 0 ${fmt(cx - inner)} ${fmt(cy)} Z" ` +
      `fill="${color}" fill-rule="evenodd" stroke="#fff"/>`
    );
  }
  const large = sweep > Math.PI ? 1 : 0;
  const p1 = polar(cx, cy, outer, start);
  const p2 = polar(cx, cy, outer, end);
  const p3 = polar(cx, cy, inner, end);
  const p4 = polar(cx, cy, inner, start);
  return `<path d="M ${fmt(p1.x)} ${fmt(p1.y)} A ${fmt(outer)} ${fmt(outer)} 0 ${large} 1 ${fmt(p2.x)} ${fmt(p2.y)} L ${fmt(p3.x)} ${fmt(p3.y)} A ${fmt(inner)} ${fmt(inner)} 0 ${large} 0 ${fmt(p4.x)} ${fmt(p4.y)} Z" fill="${color}" stroke="#fff"/>`;
}

function hierarchyDepth(node: HierarchyNode): number {
  // Fold via a loop rather than `Math.max(0, ...arr)` spread — for
  // pathologically-wide hierarchies (>~100k siblings) the spread blows
  // the JS call stack. Every other per-array fold in this file
  // (valueRange, boxStats, funnel max) takes this shape; keep the
  // defensive style consistent here too.
  let maxChild = 0;
  for (const child of node.children) {
    const d = hierarchyDepth(child);
    if (d > maxChild) {
      maxChild = d;
    }
  }
  return 1 + maxChild;
}

function renderAxes(parts: string[], plot: SvgRect, range: { min: number; max: number }): void {
  // Match `drawAxesPdf` (`chart-ex-renderer.ts:330`): emit four
  // interior gridlines and a numeric label on each, so SVG and PDF
  // readers both see a tick value at every grid mark. The top /
  // bottom rails — `range.max` and `range.min` — are drawn as the
  // baseline and left-frame strokes plus their own labels below.
  for (let i = 1; i < 5; i++) {
    const t = i / 5;
    const y = plot.y + plot.height * t;
    parts.push(
      `<line x1="${fmt(plot.x)}" y1="${fmt(y)}" x2="${fmt(plot.x + plot.width)}" y2="${fmt(y)}" stroke="#D9D9D9"/>`
    );
    const value = range.max + (range.min - range.max) * t;
    parts.push(svgText(plot.x - 8, y + 3, formatNumber(value), 10, "#555", "end"));
  }
  parts.push(
    `<line x1="${fmt(plot.x)}" y1="${fmt(plot.y + plot.height)}" x2="${fmt(plot.x + plot.width)}" y2="${fmt(plot.y + plot.height)}" stroke="#444"/>`
  );
  parts.push(
    `<line x1="${fmt(plot.x)}" y1="${fmt(plot.y)}" x2="${fmt(plot.x)}" y2="${fmt(plot.y + plot.height)}" stroke="#444"/>`
  );
  parts.push(
    svgText(
      plot.x - 8,
      valueToY(range.max, range.min, range.max, plot) + 3,
      formatNumber(range.max),
      10,
      "#555",
      "end"
    )
  );
  parts.push(
    svgText(
      plot.x - 8,
      valueToY(range.min, range.min, range.max, plot) + 3,
      formatNumber(range.min),
      10,
      "#555",
      "end"
    )
  );
}

function renderChartExLegend(
  parts: string[],
  series: ChartExSeries[],
  width: number,
  height: number,
  hasTitle: boolean,
  legendPos: "b" | "l" | "r" | "t" | "tr" | undefined
): void {
  // Honour the ChartEx `legendPos` attribute when picking where to
  // stack the swatches. Previously every preview pinned the legend to
  // the right edge regardless of authored position — a chart with
  // `legendPos="b"` rendered with its legend on the right, visibly
  // misrepresenting the authored layout.
  //
  // Layout budget: allow ~18 px per row for vertical stacks and
  // ~96 px per inline item for horizontal stacks. Unknown positions
  // fall back to the right side for backward compatibility.
  const pos = legendPos ?? "r";
  const rowHeight = 18;
  const swatchSize = 10;
  const hGap = 8; // gap between swatch and text
  const itemPadding = 14; // gap between horizontal entries
  const estItemWidth = (label: string): number =>
    swatchSize + hGap + Math.max(28, label.length * 6) + itemPadding;
  const labels = series.map((s, i) => seriesLabelText(s) ?? `Series ${i + 1}`);

  if (pos === "b" || pos === "t") {
    // Horizontal row. Put `b` at the bottom of the canvas, `t`
    // immediately below the chart title (or near the top when none).
    const totalW = labels.reduce((sum, l) => sum + estItemWidth(l), 0) - itemPadding;
    const startX = Math.max(6, (width - totalW) / 2);
    const baseY = pos === "b" ? height - 22 : hasTitle ? 44 : 10;
    let offsetX = startX;
    series.forEach((_s, i) => {
      const label = labels[i];
      const itemWidth = estItemWidth(label);
      parts.push(
        `<rect x="${fmt(offsetX)}" y="${fmt(baseY)}" width="${swatchSize}" height="${swatchSize}" fill="${COLORS[i % COLORS.length]}"/>`
      );
      parts.push(
        svgText(offsetX + swatchSize + hGap, baseY + swatchSize - 1, label, 10, "#555", "start")
      );
      offsetX += itemWidth;
    });
    return;
  }

  if (pos === "l") {
    // Left vertical stack. Place inside the left margin; getPlotRect
    // leaves `plot.x` at 58 so a 10 px swatch + ~40 px label fits.
    series.forEach((_s, i) => {
      const y = (hasTitle ? 44 : 20) + i * rowHeight;
      parts.push(
        `<rect x="${fmt(8)}" y="${fmt(y)}" width="${swatchSize}" height="${swatchSize}" fill="${COLORS[i % COLORS.length]}"/>`
      );
      parts.push(svgText(22, y + 9, labels[i], 10, "#555", "start"));
    });
    return;
  }

  // Default: right / top-right vertical stack. Top-right lifts the
  // stack up near the title baseline; plain right sits slightly below
  // it. Both land inside the 128 px right margin from `getPlotRect`.
  const baseY = pos === "tr" ? (hasTitle ? 44 : 16) : hasTitle ? 44 : 20;
  series.forEach((_s, i) => {
    const y = baseY + i * rowHeight;
    parts.push(
      `<rect x="${fmt(width - 116)}" y="${fmt(y)}" width="${swatchSize}" height="${swatchSize}" fill="${COLORS[i % COLORS.length]}"/>`
    );
    parts.push(svgText(width - 102, y + 9, labels[i], 10, "#555", "start"));
  });
}

/**
 * Extract a plain-text caption from a {@link ChartExSeries.tx} in the
 * canonical preference order: `rich` (walk paragraphs → runs) →
 * `strRef` (the formula string itself, since cached points live on
 * the series-data side of the model not on `tx`) → `value`. Mirrors
 * `Chart.title` resolution on the classic side.
 */
function seriesLabelText(s: ChartExSeries): string | undefined {
  const tx = s.tx;
  if (!tx) {
    return undefined;
  }
  if (tx.rich) {
    return tx.rich.paragraphs.map(p => (p.runs ?? []).map(r => r.text).join("")).join("\n");
  }
  if (tx.value !== undefined) {
    return tx.value;
  }
  if (tx.strRef) {
    // Prefer the cached resolved value when the parser captured one;
    // it's what Excel shows in the legend before formula evaluation.
    // Fall back to the raw formula string so the series remains
    // visually identifiable when no cache was stored.
    if (typeof tx.strRef === "object") {
      return tx.strRef.cached ?? tx.strRef.formula;
    }
    return tx.strRef;
  }
  return undefined;
}

function getPlotRect(width: number, height: number, hasTitle: boolean, titleLines = 1): SvgRect {
  const left = 58;
  const right = 128;
  // Scale the top margin by `titleLines` so multi-paragraph titles
  // (emitted as stacked tspans) do not collide with the plot area.
  // Baseline is 52 for a single-line title (matches the pre-tspan
  // behaviour byte-for-byte); each additional line adds ~22px which
  // matches the `<tspan dy="1.2em">` spacing at the 18px title font.
  const extraLines = Math.max(0, titleLines - 1);
  const top = hasTitle ? 52 + extraLines * 22 : 24;
  const bottom = 46;
  return {
    x: left,
    y: top,
    width: Math.max(10, width - left - right),
    height: Math.max(10, height - top - bottom)
  };
}

function valueRange(values: number[], axis?: ChartExAxis): { min: number; max: number } {
  // Honour author-supplied bounds first. `valScaling.min` / `valScaling.max`
  // (Chart2014 `CT_AxisUnit`) let the user pin the axis regardless of the
  // observed data range; the renderer previously ignored them, so a
  // histogram / pareto / waterfall / boxWhisker authored with an explicit
  // axis bound rendered at auto-computed bounds (user intent lost on every
  // preview, even though the serialised `.xlsx` carried the value).
  //
  // Handle each side independently: a single-sided bound (`min` without
  // `max`, or vice versa) still needs the other side computed from data.
  const authoredMin = axis?.valScaling?.min;
  const authoredMax = axis?.valScaling?.max;
  const hasAuthoredMin = typeof authoredMin === "number" && Number.isFinite(authoredMin);
  const hasAuthoredMax = typeof authoredMax === "number" && Number.isFinite(authoredMax);

  const finite = values.filter(Number.isFinite);
  if (finite.length === 0) {
    // Completely empty / non-finite dataset — fall back to a safe
    // [0, 1] range so downstream `valueToY` doesn't divide by zero.
    // Still honour explicit bounds so the plot still scales correctly
    // when the caller hands us a bound dataset with no finite points.
    const min = hasAuthoredMin ? (authoredMin as number) : 0;
    const max = hasAuthoredMax ? (authoredMax as number) : 1;
    return max > min ? { min, max } : { min, max: min + 1 };
  }
  // Fold over `reduce` instead of `Math.min(...arr)` / `Math.max(...arr)`
  // — the spread form blows the JS call stack past ~100k entries, and
  // waterfall / histogram / pareto series are exactly the workloads
  // that hit this limit.
  //
  // Compute the true data range first. The previous implementation
  // anchored `rawMax = 1` and `rawMin = 0`, which was correct for
  // bin-count axes but wrong for wholly-negative datasets (e.g.
  // `[-5, -3, -1]` produced `[-5, 1]`, wasting the upper half of
  // the plot area on whitespace). Widen zero-anchoring only when the
  // data actually straddles zero on that side; keep the `max >= 1`
  // pad ONLY when the data's natural max is positive — for negative
  // datasets the axis now ends at the max observed value.
  let dataMin = finite[0];
  let dataMax = finite[0];
  for (let i = 1; i < finite.length; i++) {
    const v = finite[i];
    if (v < dataMin) {
      dataMin = v;
    }
    if (v > dataMax) {
      dataMax = v;
    }
  }
  // Zero-anchor symmetry per ECMA-376 bar-like chart conventions:
  //   - mixed data (crosses zero)         → min = dataMin, max = max(1, dataMax)
  //   - wholly non-negative (dataMin ≥ 0) → min = 0, max = max(dataMax, dataMin + 1)
  //                                         (the +1 widens degenerate
  //                                          "all values equal" sets
  //                                          so ticks land on legible
  //                                          round numbers)
  //   - wholly negative (dataMax < 0)     → min = dataMin, max = 0 so
  //                                         the negative values sweep
  //                                         up to the axis
  //
  // The previous implementation forced `rawMax = Math.max(1, dataMax)`
  // whenever `dataMin <= 0`, which was correct for mixed data but
  // buggy for wholly-negative datasets (`[-5, -3, -1]` → `{min: -5,
  // max: 1}`, wasting the upper half of the plot area on whitespace
  // above zero). Zero-anchor to the top only when the data reaches or
  // exceeds 0; end the axis at the observed max otherwise.
  let rawMin: number;
  let rawMax: number;
  if (dataMin >= 0) {
    // All non-negative: anchor to 0, pad top for degenerate ranges.
    rawMin = 0;
    rawMax = Math.max(dataMax, dataMin + 1);
  } else if (dataMax < 0) {
    // All negative: anchor top to 0 so the bars sweep upward to the
    // zero line, showing the magnitude.
    rawMin = dataMin;
    rawMax = 0;
  } else {
    // Mixed (straddles zero): preserve the data range with a floor of
    // 1 on the positive side so small-magnitude bin counts still have
    // legible ticks.
    rawMin = dataMin;
    rawMax = Math.max(1, dataMax);
  }
  // Override with author-supplied bounds. Do this AFTER the data-derived
  // calculation so a single-sided authored bound still lets the other side
  // follow the data.
  const finalMin = hasAuthoredMin ? (authoredMin as number) : rawMin;
  const finalMax = hasAuthoredMax ? (authoredMax as number) : rawMax;
  if (finalMax <= finalMin) {
    // Degenerate or inverted bounds — widen so `valueToY` has a non-zero
    // span. When both bounds were authored we preserve the authored min
    // and invent a top; when one side is data-derived we nudge it to
    // exceed the authored side instead.
    if (hasAuthoredMin && hasAuthoredMax) {
      return { min: finalMin, max: finalMin + 1 };
    }
    if (hasAuthoredMin) {
      return { min: finalMin, max: finalMin + 1 };
    }
    if (hasAuthoredMax) {
      return { min: finalMax - 1, max: finalMax };
    }
    return { min: finalMin, max: finalMin + 1 };
  }
  return { min: finalMin, max: finalMax };
}

/**
 * Locate the "value" axis on a ChartEx plot area — the one whose
 * `valScaling` bounds (`min` / `max`) define the numeric range of the
 * displayed data. Chart2014 `CT_PlotArea` allows zero or more axes; for
 * the layouts that live in this renderer (histogram, pareto, waterfall,
 * boxWhisker) exactly one axis should carry `type === "val"`.
 *
 * Returns `undefined` when no such axis is authored, in which case the
 * caller should fall back to fully data-derived bounds.
 */
function findValueAxis(model: ChartExModel): ChartExAxis | undefined {
  const axes = model.chartSpace.chart.plotArea.axis;
  if (!axes || axes.length === 0) {
    return undefined;
  }
  return axes.find(a => a.type === "val");
}

function svgText(
  x: number,
  y: number,
  text: string,
  fontSize: number,
  color: string,
  anchor: "start" | "middle" | "end"
): string {
  // `color` and `anchor` flow into attribute values — escape defensively
  // even though all current callers pass validated strings, so a future
  // caller can't accidentally inject a `"` that breaks the SVG parse.
  return `<text x="${fmt(x)}" y="${fmt(y)}" text-anchor="${escapeXmlAttr(anchor)}" font-family="Arial" font-size="${fontSize}" fill="${escapeXmlAttr(color)}">${escapeXml(text)}</text>`;
}

function chartTitleText(title: ChartTitle | undefined): string | undefined {
  if (!title) {
    return undefined;
  }
  // Structured rich-text takes precedence — it's the authored form
  // when the builder created the title from a string literal.
  const richText = title.text?.paragraphs
    .map(p => (p.runs ?? []).map(r => r.text).join(""))
    .join("\n");
  if (richText) {
    return richText;
  }
  // Formula-linked titles resolve via the strRef cache — cache-populator
  // fills `strRef.cache.points[0].value` from the referenced cell at
  // workbook save time (and the parser carries `<cx:v>` through on
  // round-trip). Only expose a non-empty cached value so callers can
  // tell "unresolved" from "intentionally empty".
  const cached = title.strRef?.cache?.points?.[0]?.value;
  if (typeof cached === "string" && cached.length > 0) {
    return cached;
  }
  return undefined;
}

/**
 * Extract a hex-encoded fill colour from a shape's {@link ShapeProperties}
 * for rendering in the preview / vector PDF paths. Thin wrapper around
 * {@link previewShapeFillColor} kept for readability inside this file
 * (all the ChartEx renderers take `spPr` directly). Routes through
 * `getSpPrFill` so chart parts that were captured as raw XML by the
 * xform layer (the common case for loaded `.xlsx` files) still resolve
 * their fill correctly — before this fix the helper read
 * `spPr?.fill` directly, which is `undefined` for the `_rawXml` path
 * and dropped the authored colour back to the caller's fallback.
 */
function shapeFillColor(spPr: ShapeProperties | undefined, fallback: string): string {
  return previewShapeFillColor(spPr ? getSpPrFill(spPr) : undefined, fallback);
}
