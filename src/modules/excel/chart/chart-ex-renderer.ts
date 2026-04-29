/**
 * ChartEx renderer — serialises a ChartExModel to `cx:chart` XML.
 *
 * This is a standalone renderer (not a full SAX parser): it produces byte
 * output for a programmatically-built chartEx. Round-trip of existing cx:chart
 * files is handled by raw byte passthrough (model.rawXml is preferred when set).
 */

import type { ChartExAxis, ChartExDataEntry, ChartExModel, ChartExSeries } from "./chart-ex-types";
import {
  renderSvgToPng,
  type ChartPdfDrawingSurface,
  type ChartPdfPathOp,
  type ChartRenderOptions,
  type PdfColor,
  type RegionMapDataOptions,
  type RegionMapMatchRule
} from "./chart-renderer";
import { resolveTopologyObject, type ResolvedRing, type TopologyLike } from "./topojson";
import type {
  ChartColor,
  ChartTextProperties,
  ChartTitle,
  EffectList,
  Scene3D,
  Shadow,
  ShapeProperties,
  ShapeProperties3D
} from "./types";

const DEFAULT_WIDTH = 640;
const DEFAULT_HEIGHT = 360;
const COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

/**
 * Render a ChartExModel to the full XML string representation of cx:chart.
 *
 * By default the function prefers the raw XML captured at parse time
 * (`model.rawXml`) when present — that path gives byte-perfect
 * round-trip for files that haven't been mutated since load. This is
 * the fastest and safest mode for the common "load → save" pipeline.
 *
 * Pass `{ forceStructural: true }` when the caller has mutated
 * structured fields on the model (e.g. via `setSpPrFill` /
 * `setSpPrLine` / direct assignment) and needs the writer to
 * rebuild from the model tree instead of reusing the stale raw bytes.
 * The `Chart.mutateChartEx` API already clears `model.rawXml` for
 * normal mutations; `forceStructural` is the escape hatch for callers
 * that bypass that helper (e.g. low-level tests, ad-hoc scripts).
 */
export function renderChartEx(
  model: ChartExModel,
  options: { forceStructural?: boolean } = {}
): string {
  // Prefer raw XML for existing round-tripped charts unless the caller
  // explicitly asked for a structural rebuild.
  if (model.rawXml && !options.forceStructural) {
    return model.rawXml;
  }

  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push(
    [
      "<cx:chartSpace",
      '  xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"',
      '  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
      '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      '  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    ].join("\n")
  );
  const space = model.chartSpace;
  parts.push(renderChartData(space.chartData));
  parts.push(renderChart(space.chart));
  if (space.clrMapOvr) {
    parts.push(space.clrMapOvr);
  }
  if (space.extLst) {
    parts.push(space.extLst);
  }
  parts.push("</cx:chartSpace>");
  return parts.join("\n");
}

export function renderChartExSvg(model: ChartExModel, options: ChartRenderOptions = {}): string {
  const width = options.width ?? DEFAULT_WIDTH;
  const height = options.height ?? DEFAULT_HEIGHT;
  const title = options.title ?? chartTitleText(model.chartSpace.chart.title);
  const plot = getPlotRect(width, height, !!title);
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
    parts.push(`<rect width="100%" height="100%" fill="${escapeAttr(backgroundColor)}"/>`);
  }
  if (title) {
    parts.push(
      `<text x="${fmt(width / 2)}" y="26" text-anchor="middle" font-family="Arial" font-size="18" fill="#222">${escapeXml(title)}</text>`
    );
  }
  for (let i = 0; i < series.length; i++) {
    renderChartExSeriesSvg(parts, model, series[i], plot, i, options);
  }
  renderChartExLegend(parts, series, width, !!title);
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
 * {@link ChartPdfDrawingSurface}. Only the layout IDs in
 * {@link VECTOR_PDF_CHART_EX_LAYOUT_IDS} are supported today
 * (sunburst, treemap) — the surrounding geometry lives in the
 * "collect" functions shared with the SVG renderer, so when a third
 * layout gets promoted its SVG and PDF paths stay equivalent by
 * construction.
 *
 * For other layout IDs the function throws; callers (notably
 * `@pdf/excel-bridge.chartToPdf`) must pre-filter via
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
  const titleHeight = titleText ? 28 : 0;
  const plot: SvgRect = {
    x: rect.x + 12,
    y: rect.y + titleHeight + 12,
    width: Math.max(10, rect.width - 24),
    height: Math.max(10, rect.height - titleHeight - 24)
  };

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
    surface.drawText(titleText, {
      x: rect.x + rect.width / 2,
      y: rect.y + 20,
      fontSize: 16,
      anchor: "middle",
      color: hexToPdfColor("#222222")
    });
  }

  for (const series of seriesList) {
    if (series.layoutId === "treemap") {
      drawTreemapPdf(surface, model, series, plot);
    } else if (series.layoutId === "sunburst") {
      drawSunburstPdf(surface, model, series, plot);
    } else if (series.layoutId === "funnel") {
      drawFunnelPdf(surface, model, series, plot);
    } else if (series.layoutId === "waterfall") {
      drawWaterfallPdf(surface, model, series, plot);
    } else if (series.layoutId === "boxWhisker") {
      drawBoxWhiskerPdf(surface, model, series, plot);
    } else if (series.layoutId === "regionMap") {
      drawRegionMapPdf(surface, model, series, plot, options.regionMap);
    } else if (series.layoutId === "clusteredColumn") {
      // Both `histogram` and `pareto` live under clusteredColumn after
      // builder normalisation; distinguishing them is a single runtime
      // flag (`layoutPr.paretoLine`).
      if (series.layoutPr?.paretoLine) {
        drawParetoPdf(surface, model, series, plot);
      } else {
        drawHistogramPdf(surface, model, series, plot);
      }
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
  color: string
): void {
  const range = valueRange(values);
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
    COLORS[0]
  );
}

function drawParetoPdf(
  surface: ChartPdfDrawingSurface,
  model: ChartExModel,
  series: ChartExSeries,
  plot: SvgRect
): void {
  const refs = resolveChartExRefs(model, series);
  const values = refs.values;
  const categories =
    refs.categories.length > 0 ? refs.categories : values.map((_, i) => String(i + 1));
  const sorted = values
    .map((value, i) => ({ value, category: categories[i] ?? String(i + 1) }))
    .sort((a, b) => b.value - a.value);
  const sortedValues = sorted.map(item => item.value);
  drawColumnsPdf(
    surface,
    sortedValues,
    sorted.map(item => item.category),
    plot,
    COLORS[0]
  );

  // Cumulative polyline. Rendered as connected `drawLine` segments so
  // the path stays visible on surfaces without `drawPath`, matching
  // the SVG polyline behaviour. `drawCircle` is used for the dots
  // when available; otherwise small rects serve as markers.
  const total = sortedValues.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
  let cumulative = 0;
  const count = Math.max(1, sortedValues.length);
  const step = plot.width / count;
  const points = sortedValues.map((value, i) => {
    cumulative += Math.max(0, value);
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
    const start = subtotalIdx.has(i) ? 0 : running;
    const end = subtotalIdx.has(i) ? value : running + value;
    running = end;
    return { start, end, value, total: subtotalIdx.has(i) };
  });
  const range = valueRange(spans.flatMap(s => [s.start, s.end]));
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
    const colorHex = span.total
      ? shapeFillColor(series.layoutPr?.totalSpPr, COLORS[2])
      : span.value >= 0
        ? shapeFillColor(series.layoutPr?.increaseSpPr, "#70AD47")
        : shapeFillColor(series.layoutPr?.decreaseSpPr, "#C00000");
    surface.drawRect({
      x,
      y: Math.min(y1, y2),
      width: groupWidth * 0.64,
      height: Math.max(1, Math.abs(y1 - y2)),
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
  const max = Math.max(1, ...values.map(v => Math.abs(v)));
  const count = Math.max(1, values.length);
  const h = plot.height / count;
  const labelColor: PdfColor = { r: 1, g: 1, b: 1 };
  const whiteStroke: PdfColor = { r: 1, g: 1, b: 1 };
  values.forEach((value, i) => {
    const topW = (Math.abs(value) / max) * plot.width;
    const bottomW = (Math.abs(values[i + 1] ?? value) / max) * plot.width;
    const y = plot.y + i * h;
    const cx = plot.x + plot.width / 2;
    const fill = hexToPdfColor(COLORS[i % COLORS.length]);
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
  const range = valueRange(allValues);
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
    // Inner points — small translucent dots, one per raw value.
    if (series.layoutPr?.showInnerPoints) {
      const innerFill = { ...seriesColor, a: 0.55 };
      for (const value of groups.get(key) ?? []) {
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
  const range = valueRange(values);

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
    const t = (record.value - range.min) / (range.max - range.min);
    const radius = 6 + Math.sqrt(Math.max(0, t)) * 14;
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

  const strokeHex = mapOptions.strokeColor ?? "#FFFFFF";
  const strokeColor = hexToPdfColor(strokeHex);
  let matchedCount = 0;
  for (const feature of features) {
    const match = resolvedMatch.get(feature);
    const fillHex = match
      ? (() => {
          matchedCount++;
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

  return matchedCount > 0;
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
    const cx = plot.x + (i % cols) * cellW + cellW / 2;
    const cy = plot.y + Math.floor(i / cols) * cellH + cellH / 2;
    const t = (record.value - range.min) / Math.max(1e-9, range.max - range.min);
    const fill = hexToPdfColor(interpolateColor("#D9EAF7", "#2F75B5", Math.max(0, Math.min(1, t))));
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

function hexToPdfColor(hex: string): PdfColor {
  const clean = hex.replace(/^#/, "");
  const full =
    clean.length === 3
      ? clean
          .split("")
          .map(c => c + c)
          .join("")
      : clean;
  const r = parseInt(full.slice(0, 2), 16) / 255;
  const g = parseInt(full.slice(2, 4), 16) / 255;
  const b = parseInt(full.slice(4, 6), 16) / 255;
  return { r, g, b };
}

interface SvgRect {
  x: number;
  y: number;
  width: number;
  height: number;
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
  switch (series.layoutId) {
    case "funnel":
      renderFunnelSvg(parts, values, categories, plot);
      return;
    case "waterfall":
      renderWaterfallSvg(parts, values, categories, series, plot);
      return;
    case "clusteredColumn":
      if (series.layoutPr?.paretoLine) {
        renderParetoSvg(parts, values, categories, series, plot);
      } else {
        renderHistogramSvg(parts, values, series, plot);
      }
      return;
    case "boxWhisker":
      renderBoxWhiskerSvg(parts, values, categories, series, plot);
      return;
    case "treemap":
      renderTreemapSvg(parts, buildHierarchy(refs.hierarchy, categories, values), plot);
      return;
    case "sunburst":
      renderSunburstSvg(parts, buildHierarchy(refs.hierarchy, categories, values), plot);
      return;
    case "regionMap":
      renderRegionMapSvg(parts, values, categories, series, plot, renderOptions.regionMap);
      return;
    default:
      renderColumnSvg(parts, values, categories, plot, COLORS[seriesIndex % COLORS.length]);
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
  const stringDims = entries
    .filter(entry => entry.strDim)
    .map(entry => collectChartExStrings(entry));
  const numericDims = entries
    .filter(entry => entry.numDim)
    .map(entry => collectChartExNumbers(entry));
  return {
    values: numericDims[0] ?? [],
    categories: stringDims[0] ?? [],
    hierarchy: stringDims.slice(1)
  };
}

function collectChartExStrings(entry: ChartExDataEntry): string[] {
  const levels = entry.strDim?.levels ?? [];
  const first = levels[0];
  if (!first) {
    return [];
  }
  const count = first.ptCount ?? first.points.length;
  const values = Array.from({ length: count }, (_, i) => String(i + 1));
  for (const point of first.points) {
    values[point.index] = point.value;
  }
  return values;
}

function collectChartExNumbers(entry: ChartExDataEntry): number[] {
  const levels = entry.numDim?.levels ?? [];
  const first = levels[0];
  if (!first) {
    return [];
  }
  const count = first.ptCount ?? first.points.length;
  const values = Array.from({ length: count }, () => 0);
  for (const point of first.points) {
    values[point.index] = point.value;
  }
  return values;
}

function renderColumnSvg(
  parts: string[],
  values: number[],
  categories: string[],
  plot: SvgRect,
  color: string
): void {
  const range = valueRange(values);
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
  plot: SvgRect
): void {
  const bins = buildHistogramBins(values, series.layoutPr?.binning);
  renderColumnSvg(
    parts,
    bins.map(bin => bin.count),
    bins.map(bin => bin.label),
    plot,
    COLORS[0]
  );
}

function renderParetoSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect
): void {
  const sorted = values
    .map((value, i) => ({ value, category: categories[i] ?? String(i + 1) }))
    .sort((a, b) => b.value - a.value);
  const sortedValues = sorted.map(item => item.value);
  renderColumnSvg(
    parts,
    sortedValues,
    sorted.map(item => item.category),
    plot,
    COLORS[0]
  );
  const total = sortedValues.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
  let cumulative = 0;
  const count = Math.max(1, sortedValues.length);
  const step = plot.width / count;
  const points = sortedValues.map((value, i) => {
    cumulative += Math.max(0, value);
    return {
      x: plot.x + i * step + step / 2,
      y: plot.y + plot.height - (cumulative / total) * plot.height
    };
  });
  parts.push(
    `<polyline points="${points.map(p => `${fmt(p.x)},${fmt(p.y)}`).join(" ")}" fill="none" stroke="${COLORS[1]}" stroke-width="2"/>`
  );
  for (const p of points) {
    parts.push(`<circle cx="${fmt(p.x)}" cy="${fmt(p.y)}" r="3" fill="${COLORS[1]}"/>`);
  }
  if (series.layoutPr?.paretoLine) {
    parts.push(svgText(plot.x + plot.width - 4, plot.y + 12, "Cumulative %", 10, COLORS[1], "end"));
  }
}

function renderWaterfallSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect
): void {
  const subtotalIdx = new Set(series.layoutPr?.subtotals?.map(s => s.idx) ?? []);
  let running = 0;
  const spans = values.map((value, i) => {
    const start = subtotalIdx.has(i) ? 0 : running;
    const end = subtotalIdx.has(i) ? value : running + value;
    running = end;
    return { start, end, value, total: subtotalIdx.has(i) };
  });
  const range = valueRange(spans.flatMap(s => [s.start, s.end]));
  renderAxes(parts, plot, range);
  const count = Math.max(1, spans.length);
  const groupWidth = plot.width / count;
  const centers: Array<{ x: number; y: number }> = [];
  spans.forEach((span, i) => {
    const y1 = valueToY(span.start, range.min, range.max, plot);
    const y2 = valueToY(span.end, range.min, range.max, plot);
    const x = plot.x + i * groupWidth + groupWidth * 0.18;
    const color = span.total
      ? shapeFillColor(series.layoutPr?.totalSpPr, COLORS[2])
      : span.value >= 0
        ? shapeFillColor(series.layoutPr?.increaseSpPr, "#70AD47")
        : shapeFillColor(series.layoutPr?.decreaseSpPr, "#C00000");
    parts.push(
      `<rect x="${fmt(x)}" y="${fmt(Math.min(y1, y2))}" width="${fmt(groupWidth * 0.64)}" height="${fmt(Math.max(1, Math.abs(y1 - y2)))}" fill="${color}"/>`
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
  plot: SvgRect
): void {
  const max = Math.max(1, ...values.map(v => Math.abs(v)));
  const count = Math.max(1, values.length);
  const h = plot.height / count;
  values.forEach((value, i) => {
    const topW = (Math.abs(value) / max) * plot.width;
    const bottomW = (Math.abs(values[i + 1] ?? value) / max) * plot.width;
    const y = plot.y + i * h;
    const cx = plot.x + plot.width / 2;
    parts.push(
      `<polygon points="${fmt(cx - topW / 2)},${fmt(y)} ${fmt(cx + topW / 2)},${fmt(y)} ${fmt(cx + bottomW / 2)},${fmt(y + h * 0.88)} ${fmt(cx - bottomW / 2)},${fmt(y + h * 0.88)}" fill="${COLORS[i % COLORS.length]}" stroke="#fff"/>`
    );
    parts.push(svgText(cx, y + h * 0.55, categories[i] ?? String(i + 1), 11, "#fff", "middle"));
  });
}

function renderBoxWhiskerSvg(
  parts: string[],
  values: number[],
  categories: string[],
  series: ChartExSeries,
  plot: SvgRect
): void {
  const groups =
    categories.length > 0
      ? groupValuesByCategory(values, categories)
      : new Map([["Values", values]]);
  const allValues = Array.from(groups.values()).flat();
  const range = valueRange(allValues);
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
      for (const value of groups.get(key) ?? []) {
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
  return entries.map((entry, i) => ({
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
  const maxDepth = hierarchyDepth(root);
  const radius = Math.min(plot.width, plot.height) / 2.25;
  collectSunburstSlicesRecursive(slices, root, cx, cy, 0, Math.PI * 2, 0, maxDepth, radius, 0);
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
  for (const child of node.children) {
    const next = angle + (Math.max(0, child.value) / total) * (endAngle - startAngle);
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
  mapOptions?: RegionMapDataOptions
): void {
  const range = valueRange(values);

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
    const t = (record.value - range.min) / (range.max - range.min);
    const radius = 6 + Math.sqrt(Math.max(0, t)) * 14;
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

  parts.push(
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
      parts.push(
        `<path d="${path}" fill="${fill}" stroke="${escapeAttr(stroke)}" stroke-width="0.5" stroke-linejoin="round"/>`
      );
    }
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
      parts.push(svgText(projected.x, projected.y + 3, originalLabel, 9, "#1F3B53", "middle"));
    }
  }

  return matchedCount > 0;
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
  const clampedLat = Math.max(-85, Math.min(85, lat));
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

/** Geometric centroid of a ring of lon/lat pairs. */
function ringCentroid(ring: ResolvedRing): [number, number] | undefined {
  if (ring.length === 0) {
    return undefined;
  }
  let sumX = 0;
  let sumY = 0;
  for (const [x, y] of ring) {
    sumX += x;
    sumY += y;
  }
  return [sumX / ring.length, sumY / ring.length];
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
  return REGION_COORDINATES[normalizeRegionLabel(label)];
}

function normalizeRegionLabel(label: string): string {
  return label
    .trim()
    .toLowerCase()
    .replace(/\b(the|republic of)\b/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

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

function projectRegionCoordinate(
  coord: RegionCoordinate,
  projection: NonNullable<NonNullable<ChartExSeries["layoutPr"]>["projection"]>,
  plot: SvgRect
): { x: number; y: number } {
  const lon = Math.max(-180, Math.min(180, coord.lon));
  const lat = Math.max(-84, Math.min(84, coord.lat));
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
    nx = (lon + 180) / 360;
    ny = 0.5 - lat / 190;
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
  return {
    nx: (rawX - minX) / (maxX - minX),
    ny: (rawY - minY) / (maxY - minY)
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

function clamp01(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  if (value < 0) {
    return 0;
  }
  if (value > 1) {
    return 1;
  }
  return value;
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
    const cx = plot.x + (i % cols) * cellW + cellW / 2;
    const cy = plot.y + Math.floor(i / cols) * cellH + cellH / 2;
    const t = (record.value - range.min) / (range.max - range.min);
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
  if (values.length === 0) {
    return [];
  }
  const sorted = values.slice().sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  if (binning?.binType === "categories") {
    const counts = new Map<string, number>();
    for (const value of values) {
      const key = String(value);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return Array.from(counts, ([label, count]) => ({ label, count }));
  }
  const binCount =
    binning?.binCount ??
    (binning?.binType === "binCount" ? 10 : Math.ceil(Math.sqrt(values.length)));
  const rawSize = binning?.binSize ?? Math.max(1, (max - min) / Math.max(1, binCount));
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
  for (let low = start; low <= end; low += rawSize) {
    const high = low + rawSize;
    bins.push({ low, high, label: `${formatNumber(low)}-${formatNumber(high)}`, count: 0 });
    if (bins.length > 1000) {
      break;
    }
  }
  if (binning?.overflow !== undefined) {
    bins.push({
      low: binning.overflow,
      high: Infinity,
      label: `>${formatNumber(binning.overflow)}`,
      count: 0
    });
  }
  const closedLeft = binning?.intervalClosed === "l";
  for (const value of values) {
    const bin =
      bins.find(b =>
        closedLeft ? value >= b.low && value < b.high : value > b.low && value <= b.high
      ) ?? bins[bins.length - 1];
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
  return {
    q1,
    median,
    q3,
    low: Math.min(...(nonOutliers.length ? nonOutliers : safe)),
    high: Math.max(...(nonOutliers.length ? nonOutliers : safe)),
    mean: safe.reduce((sum, v) => sum + v, 0) / safe.length,
    outliers
  };
}

function percentile(values: number[], p: number, method: "inclusive" | "exclusive"): number {
  if (values.length === 1) {
    return values[0];
  }
  const rank = method === "inclusive" ? 1 + (values.length - 1) * p : (values.length + 1) * p;
  const lower = Math.max(1, Math.floor(rank));
  const upper = Math.min(values.length, Math.ceil(rank));
  const fraction = rank - lower;
  return values[lower - 1] + (values[upper - 1] - values[lower - 1]) * fraction;
}

function buildHierarchy(levels: string[][], categories: string[], values: number[]): HierarchyNode {
  const root: HierarchyNode = { name: "root", value: 0, children: [] };
  values.forEach((value, i) => {
    let node = root;
    const path = [...levels.map(level => level[i]).filter(Boolean), categories[i] ?? String(i + 1)];
    for (const name of path) {
      let child = node.children.find(c => c.name === name);
      if (!child) {
        child = { name, value: 0, children: [] };
        node.children.push(child);
      }
      child.value += value;
      node = child;
    }
    root.value += value;
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
      result.push(...sliceDice(node.children, insetRect(r, 4), !horizontal));
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
  const large = end - start > Math.PI ? 1 : 0;
  const p1 = polar(cx, cy, outer, start);
  const p2 = polar(cx, cy, outer, end);
  const p3 = polar(cx, cy, inner, end);
  const p4 = polar(cx, cy, inner, start);
  return `<path d="M ${fmt(p1.x)} ${fmt(p1.y)} A ${fmt(outer)} ${fmt(outer)} 0 ${large} 1 ${fmt(p2.x)} ${fmt(p2.y)} L ${fmt(p3.x)} ${fmt(p3.y)} A ${fmt(inner)} ${fmt(inner)} 0 ${large} 0 ${fmt(p4.x)} ${fmt(p4.y)} Z" fill="${color}" stroke="#fff"/>`;
}

function hierarchyDepth(node: HierarchyNode): number {
  return 1 + Math.max(0, ...node.children.map(hierarchyDepth));
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
  hasTitle: boolean
): void {
  series.forEach((s, i) => {
    const y = (hasTitle ? 44 : 20) + i * 18;
    parts.push(
      `<rect x="${fmt(width - 116)}" y="${fmt(y)}" width="10" height="10" fill="${COLORS[i % COLORS.length]}"/>`
    );
    parts.push(svgText(width - 102, y + 9, s.tx?.value ?? `Series ${i + 1}`, 10, "#555", "start"));
  });
}

function getPlotRect(width: number, height: number, hasTitle: boolean): SvgRect {
  const left = 58;
  const right = 128;
  const top = hasTitle ? 52 : 24;
  const bottom = 46;
  return {
    x: left,
    y: top,
    width: Math.max(10, width - left - right),
    height: Math.max(10, height - top - bottom)
  };
}

function valueRange(values: number[]): { min: number; max: number } {
  const finite = values.filter(Number.isFinite);
  const min = finite.length > 0 ? Math.min(0, ...finite) : 0;
  const max = finite.length > 0 ? Math.max(1, ...finite) : 1;
  return max <= min ? { min, max: min + 1 } : { min, max };
}

function valueToY(value: number, min: number, max: number, plot: SvgRect): number {
  return plot.y + plot.height - ((value - min) / (max - min)) * plot.height;
}

function svgText(
  x: number,
  y: number,
  text: string,
  fontSize: number,
  color: string,
  anchor: "start" | "middle" | "end"
): string {
  return `<text x="${fmt(x)}" y="${fmt(y)}" text-anchor="${anchor}" font-family="Arial" font-size="${fontSize}" fill="${color}">${escapeXml(text)}</text>`;
}

function insetRect(rect: SvgRect, amount: number): SvgRect {
  return {
    x: rect.x + amount,
    y: rect.y + amount,
    width: Math.max(0, rect.width - amount * 2),
    height: Math.max(0, rect.height - amount * 2)
  };
}

function polar(cx: number, cy: number, radius: number, angle: number): { x: number; y: number } {
  return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
}

function chartTitleText(title: ChartTitle | undefined): string | undefined {
  return title?.text?.paragraphs.map(p => (p.runs ?? []).map(r => r.text).join("")).join("\n");
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function withAlpha(hex: string, alpha: number): string {
  const clean = hex.replace(/^#/, "");
  const mix = (component: string) => {
    const value = parseInt(component, 16);
    return Math.round(value * alpha + 255 * (1 - alpha))
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${mix(clean.slice(0, 2))}${mix(clean.slice(2, 4))}${mix(clean.slice(4, 6))}`;
}

function interpolateColor(a: string, b: string, t: number): string {
  const ca = a.replace(/^#/, "");
  const cb = b.replace(/^#/, "");
  const safe = Math.max(0, Math.min(1, t));
  const mix = (i: number) => {
    const av = parseInt(ca.slice(i, i + 2), 16);
    const bv = parseInt(cb.slice(i, i + 2), 16);
    return Math.round(av + (bv - av) * safe)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${mix(0)}${mix(2)}${mix(4)}`;
}

function shapeFillColor(spPr: ShapeProperties | undefined, fallback: string): string {
  const srgb = spPr?.fill?.solid?.srgb;
  return srgb ? `#${srgb.replace(/^#/, "")}` : fallback;
}

function fmt(value: number): string {
  return value.toFixed(2).replace(/\.00$/, "");
}

function renderChartData(data: ChartExModel["chartSpace"]["chartData"]): string {
  const parts: string[] = [];
  parts.push("  <cx:chartData>");
  if (data.externalData) {
    for (const ed of data.externalData) {
      const attrs = ed.autoUpdate === undefined ? "" : ` autoUpdate="${ed.autoUpdate ? "1" : "0"}"`;
      parts.push(`    <cx:externalData r:id="${ed.id}"${attrs}/>`);
    }
  }
  for (const entry of data.data) {
    parts.push(renderDataEntry(entry));
  }
  parts.push("  </cx:chartData>");
  return parts.join("\n");
}

function renderDataEntry(entry: ChartExDataEntry): string {
  const parts: string[] = [];
  parts.push(`    <cx:data id="${entry.id}">`);
  if (entry.strDim) {
    const d = entry.strDim;
    parts.push(`      <cx:strDim type="${d.type}">`);
    if (d.formula) {
      parts.push(`        <cx:f>${escapeXml(d.formula)}</cx:f>`);
    }
    if (d.levels) {
      for (const lvl of d.levels) {
        const ptCount = lvl.ptCount ?? lvl.points.length;
        const ptAttr = ` ptCount="${ptCount}"`;
        if (lvl.points.length === 0) {
          parts.push(`        <cx:lvl${ptAttr}/>`);
        } else {
          parts.push(`        <cx:lvl${ptAttr}>`);
          for (const p of lvl.points) {
            parts.push(`          <cx:pt idx="${p.index}">${escapeXml(p.value)}</cx:pt>`);
          }
          parts.push("        </cx:lvl>");
        }
      }
    }
    parts.push("      </cx:strDim>");
  }
  if (entry.numDim) {
    const d = entry.numDim;
    parts.push(`      <cx:numDim type="${d.type}">`);
    if (d.formula) {
      parts.push(`        <cx:f>${escapeXml(d.formula)}</cx:f>`);
    }
    if (d.levels) {
      for (const lvl of d.levels) {
        const fmtAttr = lvl.formatCode ? ` formatCode="${escapeAttr(lvl.formatCode)}"` : "";
        const ptCount = lvl.ptCount ?? lvl.points.length;
        const ptAttr = ` ptCount="${ptCount}"`;
        if (lvl.points.length === 0) {
          parts.push(`        <cx:lvl${ptAttr}${fmtAttr}/>`);
        } else {
          parts.push(`        <cx:lvl${ptAttr}${fmtAttr}>`);
          for (const p of lvl.points) {
            parts.push(`          <cx:pt idx="${p.index}">${p.value}</cx:pt>`);
          }
          parts.push("        </cx:lvl>");
        }
      }
    }
    parts.push("      </cx:numDim>");
  }
  parts.push("    </cx:data>");
  return parts.join("\n");
}

function renderChart(chart: ChartExModel["chartSpace"]["chart"]): string {
  const parts: string[] = [];
  parts.push("  <cx:chart>");
  if (chart.title) {
    parts.push(renderTitle(chart.title));
  }
  if (chart.autoTitleDeleted !== undefined && !chart.title) {
    parts.push(`    <cx:autoTitleDeleted val="${chart.autoTitleDeleted ? "1" : "0"}"/>`);
  }
  parts.push(renderPlotArea(chart.plotArea));
  if (chart.legend) {
    parts.push(renderLegend(chart.legend));
  }
  if (chart.spPr) {
    parts.push(renderSpPr(chart.spPr, "    "));
  }
  parts.push("  </cx:chart>");
  return parts.join("\n");
}

function renderTitle(title: ChartTitle): string {
  const parts: string[] = [];
  parts.push("    <cx:title>");
  if (title.text) {
    parts.push("      <cx:tx>");
    parts.push("        <cx:rich>");
    parts.push("          <a:bodyPr/>");
    parts.push("          <a:lstStyle/>");
    for (const p of title.text.paragraphs) {
      parts.push("          <a:p>");
      for (const run of p.runs ?? []) {
        parts.push("            <a:r>");
        parts.push(`              <a:t>${escapeXml(run.text)}</a:t>`);
        parts.push("            </a:r>");
      }
      parts.push("          </a:p>");
    }
    parts.push("        </cx:rich>");
    parts.push("      </cx:tx>");
  }
  parts.push(`      <cx:overlay val="${title.overlay ? "1" : "0"}"/>`);
  parts.push("    </cx:title>");
  return parts.join("\n");
}

function renderPlotArea(pa: ChartExModel["chartSpace"]["chart"]["plotArea"]): string {
  const parts: string[] = [];
  parts.push("    <cx:plotArea>");
  if (pa.spPr) {
    parts.push(renderSpPr(pa.spPr, "      "));
  }
  const region = pa.plotAreaRegion;
  if (region) {
    parts.push("      <cx:plotAreaRegion>");
    if (region.layout) {
      parts.push(renderRawOrEmptyLayout(region.layout as any, "        "));
    }
    if (region.plotSurface) {
      parts.push(renderSpPr(region.plotSurface, "        "));
    }
    for (const s of region.series) {
      parts.push(renderSeries(s));
    }
    parts.push("      </cx:plotAreaRegion>");
  } else if (pa.series) {
    for (const s of pa.series) {
      parts.push(renderSeries(s));
    }
  }
  if (pa.axis) {
    for (const axis of pa.axis) {
      parts.push(renderAxis(axis));
    }
  }
  parts.push("    </cx:plotArea>");
  return parts.join("\n");
}

function renderRawOrEmptyLayout(layout: { _rawXml?: string } | undefined, indent: string): string {
  return layout?._rawXml ? indent + layout._rawXml : `${indent}<cx:layout/>`;
}

function renderSeries(s: ChartExSeries): string {
  const parts: string[] = [];
  const attrs = [`layoutId="${s.layoutId}"`];
  if (s.hidden) {
    attrs.push('hidden="1"');
  }
  if (s.ownerIdx !== undefined) {
    attrs.push(`ownerIdx="${s.ownerIdx}"`);
  }
  parts.push(`        <cx:series ${attrs.join(" ")}>`);
  if (s.tx) {
    if (s.tx.value !== undefined) {
      parts.push(
        `          <cx:tx><cx:txData><cx:v>${escapeXml(s.tx.value)}</cx:v></cx:txData></cx:tx>`
      );
    } else if (s.tx.strRef) {
      parts.push(
        `          <cx:tx><cx:txData><cx:f>${escapeXml(s.tx.strRef)}</cx:f></cx:txData></cx:tx>`
      );
    }
  }
  if (s.spPr) {
    parts.push(renderSpPr(s.spPr, "          "));
  }
  if (s.dataRefs) {
    for (const ref of s.dataRefs) {
      if (ref.dataId !== undefined) {
        parts.push(`          <cx:dataId val="${ref.dataId}"/>`);
      }
    }
  }
  if (s.layoutPr) {
    parts.push(renderLayoutProperties(s.layoutId, s.layoutPr));
  }
  if (s.axisId) {
    for (const id of s.axisId) {
      parts.push(`          <cx:axisId val="${id}"/>`);
    }
  }
  if (s.dataLabels) {
    parts.push(renderDataLabels(s.dataLabels));
  }
  if (s.dataPt) {
    for (const dp of s.dataPt) {
      parts.push(`          <cx:dataPt idx="${dp.idx}">`);
      if (dp.spPr) {
        parts.push(renderSpPr(dp.spPr, "            "));
      }
      parts.push("          </cx:dataPt>");
    }
  }
  if (s.extLst) {
    parts.push(s.extLst);
  }
  parts.push("        </cx:series>");
  return parts.join("\n");
}

function renderLayoutProperties(
  layoutId: string,
  lp: NonNullable<ChartExSeries["layoutPr"]>
): string {
  const parts: string[] = [];
  if ((lp as any)._rawXml && !hasStructuredLayoutProperties(lp)) {
    return `          ${(lp as any)._rawXml}`;
  }
  parts.push("          <cx:layoutPr>");
  if (lp.parentLabelLayout && (layoutId === "sunburst" || layoutId === "treemap")) {
    parts.push(`            <cx:parentLabelLayout val="${lp.parentLabelLayout}"/>`);
  }
  if (lp.subtotals && layoutId === "waterfall") {
    parts.push("            <cx:subtotals>");
    for (const st of lp.subtotals) {
      parts.push(`              <cx:subtotal idx="${st.idx}"/>`);
    }
    parts.push("            </cx:subtotals>");
  }
  if (layoutId === "waterfall" && lp.connectorLines !== undefined) {
    parts.push(`            <cx:connectorLines val="${lp.connectorLines ? "1" : "0"}"/>`);
  }
  if (lp.binning) {
    const b = lp.binning;
    const attrs: string[] = [];
    if (b.intervalClosed) {
      attrs.push(`intervalClosed="${b.intervalClosed}"`);
    }
    if (b.underflow !== undefined) {
      attrs.push(`underflow="${b.underflow}"`);
    }
    if (b.overflow !== undefined) {
      attrs.push(`overflow="${b.overflow}"`);
    }
    const attrStr = attrs.length > 0 ? " " + attrs.join(" ") : "";
    parts.push(`            <cx:binning${attrStr}>`);
    if (b.binType === "auto") {
      parts.push("              <cx:auto/>");
    }
    if (b.binSize !== undefined) {
      parts.push(`              <cx:binSize val="${b.binSize}"/>`);
    }
    if (b.binCount !== undefined) {
      parts.push(`              <cx:binCount val="${b.binCount}"/>`);
    }
    if (b.binType === "categories") {
      parts.push("              <cx:categories/>");
    } else if (b.binType === "manual") {
      parts.push("              <cx:manual/>");
    }
    parts.push("            </cx:binning>");
  }
  if (lp.paretoLine) {
    parts.push('            <cx:paretoLine val="1"/>');
  }
  if (layoutId === "boxWhisker") {
    if (lp.quartileMethod) {
      parts.push(`            <cx:quartileMethod val="${lp.quartileMethod}"/>`);
    }
    if (lp.showMeanLine !== undefined) {
      parts.push(`            <cx:showMeanLine val="${lp.showMeanLine ? "1" : "0"}"/>`);
    }
    if (lp.showMeanMarker !== undefined) {
      parts.push(`            <cx:showMeanMarker val="${lp.showMeanMarker ? "1" : "0"}"/>`);
    }
    if (lp.showInnerPoints !== undefined) {
      parts.push(`            <cx:showInnerPoints val="${lp.showInnerPoints ? "1" : "0"}"/>`);
    }
    if (lp.showOutlierPoints !== undefined) {
      parts.push(`            <cx:showOutlierPoints val="${lp.showOutlierPoints ? "1" : "0"}"/>`);
    }
  }
  if (layoutId === "regionMap") {
    if (lp.projection) {
      parts.push(`            <cx:projection val="${lp.projection}"/>`);
    }
    if (lp.regionLabels) {
      parts.push(`            <cx:regionLabels val="${lp.regionLabels}"/>`);
    }
    if (lp.geoMappingLevel) {
      parts.push(`            <cx:geoMappingLevel val="${lp.geoMappingLevel}"/>`);
    }
  }
  if (lp.extLst) {
    parts.push(`            ${lp.extLst}`);
  }
  parts.push("          </cx:layoutPr>");
  return parts.join("\n");
}

function hasStructuredLayoutProperties(lp: NonNullable<ChartExSeries["layoutPr"]>): boolean {
  return [
    lp.parentLabelLayout,
    lp.subtotals,
    lp.connectorLines,
    lp.increaseSpPr,
    lp.decreaseSpPr,
    lp.totalSpPr,
    lp.binning,
    lp.paretoLine,
    lp.quartileMethod,
    lp.showMeanLine,
    lp.showMeanMarker,
    lp.showInnerPoints,
    lp.showOutlierPoints,
    lp.projection,
    lp.regionLabels,
    lp.geoMappingLevel
  ].some(value => value !== undefined);
}

function renderDataLabels(dl: NonNullable<ChartExSeries["dataLabels"]>): string {
  const parts: string[] = [];
  parts.push("          <cx:dataLabels>");
  if (dl.visibility) {
    const v = dl.visibility;
    const attrs: string[] = [];
    if (v.seriesName !== undefined) {
      attrs.push(`seriesName="${v.seriesName ? "1" : "0"}"`);
    }
    if (v.categoryName !== undefined) {
      attrs.push(`categoryName="${v.categoryName ? "1" : "0"}"`);
    }
    if (v.value !== undefined) {
      attrs.push(`value="${v.value ? "1" : "0"}"`);
    }
    if (v.numFmt !== undefined) {
      attrs.push(`numFmt="${v.numFmt ? "1" : "0"}"`);
    }
    parts.push(`            <cx:visibility ${attrs.join(" ")}/>`);
  }
  if (dl.position) {
    parts.push(`            <cx:dataLabel pos="${dl.position}"/>`);
  }
  if (dl.separator) {
    parts.push(`            <cx:separator>${escapeXml(dl.separator)}</cx:separator>`);
  }
  if (dl.numFmt) {
    parts.push(`            <cx:numFmt formatCode="${escapeAttr(dl.numFmt)}"/>`);
  }
  if (dl.spPr) {
    parts.push(renderSpPr(dl.spPr, "            "));
  }
  if (dl.txPr) {
    parts.push(renderTxPr(dl.txPr, "            ", "cx:txPr"));
  }
  parts.push("          </cx:dataLabels>");
  return parts.join("\n");
}

function renderAxis(axis: ChartExAxis): string {
  const parts: string[] = [];
  parts.push(`      <cx:axis id="${axis.axisId}">`);
  if (axis.hidden) {
    parts.push('        <cx:hidden val="1"/>');
  }
  if (axis.majorTickMark) {
    parts.push(`        <cx:majorTickMark val="${axis.majorTickMark}"/>`);
  }
  if (axis.minorTickMark) {
    parts.push(`        <cx:minorTickMark val="${axis.minorTickMark}"/>`);
  }
  if (axis.numFmt) {
    const attrs = [`formatCode="${escapeAttr(axis.numFmt.formatCode)}"`];
    if (axis.numFmt.sourceLinked !== undefined) {
      attrs.push(`sourceLinked="${axis.numFmt.sourceLinked ? "1" : "0"}"`);
    }
    parts.push(`        <cx:numFmt ${attrs.join(" ")}/>`);
  }
  if (axis.title) {
    parts.push(renderTitle(axis.title));
  }
  if (axis.valScaling) {
    const vs = axis.valScaling;
    const attrs: string[] = [];
    if (vs.min !== undefined) {
      attrs.push(`min="${vs.min}"`);
    }
    if (vs.max !== undefined) {
      attrs.push(`max="${vs.max}"`);
    }
    if (vs.majorUnit !== undefined) {
      attrs.push(`majorUnit="${vs.majorUnit}"`);
    }
    if (vs.minorUnit !== undefined) {
      attrs.push(`minorUnit="${vs.minorUnit}"`);
    }
    parts.push(`        <cx:valScaling ${attrs.join(" ")}/>`);
  }
  if (axis.catScaling) {
    const cs = axis.catScaling;
    const attrs: string[] = [];
    if (cs.gapWidth !== undefined) {
      attrs.push(`gapWidth="${cs.gapWidth}"`);
    }
    parts.push(`        <cx:catScaling ${attrs.join(" ")}/>`);
  }
  if (axis.spPr) {
    parts.push(renderSpPr(axis.spPr, "        "));
  }
  if (axis.txPr) {
    parts.push(renderTxPr(axis.txPr, "        ", "cx:txPr"));
  }
  if (axis.extLst) {
    parts.push(`        ${axis.extLst}`);
  }
  parts.push("      </cx:axis>");
  return parts.join("\n");
}

function renderLegend(l: NonNullable<ChartExModel["chartSpace"]["chart"]["legend"]>): string {
  const parts: string[] = [];
  const attrs: string[] = [];
  if (l.legendPos) {
    attrs.push(`pos="${l.legendPos}"`);
  }
  if (l.overlay !== undefined) {
    attrs.push(`align="ctr" overlay="${l.overlay ? "1" : "0"}"`);
  }
  const hasChildren = !!(l.spPr || l.legendEntries);
  if (hasChildren) {
    parts.push(`    <cx:legend ${attrs.join(" ")}>`);
    if (l.spPr) {
      parts.push(renderSpPr(l.spPr, "      "));
    }
    if (l.legendEntries) {
      for (const entry of l.legendEntries) {
        parts.push(`      <cx:legendEntry idx="${entry.index}"/>`);
      }
    }
    parts.push("    </cx:legend>");
  } else {
    parts.push(`    <cx:legend ${attrs.join(" ")}/>`);
  }
  return parts.join("\n");
}

function renderSpPr(spPr: ShapeProperties, indent: string): string {
  // When the shape still carries the raw XML captured at parse time AND no
  // structured field has been re-assigned (mutation APIs in
  // `shape-properties.ts` drop `_rawXml` on `buildSpPr` / `setSpPrFill` /
  // `setSpPrLine`), emit the original bytes verbatim. This preserves
  // DrawingML elements that this renderer does not reconstruct
  // structurally (e.g. `a:xfrm`, `a:prstGeom`, `a:custGeom`, `a:ln`
  // joints/compounds the writer below doesn't cover, future extensions).
  if (spPr._rawXml) {
    return indent + spPr._rawXml;
  }
  const parts: string[] = [];
  parts.push(`${indent}<cx:spPr>`);
  if (spPr.fill) {
    if (spPr.fill.noFill) {
      parts.push(`${indent}  <a:noFill/>`);
    } else if (spPr.fill.solid) {
      parts.push(`${indent}  <a:solidFill>${renderColor(spPr.fill.solid)}</a:solidFill>`);
    } else if (spPr.fill.gradient) {
      const g = spPr.fill.gradient;
      if (g.stops.length >= 2) {
        parts.push(`${indent}  <a:gradFill>`);
        parts.push(`${indent}    <a:gsLst>`);
        for (const stop of g.stops) {
          parts.push(
            `${indent}      <a:gs pos="${Math.round(stop.position * 1000)}">${renderColor(stop.color)}</a:gs>`
          );
        }
        parts.push(`${indent}    </a:gsLst>`);
        if (g.type === "linear" || g.type === undefined) {
          parts.push(`${indent}    <a:lin ang="${(g.angle ?? 0) * 60000}" scaled="1"/>`);
        } else {
          parts.push(
            `${indent}    <a:path path="${g.type}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`
          );
        }
        parts.push(`${indent}  </a:gradFill>`);
      }
    } else if (spPr.fill.pattern) {
      const p = spPr.fill.pattern;
      parts.push(`${indent}  <a:pattFill prst="${p.preset}">`);
      if (p.foreground) {
        parts.push(`${indent}    <a:fgClr>${renderColor(p.foreground)}</a:fgClr>`);
      }
      if (p.background) {
        parts.push(`${indent}    <a:bgClr>${renderColor(p.background)}</a:bgClr>`);
      }
      parts.push(`${indent}  </a:pattFill>`);
    }
  }
  if (spPr.line) {
    const lnAttrs: string[] = [];
    if (spPr.line.width) {
      lnAttrs.push(`w="${spPr.line.width}"`);
    }
    if (spPr.line.cap) {
      lnAttrs.push(`cap="${spPr.line.cap}"`);
    }
    if (spPr.line.compound) {
      lnAttrs.push(`cmpd="${spPr.line.compound}"`);
    }
    const attrStr = lnAttrs.length > 0 ? ` ${lnAttrs.join(" ")}` : "";
    // DrawingML order: fill child first (noFill/solidFill), then dash, then
    // join. Matches `_renderSpPr` in chart-space-xform.ts:2101-2132.
    const lnChildren: string[] = [];
    if (spPr.line.noFill) {
      lnChildren.push(`<a:noFill/>`);
    } else if (spPr.line.color) {
      lnChildren.push(`<a:solidFill>${renderColor(spPr.line.color)}</a:solidFill>`);
    }
    if (spPr.line.dash) {
      lnChildren.push(`<a:prstDash val="${spPr.line.dash}"/>`);
    }
    if (spPr.line.join === "round") {
      lnChildren.push(`<a:round/>`);
    } else if (spPr.line.join === "bevel") {
      lnChildren.push(`<a:bevel/>`);
    } else if (spPr.line.join === "miter") {
      lnChildren.push(`<a:miter/>`);
    }
    if (lnChildren.length === 0) {
      parts.push(`${indent}  <a:ln${attrStr}/>`);
    } else {
      parts.push(`${indent}  <a:ln${attrStr}>${lnChildren.join("")}</a:ln>`);
    }
  }
  if (spPr.effectList) {
    renderEffectList(parts, spPr.effectList, `${indent}  `);
  }
  if (spPr.scene3d) {
    renderScene3D(parts, spPr.scene3d, `${indent}  `);
  }
  if (spPr.sp3d) {
    renderSp3D(parts, spPr.sp3d, `${indent}  `);
  }
  parts.push(`${indent}</cx:spPr>`);
  return parts.join("\n");
}

function renderEffectList(parts: string[], effects: EffectList, indent: string): void {
  parts.push(`${indent}<a:effectLst>`);
  if (effects.blur) {
    const attrs: string[] = [];
    if (effects.blur.radius !== undefined) {
      attrs.push(`rad="${effects.blur.radius}"`);
    }
    if (effects.blur.grow) {
      attrs.push(`grow="1"`);
    }
    parts.push(`${indent}  <a:blur${attrs.length > 0 ? " " + attrs.join(" ") : ""}/>`);
  }
  if (effects.outerShadow) {
    renderShadowElement(parts, "a:outerShdw", effects.outerShadow, `${indent}  `);
  }
  if (effects.innerShadow) {
    renderShadowElement(parts, "a:innerShdw", effects.innerShadow, `${indent}  `);
  }
  if (effects.presetShadow) {
    const ps = effects.presetShadow;
    const attrs: string[] = [`prst="${ps.preset}"`];
    if (ps.distance !== undefined) {
      attrs.push(`dist="${ps.distance}"`);
    }
    if (ps.direction !== undefined) {
      attrs.push(`dir="${ps.direction}"`);
    }
    if (ps.color) {
      parts.push(`${indent}  <a:prstShdw ${attrs.join(" ")}>${renderColor(ps.color)}</a:prstShdw>`);
    } else {
      parts.push(`${indent}  <a:prstShdw ${attrs.join(" ")}/>`);
    }
  }
  if (effects.glow) {
    parts.push(
      `${indent}  <a:glow rad="${effects.glow.radius}">${renderColor(effects.glow.color)}</a:glow>`
    );
  }
  if (effects.softEdge) {
    parts.push(`${indent}  <a:softEdge rad="${effects.softEdge.radius}"/>`);
  }
  if (effects.reflection) {
    const r = effects.reflection;
    const attrs: string[] = [];
    if (r.blurRadius !== undefined) {
      attrs.push(`blurRad="${r.blurRadius}"`);
    }
    if (r.startOpacity !== undefined) {
      attrs.push(`stA="${r.startOpacity}"`);
    }
    if (r.startPosition !== undefined) {
      attrs.push(`stPos="${r.startPosition}"`);
    }
    if (r.endOpacity !== undefined) {
      attrs.push(`endA="${r.endOpacity}"`);
    }
    if (r.endPosition !== undefined) {
      attrs.push(`endPos="${r.endPosition}"`);
    }
    if (r.distance !== undefined) {
      attrs.push(`dist="${r.distance}"`);
    }
    if (r.direction !== undefined) {
      attrs.push(`dir="${r.direction}"`);
    }
    if (r.fadeDirection !== undefined) {
      attrs.push(`fadeDir="${r.fadeDirection}"`);
    }
    if (r.scaleHorizontal !== undefined) {
      attrs.push(`sx="${r.scaleHorizontal}"`);
    }
    if (r.scaleVertical !== undefined) {
      attrs.push(`sy="${r.scaleVertical}"`);
    }
    if (r.skewHorizontal !== undefined) {
      attrs.push(`kx="${r.skewHorizontal}"`);
    }
    if (r.skewVertical !== undefined) {
      attrs.push(`ky="${r.skewVertical}"`);
    }
    if (r.alignment) {
      attrs.push(`algn="${r.alignment}"`);
    }
    if (r.rotateWithShape) {
      attrs.push(`rotWithShape="1"`);
    }
    parts.push(`${indent}  <a:reflection${attrs.length > 0 ? " " + attrs.join(" ") : ""}/>`);
  }
  parts.push(`${indent}</a:effectLst>`);
}

function renderShadowElement(parts: string[], tag: string, shadow: Shadow, indent: string): void {
  const attrs: string[] = [];
  if (shadow.blurRadius !== undefined) {
    attrs.push(`blurRad="${shadow.blurRadius}"`);
  }
  if (shadow.distance !== undefined) {
    attrs.push(`dist="${shadow.distance}"`);
  }
  if (shadow.direction !== undefined) {
    attrs.push(`dir="${shadow.direction}"`);
  }
  if (shadow.alignment) {
    attrs.push(`algn="${shadow.alignment}"`);
  }
  const attrStr = attrs.length > 0 ? " " + attrs.join(" ") : "";
  if (shadow.color) {
    parts.push(`${indent}<${tag}${attrStr}>${renderColor(shadow.color)}</${tag}>`);
  } else {
    parts.push(`${indent}<${tag}${attrStr}/>`);
  }
}

function renderScene3D(parts: string[], scene: Scene3D, indent: string): void {
  parts.push(`${indent}<a:scene3d>`);
  if (scene.camera) {
    const c = scene.camera;
    const camAttrs: string[] = [`prst="${c.preset}"`];
    if (c.fov !== undefined) {
      camAttrs.push(`fov="${c.fov}"`);
    }
    if (c.zoom !== undefined) {
      camAttrs.push(`zoom="${c.zoom}"`);
    }
    if (c.rotation) {
      const r = c.rotation;
      parts.push(`${indent}  <a:camera ${camAttrs.join(" ")}>`);
      parts.push(`${indent}    <a:rot lat="${r.lat}" lon="${r.lon}" rev="${r.rev}"/>`);
      parts.push(`${indent}  </a:camera>`);
    } else {
      parts.push(`${indent}  <a:camera ${camAttrs.join(" ")}/>`);
    }
  }
  if (scene.lightRig) {
    const l = scene.lightRig;
    const lightAttrs: string[] = [`rig="${l.rig}"`, `dir="${l.direction}"`];
    if (l.rotation) {
      const r = l.rotation;
      parts.push(`${indent}  <a:lightRig ${lightAttrs.join(" ")}>`);
      parts.push(`${indent}    <a:rot lat="${r.lat}" lon="${r.lon}" rev="${r.rev}"/>`);
      parts.push(`${indent}  </a:lightRig>`);
    } else {
      parts.push(`${indent}  <a:lightRig ${lightAttrs.join(" ")}/>`);
    }
  }
  if (scene.backdrop) {
    parts.push(
      `${indent}  <a:backdrop><!-- backdrop preserved verbatim from source --></a:backdrop>`
    );
  }
  parts.push(`${indent}</a:scene3d>`);
}

function renderSp3D(parts: string[], sp: ShapeProperties3D, indent: string): void {
  const attrs: string[] = [];
  if (sp.extrusionHeight !== undefined) {
    attrs.push(`extrusionH="${sp.extrusionHeight}"`);
  }
  if (sp.contourWidth !== undefined) {
    attrs.push(`contourW="${sp.contourWidth}"`);
  }
  if (sp.material) {
    attrs.push(`prstMaterial="${sp.material}"`);
  }
  const attrStr = attrs.length > 0 ? " " + attrs.join(" ") : "";
  const children: string[] = [];
  if (sp.bevelTop) {
    const b = sp.bevelTop;
    const bAttrs = [`w="${b.width ?? 0}"`, `h="${b.height ?? 0}"`];
    if (b.preset) {
      bAttrs.push(`prst="${b.preset}"`);
    }
    children.push(`<a:bevelT ${bAttrs.join(" ")}/>`);
  }
  if (sp.bevelBottom) {
    const b = sp.bevelBottom;
    const bAttrs = [`w="${b.width ?? 0}"`, `h="${b.height ?? 0}"`];
    if (b.preset) {
      bAttrs.push(`prst="${b.preset}"`);
    }
    children.push(`<a:bevelB ${bAttrs.join(" ")}/>`);
  }
  if (sp.extrusionColor) {
    children.push(`<a:extrusionClr>${renderColor(sp.extrusionColor)}</a:extrusionClr>`);
  }
  if (sp.contourColor) {
    children.push(`<a:contourClr>${renderColor(sp.contourColor)}</a:contourClr>`);
  }
  if (children.length === 0) {
    parts.push(`${indent}<a:sp3d${attrStr}/>`);
  } else {
    parts.push(`${indent}<a:sp3d${attrStr}>${children.join("")}</a:sp3d>`);
  }
}

/**
 * Render a `ChartTextProperties` as `<cx:txPr>` (or `<c:txPr>` if caller
 * overrides `wrapperName`). Follows the same "raw wins if not mutated"
 * convention as {@link renderSpPr}: when the structured object still
 * carries its original `_rawXml`, emit the captured bytes so extended
 * paragraph properties, run markup, list styles, and schemeClr references
 * survive round-trip unchanged. When a caller mutates via
 * `buildTxPr()`/`setTxPr*()` in shape-properties.ts, `_rawXml` is dropped
 * and we rebuild a minimal structured equivalent.
 */
function renderTxPr(txPr: ChartTextProperties, indent: string, wrapperName = "cx:txPr"): string {
  if (txPr._rawXml) {
    // Swap the outer wrapper to the requested namespace (chartEx uses
    // cx:txPr, classic charts use c:txPr). The raw bytes captured at
    // parse time always start with the namespace they came from.
    if (wrapperName === "cx:txPr" && txPr._rawXml.startsWith("<c:txPr")) {
      const rewrapped = txPr._rawXml
        .replace(/^<c:txPr/, "<cx:txPr")
        .replace(/<\/c:txPr>$/, "</cx:txPr>");
      return indent + rewrapped;
    }
    return indent + txPr._rawXml;
  }
  const rPrAttrs: string[] = [];
  if (txPr.size !== undefined) {
    rPrAttrs.push(`sz="${txPr.size}"`);
  }
  if (txPr.bold) {
    rPrAttrs.push(`b="1"`);
  }
  if (txPr.italic) {
    rPrAttrs.push(`i="1"`);
  }
  const rPrChildren: string[] = [];
  if (txPr.color) {
    rPrChildren.push(`<a:solidFill>${renderColor(txPr.color)}</a:solidFill>`);
  }
  if (txPr.fontFamily) {
    rPrChildren.push(`<a:latin typeface="${escapeAttr(txPr.fontFamily)}"/>`);
    rPrChildren.push(`<a:cs typeface="${escapeAttr(txPr.fontFamily)}"/>`);
  }
  const rPrAttrStr = rPrAttrs.length > 0 ? " " + rPrAttrs.join(" ") : "";
  const defRPr =
    rPrChildren.length > 0
      ? `<a:defRPr${rPrAttrStr}>${rPrChildren.join("")}</a:defRPr>`
      : `<a:defRPr${rPrAttrStr}/>`;
  const bodyPrAttrs = txPr.rotation !== undefined ? ` rot="${txPr.rotation}"` : "";
  return [
    `${indent}<${wrapperName}>`,
    `${indent}  <a:bodyPr${bodyPrAttrs}/>`,
    `${indent}  <a:lstStyle/>`,
    `${indent}  <a:p>`,
    `${indent}    <a:pPr>${defRPr}</a:pPr>`,
    `${indent}    <a:endParaRPr/>`,
    `${indent}  </a:p>`,
    `${indent}</${wrapperName}>`
  ].join("\n");
}

function renderColor(c: ChartColor): string {
  const modifiers = renderColorModifiers(c);
  if (c.srgb) {
    if (modifiers) {
      return `<a:srgbClr val="${c.srgb}">${modifiers}</a:srgbClr>`;
    }
    return `<a:srgbClr val="${c.srgb}"/>`;
  }
  if (c.theme !== undefined) {
    const themeNames = [
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
    const name = themeNames[c.theme] ?? "dk1";
    if (modifiers) {
      return `<a:schemeClr val="${name}">${modifiers}</a:schemeClr>`;
    }
    return `<a:schemeClr val="${name}"/>`;
  }
  if (c.sysClr) {
    if (modifiers) {
      return `<a:sysClr val="${c.sysClr}">${modifiers}</a:sysClr>`;
    }
    return `<a:sysClr val="${c.sysClr}"/>`;
  }
  if (c.prstClr) {
    if (modifiers) {
      return `<a:prstClr val="${c.prstClr}">${modifiers}</a:prstClr>`;
    }
    return `<a:prstClr val="${c.prstClr}"/>`;
  }
  return "";
}

function renderColorModifiers(c: ChartColor): string {
  const parts: string[] = [];
  if (c.alpha !== undefined) {
    parts.push(`<a:alpha val="${c.alpha}"/>`);
  }
  if (c.tint !== undefined) {
    parts.push(`<a:tint val="${Math.round(c.tint * 100000)}"/>`);
  }
  if (c.shade !== undefined) {
    parts.push(`<a:shade val="${c.shade}"/>`);
  }
  if (c.satMod !== undefined) {
    parts.push(`<a:satMod val="${c.satMod}"/>`);
  }
  if (c.lumMod !== undefined) {
    parts.push(`<a:lumMod val="${c.lumMod}"/>`);
  }
  if (c.lumOff !== undefined) {
    parts.push(`<a:lumOff val="${c.lumOff}"/>`);
  }
  return parts.join("");
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
