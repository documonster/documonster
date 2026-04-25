/**
 * ChartEx builder — constructs a ChartExModel from simplified AddChartExOptions.
 *
 * Each layoutId corresponds to a distinct chart type. The builder produces a
 * structured model that the renderer serialises to `cx:chart` XML.
 */

import type {
  AddChartExOptions,
  AddChartExSeriesOptions,
  ChartExAxis,
  ChartExDataEntry,
  ChartExModel,
  ChartExSeries,
  ChartExSeriesType
} from "./chart-ex-types";

/**
 * Build a structured ChartExModel from high-level options.
 */
export function buildChartExModel(opts: AddChartExOptions): ChartExModel {
  const layoutId = mapChartTypeToLayoutId(opts.type);

  // Build data entries — one cx:data per reference the series need.
  const data: ChartExDataEntry[] = [];
  const dataIdCounter = { value: 0 };

  // Categories share a single data entry for all series (like classic charts).
  let catDataId: number | undefined;
  if (opts.categories) {
    catDataId = dataIdCounter.value++;
    data.push({
      id: catDataId,
      strDim: {
        type: "cat",
        formula: opts.categories,
        levels: [{ ptCount: 0, points: [] }]
      }
    });
  }

  // Build series — each gets its own cx:data entries for values (and any
  // extras like hierarchy levels).
  const series: ChartExSeries[] = [];
  for (let i = 0; i < opts.series.length; i++) {
    const so = opts.series[i];
    const built = buildSeriesAndData(so, i, layoutId, catDataId, data, dataIdCounter, opts);
    series.push(built);
  }

  // Build axes — needed for histogram, pareto, waterfall, funnel, boxWhisker.
  const axes: ChartExAxis[] = [];
  if (needsAxes(opts.type)) {
    axes.push({ axisId: 0, type: "cat" });
    axes.push({ axisId: 1, type: "val" });
    // Wire series to axes
    for (const s of series) {
      s.axisId = [0, 1];
    }
  }

  const model: ChartExModel = {
    chartSpace: {
      chartData: { data },
      chart: {
        title: opts.title
          ? {
              text: {
                paragraphs: [{ runs: [{ text: opts.title }] }]
              },
              overlay: false
            }
          : undefined,
        autoTitleDeleted: !opts.title,
        plotArea: {
          plotAreaRegion: {
            series
          },
          axis: axes.length > 0 ? axes : undefined
        },
        legend:
          opts.showLegend !== false
            ? { legendPos: opts.legendPosition ?? "b", overlay: false }
            : undefined
      }
    }
  };

  return model;
}

function buildSeriesAndData(
  so: AddChartExSeriesOptions,
  idx: number,
  layoutId: ChartExSeriesType,
  catDataId: number | undefined,
  data: ChartExDataEntry[],
  counter: { value: number },
  opts: AddChartExOptions
): ChartExSeries {
  const series: ChartExSeries = {
    layoutId,
    seriesIndex: idx,
    dataRefs: []
  };

  if (so.name) {
    series.tx = { value: so.name };
  }

  // Categories binding — reuse the shared catDataId.
  if (catDataId !== undefined) {
    series.dataRefs!.push({ dataId: catDataId });
  }

  // Values — always create a fresh numDim entry.
  const valId = counter.value++;
  data.push({
    id: valId,
    numDim: {
      type: "val",
      formula: so.values,
      levels: [{ ptCount: 0, points: [] }]
    }
  });
  series.dataRefs!.push({ dataId: valId });

  // Hierarchy levels for sunburst/treemap
  if (so.hierarchy && (layoutId === "sunburst" || layoutId === "treemap")) {
    for (const h of so.hierarchy) {
      const hId = counter.value++;
      data.push({
        id: hId,
        strDim: {
          type: "cat",
          formula: h,
          levels: [{ ptCount: 0, points: [] }]
        }
      });
      series.dataRefs!.push({ dataId: hId });
    }
  }

  // Waterfall subtotals
  if (so.subtotals && layoutId === "waterfall") {
    series.layoutPr = {
      subtotals: so.subtotals.map(i => ({ idx: i }))
    };
  }

  // Apply chart-level layout options to this series when appropriate
  if (opts.layout) {
    series.layoutPr = { ...(series.layoutPr ?? {}), ...opts.layout };
  }

  // Shape properties
  if (so.spPr) {
    series.spPr = so.spPr;
  } else if (so.fill || so.border) {
    series.spPr = {};
    if (so.fill) {
      series.spPr.fill = { solid: { srgb: so.fill.replace(/^#/, "").toUpperCase() } };
    }
    if (so.border) {
      series.spPr.line = { color: { srgb: so.border.replace(/^#/, "").toUpperCase() } };
    }
  }

  // Data labels
  if (so.dataLabels) {
    series.dataLabels = {
      visibility: {
        value: so.dataLabels.showValue,
        categoryName: so.dataLabels.showCategory,
        seriesName: so.dataLabels.showSeriesName
      },
      position: so.dataLabels.position,
      separator: so.dataLabels.separator,
      numFmt: so.dataLabels.numFmt
    };
  }

  return series;
}

function mapChartTypeToLayoutId(type: AddChartExOptions["type"]): ChartExSeriesType {
  switch (type) {
    case "sunburst":
      return "sunburst";
    case "treemap":
      return "treemap";
    case "waterfall":
      return "waterfall";
    case "funnel":
      return "funnel";
    case "histogram":
    case "pareto":
      return "clusteredColumn"; // histogram uses clusteredColumn layoutId with binning
    case "boxWhisker":
      return "boxWhisker";
    case "regionMap":
      return "regionMap";
    default: {
      const _never: never = type;
      throw new Error(`Unsupported chartEx type: ${String(_never)}`);
    }
  }
}

function needsAxes(type: AddChartExOptions["type"]): boolean {
  // Sunburst, treemap, funnel, regionMap don't use traditional axes;
  // histogram/pareto/waterfall/boxWhisker do.
  return type === "histogram" || type === "pareto" || type === "waterfall" || type === "boxWhisker";
}
