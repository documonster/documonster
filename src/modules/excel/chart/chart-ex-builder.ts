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
  validateChartExOptions(opts);
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
        spPr: opts.spPr,
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
    },
    style: opts.chartStyle,
    colors: opts.chartColors
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
  } else if (so.literalCategories?.length) {
    const literalCatId = counter.value++;
    data.push({
      id: literalCatId,
      strDim: {
        type: "cat",
        levels: [
          {
            ptCount: so.literalCategories.length,
            points: so.literalCategories.map((value, index) => ({ index, value }))
          }
        ]
      }
    });
    series.dataRefs!.push({ dataId: literalCatId });
  }

  // Values — always create a fresh numDim entry.
  const valId = counter.value++;
  const numDimType = opts.type === "histogram" || opts.type === "pareto" ? "x" : "val";
  data.push({
    id: valId,
    numDim: {
      type: numDimType,
      formula: so.values,
      levels: so.literalValues?.length
        ? [
            {
              ptCount: so.literalValues.length,
              points: so.literalValues.map((value, index) => ({ index, value }))
            }
          ]
        : [{ ptCount: 0, points: [] }]
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

  // Literal hierarchy levels for headless sunburst/treemap previews.
  if (so.literalHierarchy && (layoutId === "sunburst" || layoutId === "treemap")) {
    for (const level of so.literalHierarchy) {
      const hId = counter.value++;
      data.push({
        id: hId,
        strDim: {
          type: "cat",
          levels: [
            {
              ptCount: level.length,
              points: level.map((value, index) => ({ index, value }))
            }
          ]
        }
      });
      series.dataRefs!.push({ dataId: hId });
    }
  }

  // Waterfall subtotals
  const subtotals = so.subtotals ?? so.subtotalPoints?.map(p => p.idx);
  if (subtotals && layoutId === "waterfall") {
    series.layoutPr = {
      subtotals: subtotals.map(i => ({ idx: i }))
    };
  }

  // Apply chart-level layout options to this series when appropriate
  if (opts.layout) {
    series.layoutPr = { ...(series.layoutPr ?? {}), ...opts.layout };
  }
  if (opts.binning) {
    series.layoutPr = { ...(series.layoutPr ?? {}), binning: opts.binning };
  }
  if ((opts.type === "histogram" || opts.type === "pareto") && !series.layoutPr?.binning) {
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      binning: { binType: "auto" }
    };
  }
  if (opts.type === "pareto") {
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      paretoLine: true
    };
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

function validateChartExOptions(opts: AddChartExOptions): void {
  if (!opts || typeof opts !== "object") {
    throw new Error("chartEx options are required");
  }
  if (!opts.type) {
    throw new Error("chartEx.type is required");
  }
  if (!Array.isArray(opts.series) || opts.series.length === 0) {
    throw new Error("chartEx.series must contain at least one series");
  }
  if (opts.binning && opts.type !== "histogram" && opts.type !== "pareto") {
    throw new Error(
      `chartEx.binning is only valid for histogram and pareto charts, not ${opts.type}`
    );
  }
  validateLayoutOptions(opts);
  const binning = opts.binning ?? opts.layout?.binning;
  if (binning) {
    validateBinning(binning);
  }
  opts.series.forEach((series, i) => {
    if (!series.values) {
      throw new Error(`chartEx.series[${i}].values is required`);
    }
    if (series.literalValues?.some(value => !Number.isFinite(value))) {
      throw new Error(`chartEx.series[${i}].literalValues must contain only finite numbers`);
    }
    if (
      series.literalCategories &&
      series.literalValues &&
      series.literalCategories.length !== series.literalValues.length
    ) {
      throw new Error(
        `chartEx.series[${i}].literalCategories length must match literalValues length`
      );
    }
    if (series.hierarchy && opts.type !== "sunburst" && opts.type !== "treemap") {
      throw new Error(
        `chartEx.series[${i}].hierarchy is only valid for sunburst and treemap charts`
      );
    }
    if (series.literalHierarchy && opts.type !== "sunburst" && opts.type !== "treemap") {
      throw new Error(
        `chartEx.series[${i}].literalHierarchy is only valid for sunburst and treemap charts`
      );
    }
    if (series.literalHierarchy && series.literalValues) {
      for (let level = 0; level < series.literalHierarchy.length; level++) {
        if (series.literalHierarchy[level].length !== series.literalValues.length) {
          throw new Error(
            `chartEx.series[${i}].literalHierarchy[${level}] length must match literalValues length`
          );
        }
      }
    }
    if ((series.subtotals || series.subtotalPoints) && opts.type !== "waterfall") {
      throw new Error(`chartEx.series[${i}].subtotals is only valid for waterfall charts`);
    }
    if (series.subtotals?.some(idx => !Number.isInteger(idx) || idx < 0)) {
      throw new Error(`chartEx.series[${i}].subtotals must contain non-negative integer indices`);
    }
    if (series.subtotalPoints?.some(point => !Number.isInteger(point.idx) || point.idx < 0)) {
      throw new Error(
        `chartEx.series[${i}].subtotalPoints must contain non-negative integer indices`
      );
    }
  });
}

function validateLayoutOptions(opts: AddChartExOptions): void {
  const layout = opts.layout;
  if (!layout) {
    return;
  }
  if (layout.binning && opts.type !== "histogram" && opts.type !== "pareto") {
    throw new Error(`chartEx.layout.binning is only valid for histogram and pareto charts`);
  }
  if (layout.subtotals && opts.type !== "waterfall") {
    throw new Error(`chartEx.layout.subtotals is only valid for waterfall charts`);
  }
  const waterfallFields = [
    layout.connectorLines,
    layout.increaseSpPr,
    layout.decreaseSpPr,
    layout.totalSpPr
  ];
  if (waterfallFields.some(value => value !== undefined) && opts.type !== "waterfall") {
    throw new Error(`chartEx.layout waterfall fields are only valid for waterfall charts`);
  }
  const hierarchyLayout = layout.parentLabelLayout !== undefined;
  if (hierarchyLayout && opts.type !== "sunburst" && opts.type !== "treemap") {
    throw new Error(
      `chartEx.layout.parentLabelLayout is only valid for sunburst and treemap charts`
    );
  }
  const boxWhiskerFields = [
    layout.quartileMethod,
    layout.showMeanLine,
    layout.showMeanMarker,
    layout.showInnerPoints,
    layout.showOutlierPoints
  ];
  if (boxWhiskerFields.some(value => value !== undefined) && opts.type !== "boxWhisker") {
    throw new Error(`chartEx.layout box-whisker fields are only valid for boxWhisker charts`);
  }
  const regionMapFields = [layout.projection, layout.regionLabels, layout.geoMappingLevel];
  if (regionMapFields.some(value => value !== undefined) && opts.type !== "regionMap") {
    throw new Error(`chartEx.layout region-map fields are only valid for regionMap charts`);
  }
  if (layout.paretoLine !== undefined && opts.type !== "pareto") {
    throw new Error(`chartEx.layout.paretoLine is only valid for pareto charts`);
  }
}

function validateBinning(binning: NonNullable<AddChartExOptions["binning"]>): void {
  if (binning.binSize !== undefined && binning.binSize <= 0) {
    throw new Error("chartEx.binning.binSize must be greater than 0");
  }
  if (
    binning.binCount !== undefined &&
    (!Number.isInteger(binning.binCount) || binning.binCount <= 0)
  ) {
    throw new Error("chartEx.binning.binCount must be a positive integer");
  }
  if (
    binning.underflow !== undefined &&
    binning.overflow !== undefined &&
    binning.underflow >= binning.overflow
  ) {
    throw new Error("chartEx.binning.underflow must be less than overflow");
  }
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
