/**
 * ChartEx builder — constructs a ChartExModel from simplified AddChartExOptions.
 *
 * Each layoutId corresponds to a distinct chart type. The builder produces a
 * structured model that the renderer serialises to `cx:chart` XML.
 */

import { ChartOptionsError } from "@excel/errors";

import { hexToColor, toShapeProperties } from "./chart-builder";
import type {
  AddChartExOptions,
  AddChartExSeriesOptions,
  ChartExAxis,
  ChartExDataEntry,
  ChartExModel,
  ChartExSeries,
  ChartExSeriesType
} from "./chart-ex-types";
import type { ChartRichText } from "./types";

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
  } else {
    // No chart-level `categories` formula — check whether the series
    // bundle carries a `literalCategories` array we can hoist up and
    // share. Classic charts and Excel-authored ChartEx both store
    // categories as a single axis-wide entry; the previous code
    // allocated a fresh cat `cx:data` entry inside each series' branch
    // below, producing `N` parallel category dimensions for `N` series
    // and breaking the "shared category axis" invariant Excel readers
    // expect. Find the first series that carries a non-empty
    // `literalCategories` and use it as the shared entry. Subsequent
    // series that redefine `literalCategories` to a different shape
    // throw — mixing-and-matching per-series category axes is not
    // supported by Chart2014.
    const firstLiteral = opts.series.find(so => so.literalCategories?.length)?.literalCategories;
    if (firstLiteral && firstLiteral.length > 0) {
      // Validate that every other series with `literalCategories`
      // carries the same array (by deep equality). Divergence is a
      // caller bug — the schema has no way to express per-series
      // category axes in a ChartEx plot area region.
      for (const so of opts.series) {
        if (!so.literalCategories) {
          continue;
        }
        if (!arraysEqual(so.literalCategories, firstLiteral)) {
          throw new ChartOptionsError(
            "AddChartExOptions.series[*].literalCategories must be the same across all series; " +
              "ChartEx stores categories on a shared category axis (see `opts.categories` for the " +
              "formula-based alternative)."
          );
        }
      }
      catDataId = dataIdCounter.value++;
      data.push({
        id: catDataId,
        strDim: {
          type: "cat",
          levels: [
            {
              ptCount: firstLiteral.length,
              points: firstLiteral.map((value, index) => ({ index, value }))
            }
          ]
        }
      });
    }
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
    // Seed axisIds with a large base so builder-created axes don't
    // collide with numeric ids allocated by the classic `chart.xml`
    // side when a ChartEx chart is merged into a workbook alongside
    // classic charts (e.g. via `copyTo`). Classic `chart-builder.ts`
    // uses `100000000` for the same reason; ChartEx previously used
    // `0` and `1`, which would collide with any axis allocated by the
    // classic side via `chart.getOrCreateAxis(0)`.
    const CHARTEX_AXIS_ID_SEED = 100000000;
    const catAxisId = CHARTEX_AXIS_ID_SEED;
    const valAxisId = CHARTEX_AXIS_ID_SEED + 1;
    axes.push({ axisId: catAxisId, type: "cat" });
    axes.push({ axisId: valAxisId, type: "val" });
    // Wire series to axes
    for (const s of series) {
      s.axisId = [catAxisId, valAxisId];
    }
  }

  const model: ChartExModel = {
    chartSpace: {
      chartData: { data },
      // Chart-frame styling lives on `CT_ChartSpace/spPr` per
      // ECMA-376 / Chart2014 — `CT_Chart` itself has no `spPr` child.
      // Route `opts.spPr` here so the writer's schema-compliant path
      // receives it. (Users can still pass either the structured
      // `ShapeProperties` form or the ergonomic hex-colour shorthand;
      // `toShapeProperties` normalises both.)
      spPr: toShapeProperties(opts.spPr),
      chart: {
        title: buildChartExTitle(opts.title),
        // Leave `autoTitleDeleted` undefined when the caller did not
        // supply a title. Setting it to `true` whenever `opts.title`
        // is absent tells Excel "the author explicitly deleted the
        // auto-title", suppressing the automatic chart title Excel
        // normally shows for single-series charts. Pass
        // `title: null` (or an empty rich-text wrapper) to request
        // explicit suppression; omitting the option preserves Excel's
        // default auto-title behaviour.
        autoTitleDeleted: opts.title === null ? true : undefined,
        plotArea: {
          plotAreaRegion: {
            series
          },
          axis: axes.length > 0 ? axes : undefined
        },
        legend:
          opts.showLegend !== false
            ? // Leave `overlay` undefined so the writer omits the
              // attribute entirely. Chart2014 `CT_Legend/@overlay`
              // defaults to `false` when absent; emitting `overlay="0"`
              // here pollutes round-trip byte-comparison and differs
              // from what Excel produces for a fresh legend.
              { legendPos: opts.legendPosition ?? "b" }
            : undefined
      }
    },
    style: opts.chartStyle,
    colors: opts.chartColors
  };

  return model;
}

/**
 * Normalise a ChartEx title option into the structured `ChartTitle` shape.
 * Accepts the same three forms as the classic chart builder: a plain string,
 * a rich-text description, or a `{ formula: "Sheet1!$A$1" }` reference. This
 * keeps the classic / ChartEx APIs symmetric — previously ChartEx only
 * supported strings and silently ignored formula / rich-text input.
 */
function buildChartExTitle(
  input: AddChartExOptions["title"]
): ChartExModel["chartSpace"]["chart"]["title"] {
  // `null` signals "explicitly suppress auto-title" — the caller has
  // already promoted this to `autoTitleDeleted: true` at the call site,
  // so we emit no title element.
  if (input === undefined || input === null) {
    return undefined;
  }
  if (typeof input === "string") {
    return {
      text: { paragraphs: [{ runs: [{ text: input }] }] },
      overlay: false
    };
  }
  if ("formula" in input && typeof input.formula === "string") {
    return {
      // `ChartTitle.strRef` is the classic `StringReference` shape
      // (`{ formula, cache: { points: [] } }`) because ChartEx titles
      // reuse the classic `ChartTitle` type. Distinct from the ChartEx
      // series `tx.strRef` which uses `{ formula, cached? }`.
      strRef: { formula: input.formula, cache: { points: [] } },
      overlay: false
    };
  }
  if ("paragraphs" in input) {
    return {
      text: input as ChartRichText,
      overlay: false
    };
  }
  throw new ChartOptionsError(
    "chartEx.title must be a string, a { formula: string } reference, or a ChartRichText object."
  );
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
  // `idx` is the caller-facing 0-based series index within the chart.
  // The OOXML Chart2014 schema routes series ordering through
  // `<cx:dataId>` / `<cx:axisId>` references rather than a dedicated
  // series-index attribute, so there is no field to populate on the
  // model — the parser never reads it and both the renderer and the
  // writer key everything off `layoutId` / `dataRefs`. Keeping `idx`
  // as a documented parameter so the caller's intent is preserved in
  // tracing output if a future schema extension ever needs it.
  void idx;
  const series: ChartExSeries = {
    layoutId,
    dataRefs: []
  };

  if (so.name !== undefined) {
    // Dispatch on the `name` shape. Plain strings become literal
    // captions; `{ formula }` routes through `tx.strRef` so the writer
    // emits `<cx:f>` (worksheet reference); rich-text structures pass
    // through unchanged. Matches `AddChartSeriesOptions.name` handling
    // on the classic side — previously the chartEx builder only
    // accepted plain strings and silently dropped the other forms.
    if (typeof so.name === "string") {
      series.tx = { value: so.name };
    } else if ("formula" in so.name) {
      // Use the canonical object form `{ formula, cached? }` rather
      // than the legacy bare-string variant. The writer at
      // `chart-ex-renderer.ts:renderSeries` handles both, but
      // storing the object uniformly keeps the typed model
      // consistent across builder-produced and parser-produced
      // series — downstream consumers can treat `tx.strRef` as
      // always object-shaped after this round of builder work.
      series.tx = { strRef: { formula: so.name.formula } };
    } else {
      series.tx = { rich: so.name };
    }
  }

  // ------------------------------------------------------------------
  // Data ref order for sunburst / treemap
  //
  // Excel-authored sunburst and treemap charts group ALL string
  // dimensions (categories + hierarchy levels) before the numeric
  // value dimension. The previous builder order was
  //   [category, value, hierarchy...]
  // which opens in Excel but doesn't match the canonical layout —
  // some readers (Excel's own chartEx parser that wraps the charts API)
  // expect contiguous strDim refs before numDim.
  //
  // For non-hierarchical types (funnel, waterfall, histogram, pareto,
  // boxWhisker, regionMap, clusteredColumn) the order category, value
  // is already canonical, so the hierarchy branch below is a no-op.
  // ------------------------------------------------------------------
  const isHierarchical = layoutId === "sunburst" || layoutId === "treemap";

  // Categories binding. `buildChartExModel` hoists any shared
  // `literalCategories` into a single `catDataId` before calling us, so
  // by the time we're here the only cases are:
  //   1. `catDataId !== undefined`      → the chart has a shared cat axis
  //      (either from `opts.categories` or hoisted literals); reuse it.
  //   2. `catDataId === undefined`      → no category axis; omit the ref.
  // The previous fallback that allocated a fresh per-series `cx:data`
  // entry from `so.literalCategories` has been removed — it produced
  // N duplicated cat axes for N series, breaking the shared-axis
  // invariant. Any literal-category hoisting / validation is done in
  // `buildChartExModel`.
  if (catDataId !== undefined) {
    series.dataRefs!.push({ dataId: catDataId });
  }

  // For sunburst / treemap, emit hierarchy levels BEFORE the numeric
  // value dimension (string dimensions contiguous). For other types,
  // emit the value dimension now.
  if (isHierarchical) {
    if (so.hierarchy) {
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
    if (so.literalHierarchy) {
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

  // Waterfall subtotals. `subtotalPoints` is an alternative spelling
  // accepted by the public API; we normalise both into the
  // `layoutPr.subtotals` array expected by the renderer. Any additional
  // fields present on `subtotalPoints[]` (beyond `idx`) are currently not
  // preserved because the structured model only carries the index — we
  // surface this limitation clearly rather than letting extra metadata
  // disappear silently.
  //
  // When both `subtotals` and `subtotalPoints` are provided, merge
  // them (deduped by `idx`) rather than silently dropping one.
  // Previously `subtotalPoints` was ignored entirely when `subtotals`
  // was set — users mixing the two forms lost state with no
  // diagnostic.
  const subtotalPoints = so.subtotalPoints;
  const mergedSubtotals: number[] | undefined = (() => {
    if (so.subtotals && subtotalPoints) {
      const indexSet = new Set<number>([...so.subtotals, ...subtotalPoints.map(p => p.idx)]);
      return Array.from(indexSet).sort((a, b) => a - b);
    }
    return so.subtotals ?? subtotalPoints?.map(p => p.idx);
  })();
  if (mergedSubtotals && layoutId === "waterfall") {
    series.layoutPr = {
      subtotals: mergedSubtotals.map(i => ({ idx: i }))
    };
  }

  // Apply chart-level layout options to this series when appropriate.
  // Note: `{ ...a, ...b }` does **not** skip `undefined` values on `b`,
  // which means a user passing `opts.layout = { subtotals: undefined }`
  // would wipe out an earlier `series.layoutPr.subtotals` set from
  // `so.subtotals`. Use `mergeDefined` so only keys with defined values
  // overwrite the base.
  if (opts.layout) {
    series.layoutPr = mergeDefined(series.layoutPr ?? {}, opts.layout);
  }
  if (opts.binning) {
    // Merge user-supplied binning on top of anything the `opts.layout`
    // branch above carried through (e.g. `layout.binning.intervalClosed`).
    // The previous spread-replace (`{ ...existing, binning: opts.binning }`)
    // discarded every field the user didn't re-specify, so
    // `{ layout: { binning: { intervalClosed: "r" } }, binning: { binType: "manual" } }`
    // silently dropped `intervalClosed`.
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      binning: mergeDefined(series.layoutPr?.binning ?? {}, opts.binning)
    };
  }
  if ((opts.type === "histogram" || opts.type === "pareto") && !series.layoutPr?.binning) {
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      binning: { binType: "auto" }
    };
  }
  if (opts.type === "pareto") {
    // Only default `paretoLine` to `true` when the caller hasn't
    // explicitly set it. A user passing `layout: { paretoLine: false }`
    // (valid per the validator) was previously overridden here,
    // re-enabling the cumulative-percent line against explicit intent.
    if (series.layoutPr?.paretoLine === undefined) {
      series.layoutPr = {
        ...(series.layoutPr ?? {}),
        paretoLine: true
      };
    }
  }

  // Shape properties. `spPr` takes precedence over the `fill` / `border`
  // hex shortcuts. Both inputs accept either a structured `ShapeProperties`
  // or an `AddShapeFillOptions` bundle (normalised by `toShapeProperties`).
  const normalisedSpPr = toShapeProperties(so.spPr);
  if (normalisedSpPr) {
    series.spPr = normalisedSpPr;
  } else if (so.fill || so.border) {
    series.spPr = {};
    if (so.fill) {
      series.spPr.fill = { solid: hexToColor(so.fill) };
    }
    if (so.border) {
      // Default to 9525 EMU (0.75pt) — Excel's default line width
      // for chart series borders. Without an explicit `w`, DrawingML
      // readers treat `<a:ln>` as hairline (effectively invisible
      // on screen, sometimes 1px on print). Users setting `border:
      // "#FF0000"` expect a visible red border, not a transparent
      // colour on an invisible line.
      series.spPr.line = {
        color: hexToColor(so.border),
        width: 9525
      };
    }
  }

  // Data labels
  if (so.dataLabels) {
    const dl = so.dataLabels;
    series.dataLabels = {
      visibility: {
        value: dl.showValue,
        categoryName: dl.showCategory,
        seriesName: dl.showSeriesName,
        // Map the public `showNumFmt` flag onto the internal
        // visibility slot so readers can distinguish "render the
        // number with its format" from "no numeric display at all".
        // Previously the field had no way to be set from the public
        // builder — the internal model exposed it, but the public
        // options dropped it silently.
        numFmt: dl.showNumFmt
      },
      position: dl.position,
      separator: dl.separator,
      numFmt: dl.numFmt,
      // `spPr` / `txPr` route through to the rendered `<cx:spPr>` /
      // `<cx:txPr>` children. Accept both the structured
      // `ShapeProperties` and the ergonomic `AddShapeFillOptions`
      // bundle via `toShapeProperties`, matching the classic
      // `buildDataLabelsFromOpts` helper.
      spPr: toShapeProperties(dl.spPr),
      txPr: dl.txPr
    };
  }

  return series;
}

/**
 * Merge `patch` onto `base`, ignoring keys on `patch` whose value is
 * `undefined`. JavaScript's native object spread (`{ ...a, ...b }`) does
 * NOT skip `undefined` values on `b`, so `{ x: 1, ...{ x: undefined } }`
 * produces `{ x: undefined }` — unintuitive and a frequent source of
 * layering bugs. Use this helper whenever merging option objects where
 * "undefined means don't change" is the expected semantics.
 */
function mergeDefined<T extends object>(base: T, patch: Partial<T>): T {
  const result = { ...base } as { [K in keyof T]: T[K] };
  for (const [key, value] of Object.entries(patch)) {
    if (value !== undefined) {
      (result as Record<string, unknown>)[key] = value;
    }
  }
  return result;
}

/**
 * Shallow element-wise string array equality. Used to verify that
 * `literalCategories` arrays declared on multiple series in the same
 * ChartEx bundle are identical, so the builder can safely hoist them
 * into a single shared category `cx:data` entry.
 */
function arraysEqual(a: readonly string[], b: readonly string[]): boolean {
  if (a === b) {
    return true;
  }
  if (a.length !== b.length) {
    return false;
  }
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      return false;
    }
  }
  return true;
}

function validateChartExOptions(opts: AddChartExOptions): void {
  if (!opts || typeof opts !== "object") {
    throw new ChartOptionsError("chartEx options are required.");
  }
  if (!opts.type) {
    throw new ChartOptionsError("chartEx.type is required.");
  }
  if (!Array.isArray(opts.series) || opts.series.length === 0) {
    throw new ChartOptionsError("chartEx.series must contain at least one series.");
  }
  if (opts.binning && opts.type !== "histogram" && opts.type !== "pareto") {
    throw new ChartOptionsError(
      `chartEx.binning is only valid for histogram and pareto charts, not ${opts.type}.`
    );
  }
  validateLayoutOptions(opts);
  const binning = opts.binning ?? opts.layout?.binning;
  if (binning) {
    validateBinning(binning);
  }
  opts.series.forEach((series, i) => {
    if (!series.values) {
      throw new ChartOptionsError(`chartEx.series[${i}].values is required.`);
    }
    if (series.literalValues?.some(value => !Number.isFinite(value))) {
      throw new ChartOptionsError(
        `chartEx.series[${i}].literalValues must contain only finite numbers.`
      );
    }
    if (
      series.literalCategories &&
      series.literalValues &&
      series.literalCategories.length !== series.literalValues.length
    ) {
      throw new ChartOptionsError(
        `chartEx.series[${i}].literalCategories length must match literalValues length.`
      );
    }
    if (series.hierarchy && opts.type !== "sunburst" && opts.type !== "treemap") {
      throw new ChartOptionsError(
        `chartEx.series[${i}].hierarchy is only valid for sunburst and treemap charts.`
      );
    }
    if (series.literalHierarchy && opts.type !== "sunburst" && opts.type !== "treemap") {
      throw new ChartOptionsError(
        `chartEx.series[${i}].literalHierarchy is only valid for sunburst and treemap charts.`
      );
    }
    if (series.literalHierarchy && series.literalValues) {
      for (let level = 0; level < series.literalHierarchy.length; level++) {
        if (series.literalHierarchy[level].length !== series.literalValues.length) {
          throw new ChartOptionsError(
            `chartEx.series[${i}].literalHierarchy[${level}] length must match literalValues length.`
          );
        }
      }
    }
    if ((series.subtotals || series.subtotalPoints) && opts.type !== "waterfall") {
      throw new ChartOptionsError(
        `chartEx.series[${i}].subtotals is only valid for waterfall charts.`
      );
    }
    if (series.subtotals?.some(idx => !Number.isInteger(idx) || idx < 0)) {
      throw new ChartOptionsError(
        `chartEx.series[${i}].subtotals must contain non-negative integer indices.`
      );
    }
    if (series.subtotalPoints?.some(point => !Number.isInteger(point.idx) || point.idx < 0)) {
      throw new ChartOptionsError(
        `chartEx.series[${i}].subtotalPoints must contain non-negative integer indices.`
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
    throw new ChartOptionsError(
      `chartEx.layout.binning is only valid for histogram and pareto charts.`
    );
  }
  if (layout.subtotals && opts.type !== "waterfall") {
    throw new ChartOptionsError(`chartEx.layout.subtotals is only valid for waterfall charts.`);
  }
  const waterfallFields = [
    layout.connectorLines,
    layout.increaseSpPr,
    layout.decreaseSpPr,
    layout.totalSpPr
  ];
  if (waterfallFields.some(value => value !== undefined) && opts.type !== "waterfall") {
    throw new ChartOptionsError(
      `chartEx.layout waterfall fields are only valid for waterfall charts.`
    );
  }
  const hierarchyLayout = layout.parentLabelLayout !== undefined;
  if (hierarchyLayout && opts.type !== "sunburst" && opts.type !== "treemap") {
    throw new ChartOptionsError(
      `chartEx.layout.parentLabelLayout is only valid for sunburst and treemap charts.`
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
    throw new ChartOptionsError(
      `chartEx.layout box-whisker fields are only valid for boxWhisker charts.`
    );
  }
  const regionMapFields = [layout.projection, layout.regionLabels, layout.geoMappingLevel];
  if (regionMapFields.some(value => value !== undefined) && opts.type !== "regionMap") {
    throw new ChartOptionsError(
      `chartEx.layout region-map fields are only valid for regionMap charts.`
    );
  }
  if (layout.paretoLine !== undefined && opts.type !== "pareto") {
    throw new ChartOptionsError(`chartEx.layout.paretoLine is only valid for pareto charts.`);
  }
}

function validateBinning(binning: NonNullable<AddChartExOptions["binning"]>): void {
  if (binning.binSize !== undefined && binning.binSize <= 0) {
    throw new ChartOptionsError("chartEx.binning.binSize must be greater than 0.");
  }
  if (
    binning.binCount !== undefined &&
    (!Number.isInteger(binning.binCount) || binning.binCount <= 0)
  ) {
    throw new ChartOptionsError("chartEx.binning.binCount must be a positive integer.");
  }
  if (
    binning.underflow !== undefined &&
    binning.overflow !== undefined &&
    binning.underflow >= binning.overflow
  ) {
    throw new ChartOptionsError("chartEx.binning.underflow must be less than overflow.");
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
      // Histogram and pareto both project onto the same `clusteredColumn`
      // layout — pareto additionally renders a line overlay (flagged via
      // `layoutPr.paretoLine`), histogram supplies the binning metadata.
      return "clusteredColumn";
    case "boxWhisker":
      return "boxWhisker";
    case "regionMap":
      return "regionMap";
    default: {
      const _never: never = type;
      throw new ChartOptionsError(`Unsupported chartEx type: ${String(_never)}.`);
    }
  }
}

function needsAxes(type: AddChartExOptions["type"]): boolean {
  // Sunburst, treemap, funnel, regionMap don't use traditional axes;
  // histogram/pareto/waterfall/boxWhisker do.
  return type === "histogram" || type === "pareto" || type === "waterfall" || type === "boxWhisker";
}
