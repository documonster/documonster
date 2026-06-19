/**
 * ChartEx builder — constructs a ChartExModel from simplified AddChartExOptions.
 *
 * Each layoutId corresponds to a distinct chart type. The builder produces a
 * structured model that the renderer serialises to `cx:chart` XML.
 */

import { hexToColor, toShapeProperties } from "@excel/chart/build/chart-builder";
import type {
  AddChartExOptions,
  AddChartExSeriesOptions,
  ChartExAxis,
  ChartExDataEntry,
  ChartExModel,
  ChartExSeries,
  ChartExSeriesType
} from "@excel/chart/model/chart-ex-types";
import type { ChartRichText } from "@excel/chart/model/types";
import { ChartOptionsError } from "@excel/errors";
import { EMU_PER_PX } from "@utils/units";

/**
 * Parse a single-column range reference of the form `Sheet!$A$2:$A$8`
 * (or `'Sheet Name'!$A$2:$A$8`) into its components. Returns `null`
 * when the formula doesn't match the single-column absolute shape —
 * caller should treat the hierarchy as non-combinable and fall back
 * to whatever behaviour is appropriate for that chart type.
 */
interface ParsedColumnRef {
  sheet: string;
  quotedSheet: boolean;
  colStart: string;
  colEnd: string;
  rowStart: number;
  rowEnd: number;
}

function parseColumnRef(formula: string): ParsedColumnRef | null {
  // Match: optional sheet (quoted or bare) + !$COL$ROW:$COL$ROW
  const match =
    /^\s*(?:'([^'\r\n]+)'|([A-Za-z_][A-Za-z0-9_. ]*))!\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)\s*$/.exec(
      formula
    );
  if (!match) {
    return null;
  }
  // Groups: 1=quoted sheet, 2=bare sheet, 3=colStart, 4=rowStart,
  // 5=colEnd, 6=rowEnd. The row/col pair alternates — the `:`
  // delimiter is between the start address (cols + rows, i.e.
  // `$COL$ROW`) and the end address.
  return {
    sheet: match[1] ?? match[2]!,
    quotedSheet: !!match[1],
    colStart: match[3],
    colEnd: match[5],
    rowStart: parseInt(match[4], 10),
    rowEnd: parseInt(match[6], 10)
  };
}

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

function colIndexToLetters(index: number): string {
  let n = index;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * Combine multiple single-column range formulas (first item is the
 * OUTERMOST hierarchy level, last item is the leaf level / primary
 * category) into a single multi-column range that Excel's hierarchical
 * chartEx loader accepts.
 *
 * Microsoft Excel's treemap / sunburst writer emits ONE `<cx:strDim>`
 * whose `<cx:f>` points at a contiguous multi-column range spanning
 * every hierarchy level (e.g. `Sheet1!$A$2:$C$8` for a Region / Country
 * / City breakdown). The chart reader derives the hierarchy from the
 * column ORDER within that range — the first column is the root and
 * the last is the leaf. Emitting per-level `<cx:strDim>` siblings (one
 * per formula) is schema-legal but makes Excel draw an empty plot
 * area: verified against `reference-hierarchy.xlsx` (Excel-authored
 * sample) where the same hierarchy renders correctly only when the
 * columns live under a single `<cx:f>`.
 *
 * Returns the combined formula when every input formula lives on the
 * same sheet with identical row spans and contiguous ascending columns
 * (the common case when authors lay their hierarchy out in adjacent
 * worksheet columns). Returns `null` otherwise — the caller should
 * fall back to per-level caching because Excel has no way to express
 * a non-contiguous hierarchy as a single range.
 */
export function combineHierarchyFormulas(formulas: readonly string[]): string | null {
  if (formulas.length === 0) {
    return null;
  }
  if (formulas.length === 1) {
    return formulas[0];
  }
  const parsed: ParsedColumnRef[] = [];
  for (const f of formulas) {
    const p = parseColumnRef(f);
    if (!p) {
      return null;
    }
    // Only single-column ranges are combinable. Multi-column inputs
    // can't be meaningfully merged with another column without
    // overlap / ambiguity.
    if (p.colStart !== p.colEnd) {
      return null;
    }
    parsed.push(p);
  }
  const first = parsed[0];
  let prevCol = colLettersToIndex(first.colStart);
  for (let i = 1; i < parsed.length; i++) {
    const p = parsed[i];
    if (p.sheet !== first.sheet) {
      return null;
    }
    if (p.rowStart !== first.rowStart || p.rowEnd !== first.rowEnd) {
      return null;
    }
    const curCol = colLettersToIndex(p.colStart);
    if (curCol !== prevCol + 1) {
      return null;
    }
    prevCol = curCol;
  }
  const startCol = first.colStart;
  const endCol = colIndexToLetters(prevCol);
  const sheetRef = first.quotedSheet ? `'${first.sheet}'` : first.sheet;
  return `${sheetRef}!$${startCol}$${first.rowStart}:$${endCol}$${first.rowEnd}`;
}

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
    // Match Microsoft Excel's convention: chartEx axes are numbered
    // 0 and 1 (cat / val), not `100000000+` like classic charts.
    // Excel writes axes this way for every chartEx it authors —
    // using small sequential ids keeps our output byte-compatible
    // with Excel 2016+ for a fresh chartEx. Classic charts still
    // use the `100000000` seed to avoid collisions with loaded
    // files that happened to allocate small ids; chartEx is a
    // separate part and cannot collide with classic chart axes.
    const catAxisId = 0;
    const valAxisId = 1;
    // Match Excel's default cat-axis `gapWidth` + `<cx:tickLabels/>`
    // and val-axis `<cx:majorGridlines/>` + `<cx:tickLabels/>` so
    // freshly-built chartEx charts look identical to what Excel
    // itself emits for the same chart type. Excel uses
    // `gapWidth="0"` for histogram + pareto (bars touching —
    // histograms traditionally render with zero gap between bins)
    // and `gapWidth="0.5"` for waterfall / boxWhisker / funnel
    // (narrower categorical bars). Users who set `layout.gapWidth`
    // override the default.
    const defaultGapWidth = opts.type === "histogram" || opts.type === "pareto" ? 0 : 0.5;
    axes.push({
      axisId: catAxisId,
      type: "cat",
      catScaling: { gapWidth: defaultGapWidth },
      tickLabels: {}
    });
    axes.push({
      axisId: valAxisId,
      type: "val",
      majorGridlines: {},
      tickLabels: {}
    });
    // Excel does NOT emit `<cx:axisId>` children on `<cx:series>` —
    // it expects the axis binding to come from the axes in
    // `cx:plotArea/cx:axis` directly. Leaving `s.axisId` unset
    // matches Excel's output exactly. (The renderer still emits
    // any `axisId` array populated by the parser for round-tripped
    // files, so existing files that DO carry these references
    // survive load → save without drift.)
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
        // Legend: Excel omits the `<cx:legend>` element entirely for
        // histogram charts (a histogram has a single unnamed series,
        // nothing to legend — emitting an empty legend placeholder
        // causes Excel 2016+ to render the chart as a blank frame).
        // Respect `opts.showLegend === false` for other types.
        // Match Excel's legend attribute defaults otherwise:
        // `pos="t" align="ctr" overlay="0"` — Excel's own output
        // for a newly-inserted waterfall / funnel / boxWhisker uses
        // top-centre placement. Users who pass `opts.legendPosition`
        // still override `pos`.
        legend:
          opts.showLegend === false || opts.type === "histogram"
            ? undefined
            : {
                legendPos: opts.legendPosition ?? "t",
                align: "ctr",
                overlay: false
              }
      }
    },
    // ChartEx ALWAYS ships with a `chartStyle` + `chartColorStyle`
    // sidecar, linked from `chartEx1.xml.rels`. Without them Excel
    // 2016+ discards the chartEx part on load ("Removed Part:
    // /xl/drawings/drawingN.xml (Drawing shape)"). When the caller
    // hasn't supplied structured style / colors, we emit the
    // id-only minimal form — `<cs:chartStyle id="395"/>` and
    // `<cs:colorStyle meth="cycle" id="10"/>` — which Excel's
    // default style table resolves to a sensible built-in palette.
    // Style id 395 + colors id 10 are the defaults Excel itself
    // uses for a freshly-inserted waterfall (verified against a
    // reference xlsx authored by Excel 2021).
    //
    // Callers who want custom styling pass `opts.chartStyle` /
    // `opts.chartColors`; those paths short-circuit this default.
    style: opts.chartStyle ?? { id: 395 },
    colors: opts.chartColors ?? { id: 10, method: "cycle" }
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

  // For sunburst / treemap, Excel expects a SINGLE `<cx:strDim>`
  // whose `<cx:f>` points to a contiguous MULTI-COLUMN range that
  // spans every hierarchy level (outer → inner, root column first,
  // leaf column last). See `combineHierarchyFormulas` for the
  // rationale + the reference fixture (`reference-hierarchy.xlsx`)
  // where the same layout rendered blank under the previous
  // "one data entry per level" approach. When the caller's
  // `categories` + `hierarchy` live in adjacent worksheet columns
  // with matching row ranges we rewrite `catDataId`'s strDim with
  // the combined range and skip per-level data entries. When the
  // ranges can't be combined (different sheets, non-contiguous
  // columns, …) we leave the entries as-is and accept that Excel
  // may draw the chart blank — manual reshaping by the author is
  // the only way to recover.
  let hierarchyCombined = false;
  if (isHierarchical) {
    const hierarchyFormulas = so.hierarchy;
    if (hierarchyFormulas && catDataId !== undefined && opts.categories) {
      // Root → leaf: `[...hierarchy, categories]` matches Excel's
      // left-to-right column reading order.
      const combined = combineHierarchyFormulas([...hierarchyFormulas, opts.categories]);
      if (combined) {
        hierarchyCombined = true;
        const catEntry = data.find(d => d.id === catDataId);
        if (catEntry?.strDim) {
          catEntry.strDim.formula = combined;
          // Clear any seeded empty-level placeholder: Excel's writer
          // emits no `<cx:lvl>` cache for hierarchical charts and
          // reads fresh from the referenced cells on open.
          catEntry.strDim.levels = undefined;
          // Suppress cache population for this dimension. A flat
          // `<cx:lvl>` of width × height points across the multi-column
          // range would confuse Excel's hierarchical renderer (the
          // chart draws empty). See `hasChartExStringPoints`.
          catEntry.strDim._skipCache = true;
        }
      } else {
        // Ranges aren't combinable — fall back to per-level data
        // entries in OUTER→INNER order so at least the schema is
        // valid. (Excel renders blank in this fallback; emit with a
        // console warning in non-production so the author can
        // diagnose why their hierarchical chart is empty.)
        if (typeof process !== "undefined" && process.env?.NODE_ENV !== "production") {
          console.warn(
            "[documonster] Treemap/sunburst hierarchy + categories could not be combined into a " +
              "contiguous multi-column range. Excel will render the chart as an empty plot area. " +
              "Lay your leaf + hierarchy columns contiguously on the same sheet with matching rows."
          );
        }
        for (let i = hierarchyFormulas.length - 1; i >= 0; i--) {
          const h = hierarchyFormulas[i];
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
    }
    if (so.literalHierarchy) {
      // Literal hierarchy (no worksheet cells to reference) — cache
      // every level's values inside a single strDim's `levels` array
      // so `consolidateDataForRender` can flatten them into one
      // `<cx:strDim>` with multiple `<cx:lvl>` children at render
      // time. This does not match Excel's own "single multi-column
      // <cx:f>" output verbatim but is what the schema allows when
      // there are no real cells to point at.
      for (let i = so.literalHierarchy.length - 1; i >= 0; i--) {
        const level = so.literalHierarchy[i];
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
  // Microsoft Excel uses `type="val"` for MOST chartEx numeric
  // dimensions (histogram + pareto binning inputs, funnel values,
  // box-whisker samples, waterfall deltas, regionMap values), but
  // **sunburst and treemap use `type="size"`** — the hierarchy rings
  // / tiles are sized by the numeric column, and Excel's loader
  // keys the hierarchical layout engine off this attribute. Reference
  // fixture `tmp/reference-hierarchy.xlsx` (Excel-authored) writes
  // `<cx:numDim type="size">` on both its treemap and sunburst
  // charts; the previous `type="val"` caused Excel to open the
  // chart with an empty plot area (schema-valid but semantically
  // wrong for the hierarchical renderer).
  //
  // Other layouts still use `"val"` — the alternative `"x"` is in
  // `ST_NumDimType` but Excel 2016+ renders the chart as a blank
  // frame when the dimension is labelled `"x"` (verified against
  // an Excel 2021-authored histogram reference, `tmp/aaaaa.xlsx`).
  const numDimType = isHierarchical ? "size" : "val";
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
        : [{ ptCount: 0, points: [] }],
      // Hierarchical charts that successfully combined their leaf +
      // hierarchy formulas into a single multi-column `<cx:f>` ship
      // the numeric dimension without a `<cx:lvl>` cache — same
      // rationale as the `_skipCache` marker on the category strDim
      // above. A simple sunburst with no hierarchy, or a combine
      // that failed (non-contiguous columns), still caches normally
      // so the chart at least has something to paint from.
      ...(hierarchyCombined ? { _skipCache: true } : {})
    }
  });
  series.dataRefs!.push({ dataId: valId });
  if (hierarchyCombined) {
    // Skipping the cache populator means the placeholder level we seed
    // above never gets replaced with real values. Clear it so the
    // renderer emits just the `<cx:f>` reference, matching Excel's
    // writer output for treemap + sunburst.
    const valEntry = data.find(d => d.id === valId);
    if (valEntry?.numDim && !so.literalValues?.length) {
      valEntry.numDim.levels = undefined;
    }
  }

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
  if (layoutId === "waterfall") {
    // Excel emits `<cx:layoutPr><cx:subtotals/></cx:layoutPr>` on
    // EVERY waterfall series, even when the user has not marked any
    // subtotal points. The empty `subtotals` element is treated as
    // a "no subtotals but waterfall-aware" marker; without it Excel
    // falls back to generic series rendering and, at load time,
    // may reject the chartEx as malformed. Seed `layoutPr.subtotals`
    // with an empty array here so the writer always emits the
    // element — then merge in any explicit user-provided subtotal
    // indices on top.
    series.layoutPr = {
      subtotals: mergedSubtotals ? mergedSubtotals.map(i => ({ idx: i })) : []
    };
  }
  if (layoutId === "treemap" || layoutId === "sunburst") {
    // Excel 2016+ will NOT render a sunburst / treemap that lacks a
    // `<cx:layoutPr>` child — the frame is drawn but the plot area
    // shows a completely blank canvas (the hierarchy arcs / tiles
    // are computed but never painted). Verified against
    // `tmp/ttttt.xlsx` (Excel-authored sunburst reference) — every
    // Excel-authored hierarchical chartEx carries at least an empty
    // `<cx:layoutPr/>`, and treemap charts additionally carry a
    // `<cx:parentLabelLayout val="overlapping"/>` (the default) so
    // the engine has something to reach for when deciding how to
    // stack parent labels. Seed a minimal default here so freshly
    // authored sunburst/treemap charts render on open; the caller's
    // `layout.parentLabelLayout` (applied below via `opts.layout`
    // merge) overrides the default when set.
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      ...(layoutId === "treemap" ? { parentLabelLayout: "overlapping" as const } : {})
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
    // Excel's default `<cx:binning>` for a freshly-inserted
    // histogram carries `intervalClosed="r"` — meaning "right-
    // closed bin boundaries" (bins are `[a, b]`, not `[a, b)`).
    // Emitting the element without the attribute is spec-legal
    // (ST_IntervalClosedSide has a schema default of `"r"`) but
    // Excel 2016+ treats the absence as "auto-compute binning"
    // AND sometimes renders the chart as a blank frame when the
    // binning has no explicit side. Verified against
    // `tmp/aaaaa.xlsx` (Excel-authored histogram reference).
    series.layoutPr = {
      ...(series.layoutPr ?? {}),
      binning: { binType: "auto", intervalClosed: "r" }
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
        width: EMU_PER_PX
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

  // Default dataLabels for chartEx types where Excel always emits
  // them. Match Excel's exact defaults so our output is byte-close
  // to what a freshly-inserted chart looks like:
  //
  //   <cx:dataLabels pos="outEnd">
  //     <cx:visibility seriesName="0" categoryName="0" value="1"/>
  //   </cx:dataLabels>
  //
  // Only applied when the caller did NOT supply `so.dataLabels` —
  // user-driven customisation is preserved. The branches cover
  // every 2014+ layout for which Excel emits a default block;
  // sunburst / treemap typically show only the category name, so
  // their defaults differ. (Pareto is handled on the pareto line
  // series via a separate pass.)
  if (!series.dataLabels) {
    if (layoutId === "waterfall" || layoutId === "funnel") {
      series.dataLabels = {
        position: "outEnd",
        visibility: { seriesName: false, categoryName: false, value: true }
      };
    } else if (layoutId === "sunburst" || layoutId === "treemap") {
      series.dataLabels = {
        visibility: { seriesName: false, categoryName: true, value: false }
      };
    }
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
    // A series must reference its data either via a worksheet formula
    // (`values`) or via cached literal values (`literalValues`). Headless
    // chart pipelines (e.g. the Word `buildWordChartExXml` bridge) ship
    // only `literalValues` because there is no underlying worksheet to
    // reference.
    if (!series.values && !series.literalValues?.length) {
      throw new ChartOptionsError(`chartEx.series[${i}].values or literalValues is required.`);
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
