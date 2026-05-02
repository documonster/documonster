/**
 * Best-effort ChartEx parser for Office 2016+ cx: chart XML.
 *
 * The parser captures the structured fields that ExcelTS can render and keeps
 * the original XML on the model so clean round-trips can still use raw bytes.
 */

import { parseXml, findChild, findChildren, textContent } from "@xml/dom";
import type { XmlElement, XmlNode } from "@xml/types";

import type {
  ChartExAxis,
  ChartExDataEntry,
  ChartExDimensionType,
  ChartExModel,
  ChartExSeries,
  ChartExUnknownElement
} from "./chart-ex-types";
import { escapeXml, escapeXmlAttr } from "./chart-utils";
import { parseSpPr, parseTxPr } from "./shape-properties";
import type {
  ChartLegend,
  ChartRichText,
  ChartTextProperties,
  ChartTitle,
  ShapeProperties
} from "./types";

// Element-name whitelists for unknown-child detection. Parents not listed here
// are either (a) purely raw-passthrough containers already captured via
// `_rawXml` / `rawXml` / `extLst`, or (b) leaf elements whose content is
// attributes only. Update these lists when a new cx: element is promoted into
// the structured model so it stops being flagged as "unknown".
const KNOWN_CHILDREN_CHART_SPACE = new Set([
  "cx:chartData",
  "cx:chart",
  "cx:clrMapOvr",
  "cx:spPr",
  "cx:txPr",
  "cx:protection",
  "cx:printSettings",
  "cx:externalData",
  "cx:extLst"
]);
const KNOWN_CHILDREN_CHART = new Set([
  "cx:title",
  "cx:autoTitleDeleted",
  "cx:plotArea",
  "cx:legend",
  "cx:spPr",
  "cx:extLst"
]);
const KNOWN_CHILDREN_PLOT_AREA = new Set([
  "cx:plotAreaRegion",
  "cx:series",
  "cx:axis",
  "cx:spPr",
  "cx:extLst"
]);
const KNOWN_CHILDREN_PLOT_AREA_REGION = new Set([
  "cx:layout",
  "cx:plotSurface",
  "cx:spPr",
  "cx:series",
  "cx:extLst"
]);
const KNOWN_CHILDREN_SERIES = new Set([
  "cx:tx",
  "cx:spPr",
  "cx:txPr",
  "cx:valueColors",
  "cx:valueColorPositions",
  "cx:dataPt",
  "cx:dataLabels",
  "cx:dataId",
  "cx:layoutPr",
  "cx:axisId",
  "cx:extLst"
]);
const KNOWN_CHILDREN_AXIS = new Set([
  "cx:catScaling",
  "cx:valScaling",
  // `cx:hidden`, `cx:majorTickMark`, and `cx:minorTickMark` were
  // previously emitted by this library in violation of the Chart2014
  // schema (hidden is an attribute; the element names are plural).
  // Keep them in the known-children set so legacy files produced by
  // older versions don't trip the `unknownElements` collector.
  "cx:hidden",
  "cx:majorTickMark",
  "cx:majorTickMarks",
  "cx:minorTickMark",
  "cx:minorTickMarks",
  "cx:numFmt",
  "cx:title",
  "cx:units",
  "cx:majorGridlines",
  "cx:minorGridlines",
  "cx:spPr",
  "cx:txPr",
  "cx:extLst"
]);

/**
 * Accumulator threaded through parse helpers to collect every child element
 * that falls outside the structured model's whitelist. Consumed by
 * {@link parseChartEx} to populate {@link ChartExModel.unknownElements}.
 */
interface UnknownCollector {
  entries: ChartExUnknownElement[];
}

function collectUnknownChildren(
  parent: XmlElement | undefined,
  knownNames: Set<string>,
  parentPath: string,
  collector: UnknownCollector
): void {
  if (!parent) {
    return;
  }
  for (const child of parent.children ?? []) {
    if (child.type !== "element") {
      continue;
    }
    if (!knownNames.has(child.name)) {
      collector.entries.push({
        name: child.name,
        path: `${parentPath}/${child.name}`
      });
    }
  }
}

export function parseChartEx(rawXml: string): ChartExModel {
  const root = parseXml(rawXml).root;
  const chartDataEl = findChild(root, "cx:chartData");
  const chartEl = findChild(root, "cx:chart");
  const collector: UnknownCollector = { entries: [] };

  collectUnknownChildren(root, KNOWN_CHILDREN_CHART_SPACE, "cx:chartSpace", collector);
  if (chartEl) {
    collectUnknownChildren(chartEl, KNOWN_CHILDREN_CHART, "cx:chartSpace/cx:chart", collector);
    const plotAreaEl = findChild(chartEl, "cx:plotArea");
    if (plotAreaEl) {
      collectUnknownChildren(
        plotAreaEl,
        KNOWN_CHILDREN_PLOT_AREA,
        "cx:chartSpace/cx:chart/cx:plotArea",
        collector
      );
      const regionEl = findChild(plotAreaEl, "cx:plotAreaRegion");
      if (regionEl) {
        collectUnknownChildren(
          regionEl,
          KNOWN_CHILDREN_PLOT_AREA_REGION,
          "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion",
          collector
        );
        for (const seriesEl of findChildren(regionEl, "cx:series")) {
          collectUnknownChildren(
            seriesEl,
            KNOWN_CHILDREN_SERIES,
            "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series",
            collector
          );
        }
      }
      for (const seriesEl of findChildren(plotAreaEl, "cx:series")) {
        collectUnknownChildren(
          seriesEl,
          KNOWN_CHILDREN_SERIES,
          "cx:chartSpace/cx:chart/cx:plotArea/cx:series",
          collector
        );
      }
      for (const axisEl of findChildren(plotAreaEl, "cx:axis")) {
        collectUnknownChildren(
          axisEl,
          KNOWN_CHILDREN_AXIS,
          "cx:chartSpace/cx:chart/cx:plotArea/cx:axis",
          collector
        );
      }
    }
  }

  // `CT_Chart` has no `spPr` child in Chart2014, but earlier versions
  // of this library emitted chart-frame styling there. Read it
  // defensively, then merge into the authoritative `chartSpace.spPr`
  // slot when that is absent — preserving legacy files without
  // re-emitting schema-invalid XML. The write path now always uses
  // `chartSpace.spPr` (`CT_ChartSpace/cx:spPr`).
  const legacyChartSpPr = parseRawSpPr(chartEl ? findChild(chartEl, "cx:spPr") : undefined);
  const chartSpacespPr = parseRawSpPr(findChild(root, "cx:spPr")) ?? legacyChartSpPr;
  // `cx:externalData` is a child of `cx:chartSpace` per the Chart2014
  // `CT_ChartSpace` schema. Some Excel-authored files (and earlier
  // versions of this library) misplaced it inside `cx:chartData`;
  // read from the schema-correct location first, fall back to the
  // legacy location. Prefer one or the other — concatenating both
  // produces duplicate `r:id` references when a file happens to carry
  // the element in BOTH places (rare but observed in files that
  // migrated through multiple writers), which fails strict-open
  // validation.
  const externalDataAtSpace = findChildren(root, "cx:externalData").map(el => ({
    id: el.attributes["r:id"] ?? "",
    autoUpdate: parseBoolAttr(el, "autoUpdate")
  }));
  const externalDataAtChartData = chartDataEl
    ? findChildren(chartDataEl, "cx:externalData").map(el => ({
        id: el.attributes["r:id"] ?? "",
        autoUpdate: parseBoolAttr(el, "autoUpdate")
      }))
    : [];
  const externalData =
    externalDataAtSpace.length > 0 ? externalDataAtSpace : externalDataAtChartData;
  return {
    chartSpace: {
      chartData: parseChartData(chartDataEl, collector),
      chart: parseChart(chartEl, collector),
      clrMapOvr: rawElement(root, "cx:clrMapOvr"),
      // ChartSpace-level `spPr` / `txPr` are captured as structured-
      // with-rawXml by `parseRawSpPr` / `parseRawTxPr` so downstream
      // mutations (setSpPrFill, etc.) still work without losing the
      // original XML for the unchanged field subset.
      spPr: chartSpacespPr,
      txPr: parseRawTxPr(root, "cx:txPr"),
      // `cx:protection` / `cx:printSettings` are kept verbatim — the
      // structured model does not (yet) break them into typed fields.
      protection: rawElement(root, "cx:protection"),
      printSettings: rawElement(root, "cx:printSettings"),
      extLst: rawElement(root, "cx:extLst"),
      ...(externalData.length > 0 ? { externalData } : {})
    },
    rawXml,
    ...(collector.entries.length > 0 ? { unknownElements: collector.entries } : {})
  };
}

function parseChartData(
  chartDataEl: XmlElement | undefined,
  collector?: UnknownCollector
): ChartExModel["chartSpace"]["chartData"] {
  if (!chartDataEl) {
    return { data: [] };
  }

  const data = findChildren(chartDataEl, "cx:data").map(el => parseDataEntry(el, collector));

  // `cx:externalData` now lives on `ChartExSpace` (schema-correct
  // location). The caller (`parseChartEx`) migrates any legacy
  // occurrences found inside `cx:chartData` up to chartSpace — this
  // function no longer returns the field so the migration is
  // authoritative.
  return { data };
}

function parseDataEntry(el: XmlElement, collector?: UnknownCollector): ChartExDataEntry {
  const strDim = findChild(el, "cx:strDim");
  const numDim = findChild(el, "cx:numDim");
  // `parseIndexAttr` coerces empty / missing / non-finite / negative
  // ids to `0` (see {@link parseIndexAttr}). OOXML treats absent and
  // empty `id=""` identically, and sparse-array-hostile sentinels
  // (`-1`, `NaN`) would poison downstream point indexing.
  return {
    id: parseIndexAttr(el.attributes.id),
    ...(strDim ? { strDim: parseStringDimension(strDim, collector) } : {}),
    ...(numDim ? { numDim: parseNumericDimension(numDim, collector) } : {})
  };
}

function parseStringDimension(
  el: XmlElement,
  collector?: UnknownCollector
): NonNullable<ChartExDataEntry["strDim"]> {
  return {
    type: parseDimensionType(el.attributes.type, collector, "cx:strDim/@type"),
    formula: childText(el, "cx:f"),
    levels: findChildren(el, "cx:lvl").map(lvl => ({
      ptCount: parseOptionalInt(lvl.attributes.ptCount),
      points: findChildren(lvl, "cx:pt").map(pt => ({
        index: parseIndexAttr(pt.attributes.idx),
        value: textContent(pt)
      }))
    }))
  };
}

function parseNumericDimension(
  el: XmlElement,
  collector?: UnknownCollector
): NonNullable<ChartExDataEntry["numDim"]> {
  return {
    type: parseDimensionType(el.attributes.type, collector, "cx:numDim/@type"),
    formula: childText(el, "cx:f"),
    levels: findChildren(el, "cx:lvl").map(lvl => ({
      ptCount: parseOptionalInt(lvl.attributes.ptCount),
      formatCode: lvl.attributes.formatCode,
      // Skip `<cx:pt>` whose text content is blank / non-numeric / an
      // Excel error code (`"#N/A"` etc.). Previously these flowed into
      // the model as `NaN`, polluted axis-range calculations, and
      // produced non-numeric `<c:v>` on the round-trip write.
      points: findChildren(lvl, "cx:pt")
        .map(pt => {
          const value = parseFloat(textContent(pt));
          if (!Number.isFinite(value)) {
            return undefined;
          }
          return {
            index: parseIndexAttr(pt.attributes.idx),
            value
          };
        })
        .filter((p): p is { index: number; value: number } => p !== undefined)
    }))
  };
}

function parseChart(
  chartEl: XmlElement | undefined,
  collector: UnknownCollector
): ChartExModel["chartSpace"]["chart"] {
  const plotAreaEl = chartEl ? findChild(chartEl, "cx:plotArea") : undefined;
  return {
    title: chartEl ? parseTitle(findChild(chartEl, "cx:title")) : undefined,
    autoTitleDeleted: chartEl
      ? parseBoolAttr(findChild(chartEl, "cx:autoTitleDeleted"), "val")
      : undefined,
    plotArea: parsePlotArea(plotAreaEl, collector),
    legend: chartEl ? parseLegend(findChild(chartEl, "cx:legend"), collector) : undefined
    // NOTE: `cx:spPr` on `<cx:chart>` is a schema violation (CT_Chart
    // has no `spPr` child); the outer caller (`parseChartExXml`)
    // migrates any legacy `chart.spPr` bytes into `chartSpace.spPr`
    // before this return value is consumed. See the block above the
    // `parseChart` caller for the migration path.
  };
}

function parsePlotArea(
  plotAreaEl: XmlElement | undefined,
  collector: UnknownCollector
): ChartExModel["chartSpace"]["chart"]["plotArea"] {
  if (!plotAreaEl) {
    return {};
  }

  const regionEl = findChild(plotAreaEl, "cx:plotAreaRegion");
  const axis = findChildren(plotAreaEl, "cx:axis").map(el => parseAxis(el, collector));

  if (regionEl) {
    const regionSeriesPath = "cx:chartSpace/cx:chart/cx:plotArea/cx:plotAreaRegion/cx:series";
    // `CT_PlotAreaRegion` (Chart2014) has children `plotSurface?` and
    // `series*` — there is no `layout` or `spPr` child. The parser
    // previously pulled `cx:spPr` directly from the region, which was a
    // schema violation. `plotSurface` wraps an inner `cx:spPr`, so we
    // must dig one level deeper.
    const plotSurfaceEl = findChild(regionEl, "cx:plotSurface");
    return {
      plotAreaRegion: {
        // `layout` is not a real `plotAreaRegion` child, but earlier
        // versions of this library emitted `<cx:layout/>` there — keep
        // parsing it so round-trip of those files doesn't drop state.
        layout: parseRawLayout(findChild(regionEl, "cx:layout")),
        plotSurface: plotSurfaceEl ? parseRawSpPr(findChild(plotSurfaceEl, "cx:spPr")) : undefined,
        series: findChildren(regionEl, "cx:series").map(el =>
          parseSeries(el, collector, regionSeriesPath)
        ),
        // `<cx:extLst>` nested inside `<cx:plotAreaRegion>` is a
        // Chart2014 extension point (Excel writes `cx14:*` annotations
        // here for pivot-backed region maps). Preserve the raw bytes
        // for round-trip; the parser never interprets them.
        extLst: rawElement(regionEl, "cx:extLst")
      },
      spPr: parseRawSpPr(findChild(plotAreaEl, "cx:spPr")),
      extLst: rawElement(plotAreaEl, "cx:extLst"),
      ...(axis.length > 0 ? { axis } : {})
    };
  }

  const plainSeriesPath = "cx:chartSpace/cx:chart/cx:plotArea/cx:series";
  const series = findChildren(plotAreaEl, "cx:series").map(el =>
    parseSeries(el, collector, plainSeriesPath)
  );
  return {
    ...(series.length > 0 ? { series } : {}),
    spPr: parseRawSpPr(findChild(plotAreaEl, "cx:spPr")),
    extLst: rawElement(plotAreaEl, "cx:extLst"),
    ...(axis.length > 0 ? { axis } : {})
  };
}

// ECMA-376 §21.3 enumerates the legal `cx:series/@layoutId` values. Preserve
// the attribute when it matches the structured union; otherwise fall back to
// "clusteredColumn" (the most permissive single-axis layout) and surface the
// original string via `unknownElements` so callers can diagnose. Silently
// defaulting to e.g. "sunburst" would mis-route every unknown chart to the
// sunburst vector renderer.
const CHART_EX_SERIES_LAYOUT_IDS: ReadonlySet<ChartExSeries["layoutId"]> = new Set([
  "clusteredColumn",
  "waterfall",
  "funnel",
  "boxWhisker",
  "paretoLine",
  "regionMap",
  "sunburst",
  "treemap"
]);

function parseSeriesLayoutId(
  value: string | undefined,
  collector: UnknownCollector | undefined,
  seriesPath: string
): { layoutId: ChartExSeries["layoutId"]; rawLayoutId?: string } {
  if (value !== undefined && CHART_EX_SERIES_LAYOUT_IDS.has(value as ChartExSeries["layoutId"])) {
    return { layoutId: value as ChartExSeries["layoutId"] };
  }
  collector?.entries.push({
    name: value === undefined ? "@layoutId(missing)" : `@layoutId=${value}`,
    path: seriesPath
  });
  // "clusteredColumn" is the neutral fallback: a plain column chart will at
  // least render bars for the available data rather than being reinterpreted
  // as a sunburst / treemap, which have different data-shape requirements.
  //
  // When the original attribute was present but unknown (e.g. a future
  // layoutId like `"funnel3D"` that this version doesn't model), keep it
  // under `rawLayoutId` so the writer can re-emit the byte-original
  // attribute during round-trip. Without this, re-saving a file with
  // a vendor extension silently downgrades the chart to a plain column
  // chart — worse than leaving the bytes alone.
  return {
    layoutId: "clusteredColumn",
    rawLayoutId: value
  };
}

function parseSeries(
  el: XmlElement,
  collector?: UnknownCollector,
  seriesPath?: string
): ChartExSeries {
  const dataRefs = findChildren(el, "cx:dataId").map(dataId => ({
    dataId: parseIndexAttr(dataId.attributes.val)
  }));
  const axisId = findChildren(el, "cx:axisId").map(axis => parseIndexAttr(axis.attributes.val));
  const tx = parseSeriesText(findChild(el, "cx:tx"));
  const layoutPr = parseLayoutProperties(findChild(el, "cx:layoutPr"));

  const layoutIdResult = parseSeriesLayoutId(
    el.attributes.layoutId,
    collector,
    seriesPath ?? "cx:series"
  );

  return {
    layoutId: layoutIdResult.layoutId,
    ...(layoutIdResult.rawLayoutId !== undefined
      ? { rawLayoutId: layoutIdResult.rawLayoutId }
      : {}),
    hidden: parseBoolAttr(el, "hidden"),
    ownerIdx: parseOptionalInt(el.attributes.ownerIdx),
    ...(tx ? { tx } : {}),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr"),
    // `<cx:valueColors>` and `<cx:valueColorPositions>` carry
    // value-driven colour gradients (region map / treemap). The
    // structured model doesn't yet interpret the stops, but we must
    // round-trip them or the gradient evaporates on save. `rawElement`
    // preserves the full outer element including attributes.
    valueColors: rawElement(el, "cx:valueColors"),
    valueColorPositions: rawElement(el, "cx:valueColorPositions"),
    ...(dataRefs.length > 0 ? { dataRefs } : {}),
    ...(layoutPr ? { layoutPr } : {}),
    ...(axisId.length > 0 ? { axisId } : {}),
    dataLabels: parseDataLabels(findChild(el, "cx:dataLabels"), collector, seriesPath),
    dataPt: findChildren(el, "cx:dataPt").map(pt => ({
      idx: parseIndexAttr(pt.attributes.idx),
      spPr: parseRawSpPr(findChild(pt, "cx:spPr"))
    })),
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseSeriesText(el: XmlElement | undefined): ChartExSeries["tx"] | undefined {
  if (!el) {
    return undefined;
  }
  // Check for structured rich-text first. `<cx:tx>/<cx:rich>…</cx:rich>`
  // is how Excel authors encode per-run formatting on a series name
  // (bold accent series, colour-coded series, etc.); previously the
  // parser ignored this branch entirely and left `tx` undefined, so
  // round-tripping such a series lost the entire name unless the raw
  // bytes path was taken.
  const richEl = findChild(el, "cx:rich");
  if (richEl) {
    const rich = parseRichText(richEl);
    if (rich) {
      return { rich };
    }
  }
  const txData = findChild(el, "cx:txData");
  if (!txData) {
    return undefined;
  }
  const formula = childText(txData, "cx:f");
  if (formula) {
    // Capture the cached resolved value alongside the formula so
    // reading clients (legend preview, title fallback, saved-chart
    // displays before recalculation) can show the user's label
    // immediately. Previously only the formula string was stored,
    // so Excel-authored `<cx:tx><cx:txData><cx:f>...</cx:f><cx:v>
    // Label</cx:v></cx:txData></cx:tx>` lost `Label` on round-trip.
    const cachedValue = childText(txData, "cx:v");
    return {
      strRef: cachedValue !== undefined ? { formula, cached: cachedValue } : { formula }
    };
  }
  const value = childText(txData, "cx:v");
  return value !== undefined ? { value } : undefined;
}

function parseLayoutProperties(el: XmlElement | undefined): ChartExSeries["layoutPr"] | undefined {
  if (!el) {
    return undefined;
  }

  const binning = findChild(el, "cx:binning");
  const subtotalsEl = findChild(el, "cx:subtotals");
  const subtotals = subtotalsEl
    ? findChildren(subtotalsEl, "cx:subtotal").map(st => ({
        idx: parseIndexAttr(st.attributes.idx)
      }))
    : undefined;
  const parentLabelLayout = childVal(el, "cx:parentLabelLayout");
  const quartileMethod = childVal(el, "cx:quartileMethod");
  const projection = childVal(el, "cx:projection");
  const regionLabels = childVal(el, "cx:regionLabels");
  const geoMappingLevel = childVal(el, "cx:geoMappingLevel");
  // Collect each child only when the XML actually contains it. Emitting
  // `subtotals: []` / `parentLabelLayout: undefined` triggers
  // `hasStructuredLayoutProperties` (see chart-ex-renderer.ts) and forces the
  // writer down the structured path, discarding `_rawXml`. That produced
  // empty / corrupt `<cx:layoutPr>` sections for any ChartEx that was loaded
  // and re-written without an explicit structured mutation.
  return {
    _rawXml: serializeElement(el),
    ...(parentLabelLayout !== undefined
      ? {
          parentLabelLayout: parentLabelLayout as NonNullable<
            ChartExSeries["layoutPr"]
          >["parentLabelLayout"]
        }
      : {}),
    ...(subtotals !== undefined ? { subtotals } : {}),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:connectorLines"));
      return v !== undefined ? { connectorLines: v } : {};
    })(),
    ...(binning
      ? {
          binning: {
            // Validate against `CT_Binning/@intervalClosed`
            // enumeration (`"l" | "r"`). Unknown values are dropped
            // rather than cast through — silently storing garbage
            // here let injection attacks break out of the attribute
            // on round-trip.
            intervalClosed: (binning.attributes.intervalClosed === "l" ||
            binning.attributes.intervalClosed === "r"
              ? binning.attributes.intervalClosed
              : undefined) as NonNullable<
              NonNullable<ChartExSeries["layoutPr"]>["binning"]
            >["intervalClosed"],
            underflow: parseOptionalFloat(binning.attributes.underflow),
            overflow: parseOptionalFloat(binning.attributes.overflow),
            binSize: parseOptionalFloat(childVal(binning, "cx:binSize")),
            binCount: parseOptionalInt(childVal(binning, "cx:binCount")),
            binType: findChild(binning, "cx:auto")
              ? "auto"
              : findChild(binning, "cx:categories")
                ? "categories"
                : findChild(binning, "cx:manual")
                  ? "manual"
                  : findChild(binning, "cx:binSize")
                    ? "binSize"
                    : findChild(binning, "cx:binCount")
                      ? "binCount"
                      : undefined
          }
        }
      : {}),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:paretoLine"));
      return v !== undefined ? { paretoLine: v } : {};
    })(),
    ...(quartileMethod !== undefined
      ? {
          quartileMethod: quartileMethod as NonNullable<ChartExSeries["layoutPr"]>["quartileMethod"]
        }
      : {}),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:showMeanLine"));
      return v !== undefined ? { showMeanLine: v } : {};
    })(),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:showMeanMarker"));
      return v !== undefined ? { showMeanMarker: v } : {};
    })(),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:showInnerPoints"));
      return v !== undefined ? { showInnerPoints: v } : {};
    })(),
    ...(() => {
      const v = parseCtBoolean(findChild(el, "cx:showOutlierPoints"));
      return v !== undefined ? { showOutlierPoints: v } : {};
    })(),
    ...(projection !== undefined
      ? { projection: projection as NonNullable<ChartExSeries["layoutPr"]>["projection"] }
      : {}),
    ...(regionLabels !== undefined
      ? { regionLabels: regionLabels as NonNullable<ChartExSeries["layoutPr"]>["regionLabels"] }
      : {}),
    ...(geoMappingLevel !== undefined
      ? {
          geoMappingLevel: geoMappingLevel as NonNullable<
            ChartExSeries["layoutPr"]
          >["geoMappingLevel"]
        }
      : {}),
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseDataLabels(
  el: XmlElement | undefined,
  collector?: UnknownCollector,
  parentPath?: string
): ChartExSeries["dataLabels"] | undefined {
  if (!el) {
    return undefined;
  }
  const visibility = findChild(el, "cx:visibility");
  // ECMA-376 `ST_DLblPosition` full enumeration (Chart2014 §21.3.2.30).
  // Values beyond the classic-chart subset — `ctrEnd`, `inStart`,
  // `outStart` — are emitted by Excel for waterfall / funnel / box &
  // whisker layouts where label anchoring is richer than the basic
  // "inside/outside/centre/best-fit" vocabulary. Previously the
  // enum was restricted to the classic-chart subset, causing the
  // validator to strip (and log) every richer value encountered in
  // user files, silently reverting those labels to default placement.
  const positionValues = [
    "ctr",
    "l",
    "r",
    "t",
    "b",
    "inBase",
    "inEnd",
    "inStart",
    "outEnd",
    "outStart",
    "ctrEnd",
    "bestFit"
  ] as const;
  const pathPrefix = parentPath ?? "cx:dataLabels";
  return {
    visibility: visibility
      ? {
          seriesName: parseBoolAttr(visibility, "seriesName"),
          categoryName: parseBoolAttr(visibility, "categoryName"),
          value: parseBoolAttr(visibility, "value"),
          numFmt: parseBoolAttr(visibility, "numFmt")
        }
      : undefined,
    position: validateEnum(
      findChild(el, "cx:dataLabel")?.attributes.pos,
      positionValues,
      collector,
      `${pathPrefix}/cx:dataLabel/@pos`
    ),
    separator: childText(el, "cx:separator"),
    numFmt: findChild(el, "cx:numFmt")?.attributes.formatCode,
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr")
  };
}

/**
 * Parse `<cx:majorGridlines>` / `<cx:minorGridlines>` — each element
 * wraps a single optional `<cx:spPr>` child describing the gridline
 * stroke. Returns an empty `ShapeProperties` when the element is
 * present but has no `<cx:spPr>` (meaning "default-styled gridlines");
 * `undefined` when the parent element is absent entirely.
 */
function parseGridlines(el: XmlElement | undefined): ShapeProperties | undefined {
  if (!el) {
    return undefined;
  }
  const spPrEl = findChild(el, "cx:spPr");
  if (!spPrEl) {
    return {};
  }
  return parseRawSpPr(spPrEl);
}

function parseAxis(el: XmlElement, collector?: UnknownCollector): ChartExAxis {
  const numFmt = findChild(el, "cx:numFmt");
  const valScaling = findChild(el, "cx:valScaling");
  const catScaling = findChild(el, "cx:catScaling");
  // OOXML `ST_TickMark` values are `cross | in | out | none`; the public
  // API uses friendlier `inside` / `outside` aliases (matching classic
  // `types.ts:TickMark`). Map during parse so callers see a consistent
  // vocabulary regardless of whether they came from classic or ChartEx.
  const tickMarkOoxml = ["cross", "in", "out", "none"] as const;
  type TickMarkOoxml = (typeof tickMarkOoxml)[number];
  const mapTickMark = (v: TickMarkOoxml | undefined): ChartExAxis["majorTickMark"] => {
    if (v === undefined) {
      return undefined;
    }
    if (v === "in") {
      return "inside";
    }
    if (v === "out") {
      return "outside";
    }
    return v;
  };
  return {
    axisId: parseIndexAttr(el.attributes.id),
    type: findChild(el, "cx:valScaling") ? "val" : "cat",
    // Per Chart2014 `CT_Axis`, `hidden` is an **attribute** on
    // `<cx:axis>`, not a child element. Accept both forms so we can
    // round-trip files our earlier versions incorrectly produced: check
    // the attribute first, fall back to the legacy child element.
    // Per Chart2014 `CT_Axis`, `hidden` is an **attribute** on
    // `<cx:axis>`, not a child element. Accept both forms so we can
    // round-trip files our earlier versions incorrectly produced: check
    // the attribute first, fall back to the legacy child element.
    // When using the legacy child form, treat a missing `val` as
    // `true` per `CT_Boolean` (previously the library only read an
    // explicit `val` and dropped the hidden flag when absent).
    hidden: parseBoolAttr(el, "hidden") ?? parseCtBoolean(findChild(el, "cx:hidden")),
    // Element names are plural in the schema (`majorTickMarks`,
    // `minorTickMarks`). Again accept the singular form for backward
    // compatibility with files we produced before the rename.
    majorTickMark: mapTickMark(
      validateEnum(
        childVal(el, "cx:majorTickMarks") ?? childVal(el, "cx:majorTickMark"),
        tickMarkOoxml,
        collector,
        "cx:axis/cx:majorTickMarks/@val"
      )
    ),
    minorTickMark: mapTickMark(
      validateEnum(
        childVal(el, "cx:minorTickMarks") ?? childVal(el, "cx:minorTickMark"),
        tickMarkOoxml,
        collector,
        "cx:axis/cx:minorTickMarks/@val"
      )
    ),
    numFmt: numFmt
      ? {
          // `formatCode` is a required attribute per
          // `CT_NumFmt/@formatCode`. When the source omits it,
          // default to `"General"` (Excel's behaviour) rather than
          // an empty string — the writer round-trips that default
          // without producing a schema-invalid empty attribute.
          formatCode: numFmt.attributes.formatCode || "General",
          sourceLinked: parseBoolAttr(numFmt, "sourceLinked")
        }
      : undefined,
    title: parseTitle(findChild(el, "cx:title")),
    // `<cx:majorGridlines>` / `<cx:minorGridlines>` each wrap a
    // single `<cx:spPr>` describing gridline stroke / colour. The
    // structured `ShapeProperties` is extracted via `parseRawSpPr`
    // so downstream consumers can read the colour without touching
    // raw XML. Absent element → `undefined` on the model.
    majorGridlines: parseGridlines(findChild(el, "cx:majorGridlines")),
    minorGridlines: parseGridlines(findChild(el, "cx:minorGridlines")),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr"),
    valScaling: valScaling
      ? {
          min: parseOptionalFloat(valScaling.attributes.min),
          max: parseOptionalFloat(valScaling.attributes.max),
          majorUnit: parseOptionalFloat(valScaling.attributes.majorUnit),
          minorUnit: parseOptionalFloat(valScaling.attributes.minorUnit)
        }
      : undefined,
    catScaling: catScaling
      ? {
          gapWidth: parseOptionalFloat(catScaling.attributes.gapWidth)
        }
      : undefined,
    // `<cx:units>` is a value-axis display-unit scaler (thousands /
    // millions / custom). The library doesn't yet interpret its
    // `cx:unit` / `cx:dispUnitsLbl` children, but preserving the raw
    // bytes keeps round-trip byte-equal for files that carry the
    // element. Without this the element evaporated on save.
    units: rawElement(el, "cx:units"),
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseLegend(
  el: XmlElement | undefined,
  collector?: UnknownCollector
): ChartLegend | undefined {
  if (!el) {
    return undefined;
  }
  const legendPositions = ["b", "l", "r", "t", "tr"] as const;
  const legendAligns = ["ctr", "l", "r", "t", "b"] as const;
  // `CT_Legend` (Chart2014) sequence: `legendEntry* → spPr? → txPr? →
  // extLst?`. The parser previously only read `pos` / `align` /
  // `overlay` / `spPr` / legendEntry idx; every other field was
  // silently dropped, breaking round-trip of legends with author-
  // supplied text styling, per-entry deletions, and vendor
  // extensions.
  return {
    legendPos: validateEnum(el.attributes.pos, legendPositions, collector, "cx:legend/@pos"),
    align: validateEnum(el.attributes.align, legendAligns, collector, "cx:legend/@align"),
    overlay: parseBoolAttr(el, "overlay"),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr"),
    extLst: rawElement(el, "cx:extLst"),
    legendEntries: findChildren(el, "cx:legendEntry").map(entry => ({
      index: parseIndexAttr(entry.attributes.idx),
      // Per `CT_LegendEntry` in the Chart2014 XSD, both `idx` and
      // `delete` are attributes (`xsd:unsignedInt` and `xsd:boolean`
      // respectively) — despite the type name `CT_Boolean` appearing
      // elsewhere in the schema for CHILD-element booleans. Excel's
      // own output is `<cx:legendEntry idx="0" delete="1"/>` with
      // attribute form, and the matching renderer emits the same
      // shape. Default per schema when `delete` is absent is `false`.
      delete: parseBoolAttr(entry, "delete"),
      txPr: parseRawTxPr(entry, "cx:txPr"),
      extLst: rawElement(entry, "cx:extLst")
    }))
  };
}

function parseTitle(el: XmlElement | undefined): ChartTitle | undefined {
  if (!el) {
    return undefined;
  }
  const txEl = findChild(el, "cx:tx");
  const text = parseRichText(findChild(txEl ?? emptyElement(), "cx:rich"));
  // `<cx:tx>` can hold EITHER `<cx:rich>` (structured rich text) OR
  // `<cx:txData>` (formula-backed text with cached value). The parser
  // previously only handled the rich variant, so a title authored as
  // `<cx:tx><cx:txData><cx:f>Sheet1!$A$1</cx:f><cx:v>Q1 Sales</cx:v></cx:txData></cx:tx>`
  // silently dropped the formula on round-trip. Map `txData` onto the
  // classic `ChartTitle.strRef` shape so writers / consumers can use
  // one API for both formula-linked titles.
  const txDataEl = findChild(txEl ?? emptyElement(), "cx:txData");
  const formula = txDataEl ? childText(txDataEl, "cx:f") : undefined;
  const cachedValue = txDataEl ? childText(txDataEl, "cx:v") : undefined;
  const strRef = formula
    ? {
        formula,
        cache:
          cachedValue !== undefined
            ? { points: [{ index: 0, value: cachedValue }] }
            : { points: [] }
      }
    : undefined;
  const overlay = parseBoolAttr(findChild(el, "cx:overlay"), "val");
  const spPr = parseRawSpPr(findChild(el, "cx:spPr"));
  const txPr = parseRawTxPr(el, "cx:txPr");
  const rawTx = rawElement(el, "cx:tx");
  // Return undefined for an empty `<cx:title/>` so downstream consumers
  // (snapshot tests, change detection) don't see a phantom title
  // object with all fields undefined. Parsing a self-closing title
  // element produces `rawTx` that's itself a self-closing tag —
  // treat that as equivalent to absent unless structured state is
  // present.
  if (
    text === undefined &&
    strRef === undefined &&
    overlay === undefined &&
    spPr === undefined &&
    txPr === undefined &&
    (rawTx === undefined || /^<cx:tx\s*\/>$/.test(rawTx.trim()))
  ) {
    return undefined;
  }
  return { text, strRef, overlay, spPr, txPr, rawTx };
}

function parseRichText(el: XmlElement | undefined): ChartRichText | undefined {
  if (!el) {
    return undefined;
  }
  const paragraphs = findChildren(el, "a:p").map(p => {
    // Paragraph properties — `<a:pPr>` with optional `<a:defRPr>` child
    // carrying default run formatting for the paragraph. Not fully
    // modelled in `ChartParagraph` today; capturing it as raw XML so
    // mutation-aware writers can still preserve author intent.
    const runs: ChartRichText["paragraphs"][number]["runs"] = findChildren(p, "a:r").map(r => {
      const runText = childText(r, "a:t") ?? "";
      // Parse `<a:rPr>` into structured `ChartTextProperties`. Wrapping
      // `<a:defRPr>` around the rPr attributes is the trick that lets
      // us reuse `parseTxPr`, which expects a `defRPr`-style XML block.
      // Previously the parser dropped run-level properties entirely —
      // any bold/italic/colour/font on a rich-text title or series
      // caption silently disappeared on round-trip unless the caller
      // explicitly used the raw-bytes path.
      const rPr = findChild(r, "a:rPr");
      if (!rPr) {
        return { text: runText };
      }
      const rPrRaw = rawElementFromNode(rPr);
      if (!rPrRaw) {
        return { text: runText };
      }
      // `parseTxPr` reads `<a:defRPr>` or `<a:rPr>` fragments from
      // `_rawXml`; the element name doesn't matter, the attribute /
      // child extraction does. Emit a thin `<a:defRPr>` envelope and
      // feed the inner bytes through.
      const innerBytes = rPrRaw
        .replace(/^<a:rPr\b/, "<a:defRPr")
        .replace(/<\/a:rPr>$/, "</a:defRPr>");
      const envelope = `<a:txPr><a:p>${innerBytes}</a:p></a:txPr>`;
      const parsed = parseTxPr({ _rawXml: envelope });
      if (!parsed || Object.keys(parsed).length === 0) {
        return { text: runText };
      }
      // Strip the `_rawXml` that `parseTxPr` may have carried through;
      // run properties are flat, not a full `ChartTextProperties`
      // wrapper.
      const { _rawXml, ...structured } = parsed as ChartTextProperties & { _rawXml?: string };
      return {
        text: runText,
        properties: structured as ChartTextProperties
      };
    });
    return { runs };
  });
  return paragraphs.length > 0 ? { paragraphs } : undefined;
}

function parseRawLayout(el: XmlElement | undefined): ChartTitle["layout"] | undefined {
  const raw = rawElementFromNode(el);
  // `ChartLayout._rawXml` was added so this path no longer needs `as any`.
  return raw ? { _rawXml: raw } : undefined;
}

function parseRawSpPr(el: XmlElement | undefined): ShapeProperties | undefined {
  const raw = rawElementFromNode(el);
  if (!raw) {
    return undefined;
  }
  // Dual representation: keep the original XML for lossless round-trip of
  // elements this parser does not structurally understand (a:xfrm,
  // a:prstGeom, a:custGeom, future DrawingML extensions…) while also
  // populating structured fields (fill / line / effectList / scene3d / sp3d)
  // so consumers can read and mutate them programmatically.
  //
  // Writers (`chart-ex-renderer.renderSpPr`, `chart-space-xform._renderSpPr`)
  // prefer `_rawXml` when present. Mutation APIs in `shape-properties.ts`
  // (`setSpPrFill`, `setSpPrLine`, `buildSpPr`) return a new object without
  // `_rawXml`, which causes the writer to take the structured path. This
  // matches the contract documented at `shape-properties.ts:927`.
  const structured = parseSpPr({ _rawXml: raw });
  return { ...structured, _rawXml: raw };
}

function parseRawTxPr(parent: XmlElement, name: string): ChartTextProperties | undefined {
  const raw = rawElement(parent, name);
  if (!raw) {
    return undefined;
  }
  // Same dual-representation trick as `parseRawSpPr`. `parseTxPr` extracts
  // font size / bold / italic / colour / font family / rotation from the raw
  // XML so callers can read and mutate text properties structurally; the
  // raw XML wins at write time unless the caller rebuilds via `buildTxPr`.
  const structured = parseTxPr({ _rawXml: raw });
  return { ...structured, _rawXml: raw };
}

function rawElement(parent: XmlElement, name: string): string | undefined {
  return rawElementFromNode(findChild(parent, name));
}

function rawElementFromNode(el: XmlElement | undefined): string | undefined {
  if (!el) {
    return undefined;
  }
  return serializeElement(el);
}

function serializeElement(el: XmlElement): string {
  const attrs = Object.entries(el.attributes ?? {})
    .map(([key, value]) => ` ${key}="${escapeXmlAttr(value)}"`)
    .join("");
  const children = el.children ?? [];
  if (children.length === 0) {
    return `<${el.name}${attrs}/>`;
  }
  return `<${el.name}${attrs}>${children.map(serializeNode).join("")}</${el.name}>`;
}

function serializeNode(node: XmlNode): string {
  if (node.type === "text") {
    return escapeXml(node.value ?? "");
  }
  if (node.type === "element") {
    return serializeElement(node);
  }
  if (node.type === "cdata") {
    return `<![CDATA[${node.value ?? ""}]]>`;
  }
  if (node.type === "comment") {
    return `<!--${node.value ?? ""}-->`;
  }
  if (node.type === "processing-instruction") {
    return `<?${node.target}${node.body ? ` ${node.body}` : ""}?>`;
  }
  return "";
}

function childText(el: XmlElement, name: string): string | undefined {
  const child = findChild(el, name);
  return child ? textContent(child) : undefined;
}

function childVal(el: XmlElement, name: string): string | undefined {
  return findChild(el, name)?.attributes.val;
}

function parseBoolAttr(el: XmlElement | undefined, name: string): boolean | undefined {
  const value = el?.attributes[name];
  if (value === undefined) {
    return undefined;
  }
  // OOXML `xsd:boolean` admits the four canonical forms. Hand-authored
  // or LibreOffice-produced XML uses the textual `"true"` / `"false"`;
  // the previous implementation only recognised numeric forms and
  // silently converted every unknown value to `false`.
  if (value === "1" || value === "true") {
    return true;
  }
  if (value === "0" || value === "false") {
    return false;
  }
  return undefined;
}

/**
 * Parse a Chart2014 `CT_Boolean` child element. The element's `val`
 * attribute defaults to `true` when absent — see ECMA-376 Part 1
 * §L.4.3.2.8. Callers like `<cx:connectorLines/>`, `<cx:paretoLine/>`,
 * `<cx:showMeanLine/>` depend on this default; previously the parser
 * required an explicit `val` attribute and silently dropped the field
 * otherwise, losing user intent for waterfall / box-and-whisker /
 * pareto charts on round-trip.
 *
 * Semantics by input:
 *   - Element missing entirely          → `undefined`
 *   - `<cx:foo/>` (no `val` attr)       → `true` (schema default)
 *   - `<cx:foo val="1"/>` / `"true"`    → `true`
 *   - `<cx:foo val="0"/>` / `"false"`   → `false`
 *   - `<cx:foo val="yes"/>` (garbage)   → `undefined` (don't promote)
 *
 * The last case previously coalesced to `true` via `?? true`, silently
 * treating an invalid attribute as opt-in. Garbage now falls through
 * so callers can distinguish "attribute absent" from "attribute
 * invalid" and react accordingly (typically: drop the field, leaving
 * the caller's default in place).
 */
function parseCtBoolean(el: XmlElement | undefined): boolean | undefined {
  if (!el) {
    return undefined;
  }
  // Distinguish "no val attribute" (schema default true) from
  // "val attribute present but unrecognised" (undefined).
  if (el.attributes.val === undefined) {
    return true;
  }
  return parseBoolAttr(el, "val");
}

/**
 * Narrow an attribute / child-element string to a known enum value.
 * When the input is missing returns `undefined`; when it isn't one of
 * the `allowed` variants returns `undefined` and appends a diagnostic
 * entry to the {@link UnknownCollector}, so callers can discover
 * unfamiliar OOXML values via `model.unknownElements` without replacing
 * them with an arbitrary fallback.
 */
function validateEnum<T extends string>(
  value: string | undefined,
  allowed: readonly T[],
  collector: UnknownCollector | undefined,
  path: string
): T | undefined {
  if (value === undefined) {
    return undefined;
  }
  if ((allowed as readonly string[]).includes(value)) {
    return value as T;
  }
  collector?.entries.push({ name: `${path}=${value}`, path });
  return undefined;
}

function parseOptionalInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = parseInt(value, 10);
  // `parseInt("Infinity", 10)` returns `NaN` — caught by the finite
  // check — but explicit finite testing guards against future
  // parser additions (e.g. `Number()` migration) still rejecting
  // non-finite slips. See `parseOptionalFloat`.
  return Number.isFinite(parsed) ? parsed : undefined;
}

/**
 * Parse a non-negative OOXML index attribute, falling back to `0`
 * when the value is absent / malformed. Use this for `idx`, `id`,
 * `dataId` and similar slots where `<c:pt idx="">` (empty string),
 * `<c:pt idx="abc">` (garbage), or `<c:pt idx="-1">` (negative)
 * should all coerce to 0 rather than land a poison value on the
 * model — `-1` silently corrupts dense arrays via property
 * assignment on the length slot, and `NaN` creates a string-indexed
 * property the rest of the code ignores.
 */
function parseIndexAttr(value: string | undefined): number {
  if (value === undefined || value === "") {
    return 0;
  }
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return 0;
  }
  return parsed;
}

function parseOptionalFloat(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = parseFloat(value);
  // Reject non-finite parses — `parseFloat("Infinity")` returns
  // `Infinity`, which propagates through axis scaling and collapses
  // the plot to an unrenderable degenerate range. Malformed input
  // should fall back to "absent" rather than poison downstream math.
  return Number.isFinite(parsed) ? parsed : undefined;
}

/**
 * Validate an OOXML `cx:strDim/@type` / `cx:numDim/@type` value
 * against {@link ChartExDimensionType}. Unknown values fall back to
 * `"val"` (the schema default) and are reported to the optional
 * {@link UnknownCollector} so `model.unknownElements` surfaces
 * vendor / forward-compat extensions instead of silently coercing
 * them into a real enum member.
 */
function parseDimensionType(
  value: string | undefined,
  collector?: UnknownCollector,
  path?: string
): ChartExDimensionType {
  const allowed = [
    "cat",
    "val",
    "x",
    "y",
    "size",
    "colorVal",
    "from",
    "to",
    "classification"
  ] as const;
  if (value === undefined) {
    return "val";
  }
  if ((allowed as readonly string[]).includes(value)) {
    return value as ChartExDimensionType;
  }
  const diagnosticPath = path ?? "cx:dim/@type";
  collector?.entries.push({
    name: `${diagnosticPath}=${value}`,
    path: diagnosticPath
  });
  return "val";
}

function emptyElement(): XmlElement {
  return { type: "element", name: "", attributes: {}, children: [] as XmlNode[] };
}
