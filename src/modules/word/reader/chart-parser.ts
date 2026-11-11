/**
 * DOCX Reader - Chart Parser (c: and cx: namespaces)
 *
 * Parses traditional and ChartEx (Microsoft 365) chart XML into Chart model
 * objects. Also resolves chart references in the document body, replacing
 * opaque chart drawings with parsed Chart instances.
 */

import { NS_A, NS_C_CHART, NS_R } from "@word/constants";
import { WRITTEN, readToken } from "@word/core/argument-checks";
import {
  axisKeys,
  chartGroupKey,
  chartKeys,
  chartSpaceKeys,
  groupDataLabelsKeys,
  legendKeys,
  plotAreaKeys,
  pointKey,
  scalingKeys,
  seriesDataLabelsKeys,
  seriesKeys,
  titleKeys
} from "@word/core/chart-kept";
import type { ChartFieldKeys } from "@word/core/chart-kept";
import type { Mutable } from "@word/core/internal-utils";
import { DocxLimitExceededError } from "@word/errors";
import { frameFieldsOf, readDrawingFrame } from "@word/reader/drawing-frame";
import {
  findChildLocal as findChartChild,
  findChildrenLocal as findAllChartChildren,
  localName
} from "@word/reader/parse-utils";
import { parsePartXml } from "@word/reader/part-xml";
import {
  namespacesInScope,
  preservedRootXml,
  serializeSelfContained
} from "@word/reader/preserved-xml";
import type {
  BodyContent,
  Chart,
  ChartAxis,
  ChartContent,
  ChartKept,
  ChartKeptAxis,
  ChartKeptElement,
  ChartKeptNode,
  ChartDataLabelPosition,
  ChartDataLabels,
  ChartErrorBarDirection,
  ChartErrorBarType,
  ChartErrorBars,
  ChartExContent,
  ChartExData,
  ChartExSeriesData,
  ChartLegendPosition,
  ChartSeries,
  ChartTrendline,
  ChartTrendlineType,
  ChartType,
  Emu,
  OpaqueDrawing,
  OpaqueRelationship,
  PreservedXml
} from "@word/types";
import { textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

/** The `ST_DLblPos` values the model names, by the part's value. */
const DATA_LABEL_POSITION_FROM_PART: Readonly<Record<string, ChartDataLabelPosition>> = {
  outEnd: "outsideEnd",
  ctr: "center",
  inEnd: "insideEnd",
  bestFit: "bestFit"
};
const DATA_LABEL_POSITIONS_READ: ReadonlySet<string> = new Set(
  Object.keys(DATA_LABEL_POSITION_FROM_PART)
);

// =============================================================================
// Chart Reader
// =============================================================================

/**
 * Recover the `<wp:extent cx="…" cy="…"/>` dimensions from the raw
 * `<w:drawing>` markup that was preserved on `OpaqueDrawing`. The
 * reader serialises the parent `wp:anchor` / `wp:inline` element when
 * it doesn't recognise the inner content; chart parts re-parse it
 * here so the size makes it onto the upgraded `ChartContent`.
 *
 * Returns `undefined` when no `<wp:extent>` element is present (the
 * drawing might be malformed) or when its dimensions are zero —
 * downstream callers then fall back to default sizing.
 *
 * Attribute order is unspecified in ECMA-376 / XML, so we extract
 * `cx` and `cy` independently.
 */
function extractDrawingExtent(rawXml: string): { width: Emu; height: Emu } | undefined {
  const extentMatch = /<wp:extent\b([^>]*)>/.exec(rawXml);
  if (!extentMatch) {
    return undefined;
  }
  const attrs = extentMatch[1];
  const cxMatch = /\bcx="(\d+)"/.exec(attrs);
  const cyMatch = /\bcy="(\d+)"/.exec(attrs);
  if (!cxMatch || !cyMatch) {
    return undefined;
  }
  const cx = parseInt(cxMatch[1], 10);
  const cy = parseInt(cyMatch[1], 10);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
    return undefined;
  }
  return { width: cx, height: cy };
}

/**
 * `content` with the name, alt text and frame of the drawing `rawXml` (a
 * `w:drawing`) it was read from, so the chart is written back in the frame it
 * had — anchored or inline, with its spacing and extent.
 */
function withChartFrame<T extends ChartContent | ChartExContent>(
  content: T,
  rawXml: string,
  size: (content: T) => { width?: number; height?: number }
): T {
  let frameEl: XmlElement | undefined;
  try {
    const root = parsePartXml(rawXml, undefined).root;
    frameEl = root.children.find(
      (c): c is XmlElement =>
        c.type === "element" && (c.name === "wp:inline" || c.name === "wp:anchor")
    );
  } catch {
    return content;
  }
  if (!frameEl) {
    return content;
  }
  const docPr = frameEl.children.find(
    (c): c is XmlElement => c.type === "element" && c.name === "wp:docPr"
  );
  const named: T = {
    ...content,
    ...(content.name === undefined && docPr?.attributes["name"] !== undefined
      ? { name: docPr.attributes["name"] }
      : {}),
    ...(content.altText === undefined && docPr?.attributes["descr"] !== undefined
      ? { altText: docPr.attributes["descr"] }
      : {})
  };
  const fields = { ...size(named), name: named.name, altText: named.altText };
  return { ...named, _frame: readDrawingFrame(frameEl, frameFieldsOf(fields)) };
}

/** Replace OpaqueDrawing items referencing chart rIds with ChartContent. */
function replaceOpaqueCharts(body: BodyContent[], chartRIdToChart: Map<string, Chart>): void {
  for (let i = 0; i < body.length; i++) {
    const item = body[i];
    if (item.type === "opaqueDrawing") {
      const chartRId = findChartRIdInDrawing(item, chartRIdToChart);
      if (chartRId) {
        const chart = chartRIdToChart.get(chartRId)!;
        // Recover the drawing extent so the chart is sized correctly
        // when laid out / rendered. The chart-writer reads
        // `chart.width` / `chart.height` to emit `<wp:extent>` so we
        // close the loop by writing them back into the same fields.
        // Don't clobber values the chart parser may have already
        // deduced from the inner `c:chart` payload.
        const extent = extractDrawingExtent(item.rawXml);
        const enrichedChart =
          extent && chart.width == null && chart.height == null
            ? { ...chart, width: extent.width, height: extent.height }
            : chart;
        body[i] = withChartFrame(
          { type: "chart", chart: enrichedChart } as ChartContent,
          item.rawXml,
          content => ({
            width: content.chart.width,
            height: content.chart.height
          })
        );
      }
    }
  }
}

/** Check if an OpaqueDrawing references a chart rId. */
function findChartRIdInDrawing(
  od: OpaqueDrawing,
  chartRIdToChart: Map<string, Chart>
): string | undefined {
  for (const rId of od.referencedRIds) {
    if (chartRIdToChart.has(rId)) {
      return rId;
    }
  }
  return undefined;
}

/** Replace OpaqueDrawing items referencing ChartEx rIds with ChartExContent. */
function replaceOpaqueChartExDrawings(
  body: BodyContent[],
  chartExRIdToContent: Map<string, ChartExContent>
): void {
  for (let i = 0; i < body.length; i++) {
    const item = body[i];
    if (item.type === "opaqueDrawing") {
      for (const rId of item.referencedRIds) {
        if (chartExRIdToContent.has(rId)) {
          const content = chartExRIdToContent.get(rId)!;
          body[i] = withChartFrame(content, item.rawXml, c => ({
            width: c.width,
            height: c.height
          }));
          break;
        }
      }
    }
  }
}

/** Chart type tag → our ChartType mapping (inverse of chart-writer's CHART_TYPE_CATEGORY). */
function resolveChartType(
  tag: string,
  grouping?: string,
  direction?: string,
  extra?: { radarStyle?: string; wireframe?: boolean }
): ChartType {
  switch (tag) {
    case "barChart":
    case "bar3DChart": {
      const stacked = grouping === "stacked";
      const percent = grouping === "percentStacked";
      if (direction === "bar") {
        return percent ? "barPercentStacked" : stacked ? "barStacked" : "bar";
      }
      return percent ? "columnPercentStacked" : stacked ? "columnStacked" : "column";
    }
    case "lineChart":
    case "line3DChart":
      return grouping === "stacked" ? "lineStacked" : "line";
    case "areaChart":
    case "area3DChart":
      return grouping === "stacked" ? "areaStacked" : "area";
    case "pieChart":
      return "pie";
    case "pie3DChart":
      return "pie3D";
    case "doughnutChart":
      return "doughnut";
    case "scatterChart":
      return "scatter";
    case "radarChart":
      if (extra?.radarStyle === "filled") {
        return "radarFilled";
      }
      return "radar";
    case "bubbleChart":
      return "bubble";
    case "stockChart":
      return "stock";
    case "surfaceChart":
      return extra?.wireframe ? "surfaceWireframe" : "surface";
    case "surface3DChart":
      return extra?.wireframe ? "surfaceWireframe3D" : "surface3D";
    default:
      return "column";
  }
}

/** The chart group elements the model reads, by local name. */
const CHART_TAGS: ReadonlySet<string> = new Set([
  "barChart",
  "bar3DChart",
  "lineChart",
  "line3DChart",
  "areaChart",
  "area3DChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "scatterChart",
  "radarChart",
  "bubbleChart",
  "stockChart",
  "surfaceChart",
  "surface3DChart"
]);

const AXIS_TAGS: ReadonlySet<string> = new Set(["catAx", "valAx", "dateAx", "serAx"]);

type Scope = Readonly<Record<string, string>>;

/** A chart group as read: its element, type, axes and series elements. */
interface ReadGroup {
  readonly el: XmlElement;
  readonly scope: Scope;
  readonly type: ChartType;
  readonly axIds: readonly string[];
  readonly secondary: boolean;
  readonly seriesEls: readonly XmlElement[];
}

function readGroupType(el: XmlElement, seriesEls: readonly XmlElement[]): ChartType {
  const tag = localName(el.name);
  const val = (name: string) => findChartChild(el, name)?.attributes["val"];
  const wireframe = val("wireframe") === "1" || val("wireframe") === "true";
  let type = resolveChartType(tag, val("grouping"), val("barDir"), {
    radarStyle: val("radarStyle"),
    wireframe
  });
  const marked = seriesEls.some(ser => {
    const symbol = findChartChild(findChartChild(ser, "marker") ?? ser, "symbol");
    return symbol !== undefined && symbol.attributes["val"] !== "none";
  });
  if (type === "line" && marked) {
    type = "lineMarked";
  }
  if (tag === "scatterChart" && val("scatterStyle") === "smoothMarker") {
    type = "scatterSmooth";
  }
  return type;
}

/** The `val` of `el`'s `c:style`, `c14:style` or `mc:AlternateContent`-wrapped style. */
function readStyle(root: XmlElement): number | undefined {
  const parse = (el: XmlElement | undefined): number | undefined => {
    const n = parseInt(el?.attributes["val"] ?? "", 10);
    return Number.isFinite(n) ? n : undefined;
  };
  const direct = parse(findChartChild(root, "style"));
  if (direct !== undefined) {
    return direct;
  }
  const alternate = findChartChild(root, "AlternateContent");
  if (!alternate) {
    return undefined;
  }
  return (
    parse(findChartChild(findChartChild(alternate, "Choice") ?? alternate, "style")) ??
    parse(findChartChild(findChartChild(alternate, "Fallback") ?? alternate, "style"))
  );
}

/**
 * Parse a chart part XML string into a Chart object. Every element the model
 * has no field for is kept on `_kept` (see `ChartKept`), with
 * `relationships`, the part's own relationships.
 */
function parseChartXml(
  xmlStr: string,
  maxDepth?: number,
  relationships?: readonly OpaqueRelationship[]
): Chart | undefined {
  let doc: ReturnType<typeof parsePartXml>;
  try {
    doc = parsePartXml(xmlStr, maxDepth);
  } catch (error) {
    // A resource limit is a verdict on the package, not a malformed chart.
    if (error instanceof DocxLimitExceededError) {
      throw error;
    }
    return undefined;
  }

  const root = doc.root; // c:chartSpace
  const chartEl = findChartChild(root, "chart");
  if (!chartEl) {
    return undefined;
  }
  const plotAreaEl = findChartChild(chartEl, "plotArea");
  if (!plotAreaEl) {
    return undefined;
  }
  const rootScope = namespacesInScope({}, root);
  const keeper = createKeeper(rootScope);
  const chartScope = namespacesInScope(rootScope, chartEl);
  const plotAreaScope = namespacesInScope(chartScope, plotAreaEl);

  // Chart groups: the first gives the chart's type and primary axes; a group
  // on other axes is on the secondary ones.
  const groups: ReadGroup[] = [];
  for (const el of elementChildren(plotAreaEl)) {
    if (!CHART_TAGS.has(localName(el.name))) {
      continue;
    }
    const seriesEls = elementChildren(el).filter(child => localName(child.name) === "ser");
    const axIds = elementChildren(el)
      .filter(child => localName(child.name) === "axId")
      .map(child => child.attributes["val"] ?? "");
    const primary = groups[0]?.axIds;
    const secondary = primary !== undefined && (axIds[0] !== primary[0] || axIds[1] !== primary[1]);
    groups.push({
      el,
      scope: namespacesInScope(plotAreaScope, el),
      type: readGroupType(el, seriesEls),
      axIds,
      secondary,
      seriesEls
    });
  }
  const first = groups[0];
  if (!first) {
    return undefined;
  }
  const chartType = first.type;

  // Series, with the type and axis of their group where it is not the chart's.
  const series: { series: ChartSeries; el: XmlElement; group: ReadGroup }[] = [];
  for (const group of groups) {
    for (const el of group.seriesEls) {
      const s = parseChartSeries(el, localName(group.el.name));
      if (s) {
        series.push({
          series: {
            ...s,
            ...(group.type !== chartType && { chartType: group.type }),
            ...(group.secondary && { plotOnSecondaryAxis: true })
          },
          el,
          group
        });
      }
    }
  }

  // Axes, by the role their id plays in the groups.
  const secondaryIds = groups.find(g => g.secondary)?.axIds;
  const roleOf = (id: string): ChartKeptAxis["role"] =>
    id === first.axIds[0]
      ? "category"
      : id === first.axIds[1]
        ? "value"
        : secondaryIds && id === secondaryIds[0]
          ? "secondaryCategory"
          : secondaryIds && id === secondaryIds[1]
            ? "secondaryValue"
            : undefined;
  const axisEls = elementChildren(plotAreaEl)
    .filter(el => AXIS_TAGS.has(localName(el.name)))
    .map(el => {
      const id = findChartChild(el, "axId")?.attributes["val"] ?? "";
      const role = localName(el.name) === "serAx" ? undefined : roleOf(id);
      return {
        el,
        id,
        role,
        axis: role === "secondaryCategory" ? undefined : parseChartAxis(el)
      };
    });
  const axisFor = (role: ChartKeptAxis["role"]) => axisEls.find(a => a.role === role)?.axis;
  const categoryAxis = axisFor("category");
  const valueAxis = axisFor("value");
  const secondaryValueAxis = axisFor("secondaryValue");

  // Title
  const titleEl = findChartChild(chartEl, "title");
  const title = titleEl ? extractChartText(titleEl) : undefined;

  // 3D view
  const view3d = !!findChartChild(chartEl, "view3D");

  // Chart-level data labels (from the first group)
  const dataLabels = readGroupDataLabels(findChartChild(first.el, "dLbls"));

  // Legend
  const legendEl = findChartChild(chartEl, "legend");
  // A chart without `c:legend` shows none: read as "none", or the writer's
  // default would give it one.
  let legend: ChartLegendPosition | undefined = "none";
  const unrecognized: object = {};
  if (legendEl) {
    const posEl = findChartChild(legendEl, "legendPos");
    legend = readToken(
      posEl?.attributes["val"] ?? "r",
      WRITTEN.legendPosition,
      unrecognized,
      "legend"
    ) as ChartLegendPosition | undefined;
  }

  // Plot area color
  const plotSpPrEl = findChartChild(plotAreaEl, "spPr");
  const plotAreaColor = plotSpPrEl ? extractSolidFillColor(plotSpPrEl) : undefined;

  // Chart area color (from c:chartSpace > c:spPr)
  const chartSpPrEl = findChartChild(root, "spPr");
  const chartAreaColor = chartSpPrEl ? extractSolidFillColor(chartSpPrEl) : undefined;

  const style = readStyle(root);

  const chart: Chart = {
    ...unrecognized,
    type: chartType,
    series: series.map(s => s.series),
    ...(title !== undefined && { title }),
    ...(legend !== undefined && { legend }),
    ...(dataLabels !== undefined && { dataLabels }),
    ...(categoryAxis !== undefined && { categoryAxis }),
    ...(valueAxis !== undefined && { valueAxis }),
    ...(secondaryValueAxis !== undefined && { secondaryValueAxis }),
    ...(plotAreaColor !== undefined && { plotAreaColor }),
    ...(chartAreaColor !== undefined && { chartAreaColor }),
    ...(view3d && { view3d }),
    ...(style !== undefined && { style })
  };

  // Everything else, keyed by the fields just read (see `core/chart-kept.ts`).
  const keptSeries = series.map(({ series: s, el, group }, index) => {
    const node = keeper.node(el, group.scope, ["dLbls"], seriesKeys(s, index, group.type));
    const serScope = namespacesInScope(group.scope, el);
    const dLblsEl = findChartChild(el, "dLbls");
    const dataLabelsNode = dLblsEl
      ? keeper.node(dLblsEl, serScope, [], seriesDataLabelsKeys(s))
      : undefined;
    const withPoints: ChartKeptNode = {
      ...node,
      children: node.children.map((child, i) => {
        if (child.name !== "dPt") {
          return child;
        }
        const dPt = elementChildren(el)[i]!;
        const pointIdx = pointIndex(dPt);
        const spPr = findChartChild(dPt, "spPr");
        const spPrXml = spPr
          ? keeper.element(spPr, namespacesInScope(serScope, dPt)).xml
          : undefined;
        return {
          ...child,
          ...(pointIdx !== undefined && {
            index: pointIdx,
            read: pointKey(s, pointIdx)
          }),
          ...(spPrXml !== undefined && { spPr: spPrXml })
        };
      })
    };
    return {
      ...s,
      _kept: {
        node: withPoints,
        ...(dataLabelsNode && { dataLabels: dataLabelsNode })
      }
    };
  });

  const legendNode = legendEl
    ? keeper.node(legendEl, chartScope, [], legendKeys(chart))
    : undefined;
  const titleNode = titleEl ? keeper.node(titleEl, chartScope, [], titleKeys(title)) : undefined;
  const declarations = rootDeclarations(root);
  const kept: ChartKept = {
    ...(declarations && { root: declarations }),
    space: keeper.node(root, {}, ["chart"], chartSpaceKeys(chart)),
    chart: keeper.node(chartEl, rootScope, ["title", "plotArea", "legend"], chartKeys(chart)),
    ...(titleNode && { title: titleNode }),
    plotArea: keeper.node(
      plotAreaEl,
      chartScope,
      [...CHART_TAGS, ...AXIS_TAGS],
      plotAreaKeys(chart)
    ),
    ...(legendNode && { legend: legendNode }),
    groups: groups.map(group => {
      const dLblsEl = findChartChild(group.el, "dLbls");
      return {
        key: chartGroupKey(group.type, group.secondary),
        node: keeper.node(group.el, plotAreaScope, ["ser", "dLbls"], {}),
        ...(dLblsEl && {
          dataLabels: keeper.node(dLblsEl, group.scope, [], groupDataLabelsKeys(dataLabels))
        })
      };
    }),
    axes: axisEls.map(({ el, id, role, axis }) => {
      if (role === undefined) {
        const verbatim = keeper.element(el, plotAreaScope);
        return {
          id,
          node: keeper.node(el, plotAreaScope, [], {}),
          xml: verbatim.xml,
          ...(verbatim.rIds && { rIds: verbatim.rIds })
        };
      }
      const axisScope = namespacesInScope(plotAreaScope, el);
      const scalingEl = findChartChild(el, "scaling");
      const axisTitleEl = findChartChild(el, "title");
      return {
        id,
        role,
        node: keeper.node(el, plotAreaScope, ["scaling", "title"], axisKeys(axis, role)),
        ...(scalingEl && {
          scaling: keeper.node(scalingEl, axisScope, [], scalingKeys(axis))
        }),
        ...(axisTitleEl && {
          title: keeper.node(axisTitleEl, axisScope, [], titleKeys(axis?.title))
        })
      };
    }),
    ...(relationships && relationships.length > 0 && { relationships })
  };

  return { ...chart, series: keptSeries, _kept: kept };
}

/** The chart-level data labels a group's `c:dLbls` sets. */
function readGroupDataLabels(chartDLblsEl: XmlElement | undefined): ChartDataLabels | undefined {
  if (!chartDLblsEl) {
    return undefined;
  }
  const dl: Mutable<ChartDataLabels> = {};
  let hasDl = false;
  const on = (name: string) => findChartChild(chartDLblsEl, name)?.attributes["val"] === "1";
  if (on("showVal")) {
    dl.showValue = true;
    hasDl = true;
  }
  if (on("showCatName")) {
    dl.showCategory = true;
    hasDl = true;
  }
  if (on("showSerName")) {
    dl.showSerName = true;
    hasDl = true;
  }
  if (on("showPercent")) {
    dl.showPercent = true;
    hasDl = true;
  }
  const dlPosEl = findChartChild(chartDLblsEl, "dLblPos");
  if (dlPosEl?.attributes["val"]) {
    // The part holds `ST_DLblPos`; the model names four of its values.
    const position = readToken(
      dlPosEl.attributes["val"],
      DATA_LABEL_POSITIONS_READ,
      dl,
      "position"
    );
    if (position !== undefined) {
      dl.position = DATA_LABEL_POSITION_FROM_PART[position];
    }
    hasDl = true;
  }
  return hasDl ? dl : undefined;
}

/** Extract text from a chart title element. */
function extractChartText(titleEl: XmlElement): string | undefined {
  const txEl = findChartChild(titleEl, "tx");
  if (!txEl) {
    return undefined;
  }
  const richEl = findChartChild(txEl, "rich");
  if (!richEl) {
    return undefined;
  }

  let text = "";
  for (const child of richEl.children) {
    if (child.type !== "element") {
      continue;
    }
    if (localName(child.name) !== "p") {
      continue;
    }
    for (const rChild of child.children) {
      if (rChild.type !== "element") {
        continue;
      }
      if (localName(rChild.name) !== "r") {
        continue;
      }
      const tEl = findChartChild(rChild, "t");
      if (tEl) {
        text += textContent(tEl);
      }
    }
  }
  return text || undefined;
}

/** Parse a chart series element. */
function parseChartSeries(serEl: XmlElement, chartTag: string): ChartSeries | undefined {
  // Series name from c:tx
  let name = "";
  const txEl = findChartChild(serEl, "tx");
  if (txEl) {
    const strRefEl = findChartChild(txEl, "strRef");
    if (strRefEl) {
      const cacheEl = findChartChild(strRefEl, "strCache");
      if (cacheEl) {
        const ptEl = findChartChild(cacheEl, "pt");
        if (ptEl) {
          const vEl = findChartChild(ptEl, "v");
          if (vEl) {
            name = textContent(vEl);
          }
        }
      }
    }
    if (!name) {
      const vEl = findChartChild(txEl, "v");
      if (vEl) {
        name = textContent(vEl);
      }
    }
  }

  // Categories (c:cat or c:xVal)
  const categories: string[] = [];
  const catEl = findChartChild(serEl, "cat") ?? findChartChild(serEl, "xVal");
  if (catEl) {
    const strRef = findChartChild(catEl, "strRef");
    const numRef = findChartChild(catEl, "numRef");
    const refEl = strRef ?? numRef;
    if (refEl) {
      const cacheEl = findChartChild(refEl, "strCache") ?? findChartChild(refEl, "numCache");
      if (cacheEl) {
        extractPointValues(cacheEl, categories);
      }
    }
    // Fallback: inline string literal
    if (categories.length === 0) {
      const strLitEl = findChartChild(catEl, "strLit") ?? findChartChild(catEl, "numLit");
      if (strLitEl) {
        extractPointValues(strLitEl, categories);
      }
    }
  }

  // Values (c:val or c:yVal)
  const values: number[] = [];
  const valEl = findChartChild(serEl, "val") ?? findChartChild(serEl, "yVal");
  if (valEl) {
    const numRef = findChartChild(valEl, "numRef");
    if (numRef) {
      const cacheEl = findChartChild(numRef, "numCache");
      if (cacheEl) {
        const strVals: string[] = [];
        extractPointValues(cacheEl, strVals);
        for (const v of strVals) {
          const n = parseFloat(v);
          values.push(Number.isFinite(n) ? n : 0);
        }
      }
    }
    if (values.length === 0) {
      const numLitEl = findChartChild(valEl, "numLit");
      if (numLitEl) {
        const strVals: string[] = [];
        extractPointValues(numLitEl, strVals);
        for (const v of strVals) {
          const n = parseFloat(v);
          values.push(Number.isFinite(n) ? n : 0);
        }
      }
    }
  }

  // Series color
  const spPrEl = findChartChild(serEl, "spPr");
  const color = spPrEl ? extractSolidFillColor(spPrEl) : undefined;

  // Per-point colors, by `c:idx`: entry i is point i's colour, a point with
  // none of its own is a hole. An index past the series' points colours
  // nothing (and would grow the array without bound); the `c:dPt` is kept.
  const pointCount = Math.max(values.length, categories.length);
  const pointColors: string[] = [];
  let hasPointColor = false;
  for (const child of serEl.children) {
    if (child.type !== "element" || localName(child.name) !== "dPt") {
      continue;
    }
    const index = pointIndex(child);
    const ptSpPr = findChartChild(child, "spPr");
    const c = ptSpPr ? extractSolidFillColor(ptSpPr) : undefined;
    if (c !== undefined && index !== undefined && index < pointCount) {
      pointColors[index] = c;
      hasPointColor = true;
    }
  }

  // Data labels
  let showDataLabels = false;
  const dLblsEl = findChartChild(serEl, "dLbls");
  if (dLblsEl) {
    const showValEl = findChartChild(dLblsEl, "showVal");
    if (showValEl?.attributes["val"] === "1") {
      showDataLabels = true;
    }
  }

  // Trendline
  let trendline: ChartTrendline | undefined;
  const trendlineEl = findChartChild(serEl, "trendline");
  if (trendlineEl) {
    const trendlineTypeEl = findChartChild(trendlineEl, "trendlineType");
    const trendlineType = (trendlineTypeEl?.attributes["val"] ?? "linear") as ChartTrendlineType;
    const orderEl = findChartChild(trendlineEl, "order");
    const order = orderEl?.attributes["val"] ? parseInt(orderEl.attributes["val"], 10) : undefined;
    const periodEl = findChartChild(trendlineEl, "period");
    const period = periodEl?.attributes["val"]
      ? parseInt(periodEl.attributes["val"], 10)
      : undefined;
    const dispEqEl = findChartChild(trendlineEl, "dispEq");
    const displayEquation = dispEqEl?.attributes["val"] === "1";
    const dispRSqEl = findChartChild(trendlineEl, "dispRSqr");
    const displayRSquared = dispRSqEl?.attributes["val"] === "1";
    trendline = {
      type: trendlineType,
      ...(order !== undefined && Number.isFinite(order) && { order }),
      ...(period !== undefined && Number.isFinite(period) && { period }),
      ...(displayEquation && { displayEquation }),
      ...(displayRSquared && { displayRSquared })
    };
  }

  // Error bars
  let errorBars: ChartErrorBars | undefined;
  const errBarsEl = findChartChild(serEl, "errBars");
  if (errBarsEl) {
    const errBarTypeEl = findChartChild(errBarsEl, "errBarType");
    const errValTypeEl = findChartChild(errBarsEl, "errValType");
    const errDirEl = findChartChild(errBarsEl, "errDir");
    const valEl2 = findChartChild(errBarsEl, "val");
    const direction = (errDirEl?.attributes["val"] ?? "y") as ChartErrorBarDirection;
    const errType = (errValTypeEl?.attributes["val"] ?? "fixedVal") as ChartErrorBarType;
    const errValue = valEl2?.attributes["val"] ? parseFloat(valEl2.attributes["val"]) : undefined;
    errorBars = {
      direction,
      type: errType,
      ...(errValue !== undefined && Number.isFinite(errValue) && { value: errValue })
    };
    // Suppress unused variable warning for errBarTypeEl
    void errBarTypeEl;
  }

  return {
    name,
    categories,
    values,
    ...(color !== undefined && { color }),
    ...(hasPointColor && { pointColors }),
    ...(showDataLabels && { showDataLabels }),
    ...(trendline !== undefined && { trendline }),
    ...(errorBars !== undefined && { errorBars })
  };
}

/** Extract point values (c:pt > c:v) from a cache/literal element. */
function extractPointValues(cacheEl: XmlElement, out: string[]): void {
  const pts: { idx: number; val: string }[] = [];
  for (const child of cacheEl.children) {
    if (child.type !== "element") {
      continue;
    }
    if (localName(child.name) !== "pt") {
      continue;
    }
    const idxStr = child.attributes["idx"];
    const idx = idxStr !== undefined ? parseInt(idxStr, 10) : pts.length;
    const vEl = findChartChild(child, "v");
    const val = vEl ? textContent(vEl) : "";
    pts.push({ idx: Number.isFinite(idx) ? idx : pts.length, val });
  }
  // Sort by index and fill output
  pts.sort((a, b) => a.idx - b.idx);
  for (const pt of pts) {
    // Fill gaps with empty strings
    while (out.length < pt.idx) {
      out.push("");
    }
    out.push(pt.val);
  }
}

/** Parse a chart axis element. */
function parseChartAxis(axEl: XmlElement): ChartAxis | undefined {
  const axis: Mutable<ChartAxis> = {};
  let hasAny = false;

  // Title
  const titleEl = findChartChild(axEl, "title");
  if (titleEl) {
    const t = extractChartText(titleEl);
    if (t) {
      axis.title = t;
      hasAny = true;
    }
  }

  // Hidden
  const deleteEl = findChartChild(axEl, "delete");
  if (deleteEl?.attributes["val"] === "1") {
    axis.hidden = true;
    hasAny = true;
  }

  // Scaling
  const scalingEl = findChartChild(axEl, "scaling");
  if (scalingEl) {
    const minEl = findChartChild(scalingEl, "min");
    if (minEl?.attributes["val"]) {
      const n = parseFloat(minEl.attributes["val"]);
      if (Number.isFinite(n)) {
        axis.min = n;
        hasAny = true;
      }
    }
    const maxEl = findChartChild(scalingEl, "max");
    if (maxEl?.attributes["val"]) {
      const n = parseFloat(maxEl.attributes["val"]);
      if (Number.isFinite(n)) {
        axis.max = n;
        hasAny = true;
      }
    }
  }

  // Number format
  const numFmtEl = findChartChild(axEl, "numFmt");
  if (numFmtEl) {
    const fc = numFmtEl.attributes["formatCode"];
    if (fc && fc !== "General") {
      axis.numberFormat = fc;
      hasAny = true;
    }
  }

  // Major unit
  const majorUnitEl = findChartChild(axEl, "majorUnit");
  if (majorUnitEl?.attributes["val"]) {
    const n = parseFloat(majorUnitEl.attributes["val"]);
    if (Number.isFinite(n)) {
      axis.majorUnit = n;
      hasAny = true;
    }
  }

  return hasAny ? axis : undefined;
}

/** The namespaces the chart writer declares on `c:chartSpace` itself. */
const WRITER_ROOT_NAMESPACES: Readonly<Record<string, string>> = {
  c: NS_C_CHART,
  a: NS_A,
  r: NS_R
};

function elementChildren(el: XmlElement): XmlElement[] {
  return el.children.filter((child): child is XmlElement => child.type === "element");
}

/** A `c:dPt`'s `c:idx`, or `undefined` when it has no usable one. */
function pointIndex(dPt: XmlElement): number | undefined {
  const raw = findChartChild(dPt, "idx")?.attributes["val"];
  const index = raw === undefined || !/^\d+$/.test(raw) ? Number.NaN : Number(raw);
  return Number.isSafeInteger(index) ? index : undefined;
}

/** The part root's namespace declarations and `mc:Ignorable`, which the writer re-declares. */
function rootDeclarations(root: XmlElement): PreservedXml | undefined {
  return preservedRootXml(root, () => true);
}

/**
 * Serializes kept chart elements. Each declares the namespace prefixes it
 * uses that the written root will not — the writer declares `c`, `a`, `r`
 * and the source root's own declarations — and records the relationships it
 * refers to, so the writer can drop it when a target is gone.
 */
function createKeeper(rootScope: Scope) {
  const written: Record<string, string> = {
    ...rootScope,
    ...WRITER_ROOT_NAMESPACES
  };
  const element = (el: XmlElement, scope: Scope): ChartKeptElement => {
    const inScope = namespacesInScope(scope, el);
    const undeclared: Record<string, string> = {};
    for (const [prefix, uri] of Object.entries(scope)) {
      if (written[prefix] !== uri) {
        undeclared[prefix] = uri;
      }
    }
    const rIds = new Set<string>();
    collectRelationshipIds(el, inScope, rIds);
    return {
      name: localName(el.name),
      xml: serializeSelfContained(el, undeclared),
      ...(rIds.size > 0 && { rIds: [...rIds] })
    };
  };
  const node = (
    el: XmlElement,
    outerScope: Scope,
    containers: readonly string[],
    read: ChartFieldKeys
  ): ChartKeptNode => {
    const scope = namespacesInScope(outerScope, el);
    const attrs = Object.fromEntries(
      Object.entries(el.attributes).filter(
        ([name]) => name !== "xmlns" && !name.startsWith("xmlns:")
      )
    );
    return {
      name: el.name,
      ...(Object.keys(attrs).length > 0 && { attrs }),
      children: elementChildren(el).map(child =>
        containers.includes(localName(child.name))
          ? { name: localName(child.name), xml: "" }
          : element(child, scope)
      ),
      read
    };
  };
  return { element, node };
}

/** The values of every attribute under `el` in the relationships namespace. */
function collectRelationshipIds(el: XmlElement, outerScope: Scope, out: Set<string>): void {
  const scope = namespacesInScope(outerScope, el);
  for (const [name, value] of Object.entries(el.attributes)) {
    const colon = name.indexOf(":");
    if (colon > 0 && name.slice(0, colon) !== "xmlns" && scope[name.slice(0, colon)] === NS_R) {
      out.add(value);
    }
  }
  for (const child of elementChildren(el)) {
    collectRelationshipIds(child, scope, out);
  }
}

/** Extract solid fill color from a spPr element. */
function extractSolidFillColor(spPrEl: XmlElement): string | undefined {
  const solidFillEl = findChartChild(spPrEl, "solidFill");
  if (!solidFillEl) {
    return undefined;
  }
  const srgbEl = findChartChild(solidFillEl, "srgbClr");
  return srgbEl?.attributes["val"] ?? undefined;
}

// =============================================================================
// ChartEx Reader (cx: namespace)
// =============================================================================

/**
 * Parse a ChartEx XML string (cx:chartSpace) into structured ChartExData.
 * Returns undefined if parsing fails or the structure is unrecognizable.
 */
function parseChartExXml(xmlStr: string, maxDepth?: number): ChartExData | undefined {
  let doc: ReturnType<typeof parsePartXml>;
  try {
    doc = parsePartXml(xmlStr, maxDepth);
  } catch (error) {
    // A resource limit is a verdict on the package, not a malformed chart.
    if (error instanceof DocxLimitExceededError) {
      throw error;
    }
    return undefined;
  }

  const root = doc.root; // cx:chartSpace
  if (!root) {
    return undefined;
  }

  // Find cx:chart
  const chartEl = findChartChild(root, "chart");
  if (!chartEl) {
    return undefined;
  }

  // Extract title from cx:chart > cx:title
  let title: string | undefined;
  const titleEl = findChartChild(chartEl, "title");
  if (titleEl) {
    title = extractChartExTitleText(titleEl);
  }

  // Find cx:plotArea > cx:plotAreaRegion > cx:series
  const plotAreaEl = findChartChild(chartEl, "plotArea");
  if (!plotAreaEl) {
    return undefined;
  }

  const plotAreaRegionEl = findChartChild(plotAreaEl, "plotAreaRegion");
  if (!plotAreaRegionEl) {
    return undefined;
  }

  // Determine chart type from the plotAreaRegion's series layoutId or from series elements
  let chartType = "";
  const seriesElements = findAllChartChildren(plotAreaRegionEl, "series");

  // The chart type is typically in the series @layoutId attribute
  if (seriesElements.length > 0) {
    chartType = seriesElements[0].attributes["layoutId"] ?? "";
  }

  if (!chartType) {
    return undefined;
  }

  // Parse cx:chartData to build a map of dataId → { strDim, numDim }
  const chartDataEl = findChartChild(root, "chartData");
  const dataMap = new Map<string, { strings: string[]; numbers: number[] }>();
  if (chartDataEl) {
    const dataElements = findAllChartChildren(chartDataEl, "data");
    for (const dataEl of dataElements) {
      const dataId = dataEl.attributes["id"];
      if (dataId === undefined) {
        continue;
      }
      const strings: string[] = [];
      const numbers: number[] = [];

      // cx:strDim — string dimension (categories)
      const strDimEl = findChartChild(dataEl, "strDim");
      if (strDimEl) {
        extractChartExDimValues(strDimEl, strings);
      }

      // cx:numDim — numeric dimension (values)
      const numDimEl = findChartChild(dataEl, "numDim");
      if (numDimEl) {
        const numStrings: string[] = [];
        extractChartExDimValues(numDimEl, numStrings);
        for (const v of numStrings) {
          const n = parseFloat(v);
          numbers.push(Number.isFinite(n) ? n : 0);
        }
      }

      dataMap.set(dataId, { strings, numbers });
    }
  }

  // Build series data
  const series: ChartExSeriesData[] = [];
  for (const serEl of seriesElements) {
    const seriesData = parseChartExSeries(serEl, dataMap);
    if (seriesData) {
      series.push(seriesData);
    }
  }

  return {
    chartType,
    ...(title !== undefined && { title }),
    series
  };
}

/** Extract text content from a cx:title element. */
function extractChartExTitleText(titleEl: XmlElement): string | undefined {
  // cx:title > cx:tx > cx:rich > a:p > a:r > a:t
  const txEl = findChartChild(titleEl, "tx");
  if (!txEl) {
    return undefined;
  }
  const richEl = findChartChild(txEl, "rich");
  if (!richEl) {
    return undefined;
  }

  let text = "";
  for (const child of richEl.children) {
    if (child.type !== "element") {
      continue;
    }
    if (localName(child.name) === "p") {
      for (const rChild of child.children) {
        if (rChild.type !== "element") {
          continue;
        }
        if (localName(rChild.name) === "r") {
          const tEl = findChartChild(rChild, "t");
          if (tEl) {
            text += textContent(tEl);
          }
        }
      }
    }
  }

  return text || undefined;
}

/** Parse a single cx:series element into ChartExSeriesData. */
function parseChartExSeries(
  serEl: XmlElement,
  dataMap: Map<string, { strings: string[]; numbers: number[] }>
): ChartExSeriesData | undefined {
  // Series name: cx:series > cx:tx > cx:txData > cx:v
  let name: string | undefined;
  const txEl = findChartChild(serEl, "tx");
  if (txEl) {
    const txDataEl = findChartChild(txEl, "txData");
    if (txDataEl) {
      const vEl = findChartChild(txDataEl, "v");
      if (vEl) {
        name = textContent(vEl) || undefined;
      }
    }
  }

  // Resolve data reference: cx:series > cx:dataId @val
  const dataIdEl = findChartChild(serEl, "dataId");
  const dataId = dataIdEl?.attributes["val"];

  let categories: string[] | undefined;
  let values: number[] | undefined;

  if (dataId !== undefined && dataMap.has(dataId)) {
    const data = dataMap.get(dataId)!;
    if (data.strings.length > 0) {
      categories = data.strings;
    }
    if (data.numbers.length > 0) {
      values = data.numbers;
    }
  }

  return {
    ...(name !== undefined && { name }),
    ...(categories !== undefined && { categories }),
    ...(values !== undefined && { values })
  };
}

/** Extract values from a cx:strDim or cx:numDim element. */
function extractChartExDimValues(dimEl: XmlElement, out: string[]): void {
  // Possible structures:
  // cx:strDim/cx:numDim > cx:lvl > cx:pt (with idx attribute) containing text
  // cx:strDim/cx:numDim > cx:f (formula reference, skip)
  // Look for cx:lvl first (level-based categories)
  const lvlEl = findChartChild(dimEl, "lvl");
  if (lvlEl) {
    const pts: { idx: number; val: string }[] = [];
    for (const child of lvlEl.children) {
      if (child.type !== "element" || localName(child.name) !== "pt") {
        continue;
      }
      const idxStr = child.attributes["idx"];
      const idx = idxStr !== undefined ? parseInt(idxStr, 10) : pts.length;
      const val = textContent(child);
      pts.push({ idx: Number.isFinite(idx) ? idx : pts.length, val });
    }
    pts.sort((a, b) => a.idx - b.idx);
    for (const pt of pts) {
      while (out.length < pt.idx) {
        out.push("");
      }
      out.push(pt.val);
    }
  }
}

export { replaceOpaqueCharts, replaceOpaqueChartExDrawings, parseChartXml, parseChartExXml };
