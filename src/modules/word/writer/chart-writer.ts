/**
 * DOCX Module - Chart Writer
 *
 * Generates DrawingML chart part XML (word/charts/chartN.xml) from
 * high-level Chart definitions.
 */

import { unrecognizedValue } from "@utils/unrecognized-values";
import { NS_A, NS_C_CHART, NS_R, STD_DOC_ATTRIBUTES } from "@word/constants";
import { MODEL_CHECK, WRITTEN, checkOneOf, writtenToken } from "@word/core/argument-checks";
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
import { drawingInteger } from "@word/core/drawing-geometry";
import { DocxRawXmlPolicyError } from "@word/errors";
import type { WordSecurityPolicy } from "@word/security/policy";
import type {
  Chart,
  ChartAxis,
  ChartDataLabels,
  ChartKept,
  ChartKeptAxis,
  ChartKeptElement,
  ChartKeptNode,
  ChartSeries,
  ChartType
} from "@word/types";
import { withPreservedNamespaces } from "@word/writer/preserved-xml";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

const CHART_TYPE_CATEGORY = {
  bar: { tag: "barChart", direction: "bar", grouping: "clustered" },
  barStacked: { tag: "barChart", direction: "bar", grouping: "stacked" },
  barPercentStacked: { tag: "barChart", direction: "bar", grouping: "percentStacked" },
  column: { tag: "barChart", direction: "col", grouping: "clustered" },
  columnStacked: { tag: "barChart", direction: "col", grouping: "stacked" },
  columnPercentStacked: { tag: "barChart", direction: "col", grouping: "percentStacked" },
  line: { tag: "lineChart", grouping: "standard" },
  lineStacked: { tag: "lineChart", grouping: "stacked" },
  lineMarked: { tag: "lineChart", grouping: "standard", marked: true },
  area: { tag: "areaChart", grouping: "standard" },
  areaStacked: { tag: "areaChart", grouping: "stacked" },
  pie: { tag: "pieChart" },
  pie3D: { tag: "pie3DChart" },
  doughnut: { tag: "doughnutChart" },
  scatter: { tag: "scatterChart", scatterStyle: "marker" },
  scatterSmooth: { tag: "scatterChart", scatterStyle: "smoothMarker" },
  radar: { tag: "radarChart" },
  radarFilled: { tag: "radarChart", radarStyle: "filled" },
  bubble: { tag: "bubbleChart" },
  stock: { tag: "stockChart" },
  surface: { tag: "surfaceChart" },
  surface3D: { tag: "surface3DChart" },
  surfaceWireframe: { tag: "surfaceChart", wireframe: true },
  surfaceWireframe3D: { tag: "surface3DChart", wireframe: true }
} as const;

/**
 * One child position of a chart part element, in schema order.
 *
 * `key` is the field key the child is written from (see `core/chart-kept.ts`);
 * absent when no field governs it. A container slot is always written by
 * `write`, which writes its own kept children.
 */
interface Slot {
  /** The local names that take this position. */
  readonly names: readonly string[];
  readonly key?: string | undefined;
  readonly write?: (kept: readonly ChartKeptElement[]) => void;
  readonly container?: boolean;
}

interface WriteContext {
  readonly xml: XmlSink;
  /** `node`, as the security policy lets it be used. */
  use<T>(node: T | undefined): T | undefined;
  /** Whether every relationship a kept element refers to is written. */
  live(element: { readonly rIds?: readonly string[] | undefined }): boolean;
}

function writeVerbatim(ctx: WriteContext, elements: readonly ChartKeptElement[]): void {
  for (const element of elements) {
    if (element.xml !== "" && ctx.live(element)) {
      ctx.xml.writeRaw(element.xml);
    }
  }
}

/**
 * Write an element's children. Without a kept node every slot is written from
 * the model. With one, a child whose field key is unchanged (or that no field
 * governs) is written as read — and one the part did not have stays absent —
 * a child whose key changed is written from its field, and a child no slot
 * names (an extension, `mc:AlternateContent`) is written after the child it
 * followed.
 */
function writeSlots(
  ctx: WriteContext,
  slots: readonly Slot[],
  node: ChartKeptNode | undefined
): void {
  if (node === undefined) {
    for (const slot of slots) {
      slot.write?.([]);
    }
    return;
  }
  const bySlot: ChartKeptElement[][] = slots.map(() => []);
  const after: ChartKeptElement[][] = slots.map(() => []);
  const before: ChartKeptElement[] = [];
  let anchor = -1;
  for (const child of node.children) {
    const i = slots.findIndex(slot => slot.names.includes(child.name));
    if (i >= 0) {
      bySlot[i]!.push(child);
      anchor = i;
    } else {
      (anchor < 0 ? before : after[anchor]!).push(child);
    }
  }
  writeVerbatim(ctx, before);
  slots.forEach((slot, i) => {
    const kept = bySlot[i]!;
    if (slot.container) {
      slot.write?.(kept);
    } else if (slot.key === undefined || node.read[slot.names[0]!] === slot.key) {
      writeVerbatim(ctx, kept);
    } else {
      slot.write?.(kept);
    }
    writeVerbatim(ctx, after[i]!);
  });
}

/** A slot no field governs, written from the model with `write` when nothing was kept. */
function fixed(name: string, write?: () => void): Slot {
  return write ? { names: [name], write } : { names: [name] };
}

function openElement(xml: XmlSink, name: string, node: ChartKeptNode | undefined): void {
  xml.openNode(node?.name ?? name, node?.attrs);
}

const GROUP_TAGS = [
  "areaChart",
  "area3DChart",
  "lineChart",
  "line3DChart",
  "stockChart",
  "radarChart",
  "scatterChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "barChart",
  "bar3DChart",
  "ofPieChart",
  "surfaceChart",
  "surface3DChart",
  "bubbleChart"
];
const AXIS_TAGS = ["catAx", "valAx", "dateAx", "serAx"];

type AxisRole = NonNullable<ChartKeptAxis["role"]>;

/** The axis ids written: the ones read, or the writer's own. */
type AxisIds = Readonly<Record<AxisRole, string>>;

const DEFAULT_AXIS_IDS: AxisIds = {
  category: "111111111",
  value: "222222222",
  secondaryCategory: "333333333",
  secondaryValue: "444444444"
};

const CROSSING_AXIS: Readonly<Record<AxisRole, AxisRole>> = {
  category: "value",
  value: "category",
  secondaryCategory: "secondaryValue",
  secondaryValue: "secondaryCategory"
};

/** A chart group to write: series of one type on one pair of axes. */
interface PlannedGroup {
  readonly type: ChartType;
  readonly secondary: boolean;
  readonly key: string;
  readonly series: { readonly series: ChartSeries; readonly index: number }[];
}

/** Group the series by the type they are drawn as and the axes they are on, in first-seen order. */
function planGroups(chart: Chart): PlannedGroup[] {
  const groups: PlannedGroup[] = [];
  const add = (series: ChartSeries, index: number, type: ChartType, secondary: boolean) => {
    let group = groups.find(g => g.type === type && g.secondary === secondary);
    if (!group) {
      group = { type, secondary, key: chartGroupKey(type, secondary), series: [] };
      groups.push(group);
    }
    group.series.push({ series, index });
  };
  chart.series.forEach((s, i) =>
    add(s, i, s.chartType ?? chart.type, s.plotOnSecondaryAxis ?? false)
  );
  const secondaryType = chart.secondaryType ?? "line";
  (chart.secondarySeries ?? []).forEach((s, i) =>
    add(s, chart.series.length + i, secondaryType, true)
  );
  if (groups.length === 0) {
    groups.push({
      type: chart.type,
      secondary: false,
      key: chartGroupKey(chart.type, false),
      series: []
    });
  }
  return groups;
}

const hasNoAxes = (type: ChartType): boolean =>
  type === "pie" || type === "pie3D" || type === "doughnut";

/** `c:spPr` with a solid fill of `color`. */
function writeSolidSpPr(xml: XmlSink, color: string | undefined, line = false): void {
  if (!color) {
    return;
  }
  xml.openNode("c:spPr");
  xml.openNode("a:solidFill");
  xml.leafNode("a:srgbClr", { val: color });
  xml.closeNode();
  if (line) {
    xml.openNode("a:ln", { w: "28575" });
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: color });
    xml.closeNode();
    xml.closeNode();
  }
  xml.closeNode();
}

/**
 * Render a chart part (word/charts/chartN.xml).
 *
 * The part is built from the chart's fields and what the reader kept beside
 * them (see `ChartKeptNode`): a child whose field still equals what was read
 * is written as read, one whose field was set is written from the field, and
 * a child no field governs is written as read. A kept element that refers to
 * a relationship not in `relationshipIds` is dropped rather than left
 * dangling.
 */
export function renderChartPart(
  xml: XmlSink,
  chart: Chart,
  options: {
    readonly rawXmlPolicy?: WordSecurityPolicy["rawXmlPolicy"];
    /** The relationships the part is written with. */
    readonly relationshipIds?: ReadonlySet<string>;
  } = {}
): void {
  const ctx: WriteContext = {
    xml,
    use(node) {
      if (node === undefined || options.rawXmlPolicy === "strip") {
        return undefined;
      }
      if (options.rawXmlPolicy === "reject") {
        throw new DocxRawXmlPolicyError("chart part markup");
      }
      return node;
    },
    live: element => element.rIds?.every(id => options.relationshipIds?.has(id) === true) ?? true
  };
  const kept = ctx.use(chart._kept);
  const ids: AxisIds = { ...DEFAULT_AXIS_IDS };
  for (const axis of kept?.axes ?? []) {
    if (axis.role !== undefined) {
      (ids as Record<AxisRole, string>)[axis.role] = axis.id;
    }
  }
  const keys = chartSpaceKeys(chart);
  const rootAttrs = { "xmlns:c": NS_C_CHART, "xmlns:a": NS_A, "xmlns:r": NS_R };

  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode(
    "c:chartSpace",
    kept ? { ...withPreservedNamespaces(rootAttrs, kept.root), ...kept.space.attrs } : rootAttrs
  );
  writeSlots(
    ctx,
    [
      fixed("date1904", () => xml.leafNode("c:date1904", { val: "0" })),
      fixed("lang", () => xml.leafNode("c:lang", { val: "en-US" })),
      fixed("roundedCorners", () => xml.leafNode("c:roundedCorners", { val: "0" })),
      {
        names: ["AlternateContent", "style"],
        key: keys.style,
        write: () => writeStyle(xml, chart)
      },
      fixed("clrMapOvr"),
      fixed("pivotSource"),
      fixed("protection"),
      { names: ["chart"], container: true, write: () => renderChart(ctx, chart, kept, ids) },
      { names: ["spPr"], key: keys.spPr, write: () => writeSolidSpPr(xml, chart.chartAreaColor) },
      fixed("txPr"),
      { names: ["externalData"], key: keys.externalData },
      fixed("printSettings"),
      fixed("userShapes"),
      fixed("extLst")
    ],
    kept?.space
  );
  xml.closeNode(); // c:chartSpace
}

/** The chart style (optional, defaults to 2). */
function writeStyle(xml: XmlSink, chart: Chart): void {
  if (chart.style === undefined) {
    return;
  }
  xml.openNode("mc:AlternateContent", {
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
  });
  xml.openNode("mc:Choice", {
    Requires: "c14",
    "xmlns:c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
  });
  xml.leafNode("c14:style", { val: xmlValue(chart.style, "val") });
  xml.closeNode();
  xml.openNode("mc:Fallback");
  xml.leafNode("c:style", { val: xmlValue(Math.min(chart.style, 48), "val") });
  xml.closeNode();
  xml.closeNode();
}

function renderChart(
  ctx: WriteContext,
  chart: Chart,
  kept: ChartKept | undefined,
  ids: AxisIds
): void {
  const { xml } = ctx;
  const keys = chartKeys(chart);
  openElement(xml, "c:chart", kept?.chart);
  writeSlots(
    ctx,
    [
      {
        names: ["title"],
        container: true,
        write: () => renderTitle(ctx, chart.title, kept?.title)
      },
      {
        names: ["autoTitleDeleted"],
        key: keys.autoTitleDeleted,
        write: () => xml.leafNode("c:autoTitleDeleted", { val: chart.title ? "0" : "1" })
      },
      fixed("pivotFmts"),
      {
        names: ["view3D"],
        key: keys.view3D,
        write: () => {
          if (chart.view3d) {
            xml.openNode("c:view3D");
            xml.leafNode("c:rotX", { val: "15" });
            xml.leafNode("c:rotY", { val: "20" });
            xml.leafNode("c:depthPercent", { val: "100" });
            xml.leafNode("c:rAngAx", { val: "1" });
            xml.closeNode();
          }
        }
      },
      fixed("floor"),
      fixed("sideWall"),
      fixed("backWall"),
      { names: ["plotArea"], container: true, write: () => renderPlotArea(ctx, chart, kept, ids) },
      { names: ["legend"], container: true, write: () => renderLegend(ctx, chart, kept?.legend) },
      fixed("plotVisOnly", () => xml.leafNode("c:plotVisOnly", { val: "1" })),
      fixed("dispBlanksAs", () => xml.leafNode("c:dispBlanksAs", { val: "gap" })),
      fixed("showDLblsOverMax"),
      fixed("extLst")
    ],
    kept?.chart
  );
  xml.closeNode(); // c:chart
}

/**
 * A chart or axis title: written while it has text, or while it is the title
 * that was read (one with no text of its own — an automatic or linked title).
 */
function renderTitle(
  ctx: WriteContext,
  text: string | undefined,
  node: ChartKeptNode | undefined
): void {
  const keys = titleKeys(text);
  if (!text && (node === undefined || node.read[""] !== keys[""])) {
    return;
  }
  const { xml } = ctx;
  openElement(xml, "c:title", node);
  writeSlots(
    ctx,
    [
      { names: ["tx"], key: keys.tx, write: () => text && writeRichText(xml, text) },
      fixed("layout"),
      fixed("overlay", () => xml.leafNode("c:overlay", { val: "0" })),
      fixed("spPr"),
      fixed("txPr"),
      fixed("extLst")
    ],
    node
  );
  xml.closeNode(); // c:title
}

function writeRichText(xml: XmlSink, text: string): void {
  xml.openNode("c:tx");
  xml.openNode("c:rich");
  xml.leafNode("a:bodyPr", {
    rot: "0",
    spcFirstLastPara: "1",
    vertOverflow: "ellipsis",
    wrap: "square",
    anchor: "ctr",
    anchorCtr: "1"
  });
  xml.leafNode("a:lstStyle");
  xml.openNode("a:p");
  xml.openNode("a:pPr");
  xml.leafNode("a:defRPr", {
    sz: "1400",
    b: "0",
    i: "0",
    u: "none",
    strike: "noStrike",
    kern: "1200",
    spc: "0",
    baseline: "0"
  });
  xml.closeNode();
  xml.openNode("a:r");
  xml.leafNode("a:rPr", { lang: "en-US" });
  xml.openNode("a:t");
  xml.writeText(text);
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
  xml.closeNode(); // c:rich
  xml.closeNode(); // c:tx
}

function renderPlotArea(
  ctx: WriteContext,
  chart: Chart,
  kept: ChartKept | undefined,
  ids: AxisIds
): void {
  const { xml } = ctx;
  const keys = plotAreaKeys(chart);
  const groups = planGroups(chart);
  const hasSecondary = groups.some(g => g.secondary);
  openElement(xml, "c:plotArea", kept?.plotArea);
  writeSlots(
    ctx,
    [
      fixed("layout", () => xml.leafNode("c:layout")),
      {
        names: GROUP_TAGS,
        container: true,
        write: source => renderGroups(ctx, chart, groups, kept, ids, source)
      },
      {
        names: AXIS_TAGS,
        container: true,
        write: () => renderAxes(ctx, chart, kept, ids, hasSecondary)
      },
      fixed("dTable"),
      { names: ["spPr"], key: keys.spPr, write: () => writeSolidSpPr(xml, chart.plotAreaColor) },
      fixed("extLst")
    ],
    kept?.plotArea
  );
  xml.closeNode(); // c:plotArea
}

/**
 * The chart groups: each kept group in its place, written while its type and
 * axis are what was read; a group the model cannot read (`c:ofPieChart`) as
 * read; then the groups the part did not have.
 */
function renderGroups(
  ctx: WriteContext,
  chart: Chart,
  groups: readonly PlannedGroup[],
  kept: ChartKept | undefined,
  ids: AxisIds,
  source: readonly ChartKeptElement[]
): void {
  const written = new Set<PlannedGroup>();
  let modelled = 0;
  for (const element of source) {
    if (element.xml !== "") {
      writeVerbatim(ctx, [element]);
      continue;
    }
    const keptGroup = kept?.groups[modelled++];
    const group = groups.find(g => !written.has(g) && g.key === keptGroup?.key);
    if (group) {
      written.add(group);
      renderGroup(ctx, chart, group, ids, keptGroup);
    }
  }
  for (const group of groups) {
    if (!written.has(group)) {
      renderGroup(ctx, chart, group, ids, undefined);
    }
  }
}

function renderGroup(
  ctx: WriteContext,
  chart: Chart,
  group: PlannedGroup,
  ids: AxisIds,
  kept: ChartKept["groups"][number] | undefined
): void {
  const { xml } = ctx;
  const chartType = group.type;
  const info = CHART_TYPE_CATEGORY[chartType] ?? { tag: "barChart" };
  const leaf = (name: string, val: string | undefined) => () => {
    if (val !== undefined) {
      xml.leafNode(`c:${name}`, { val });
    }
  };
  const stacked =
    chartType === "barStacked" ||
    chartType === "barPercentStacked" ||
    chartType === "columnStacked" ||
    chartType === "columnPercentStacked";
  const axes = group.secondary
    ? [ids.secondaryCategory, ids.secondaryValue]
    : [ids.category, ids.value];
  openElement(xml, `c:${info.tag}`, kept?.node);
  writeSlots(
    ctx,
    [
      fixed("barDir", leaf("barDir", "direction" in info ? info.direction : undefined)),
      fixed("grouping", leaf("grouping", "grouping" in info ? info.grouping : undefined)),
      fixed(
        "scatterStyle",
        leaf("scatterStyle", "scatterStyle" in info ? info.scatterStyle : undefined)
      ),
      fixed(
        "radarStyle",
        leaf(
          "radarStyle",
          "radarStyle" in info ? info.radarStyle : info.tag === "radarChart" ? "marker" : undefined
        )
      ),
      fixed(
        "wireframe",
        leaf("wireframe", "wireframe" in info && info.wireframe ? "1" : undefined)
      ),
      fixed("varyColors", leaf("varyColors", hasNoAxes(chartType) ? "1" : "0")),
      {
        names: ["ser"],
        container: true,
        write: () =>
          group.series.forEach(({ series, index }) =>
            renderSeries(ctx, series, index, chartType, "marked" in info && info.marked)
          )
      },
      fixed("bandFmts"),
      {
        names: ["dLbls"],
        container: true,
        write: () => renderGroupDataLabels(ctx, chart.dataLabels, kept?.dataLabels)
      },
      fixed("dropLines"),
      fixed("hiLowLines"),
      fixed("upDownBars"),
      fixed("marker"),
      fixed("smooth"),
      fixed("gapWidth", leaf("gapWidth", info.tag === "barChart" ? "150" : undefined)),
      fixed("gapDepth"),
      fixed("overlap", leaf("overlap", info.tag === "barChart" && stacked ? "100" : undefined)),
      fixed("shape"),
      fixed("serLines"),
      fixed("firstSliceAng"),
      fixed("holeSize"),
      fixed("bubble3D"),
      fixed("bubbleScale"),
      fixed("showNegBubbles"),
      fixed("sizeRepresents"),
      fixed("axId", () => {
        if (!hasNoAxes(chartType)) {
          axes.forEach(id => xml.leafNode("c:axId", { val: id }));
        }
      }),
      fixed("extLst")
    ],
    kept?.node
  );
  xml.closeNode(); // chart type tag
}

function renderSeries(
  ctx: WriteContext,
  series: ChartSeries,
  index: number,
  chartType: ChartType,
  markedLine?: boolean
): void {
  const { xml } = ctx;
  const node = ctx.use(series._kept?.node);
  const keys = seriesKeys(series, index, chartType);
  const scatter = chartType === "scatter" || chartType === "scatterSmooth";
  openElement(xml, "c:ser", node);
  writeSlots(
    ctx,
    [
      {
        names: ["idx"],
        key: keys.idx,
        write: () => xml.leafNode("c:idx", { val: xmlValue(index, "val") })
      },
      {
        names: ["order"],
        key: keys.order,
        write: () => xml.leafNode("c:order", { val: xmlValue(index, "val") })
      },
      { names: ["tx"], key: keys.tx, write: () => writeSeriesName(xml, series, index) },
      {
        names: ["spPr"],
        key: keys.spPr,
        write: () => writeSolidSpPr(xml, series.color, chartType.startsWith("line") || scatter)
      },
      fixed("invertIfNegative"),
      fixed("pictureOptions"),
      {
        names: ["marker"],
        key: keys.marker,
        write: () => {
          if (markedLine) {
            xml.openNode("c:marker");
            xml.leafNode("c:symbol", { val: "circle" });
            xml.leafNode("c:size", { val: "5" });
            xml.closeNode();
          } else if (chartType.startsWith("line")) {
            xml.openNode("c:marker");
            xml.leafNode("c:symbol", { val: "none" });
            xml.closeNode();
          }
        }
      },
      fixed("explosion"),
      { names: ["dPt"], container: true, write: kept => writePoints(ctx, series, kept) },
      {
        names: ["dLbls"],
        container: true,
        write: () => renderSeriesDataLabels(ctx, series, ctx.use(series._kept?.dataLabels))
      },
      {
        names: ["trendline"],
        key: keys.trendline,
        write: () => renderTrendline(xml, series.trendline)
      },
      {
        names: ["errBars"],
        key: keys.errBars,
        write: () => renderErrorBars(xml, series.errorBars)
      },
      {
        names: ["cat", "xVal"],
        key: keys.cat,
        write: () => writeCategories(xml, series, scatter)
      },
      { names: ["val", "yVal"], key: keys.val, write: () => writeValues(xml, series, scatter) },
      fixed("bubbleSize"),
      fixed("bubble3D"),
      fixed("shape"),
      {
        names: ["smooth"],
        key: keys.smooth,
        write: () => {
          if (chartType === "scatterSmooth") {
            xml.leafNode("c:smooth", { val: "1" });
          }
        }
      },
      fixed("extLst")
    ],
    node
  );
  xml.closeNode(); // c:ser
}

function writeSeriesName(xml: XmlSink, series: ChartSeries, index: number): void {
  xml.openNode("c:tx");
  xml.openNode("c:strRef");
  xml.openNode("c:f");
  xml.writeText(`Sheet1!$${String.fromCharCode(66 + index)}$1`);
  xml.closeNode();
  xml.openNode("c:strCache");
  xml.leafNode("c:ptCount", { val: "1" });
  xml.openNode("c:pt", { idx: "0" });
  xml.openNode("c:v");
  xml.writeText(series.name);
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
  xml.closeNode(); // c:tx
}

/**
 * The series' `c:dPt`, by point index: a kept point while its colour is what
 * was read from it, else one for each colour in `pointColors` (a hole or
 * `undefined` entry colours nothing).
 */
function writePoints(
  ctx: WriteContext,
  series: ChartSeries,
  kept: readonly ChartKeptElement[]
): void {
  const { xml } = ctx;
  const points = new Map<number, () => void>();
  for (const element of kept) {
    if (element.index === undefined) {
      writeVerbatim(ctx, [element]);
    } else if (element.read === pointKey(series, element.index)) {
      points.set(element.index, () => writeVerbatim(ctx, [element]));
    }
  }
  series.pointColors?.forEach((color, i) => {
    if (color !== undefined && !points.has(i)) {
      points.set(i, () => {
        xml.openNode("c:dPt");
        xml.leafNode("c:idx", { val: xmlValue(i, "val") });
        xml.leafNode("c:bubble3D", { val: "0" });
        writeSolidSpPr(xml, color);
        xml.closeNode();
      });
    }
  });
  [...points.keys()].sort((a, b) => a - b).forEach(i => points.get(i)!());
}

function writeCategories(xml: XmlSink, series: ChartSeries, scatter: boolean): void {
  xml.openNode(scatter ? "c:xVal" : "c:cat");
  xml.openNode(scatter ? "c:numRef" : "c:strRef");
  xml.openNode("c:f");
  xml.writeText(`Sheet1!$A$2:$A$${series.categories.length + 1}`);
  xml.closeNode();
  xml.openNode(scatter ? "c:numCache" : "c:strCache");
  if (scatter) {
    xml.leafNode("c:formatCode", { val: "General" });
  }
  xml.leafNode("c:ptCount", { val: xmlValue(series.categories.length, "val") });
  series.categories.forEach((cat, i) => {
    xml.openNode("c:pt", { idx: xmlValue(i, "idx") });
    xml.openNode("c:v");
    xml.writeText(cat);
    xml.closeNode();
    xml.closeNode();
  });
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
}

function writeValues(xml: XmlSink, series: ChartSeries, scatter: boolean): void {
  xml.openNode(scatter ? "c:yVal" : "c:val");
  xml.openNode("c:numRef");
  xml.openNode("c:f");
  xml.writeText(`Sheet1!$B$2:$B$${series.values.length + 1}`);
  xml.closeNode();
  xml.openNode("c:numCache");
  xml.leafNode("c:formatCode", { val: "General" });
  xml.leafNode("c:ptCount", { val: xmlValue(series.values.length, "val") });
  series.values.forEach((val, i) => {
    xml.openNode("c:pt", { idx: xmlValue(i, "idx") });
    xml.openNode("c:v");
    xml.writeText(xmlValue(val, "c:v"));
    xml.closeNode();
    xml.closeNode();
  });
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
}

/** The axes: each kept one in its place, those the model does not represent as read, then any new ones. */
function renderAxes(
  ctx: WriteContext,
  chart: Chart,
  kept: ChartKept | undefined,
  ids: AxisIds,
  hasSecondary: boolean
): void {
  if (hasNoAxes(chart.type)) {
    return;
  }
  const wanted: AxisRole[] = hasSecondary
    ? ["category", "value", "secondaryCategory", "secondaryValue"]
    : ["category", "value"];
  const done = new Set<AxisRole>();
  for (const axis of kept?.axes ?? []) {
    if (axis.role === undefined) {
      writeVerbatim(ctx, [
        { name: "", xml: axis.xml ?? "", ...(axis.rIds && { rIds: axis.rIds }) }
      ]);
    } else if (wanted.includes(axis.role) && !done.has(axis.role)) {
      done.add(axis.role);
      renderAxis(ctx, chart, axis.role, ids, axis);
    }
  }
  for (const role of wanted) {
    if (!done.has(role)) {
      renderAxis(ctx, chart, role, ids, undefined);
    }
  }
}

function renderAxis(
  ctx: WriteContext,
  chart: Chart,
  role: AxisRole,
  ids: AxisIds,
  kept: ChartKeptAxis | undefined
): void {
  const { xml } = ctx;
  const axis: ChartAxis | undefined =
    role === "category"
      ? chart.categoryAxis
      : role === "value"
        ? chart.valueAxis
        : role === "secondaryValue"
          ? chart.secondaryValueAxis
          : undefined;
  const isCategory = role === "category" || role === "secondaryCategory";
  const tag = kept?.node.name ?? (isCategory ? "c:catAx" : "c:valAx");
  const catAx = tag.endsWith("catAx");
  const valAx = tag.endsWith("valAx");
  const keys = axisKeys(axis, role);
  const leaf = (name: string, val: string) => () => xml.leafNode(`c:${name}`, { val });
  const onlyIf = (when: boolean, write: () => void) => () => {
    if (when) {
      write();
    }
  };
  xml.openNode(tag, kept?.node.attrs);
  writeSlots(
    ctx,
    [
      fixed("axId", leaf("axId", ids[role])),
      { names: ["scaling"], container: true, write: () => renderScaling(ctx, axis, kept?.scaling) },
      {
        names: ["delete"],
        key: keys.delete,
        write: leaf("delete", role === "secondaryCategory" || axis?.hidden ? "1" : "0")
      },
      fixed("axPos", leaf("axPos", isCategory ? "b" : role === "value" ? "l" : "r")),
      fixed("majorGridlines"),
      fixed("minorGridlines"),
      {
        names: ["title"],
        container: true,
        write: () => renderTitle(ctx, axis?.title, kept?.title)
      },
      {
        names: ["numFmt"],
        key: keys.numFmt,
        write: () =>
          axis?.numberFormat
            ? xml.leafNode("c:numFmt", { formatCode: axis.numberFormat, sourceLinked: "0" })
            : xml.leafNode("c:numFmt", { formatCode: "General", sourceLinked: "1" })
      },
      fixed("majorTickMark", leaf("majorTickMark", "out")),
      fixed("minorTickMark", leaf("minorTickMark", "none")),
      fixed("tickLblPos", leaf("tickLblPos", "nextTo")),
      fixed("spPr"),
      fixed("txPr"),
      fixed("crossAx", leaf("crossAx", ids[CROSSING_AXIS[role]])),
      { names: ["crosses", "crossesAt"], write: leaf("crosses", "autoZero") },
      fixed("crossBetween", onlyIf(valAx, leaf("crossBetween", "between"))),
      fixed("auto", onlyIf(catAx, leaf("auto", "1"))),
      fixed("lblAlgn", onlyIf(catAx, leaf("lblAlgn", "ctr"))),
      fixed("lblOffset", onlyIf(catAx, leaf("lblOffset", "100"))),
      fixed("baseTimeUnit"),
      fixed("tickLblSkip"),
      fixed("tickMarkSkip"),
      fixed("noMultiLvlLbl", onlyIf(catAx, leaf("noMultiLvlLbl", "0"))),
      {
        names: ["majorUnit"],
        key: keys.majorUnit,
        write: () => {
          if (valAx && axis?.majorUnit) {
            xml.leafNode("c:majorUnit", { val: xmlValue(axis.majorUnit, "val") });
          }
        }
      },
      fixed("majorTimeUnit"),
      fixed("minorUnit"),
      fixed("minorTimeUnit"),
      fixed("dispUnits"),
      fixed("extLst")
    ],
    kept?.node
  );
  xml.closeNode();
}

function renderScaling(
  ctx: WriteContext,
  axis: ChartAxis | undefined,
  node: ChartKeptNode | undefined
): void {
  const { xml } = ctx;
  const keys = scalingKeys(axis);
  openElement(xml, "c:scaling", node);
  writeSlots(
    ctx,
    [
      fixed("logBase"),
      fixed("orientation", () => xml.leafNode("c:orientation", { val: "minMax" })),
      {
        names: ["max"],
        key: keys.max,
        write: () => {
          if (axis?.max !== undefined) {
            xml.leafNode("c:max", { val: xmlValue(axis.max, "val") });
          }
        }
      },
      {
        names: ["min"],
        key: keys.min,
        write: () => {
          if (axis?.min !== undefined) {
            xml.leafNode("c:min", { val: xmlValue(axis.min, "val") });
          }
        }
      },
      fixed("extLst")
    ],
    node
  );
  xml.closeNode(); // c:scaling
}

function renderLegend(ctx: WriteContext, chart: Chart, node: ChartKeptNode | undefined): void {
  if (chart.legend === "none") {
    return;
  }
  const { xml } = ctx;
  const keys = legendKeys(chart);
  openElement(xml, "c:legend", node);
  writeSlots(
    ctx,
    [
      {
        names: ["legendPos"],
        key: keys.legendPos,
        write: () =>
          xml.leafNode("c:legendPos", {
            val: writtenToken(chart, "legend", WRITTEN.legendPosition, "legend") ?? "r"
          })
      },
      fixed("legendEntry"),
      fixed("layout"),
      fixed("overlay", () => xml.leafNode("c:overlay", { val: "0" })),
      fixed("spPr"),
      fixed("txPr"),
      fixed("extLst")
    ],
    node
  );
  xml.closeNode(); // c:legend
}

// -- Data labels position map ---------------------------------------------------

const DATA_LABEL_POSITION_MAP: Record<string, string> = {
  outsideEnd: "outEnd",
  center: "ctr",
  insideEnd: "inEnd",
  bestFit: "bestFit"
};
const DATA_LABEL_POSITIONS: ReadonlySet<string> = new Set(Object.keys(DATA_LABEL_POSITION_MAP));

/**
 * The checks the chart part's writer makes, run while the body is written —
 * where the chart's place in the document is known — so a refused value is
 * named by its path there. Fields are relative to the `Chart`.
 */
export function checkWritableChart(chart: Chart): void {
  if (chart.legend !== "none") {
    writtenToken(chart, "legend", WRITTEN.legendPosition, "legend");
  }
  const dataLabels = chart.dataLabels;
  if (
    dataLabels &&
    unrecognizedValue(dataLabels, "position") === undefined &&
    dataLabels.position
  ) {
    checkOneOf(dataLabels.position, DATA_LABEL_POSITIONS, "dataLabels.position", MODEL_CHECK);
  }
  drawingInteger(chart.width, "width", 0);
  drawingInteger(chart.height, "height", 0);
}

/**
 * The slots of a `c:dLbls`. `keys` are its field keys; `show` writes a
 * `show*` flag from the model when it is not kept.
 */
function dataLabelSlots(
  xml: XmlSink,
  keys: Readonly<Record<string, string | undefined>>,
  show: Readonly<Record<string, boolean>>,
  writePosition?: () => void
): Slot[] {
  const flag = (name: string): Slot => ({
    names: [name],
    key: keys[name],
    write: () => xml.leafNode(`c:${name}`, { val: show[name] ? "1" : "0" })
  });
  return [
    fixed("dLbl"),
    // `c:delete` excludes the flags (CT_DLbls is dLbl*, then delete | the
    // group): once the labels' field changes, the kept one goes.
    { names: ["delete"], key: keys.delete },
    fixed("numFmt"),
    fixed("spPr"),
    fixed("txPr"),
    { names: ["dLblPos"], key: keys.dLblPos, ...(writePosition && { write: writePosition }) },
    flag("showLegendKey"),
    flag("showVal"),
    flag("showCatName"),
    flag("showSerName"),
    flag("showPercent"),
    flag("showBubbleSize"),
    fixed("separator"),
    fixed("showLeaderLines"),
    fixed("leaderLines"),
    fixed("extLst")
  ];
}

/**
 * A chart group's `c:dLbls`, from the chart's `dataLabels`: written while it
 * is set, or while it is what was read (labels the model reads no flag from).
 */
function renderGroupDataLabels(
  ctx: WriteContext,
  dataLabels: ChartDataLabels | undefined,
  kept: ChartKeptNode | undefined
): void {
  const node = ctx.use(kept);
  const keys = groupDataLabelsKeys(dataLabels);
  if (dataLabels === undefined && (node === undefined || node.read[""] !== keys[""])) {
    return;
  }
  const { xml } = ctx;
  openElement(xml, "c:dLbls", node);
  writeSlots(
    ctx,
    dataLabelSlots(
      xml,
      keys,
      {
        showVal: !!dataLabels?.showValue,
        showCatName: !!dataLabels?.showCategory,
        showSerName: !!dataLabels?.showSerName,
        showPercent: !!dataLabels?.showPercent
      },
      () => {
        // A position the model names is written as its `ST_DLblPos` value; one
        // the reader kept (`l`, `t`, `inBase`, …) as read.
        const keptPosition = dataLabels ? unrecognizedValue(dataLabels, "position") : undefined;
        if (keptPosition !== undefined) {
          xml.leafNode("c:dLblPos", { val: keptPosition });
        } else if (dataLabels?.position) {
          checkOneOf(dataLabels.position, DATA_LABEL_POSITIONS, "dataLabels.position", MODEL_CHECK);
          xml.leafNode("c:dLblPos", { val: DATA_LABEL_POSITION_MAP[dataLabels.position]! });
        }
      }
    ),
    node
  );
  xml.closeNode();
}

/** A series' `c:dLbls`, from `showDataLabels`, kept as read like the chart's. */
function renderSeriesDataLabels(
  ctx: WriteContext,
  series: ChartSeries,
  node: ChartKeptNode | undefined
): void {
  const keys = seriesDataLabelsKeys(series);
  if (!series.showDataLabels && (node === undefined || node.read[""] !== keys[""])) {
    return;
  }
  const { xml } = ctx;
  openElement(xml, "c:dLbls", node);
  writeSlots(ctx, dataLabelSlots(xml, keys, { showVal: !!series.showDataLabels }), node);
  xml.closeNode();
}

// -- Trendline ------------------------------------------------------------------

function renderTrendline(xml: XmlSink, trendline: ChartSeries["trendline"]): void {
  if (!trendline) {
    return;
  }
  xml.openNode("c:trendline");
  xml.leafNode("c:trendlineType", { val: trendline.type });
  if (trendline.type === "polynomial" && trendline.order !== undefined) {
    xml.leafNode("c:order", { val: xmlValue(trendline.order, "val") });
  }
  if (trendline.type === "movingAvg" && trendline.period !== undefined) {
    xml.leafNode("c:period", { val: xmlValue(trendline.period, "val") });
  }
  if (trendline.displayEquation) {
    xml.leafNode("c:dispEq", { val: "1" });
  }
  if (trendline.displayRSquared) {
    xml.leafNode("c:dispRSqr", { val: "1" });
  }
  xml.closeNode();
}

// -- Error Bars -----------------------------------------------------------------

function renderErrorBars(xml: XmlSink, errorBars: ChartSeries["errorBars"]): void {
  if (!errorBars) {
    return;
  }

  const directions: Array<"x" | "y"> =
    errorBars.direction === "both" ? ["x", "y"] : [errorBars.direction];

  for (const dir of directions) {
    xml.openNode("c:errBars");
    xml.leafNode("c:errBarType", { val: "both" });
    xml.leafNode("c:errValType", { val: errorBars.type });
    xml.leafNode("c:errDir", { val: dir });
    if (
      errorBars.value !== undefined &&
      (errorBars.type === "fixedVal" || errorBars.type === "percentage")
    ) {
      xml.leafNode("c:val", { val: xmlValue(errorBars.value, "val") });
    }
    xml.leafNode("c:noEndCap", { val: "0" });
    xml.closeNode();
  }
}
