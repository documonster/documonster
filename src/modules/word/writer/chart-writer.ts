/**
 * DOCX Module - Chart Writer
 *
 * Generates DrawingML chart part XML (word/charts/chartN.xml) from
 * high-level Chart definitions.
 */

import { NS_A, NS_C_CHART, NS_R, STD_DOC_ATTRIBUTES } from "@word/constants";
import type { Chart, ChartAxis, ChartDataLabels, ChartSeries, ChartType } from "@word/types";
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

/** Render a chart part (word/charts/chartN.xml). */
export function renderChartPart(xml: XmlSink, chart: Chart): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("c:chartSpace", {
    "xmlns:c": NS_C_CHART,
    "xmlns:a": NS_A,
    "xmlns:r": NS_R
  });

  // Date1904 (calendar base)
  xml.leafNode("c:date1904", { val: "0" });
  // Language
  xml.leafNode("c:lang", { val: "en-US" });
  // Round text
  xml.leafNode("c:roundedCorners", { val: "0" });

  // Chart style (optional, defaults to 2)
  if (chart.style !== undefined) {
    xml.openNode("mc:AlternateContent", {
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
    });
    xml.openNode("mc:Choice", {
      Requires: "c14",
      "xmlns:c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
    });
    xml.leafNode("c14:style", { val: String(chart.style) });
    xml.closeNode();
    xml.openNode("mc:Fallback");
    xml.leafNode("c:style", { val: String(Math.min(chart.style, 48)) });
    xml.closeNode();
    xml.closeNode();
  }

  // Chart
  xml.openNode("c:chart");

  // Title
  if (chart.title) {
    renderTitle(xml, chart.title);
    xml.leafNode("c:autoTitleDeleted", { val: "0" });
  } else {
    xml.leafNode("c:autoTitleDeleted", { val: "1" });
  }

  // 3D view
  if (chart.view3d) {
    xml.openNode("c:view3D");
    xml.leafNode("c:rotX", { val: "15" });
    xml.leafNode("c:rotY", { val: "20" });
    xml.leafNode("c:depthPercent", { val: "100" });
    xml.leafNode("c:rAngAx", { val: "1" });
    xml.closeNode();
  }

  // Plot area
  xml.openNode("c:plotArea");
  xml.leafNode("c:layout");

  // Chart type-specific rendering
  const catAxId = "111111111";
  const valAxId = "222222222";
  const catAxId2 = "333333333";
  const valAxId2 = "444444444";
  const hasSecondary = renderChartTypeElement(xml, chart, catAxId, valAxId, catAxId2, valAxId2);

  // Axes (unless pie/doughnut which use no axes)
  const noAxes = chart.type === "pie" || chart.type === "pie3D" || chart.type === "doughnut";
  if (!noAxes) {
    // Primary category axis
    renderAxis(xml, "catAx", catAxId, valAxId, chart.categoryAxis, "b");
    // Primary value axis
    renderAxis(xml, "valAx", valAxId, catAxId, chart.valueAxis, "l");

    // Secondary axes (for combo charts)
    if (hasSecondary) {
      renderAxis(xml, "catAx", catAxId2, valAxId2, undefined, "b", true);
      renderAxis(xml, "valAx", valAxId2, catAxId2, chart.secondaryValueAxis, "r");
    }
  }

  // Plot area background
  if (chart.plotAreaColor) {
    xml.openNode("c:spPr");
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: chart.plotAreaColor });
    xml.closeNode();
    xml.closeNode();
  }

  xml.closeNode(); // c:plotArea

  // Legend
  if (chart.legend !== "none") {
    xml.openNode("c:legend");
    xml.leafNode("c:legendPos", { val: chart.legend ?? "r" });
    xml.leafNode("c:overlay", { val: "0" });
    xml.closeNode();
  }

  xml.leafNode("c:plotVisOnly", { val: "1" });
  xml.leafNode("c:dispBlanksAs", { val: "gap" });

  xml.closeNode(); // c:chart

  // Chart area background
  if (chart.chartAreaColor) {
    xml.openNode("c:spPr");
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: chart.chartAreaColor });
    xml.closeNode();
    xml.closeNode();
  }

  xml.closeNode(); // c:chartSpace
}

function renderTitle(xml: XmlSink, title: string): void {
  xml.openNode("c:title");
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
  xml.writeText(title);
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
  xml.closeNode(); // c:rich
  xml.closeNode(); // c:tx
  xml.leafNode("c:overlay", { val: "0" });
  xml.closeNode(); // c:title
}

function renderChartTypeElement(
  xml: XmlSink,
  chart: Chart,
  catAxId: string,
  valAxId: string,
  catAxId2: string,
  valAxId2: string
): boolean {
  // Determine if this is a combo chart
  const hasSeriesOverrides = chart.series.some(s => s.chartType !== undefined);
  const hasSecondarySeries =
    chart.secondarySeries !== undefined && chart.secondarySeries.length > 0;
  const hasSecondaryAxis = hasSecondarySeries || chart.series.some(s => s.plotOnSecondaryAxis);

  if (!hasSeriesOverrides && !hasSecondarySeries) {
    // Simple (non-combo) chart — original logic
    renderSingleChartTypeBlock(xml, chart.type, chart.series, chart, catAxId, valAxId, 0);
    return false;
  }

  // Combo chart: group series by effective chart type and axis
  interface SeriesGroup {
    chartType: ChartType;
    secondary: boolean;
    series: { series: ChartSeries; globalIndex: number }[];
  }

  const groups: SeriesGroup[] = [];
  let globalIndex = 0;

  const findOrCreateGroup = (type: ChartType, secondary: boolean): SeriesGroup => {
    let group = groups.find(g => g.chartType === type && g.secondary === secondary);
    if (!group) {
      group = { chartType: type, secondary, series: [] };
      groups.push(group);
    }
    return group;
  };

  // Primary series
  for (const s of chart.series) {
    const effectiveType = s.chartType ?? chart.type;
    const secondary = s.plotOnSecondaryAxis ?? false;
    const group = findOrCreateGroup(effectiveType, secondary);
    group.series.push({ series: s, globalIndex });
    globalIndex++;
  }

  // Secondary series (legacy approach)
  if (hasSecondarySeries) {
    const secondaryType = chart.secondaryType ?? "line";
    for (const s of chart.secondarySeries!) {
      const group = findOrCreateGroup(secondaryType, true);
      group.series.push({ series: s, globalIndex });
      globalIndex++;
    }
  }

  // Render each group as a separate chart type element
  for (const group of groups) {
    const useCatAx = group.secondary ? catAxId2 : catAxId;
    const useValAx = group.secondary ? valAxId2 : valAxId;
    const seriesItems = group.series.map(item => item.series);
    const startIndex = group.series[0]!.globalIndex;
    renderSingleChartTypeBlock(
      xml,
      group.chartType,
      seriesItems,
      chart,
      useCatAx,
      useValAx,
      startIndex
    );
  }

  return hasSecondaryAxis;
}

function renderSingleChartTypeBlock(
  xml: XmlSink,
  chartType: ChartType,
  series: readonly ChartSeries[],
  chart: Chart,
  catAxId: string,
  valAxId: string,
  startIndex: number
): void {
  const info = CHART_TYPE_CATEGORY[chartType] ?? { tag: "barChart" };
  const tagName = `c:${info.tag}`;
  xml.openNode(tagName);

  // Direction (bar vs col)
  if ("direction" in info && info.direction) {
    xml.leafNode("c:barDir", { val: info.direction });
  }

  // Grouping (stacked/clustered/percentStacked)
  if ("grouping" in info && info.grouping) {
    xml.leafNode("c:grouping", { val: info.grouping });
  }

  // Scatter style
  if ("scatterStyle" in info && info.scatterStyle) {
    xml.leafNode("c:scatterStyle", { val: info.scatterStyle });
  }

  // Radar style
  if ("radarStyle" in info && info.radarStyle) {
    xml.leafNode("c:radarStyle", { val: info.radarStyle });
  } else if (info.tag === "radarChart") {
    xml.leafNode("c:radarStyle", { val: "marker" });
  }

  // Wireframe (surface charts)
  if ("wireframe" in info && info.wireframe) {
    xml.leafNode("c:wireframe", { val: "1" });
  }

  xml.leafNode("c:varyColors", {
    val: chartType === "pie" || chartType === "pie3D" || chartType === "doughnut" ? "1" : "0"
  });

  // Series
  series.forEach((s, idx) => {
    renderSeries(xml, s, startIndex + idx, chartType, "marked" in info && info.marked);
  });

  // Chart-level data labels
  if (chart.dataLabels) {
    renderChartDataLabels(xml, chart.dataLabels);
  }

  // Gap width / overlap (bar/column)
  if (info.tag === "barChart") {
    xml.leafNode("c:gapWidth", { val: "150" });
    if (
      chartType === "barStacked" ||
      chartType === "barPercentStacked" ||
      chartType === "columnStacked" ||
      chartType === "columnPercentStacked"
    ) {
      xml.leafNode("c:overlap", { val: "100" });
    }
  }

  // Axes IDs
  const noAxes = chartType === "pie" || chartType === "pie3D" || chartType === "doughnut";
  if (!noAxes) {
    xml.leafNode("c:axId", { val: catAxId });
    xml.leafNode("c:axId", { val: valAxId });
  }

  xml.closeNode(); // chart type tag
}

function renderSeries(
  xml: XmlSink,
  series: ChartSeries,
  index: number,
  chartType: ChartType,
  markedLine?: boolean
): void {
  xml.openNode("c:ser");
  xml.leafNode("c:idx", { val: String(index) });
  xml.leafNode("c:order", { val: String(index) });

  // Series name
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

  // Series fill color
  if (series.color) {
    xml.openNode("c:spPr");
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: series.color });
    xml.closeNode();
    if (chartType.startsWith("line") || chartType === "scatter" || chartType === "scatterSmooth") {
      xml.openNode("a:ln", { w: "28575" });
      xml.openNode("a:solidFill");
      xml.leafNode("a:srgbClr", { val: series.color });
      xml.closeNode();
      xml.closeNode();
    }
    xml.closeNode();
  }

  // Per-point colors (for pie)
  if (series.pointColors && series.pointColors.length > 0) {
    series.pointColors.forEach((color, i) => {
      xml.openNode("c:dPt");
      xml.leafNode("c:idx", { val: String(i) });
      xml.leafNode("c:bubble3D", { val: "0" });
      xml.openNode("c:spPr");
      xml.openNode("a:solidFill");
      xml.leafNode("a:srgbClr", { val: color });
      xml.closeNode();
      xml.closeNode();
      xml.closeNode();
    });
  }

  // Markers for line charts
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

  // Smooth for line/scatter smooth
  if (chartType === "scatterSmooth") {
    xml.leafNode("c:smooth", { val: "1" });
  }

  // Data labels
  if (series.showDataLabels) {
    xml.openNode("c:dLbls");
    xml.leafNode("c:showLegendKey", { val: "0" });
    xml.leafNode("c:showVal", { val: "1" });
    xml.leafNode("c:showCatName", { val: "0" });
    xml.leafNode("c:showSerName", { val: "0" });
    xml.leafNode("c:showPercent", { val: "0" });
    xml.leafNode("c:showBubbleSize", { val: "0" });
    xml.closeNode();
  }

  // Categories (c:cat)
  if (chartType !== "scatter" && chartType !== "scatterSmooth") {
    xml.openNode("c:cat");
    xml.openNode("c:strRef");
    xml.openNode("c:f");
    xml.writeText(`Sheet1!$A$2:$A$${series.categories.length + 1}`);
    xml.closeNode();
    xml.openNode("c:strCache");
    xml.leafNode("c:ptCount", { val: String(series.categories.length) });
    series.categories.forEach((cat, i) => {
      xml.openNode("c:pt", { idx: String(i) });
      xml.openNode("c:v");
      xml.writeText(cat);
      xml.closeNode();
      xml.closeNode();
    });
    xml.closeNode();
    xml.closeNode();
    xml.closeNode();
  } else {
    // For scatter: xVal
    xml.openNode("c:xVal");
    xml.openNode("c:numRef");
    xml.openNode("c:f");
    xml.writeText(`Sheet1!$A$2:$A$${series.categories.length + 1}`);
    xml.closeNode();
    xml.openNode("c:numCache");
    xml.leafNode("c:formatCode", { val: "General" });
    xml.leafNode("c:ptCount", { val: String(series.categories.length) });
    series.categories.forEach((cat, i) => {
      xml.openNode("c:pt", { idx: String(i) });
      xml.openNode("c:v");
      xml.writeText(cat);
      xml.closeNode();
      xml.closeNode();
    });
    xml.closeNode();
    xml.closeNode();
    xml.closeNode();
  }

  // Values (c:val or c:yVal)
  const valTag = chartType === "scatter" || chartType === "scatterSmooth" ? "c:yVal" : "c:val";
  xml.openNode(valTag);
  xml.openNode("c:numRef");
  xml.openNode("c:f");
  xml.writeText(`Sheet1!$B$2:$B$${series.values.length + 1}`);
  xml.closeNode();
  xml.openNode("c:numCache");
  xml.leafNode("c:formatCode", { val: "General" });
  xml.leafNode("c:ptCount", { val: String(series.values.length) });
  series.values.forEach((val, i) => {
    xml.openNode("c:pt", { idx: String(i) });
    xml.openNode("c:v");
    xml.writeText(String(val));
    xml.closeNode();
    xml.closeNode();
  });
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();

  // Trendline
  if (series.trendline) {
    renderTrendline(xml, series.trendline);
  }

  // Error bars
  if (series.errorBars) {
    renderErrorBars(xml, series.errorBars);
  }

  xml.closeNode(); // c:ser
}

function renderAxis(
  xml: XmlSink,
  tagName: "catAx" | "valAx",
  axId: string,
  crossAxId: string,
  axis: ChartAxis | undefined,
  defaultPos: "b" | "l" | "r" | "t",
  forceHidden?: boolean
): void {
  xml.openNode(`c:${tagName}`);
  xml.leafNode("c:axId", { val: axId });

  // Scaling
  xml.openNode("c:scaling");
  xml.leafNode("c:orientation", { val: "minMax" });
  if (axis?.min !== undefined) {
    xml.leafNode("c:min", { val: String(axis.min) });
  }
  if (axis?.max !== undefined) {
    xml.leafNode("c:max", { val: String(axis.max) });
  }
  xml.closeNode();

  xml.leafNode("c:delete", { val: forceHidden || axis?.hidden ? "1" : "0" });
  xml.leafNode("c:axPos", { val: defaultPos });

  // Title
  if (axis?.title) {
    renderTitle(xml, axis.title);
  }

  // Number format
  if (axis?.numberFormat) {
    xml.leafNode("c:numFmt", { formatCode: axis.numberFormat, sourceLinked: "0" });
  } else {
    xml.leafNode("c:numFmt", { formatCode: "General", sourceLinked: "1" });
  }

  xml.leafNode("c:majorTickMark", { val: "out" });
  xml.leafNode("c:minorTickMark", { val: "none" });
  xml.leafNode("c:tickLblPos", { val: "nextTo" });

  xml.leafNode("c:crossAx", { val: crossAxId });
  xml.leafNode("c:crosses", { val: "autoZero" });
  if (tagName === "catAx") {
    xml.leafNode("c:auto", { val: "1" });
    xml.leafNode("c:lblAlgn", { val: "ctr" });
    xml.leafNode("c:lblOffset", { val: "100" });
    xml.leafNode("c:noMultiLvlLbl", { val: "0" });
  } else {
    xml.leafNode("c:crossBetween", { val: "between" });
    if (axis?.majorUnit) {
      xml.leafNode("c:majorUnit", { val: String(axis.majorUnit) });
    }
  }

  xml.closeNode();
}

// -- Data labels position map ---------------------------------------------------

const DATA_LABEL_POSITION_MAP: Record<string, string> = {
  outsideEnd: "outEnd",
  center: "ctr",
  insideEnd: "inEnd",
  bestFit: "bestFit"
};

function renderChartDataLabels(xml: XmlSink, dataLabels: ChartDataLabels): void {
  xml.openNode("c:dLbls");
  if (dataLabels.position) {
    xml.leafNode("c:dLblPos", { val: DATA_LABEL_POSITION_MAP[dataLabels.position] ?? "bestFit" });
  }
  xml.leafNode("c:showLegendKey", { val: "0" });
  xml.leafNode("c:showVal", { val: dataLabels.showValue ? "1" : "0" });
  xml.leafNode("c:showCatName", { val: dataLabels.showCategory ? "1" : "0" });
  xml.leafNode("c:showSerName", { val: dataLabels.showSerName ? "1" : "0" });
  xml.leafNode("c:showPercent", { val: dataLabels.showPercent ? "1" : "0" });
  xml.leafNode("c:showBubbleSize", { val: "0" });
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
    xml.leafNode("c:order", { val: String(trendline.order) });
  }
  if (trendline.type === "movingAvg" && trendline.period !== undefined) {
    xml.leafNode("c:period", { val: String(trendline.period) });
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
      xml.leafNode("c:val", { val: String(errorBars.value) });
    }
    xml.leafNode("c:noEndCap", { val: "0" });
    xml.closeNode();
  }
}
