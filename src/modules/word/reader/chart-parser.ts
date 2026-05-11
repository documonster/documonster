/**
 * DOCX Reader - Chart Parser (c: and cx: namespaces)
 *
 * Parses traditional and ChartEx (Microsoft 365) chart XML into Chart model
 * objects. Also resolves chart references in the document body, replacing
 * opaque chart drawings with parsed Chart instances.
 */

import { parseXml, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { type Mutable } from "../core/internal-utils";
import type {
  BodyContent,
  Chart,
  ChartAxis,
  ChartContent,
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
  OpaqueDrawing
} from "../types";
import {
  findChildLocal as findChartChild,
  findChildrenLocal as findAllChartChildren,
  localName
} from "./parse-utils";

// =============================================================================
// Chart Reader
// =============================================================================

/** Replace OpaqueDrawing items referencing chart rIds with ChartContent. */
function replaceOpaqueCharts(body: BodyContent[], chartRIdToChart: Map<string, Chart>): void {
  for (let i = 0; i < body.length; i++) {
    const item = body[i];
    if (item.type === "opaqueDrawing") {
      const chartRId = findChartRIdInDrawing(item, chartRIdToChart);
      if (chartRId) {
        const chart = chartRIdToChart.get(chartRId)!;
        body[i] = { type: "chart", chart } as ChartContent;
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
          body[i] = content;
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

/** Parse a chart part XML string into a Chart object. */
function parseChartXml(xmlStr: string): Chart | undefined {
  let doc: ReturnType<typeof parseXml>;
  try {
    doc = parseXml(xmlStr);
  } catch {
    return undefined;
  }

  const root = doc.root; // c:chartSpace
  const chartEl = findChartChild(root, "chart");
  if (!chartEl) {
    return undefined;
  }

  // Title
  const titleEl = findChartChild(chartEl, "title");
  let title: string | undefined;
  if (titleEl) {
    title = extractChartText(titleEl);
  }

  // 3D view
  const view3dEl = findChartChild(chartEl, "view3D");
  const view3d = !!view3dEl;

  // Plot area
  const plotAreaEl = findChartChild(chartEl, "plotArea");
  if (!plotAreaEl) {
    return undefined;
  }

  // Find the chart type element within plotArea
  const CHART_TAGS = [
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
  ];
  let chartTypeEl: XmlElement | undefined;
  let chartTag = "";
  for (const tag of CHART_TAGS) {
    chartTypeEl = findChartChild(plotAreaEl, tag);
    if (chartTypeEl) {
      chartTag = tag;
      break;
    }
  }

  if (!chartTypeEl) {
    return undefined;
  }

  // Grouping and direction
  const groupingEl = findChartChild(chartTypeEl, "grouping");
  const grouping = groupingEl?.attributes["val"] ?? undefined;
  const barDirEl = findChartChild(chartTypeEl, "barDir");
  const direction = barDirEl?.attributes["val"] ?? undefined;

  // Detect radar style (filled vs standard)
  const radarStyleEl = findChartChild(chartTypeEl, "radarStyle");
  const radarStyle = radarStyleEl?.attributes["val"] ?? undefined;

  // Detect wireframe (surface charts)
  const wireframeEl = findChartChild(chartTypeEl, "wireframe");
  const wireframe =
    wireframeEl?.attributes["val"] === "1" || wireframeEl?.attributes["val"] === "true";

  // Detect marked line (marker symbol != "none")
  let isMarked = false;

  // Parse series
  const series: ChartSeries[] = [];
  for (const child of chartTypeEl.children) {
    if (child.type !== "element") {
      continue;
    }
    if (localName(child.name) !== "ser") {
      continue;
    }
    const s = parseChartSeries(child, chartTag);
    if (s) {
      series.push(s);
      // Detect marked line
      const markerEl = findChartChild(child, "marker");
      if (markerEl) {
        const symbolEl = findChartChild(markerEl, "symbol");
        if (symbolEl && symbolEl.attributes["val"] !== "none") {
          isMarked = true;
        }
      }
    }
  }

  // Resolve chart type
  let chartType = resolveChartType(chartTag, grouping, direction, { radarStyle, wireframe });
  if (chartType === "line" && isMarked) {
    chartType = "lineMarked";
  }

  // Scatter smooth detection
  if (chartTag === "scatterChart") {
    const styleEl = findChartChild(chartTypeEl, "scatterStyle");
    if (styleEl?.attributes["val"] === "smoothMarker") {
      chartType = "scatterSmooth";
    }
  }

  // Parse axes
  const catAxEl = findChartChild(plotAreaEl, "catAx") ?? findChartChild(plotAreaEl, "dateAx");
  const valAxEl = findChartChild(plotAreaEl, "valAx");
  const categoryAxis = catAxEl ? parseChartAxis(catAxEl) : undefined;
  const valueAxis = valAxEl ? parseChartAxis(valAxEl) : undefined;

  // Secondary value axis (second c:valAx in plotArea)
  let secondaryValueAxis: ChartAxis | undefined;
  const allValAxEls: XmlElement[] = [];
  for (const child of plotAreaEl.children) {
    if (child.type === "element" && localName(child.name) === "valAx") {
      allValAxEls.push(child);
    }
  }
  if (allValAxEls.length >= 2) {
    secondaryValueAxis = parseChartAxis(allValAxEls[1]);
  }

  // Chart-level data labels (from chart type element)
  let dataLabels: ChartDataLabels | undefined;
  const chartDLblsEl = findChartChild(chartTypeEl, "dLbls");
  if (chartDLblsEl) {
    const dl: Mutable<ChartDataLabels> = {};
    let hasDl = false;
    const showValEl2 = findChartChild(chartDLblsEl, "showVal");
    if (showValEl2?.attributes["val"] === "1") {
      dl.showValue = true;
      hasDl = true;
    }
    const showCatEl = findChartChild(chartDLblsEl, "showCatName");
    if (showCatEl?.attributes["val"] === "1") {
      dl.showCategory = true;
      hasDl = true;
    }
    const showSerEl = findChartChild(chartDLblsEl, "showSerName");
    if (showSerEl?.attributes["val"] === "1") {
      dl.showSerName = true;
      hasDl = true;
    }
    const showPctEl = findChartChild(chartDLblsEl, "showPercent");
    if (showPctEl?.attributes["val"] === "1") {
      dl.showPercent = true;
      hasDl = true;
    }
    const dlPosEl = findChartChild(chartDLblsEl, "dLblPos");
    if (dlPosEl?.attributes["val"]) {
      dl.position = dlPosEl.attributes["val"] as ChartDataLabelPosition;
      hasDl = true;
    }
    if (hasDl) {
      dataLabels = dl;
    }
  }

  // Legend
  const legendEl = findChartChild(chartEl, "legend");
  let legend: ChartLegendPosition | undefined;
  if (legendEl) {
    const posEl = findChartChild(legendEl, "legendPos");
    legend = (posEl?.attributes["val"] as ChartLegendPosition) ?? "r";
  }

  // Plot area color
  let plotAreaColor: string | undefined;
  const plotSpPrEl = findChartChild(plotAreaEl, "spPr");
  if (plotSpPrEl) {
    plotAreaColor = extractSolidFillColor(plotSpPrEl);
  }

  // Chart area color (from c:chartSpace > c:spPr)
  let chartAreaColor: string | undefined;
  const chartSpPrEl = findChartChild(root, "spPr");
  if (chartSpPrEl) {
    chartAreaColor = extractSolidFillColor(chartSpPrEl);
  }

  // Style
  let style: number | undefined;
  const styleEl = findChartChild(root, "style");
  if (styleEl) {
    const v = styleEl.attributes["val"];
    if (v) {
      const n = parseInt(v, 10);
      if (Number.isFinite(n)) {
        style = n;
      }
    }
  }

  const chart: Chart = {
    type: chartType,
    series,
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

  return chart;
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

  // Per-point colors
  const pointColors: string[] = [];
  for (const child of serEl.children) {
    if (child.type !== "element") {
      continue;
    }
    if (localName(child.name) !== "dPt") {
      continue;
    }
    const ptSpPr = findChartChild(child, "spPr");
    if (ptSpPr) {
      const c = extractSolidFillColor(ptSpPr);
      if (c) {
        pointColors.push(c);
      }
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
    ...(pointColors.length > 0 && { pointColors }),
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
function parseChartExXml(xmlStr: string): ChartExData | undefined {
  let doc: ReturnType<typeof parseXml>;
  try {
    doc = parseXml(xmlStr);
  } catch {
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
