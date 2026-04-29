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
  "cx:dataId",
  "cx:layoutPr",
  "cx:axisId",
  "cx:dataLabels",
  "cx:dataPt",
  "cx:extLst"
]);
const KNOWN_CHILDREN_AXIS = new Set([
  "cx:catScaling",
  "cx:valScaling",
  "cx:hidden",
  "cx:majorTickMark",
  "cx:minorTickMark",
  "cx:numFmt",
  "cx:title",
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

  return {
    chartSpace: {
      chartData: parseChartData(chartDataEl),
      chart: parseChart(chartEl),
      extLst: rawElement(root, "cx:extLst")
    },
    rawXml,
    ...(collector.entries.length > 0 ? { unknownElements: collector.entries } : {})
  };
}

function parseChartData(
  chartDataEl: XmlElement | undefined
): ChartExModel["chartSpace"]["chartData"] {
  if (!chartDataEl) {
    return { data: [] };
  }

  const data = findChildren(chartDataEl, "cx:data").map(parseDataEntry);
  const externalData = findChildren(chartDataEl, "cx:externalData").map(el => ({
    id: el.attributes["r:id"] ?? "",
    autoUpdate: parseBoolAttr(el, "autoUpdate")
  }));

  return {
    data,
    ...(externalData.length > 0 ? { externalData } : {})
  };
}

function parseDataEntry(el: XmlElement): ChartExDataEntry {
  const strDim = findChild(el, "cx:strDim");
  const numDim = findChild(el, "cx:numDim");
  return {
    id: parseInt(el.attributes.id ?? "0", 10),
    ...(strDim ? { strDim: parseStringDimension(strDim) } : {}),
    ...(numDim ? { numDim: parseNumericDimension(numDim) } : {})
  };
}

function parseStringDimension(el: XmlElement): NonNullable<ChartExDataEntry["strDim"]> {
  return {
    type: parseDimensionType(el.attributes.type),
    formula: childText(el, "cx:f"),
    levels: findChildren(el, "cx:lvl").map(lvl => ({
      ptCount: parseOptionalInt(lvl.attributes.ptCount),
      points: findChildren(lvl, "cx:pt").map(pt => ({
        index: parseInt(pt.attributes.idx ?? "0", 10),
        value: textContent(pt)
      }))
    }))
  };
}

function parseNumericDimension(el: XmlElement): NonNullable<ChartExDataEntry["numDim"]> {
  return {
    type: parseDimensionType(el.attributes.type),
    formula: childText(el, "cx:f"),
    levels: findChildren(el, "cx:lvl").map(lvl => ({
      ptCount: parseOptionalInt(lvl.attributes.ptCount),
      formatCode: lvl.attributes.formatCode,
      points: findChildren(lvl, "cx:pt").map(pt => ({
        index: parseInt(pt.attributes.idx ?? "0", 10),
        value: parseFloat(textContent(pt))
      }))
    }))
  };
}

function parseChart(chartEl: XmlElement | undefined): ChartExModel["chartSpace"]["chart"] {
  const plotAreaEl = chartEl ? findChild(chartEl, "cx:plotArea") : undefined;
  return {
    title: chartEl ? parseTitle(findChild(chartEl, "cx:title")) : undefined,
    autoTitleDeleted: chartEl
      ? parseBoolAttr(findChild(chartEl, "cx:autoTitleDeleted"), "val")
      : undefined,
    plotArea: parsePlotArea(plotAreaEl),
    legend: chartEl ? parseLegend(findChild(chartEl, "cx:legend")) : undefined,
    spPr: parseRawSpPr(chartEl ? findChild(chartEl, "cx:spPr") : undefined)
  };
}

function parsePlotArea(
  plotAreaEl: XmlElement | undefined
): ChartExModel["chartSpace"]["chart"]["plotArea"] {
  if (!plotAreaEl) {
    return {};
  }

  const regionEl = findChild(plotAreaEl, "cx:plotAreaRegion");
  const axis = findChildren(plotAreaEl, "cx:axis").map(parseAxis);

  if (regionEl) {
    return {
      plotAreaRegion: {
        layout: parseRawLayout(findChild(regionEl, "cx:layout")),
        plotSurface: parseRawSpPr(findChild(regionEl, "cx:spPr")),
        series: findChildren(regionEl, "cx:series").map(parseSeries)
      },
      spPr: parseRawSpPr(findChild(plotAreaEl, "cx:spPr")),
      ...(axis.length > 0 ? { axis } : {})
    };
  }

  const series = findChildren(plotAreaEl, "cx:series").map(parseSeries);
  return {
    ...(series.length > 0 ? { series } : {}),
    spPr: parseRawSpPr(findChild(plotAreaEl, "cx:spPr")),
    ...(axis.length > 0 ? { axis } : {})
  };
}

function parseSeries(el: XmlElement): ChartExSeries {
  const dataRefs = findChildren(el, "cx:dataId").map(dataId => ({
    dataId: parseInt(dataId.attributes.val ?? "0", 10)
  }));
  const axisId = findChildren(el, "cx:axisId").map(axis =>
    parseInt(axis.attributes.val ?? "0", 10)
  );
  const tx = parseSeriesText(findChild(el, "cx:tx"));
  const layoutPr = parseLayoutProperties(findChild(el, "cx:layoutPr"));

  return {
    layoutId: (el.attributes.layoutId ?? "sunburst") as ChartExSeries["layoutId"],
    hidden: parseBoolAttr(el, "hidden"),
    ownerIdx: parseOptionalInt(el.attributes.ownerIdx),
    ...(tx ? { tx } : {}),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    ...(dataRefs.length > 0 ? { dataRefs } : {}),
    ...(layoutPr ? { layoutPr } : {}),
    ...(axisId.length > 0 ? { axisId } : {}),
    dataLabels: parseDataLabels(findChild(el, "cx:dataLabels")),
    dataPt: findChildren(el, "cx:dataPt").map(pt => ({
      idx: parseInt(pt.attributes.idx ?? "0", 10),
      spPr: parseRawSpPr(findChild(pt, "cx:spPr"))
    })),
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseSeriesText(el: XmlElement | undefined): ChartExSeries["tx"] | undefined {
  const txData = el ? findChild(el, "cx:txData") : undefined;
  if (!txData) {
    return undefined;
  }
  const formula = childText(txData, "cx:f");
  if (formula) {
    return { strRef: formula };
  }
  const value = childText(txData, "cx:v");
  return value !== undefined ? { value } : undefined;
}

function parseLayoutProperties(el: XmlElement | undefined): ChartExSeries["layoutPr"] | undefined {
  if (!el) {
    return undefined;
  }

  const binning = findChild(el, "cx:binning");
  return {
    _rawXml: serializeElement(el),
    parentLabelLayout: childVal(el, "cx:parentLabelLayout") as any,
    subtotals: findChildren(findChild(el, "cx:subtotals") ?? emptyElement(), "cx:subtotal").map(
      st => ({
        idx: parseInt(st.attributes.idx ?? "0", 10)
      })
    ),
    connectorLines: parseBoolAttr(findChild(el, "cx:connectorLines"), "val"),
    binning: binning
      ? {
          intervalClosed: binning.attributes.intervalClosed as any,
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
      : undefined,
    paretoLine: parseBoolAttr(findChild(el, "cx:paretoLine"), "val"),
    quartileMethod: childVal(el, "cx:quartileMethod") as any,
    showMeanLine: parseBoolAttr(findChild(el, "cx:showMeanLine"), "val"),
    showMeanMarker: parseBoolAttr(findChild(el, "cx:showMeanMarker"), "val"),
    showInnerPoints: parseBoolAttr(findChild(el, "cx:showInnerPoints"), "val"),
    showOutlierPoints: parseBoolAttr(findChild(el, "cx:showOutlierPoints"), "val"),
    projection: childVal(el, "cx:projection") as any,
    regionLabels: childVal(el, "cx:regionLabels") as any,
    geoMappingLevel: childVal(el, "cx:geoMappingLevel") as any,
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseDataLabels(el: XmlElement | undefined): ChartExSeries["dataLabels"] | undefined {
  if (!el) {
    return undefined;
  }
  const visibility = findChild(el, "cx:visibility");
  return {
    visibility: visibility
      ? {
          seriesName: parseBoolAttr(visibility, "seriesName"),
          categoryName: parseBoolAttr(visibility, "categoryName"),
          value: parseBoolAttr(visibility, "value"),
          numFmt: parseBoolAttr(visibility, "numFmt")
        }
      : undefined,
    position: findChild(el, "cx:dataLabel")?.attributes.pos,
    separator: childText(el, "cx:separator"),
    numFmt: findChild(el, "cx:numFmt")?.attributes.formatCode,
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr")
  };
}

function parseAxis(el: XmlElement): ChartExAxis {
  const numFmt = findChild(el, "cx:numFmt");
  const valScaling = findChild(el, "cx:valScaling");
  const catScaling = findChild(el, "cx:catScaling");
  return {
    axisId: parseInt(el.attributes.id ?? "0", 10),
    type: findChild(el, "cx:valScaling") ? "val" : "cat",
    hidden: parseBoolAttr(findChild(el, "cx:hidden"), "val"),
    majorTickMark: childVal(el, "cx:majorTickMark") as any,
    minorTickMark: childVal(el, "cx:minorTickMark") as any,
    numFmt: numFmt
      ? {
          formatCode: numFmt.attributes.formatCode ?? "",
          sourceLinked: parseBoolAttr(numFmt, "sourceLinked")
        }
      : undefined,
    title: parseTitle(findChild(el, "cx:title")),
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
    extLst: rawElement(el, "cx:extLst")
  };
}

function parseLegend(el: XmlElement | undefined): ChartLegend | undefined {
  if (!el) {
    return undefined;
  }
  return {
    legendPos: el.attributes.pos as any,
    overlay: parseBoolAttr(el, "overlay"),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    legendEntries: findChildren(el, "cx:legendEntry").map(entry => ({
      index: parseInt(entry.attributes.idx ?? "0", 10)
    }))
  };
}

function parseTitle(el: XmlElement | undefined): ChartTitle | undefined {
  if (!el) {
    return undefined;
  }
  return {
    text: parseRichText(findChild(findChild(el, "cx:tx") ?? emptyElement(), "cx:rich")),
    overlay: parseBoolAttr(findChild(el, "cx:overlay"), "val"),
    spPr: parseRawSpPr(findChild(el, "cx:spPr")),
    txPr: parseRawTxPr(el, "cx:txPr"),
    rawTx: rawElement(el, "cx:tx")
  };
}

function parseRichText(el: XmlElement | undefined): ChartRichText | undefined {
  if (!el) {
    return undefined;
  }
  const paragraphs = findChildren(el, "a:p").map(p => ({
    runs: findChildren(p, "a:r").map(r => ({
      text: childText(r, "a:t") ?? ""
    }))
  }));
  return paragraphs.length > 0 ? { paragraphs } : undefined;
}

function parseRawLayout(el: XmlElement | undefined): ChartTitle["layout"] | undefined {
  const raw = rawElementFromNode(el);
  return raw ? ({ _rawXml: raw } as any) : undefined;
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
    .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
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

function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(value: string): string {
  return escapeXml(value).replace(/"/g, "&quot;");
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
  return value === "1" || value === "true";
}

function parseOptionalInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) ? undefined : parsed;
}

function parseOptionalFloat(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const parsed = parseFloat(value);
  return Number.isNaN(parsed) ? undefined : parsed;
}

function parseDimensionType(value: string | undefined): ChartExDimensionType {
  return (value ?? "val") as ChartExDimensionType;
}

function emptyElement(): XmlElement {
  return { type: "element", name: "", attributes: {}, children: [] as XmlNode[] };
}
