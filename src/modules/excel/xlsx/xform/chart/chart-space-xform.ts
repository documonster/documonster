/**
 * ChartSpaceXform — Full OOXML DrawingML Chart parser and renderer.
 *
 * Design: single class with state-stack SAX parsing. Shape properties
 * (c:spPr) and text properties (c:txPr) are captured as raw XML strings
 * for perfect round-trip fidelity. Structured access is provided for the
 * most commonly used properties.
 *
 * The c: namespace is http://schemas.openxmlformats.org/drawingml/2006/chart
 * The a: namespace is http://schemas.openxmlformats.org/drawingml/2006/main
 * The r: namespace is http://schemas.openxmlformats.org/officeDocument/2006/relationships
 */

import type {
  ChartModel,
  ChartData,
  PlotArea,
  ChartTypeGroup,
  BarChartGroup,
  LineChartGroup,
  PieChartGroup,
  DoughnutChartGroup,
  AreaChartGroup,
  ScatterChartGroup,
  BubbleChartGroup,
  RadarChartGroup,
  StockChartGroup,
  SurfaceChartGroup,
  OfPieChartGroup,
  BarSeries,
  LineSeries,
  PieSeries,
  AreaSeries,
  ScatterSeries,
  BubbleSeries,
  RadarSeries,
  SurfaceSeries,
  ChartAxis,
  CategoryAxis,
  ValueAxis,
  DateAxis,
  SeriesAxis,
  ChartTitle,
  ChartLegend,
  ChartLayout,
  ManualLayout,
  View3D,
  DataTable,
  DataLabels,
  DataLabelEntry,
  Trendline,
  TrendlineLabel,
  ErrorBars,
  ErrorBarType,
  ChartMarker,
  DataPoint,
  NumberReference,
  StringReference,
  NumberCache,
  StringCache,
  NumberDataSource,
  AxisDataSource,
  NumberLiteral,
  StringLiteral,
  MultiLevelStringReference,
  ShapeProperties,
  ChartColor,
  ChartTextProperties,
  ChartRichText,
  ChartBodyProperties,
  ChartParagraphProperties,
  ChartLineSpacing,
  EffectList,
  Shadow,
  Scene3D,
  ShapeProperties3D,
  PrintSettings,
  PivotFormat,
  UpDownBars,
  DisplayUnits,
  LegendEntry,
  PictureOptions
} from "@excel/chart/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

// ============================================================================
// Raw XML capture helper
// ============================================================================

/**
 * Captures raw XML during SAX parsing for elements we want to round-trip
 * without full structural modelling (spPr, txPr, etc.).
 */
class RawXmlCapture {
  private depth = 0;
  private parts: string[] = [];
  private rootTag = "";
  /** Track self-closing tags so the subsequent parseClose is skipped */
  private skipNextClose = false;

  start(node: any): void {
    this.rootTag = node.name;
    this.depth = 1;
    this.parts = [this._openTag(node)];
  }

  handleOpen(node: any): void {
    this.depth++;
    this.parts.push(this._openTag(node));
  }

  handleText(text: string): void {
    this.parts.push(escapeXml(text));
  }

  handleClose(name: string): boolean {
    // SAX fires closetag for self-closing elements too — skip those
    if (this.skipNextClose) {
      this.skipNextClose = false;
      return true;
    }
    this.depth--;
    if (this.depth === 0) {
      this.parts.push(`</${name}>`);
      return false; // done
    }
    this.parts.push(`</${name}>`);
    return true; // still capturing
  }

  handleSelfClose(node: any): void {
    // For self-closing tags encountered during capture.
    // SAX will also fire a closetag event — flag to skip it.
    this.parts.push(this._selfCloseTag(node));
    this.skipNextClose = true;
  }

  get xml(): string {
    return this.parts.join("");
  }

  get active(): boolean {
    return this.depth > 0;
  }

  private _openTag(node: any): string {
    let s = `<${node.name}`;
    if (node.attributes) {
      for (const [k, v] of Object.entries(node.attributes)) {
        s += ` ${k}="${escapeXml(String(v))}"`;
      }
    }
    s += ">";
    return s;
  }

  private _selfCloseTag(node: any): string {
    let s = `<${node.name}`;
    if (node.attributes) {
      for (const [k, v] of Object.entries(node.attributes)) {
        s += ` ${k}="${escapeXml(String(v))}"`;
      }
    }
    return s + "/>";
  }
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ============================================================================
// Parse state
// ============================================================================

interface ParseState {
  tag: string;
  context: any;
}

// Chart type element tags
const CHART_TYPE_TAGS = new Set([
  "c:barChart",
  "c:bar3DChart",
  "c:lineChart",
  "c:line3DChart",
  "c:pieChart",
  "c:pie3DChart",
  "c:doughnutChart",
  "c:areaChart",
  "c:area3DChart",
  "c:scatterChart",
  "c:bubbleChart",
  "c:radarChart",
  "c:stockChart",
  "c:surfaceChart",
  "c:surface3DChart",
  "c:ofPieChart"
]);

// Axis element tags
const AXIS_TAGS = new Set(["c:catAx", "c:valAx", "c:dateAx", "c:serAx"]);

// ============================================================================
// ChartSpaceXform
// ============================================================================

class ChartSpaceXform extends BaseXform<ChartModel> {
  private stateStack: ParseState[] = [];
  private rawCapture: RawXmlCapture | null = null;
  private textBuf = "";

  // Temporary parse state
  private chartModel!: ChartModel;
  private chartData!: ChartData;
  private plotArea!: PlotArea;
  private currentChartTypeGroup: any = null;
  private currentSeries: any = null;
  private currentAxis: any = null;
  private currentTitle: any = null;
  private currentLegend: ChartLegend | null = null;
  private currentDataLabels: DataLabels | null = null;
  private currentDataLabelEntry: DataLabelEntry | null = null;
  private currentMarker: ChartMarker | null = null;
  private currentDataPoint: DataPoint | null = null;
  private currentTrendline: Trendline | null = null;
  private currentErrorBars: ErrorBars | null = null;
  private currentNumRef: NumberReference | null = null;
  private currentStrRef: StringReference | null = null;
  private currentNumLit: NumberLiteral | null = null;
  private currentStrLit: StringLiteral | null = null;
  private currentNumCache: NumberCache | null = null;
  private currentStrCache: StringCache | null = null;
  private currentLayout: ChartLayout | null = null;
  private currentManualLayout: ManualLayout | null = null;
  private currentView3D: View3D | null = null;
  private currentUpDownBars: UpDownBars | null = null;
  private currentLegendEntry: LegendEntry | null = null;
  private currentDataTable: DataTable | null = null;
  private currentDisplayUnits: DisplayUnits | null = null;
  private currentPivotFormat: PivotFormat | null = null;
  private currentTrendlineLbl: TrendlineLabel | null = null;
  private currentPictureOptions: PictureOptions | null = null;
  private currentPrintSettings: PrintSettings | null = null;
  private currentMultiLvlStrRef: MultiLevelStringReference | null = null;
  private currentLvl: { points: Array<{ index: number; value: string }> } | null = null;
  /** Tracks c:floor / c:sideWall / c:backWall context */
  private currentFloorWall: string | null = null;
  /** Tracks c:hiLowLines / c:dropLines / c:serLines context */
  private currentLineContext: string | null = null;
  /** Pending spPr captured inside a line context element */
  private _pendingLineSpPr: ShapeProperties | null = null;

  // For collecting spPr/txPr raw XML to attach to the correct parent
  private rawXmlTarget: string | null = null;
  /** Context hint for c:extLst routing */
  private rawCaptureContext: string | null = null;
  /** Whether we are currently inside c:plotArea during parsing */
  private insidePlotArea = false;
  private rawSpPrXml: string | null = null;
  private rawTxPrXml: string | null = null;

  get tag(): string {
    return "c:chartSpace";
  }

  reset(): void {
    super.reset();
    this.stateStack = [];
    this.rawCapture = null;
    this.rawCaptureContext = null;
    this.textBuf = "";
  }

  // ============================================================================
  // RENDER
  // ============================================================================

  render(xmlStream: XmlSink, model?: ChartModel): void {
    const m = model ?? this.model;
    if (!m) {
      return;
    }

    const nsAttrs: Record<string, string> = { ...ChartSpaceXform.CHART_SPACE_ATTRIBUTES };
    if (m.extraNamespaces) {
      Object.assign(nsAttrs, m.extraNamespaces);
    }
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode("c:chartSpace", nsAttrs);

    if (m.date1904 !== undefined) {
      xmlStream.leafNode("c:date1904", { val: m.date1904 ? "1" : "0" });
    }
    if (m.lang) {
      xmlStream.leafNode("c:lang", { val: m.lang });
    }
    if (m.roundedCorners !== undefined) {
      xmlStream.leafNode("c:roundedCorners", { val: m.roundedCorners ? "1" : "0" });
    }
    if (m.alternateContentStyle) {
      xmlStream.writeRaw(m.alternateContentStyle);
    } else if (m.style !== undefined) {
      xmlStream.leafNode("c:style", { val: String(m.style) });
    }
    if (m.pivotSource) {
      xmlStream.writeRaw(m.pivotSource);
    }
    if (m.protection) {
      xmlStream.writeRaw(m.protection);
    }

    this._renderChart(xmlStream, m.chart);

    if (m.clrMapOvr) {
      xmlStream.writeRaw(m.clrMapOvr);
    }
    if (m.spPr) {
      this._renderSpPr(xmlStream, m.spPr);
    }
    if (m.txPr) {
      this._renderTxPr(xmlStream, m.txPr);
    }
    if (m.printSettings) {
      this._renderPrintSettings(xmlStream, m.printSettings);
    }
    if (m.externalData) {
      xmlStream.openNode("c:externalData", { "r:id": m.externalData.id });
      if (m.externalData.autoUpdate !== undefined) {
        xmlStream.leafNode("c:autoUpdate", { val: m.externalData.autoUpdate ? "1" : "0" });
      }
      xmlStream.closeNode();
    }
    if (m.extLst) {
      xmlStream.writeRaw(m.extLst);
    }

    xmlStream.closeNode();
  }

  private _renderChart(xml: XmlSink, chart: ChartData): void {
    xml.openNode("c:chart");

    if (chart.title) {
      this._renderTitle(xml, chart.title, "c:title");
    }
    if (chart.autoTitleDeleted !== undefined) {
      xml.leafNode("c:autoTitleDeleted", { val: chart.autoTitleDeleted ? "1" : "0" });
    }
    if (chart.pivotFormats && chart.pivotFormats.length > 0) {
      xml.openNode("c:pivotFmts");
      for (const pf of chart.pivotFormats) {
        xml.openNode("c:pivotFmt");
        xml.leafNode("c:idx", { val: String(pf.index) });
        if (pf.spPr) {
          this._renderSpPr(xml, pf.spPr);
        }
        if (pf.marker) {
          this._renderMarker(xml, pf.marker);
        }
        if (pf.rawDLbl) {
          xml.writeRaw(pf.rawDLbl);
        } else if (pf.dataLabels) {
          this._renderDataLabels(xml, pf.dataLabels);
        }
        xml.closeNode();
      }
      xml.closeNode();
    }
    if (chart.view3D) {
      this._renderView3D(xml, chart.view3D);
    }
    if (chart.floor) {
      xml.openNode("c:floor");
      this._renderSpPr(xml, chart.floor);
      xml.closeNode();
    }
    if (chart.sideWall) {
      xml.openNode("c:sideWall");
      this._renderSpPr(xml, chart.sideWall);
      xml.closeNode();
    }
    if (chart.backWall) {
      xml.openNode("c:backWall");
      this._renderSpPr(xml, chart.backWall);
      xml.closeNode();
    }

    this._renderPlotArea(xml, chart.plotArea);

    if (chart.legend) {
      this._renderLegend(xml, chart.legend);
    }
    if (chart.plotVisOnly !== undefined) {
      xml.leafNode("c:plotVisOnly", { val: chart.plotVisOnly ? "1" : "0" });
    }
    if (chart.dispBlanksAs) {
      xml.leafNode("c:dispBlanksAs", { val: chart.dispBlanksAs });
    }
    if (chart.showDLblsOverMax !== undefined) {
      xml.leafNode("c:showDLblsOverMax", { val: chart.showDLblsOverMax ? "1" : "0" });
    }
    if (chart.extLst) {
      xml.writeRaw(chart.extLst);
    }

    xml.closeNode();
  }

  private _renderView3D(xml: XmlSink, v: View3D): void {
    xml.openNode("c:view3D");
    if (v.rotX !== undefined) {
      xml.leafNode("c:rotX", { val: String(v.rotX) });
    }
    if (v.rotY !== undefined) {
      xml.leafNode("c:rotY", { val: String(v.rotY) });
    }
    if (v.depthPercent !== undefined) {
      xml.leafNode("c:depthPercent", { val: String(v.depthPercent) });
    }
    if (v.rAngAx !== undefined) {
      xml.leafNode("c:rAngAx", { val: v.rAngAx ? "1" : "0" });
    }
    if (v.hPercent !== undefined) {
      xml.leafNode("c:hPercent", { val: String(v.hPercent) });
    }
    if (v.perspective !== undefined) {
      xml.leafNode("c:perspective", { val: String(v.perspective) });
    }
    xml.closeNode();
  }

  private _renderPlotArea(xml: XmlSink, pa: PlotArea): void {
    xml.openNode("c:plotArea");
    this._renderLayout(xml, pa.layout);

    for (const ctg of pa.chartTypes) {
      this._renderChartTypeGroup(xml, ctg);
    }
    for (const ax of pa.axes) {
      this._renderAxis(xml, ax);
    }
    if (pa.dataTable) {
      this._renderDataTable(xml, pa.dataTable);
    }
    if (pa.spPr) {
      this._renderSpPr(xml, pa.spPr);
    }
    if (pa.extLst) {
      xml.writeRaw(pa.extLst);
    }
    xml.closeNode();
  }

  // ---- Chart type groups ----

  private _renderChartTypeGroup(xml: XmlSink, ctg: ChartTypeGroup): void {
    const tagMap: Record<string, string> = {
      bar: "c:barChart",
      bar3D: "c:bar3DChart",
      line: "c:lineChart",
      line3D: "c:line3DChart",
      pie: "c:pieChart",
      pie3D: "c:pie3DChart",
      doughnut: "c:doughnutChart",
      area: "c:areaChart",
      area3D: "c:area3DChart",
      scatter: "c:scatterChart",
      bubble: "c:bubbleChart",
      radar: "c:radarChart",
      stock: "c:stockChart",
      surface: "c:surfaceChart",
      surface3D: "c:surface3DChart",
      ofPie: "c:ofPieChart"
    };
    const tag = tagMap[ctg.type];
    if (!tag) {
      return;
    }

    xml.openNode(tag);

    switch (ctg.type) {
      case "bar":
      case "bar3D":
        this._renderBarChart(xml, ctg);
        break;
      case "line":
      case "line3D":
        this._renderLineChart(xml, ctg);
        break;
      case "pie":
      case "pie3D":
        this._renderPieChart(xml, ctg);
        break;
      case "doughnut":
        this._renderDoughnutChart(xml, ctg);
        break;
      case "area":
      case "area3D":
        this._renderAreaChart(xml, ctg);
        break;
      case "scatter":
        this._renderScatterChart(xml, ctg);
        break;
      case "bubble":
        this._renderBubbleChart(xml, ctg);
        break;
      case "radar":
        this._renderRadarChart(xml, ctg);
        break;
      case "stock":
        this._renderStockChart(xml, ctg);
        break;
      case "surface":
      case "surface3D":
        this._renderSurfaceChart(xml, ctg);
        break;
      case "ofPie":
        this._renderOfPieChart(xml, ctg);
        break;
    }

    if (ctg.extLst) {
      xml.writeRaw(ctg.extLst);
    }

    xml.closeNode();
  }

  private _renderBarChart(xml: XmlSink, g: BarChartGroup): void {
    xml.leafNode("c:barDir", { val: g.barDir });
    xml.leafNode("c:grouping", { val: g.grouping });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderBarSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.gapWidth !== undefined) {
      xml.leafNode("c:gapWidth", { val: String(g.gapWidth) });
    }
    if (g.overlap !== undefined) {
      xml.leafNode("c:overlap", { val: String(g.overlap) });
    }
    if (g.serLines) {
      xml.openNode("c:serLines");
      this._renderSpPr(xml, g.serLines);
      xml.closeNode();
    }
    if (g.shape) {
      xml.leafNode("c:shape", { val: g.shape });
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderLineChart(xml: XmlSink, g: LineChartGroup): void {
    xml.leafNode("c:grouping", { val: g.grouping });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderLineSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.dropLines) {
      xml.openNode("c:dropLines");
      this._renderSpPr(xml, g.dropLines);
      xml.closeNode();
    }
    if (g.hiLowLines) {
      xml.openNode("c:hiLowLines");
      this._renderSpPr(xml, g.hiLowLines);
      xml.closeNode();
    }
    if (g.upDownBars) {
      this._renderUpDownBars(xml, g.upDownBars);
    }
    if (g.marker !== undefined) {
      xml.leafNode("c:marker", { val: g.marker ? "1" : "0" });
    }
    if (g.smooth !== undefined) {
      xml.leafNode("c:smooth", { val: g.smooth ? "1" : "0" });
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderPieChart(xml: XmlSink, g: PieChartGroup): void {
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderPieSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.firstSliceAng !== undefined) {
      xml.leafNode("c:firstSliceAng", { val: String(g.firstSliceAng) });
    }
  }

  private _renderDoughnutChart(xml: XmlSink, g: DoughnutChartGroup): void {
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderPieSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.firstSliceAng !== undefined) {
      xml.leafNode("c:firstSliceAng", { val: String(g.firstSliceAng) });
    }
    if (g.holeSize !== undefined) {
      xml.leafNode("c:holeSize", { val: String(g.holeSize) });
    }
  }

  private _renderAreaChart(xml: XmlSink, g: AreaChartGroup): void {
    xml.leafNode("c:grouping", { val: g.grouping });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderAreaSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.dropLines) {
      xml.openNode("c:dropLines");
      this._renderSpPr(xml, g.dropLines);
      xml.closeNode();
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderScatterChart(xml: XmlSink, g: ScatterChartGroup): void {
    xml.leafNode("c:scatterStyle", { val: g.scatterStyle });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderScatterSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderBubbleChart(xml: XmlSink, g: BubbleChartGroup): void {
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderBubbleSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.bubbleScale !== undefined) {
      xml.leafNode("c:bubbleScale", { val: String(g.bubbleScale) });
    }
    if (g.showNegBubbles !== undefined) {
      xml.leafNode("c:showNegBubbles", { val: g.showNegBubbles ? "1" : "0" });
    }
    if (g.sizeRepresents) {
      xml.leafNode("c:sizeRepresents", { val: g.sizeRepresents });
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderRadarChart(xml: XmlSink, g: RadarChartGroup): void {
    xml.leafNode("c:radarStyle", { val: g.radarStyle });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderRadarSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderStockChart(xml: XmlSink, g: StockChartGroup): void {
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderLineSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.hiLowLines) {
      xml.openNode("c:hiLowLines");
      this._renderSpPr(xml, g.hiLowLines);
      xml.closeNode();
    }
    if (g.upDownBars) {
      this._renderUpDownBars(xml, g.upDownBars);
    }
    if (g.dropLines) {
      xml.openNode("c:dropLines");
      this._renderSpPr(xml, g.dropLines);
      xml.closeNode();
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderSurfaceChart(xml: XmlSink, g: SurfaceChartGroup): void {
    if (g.wireframe !== undefined) {
      xml.leafNode("c:wireframe", { val: g.wireframe ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderSurfaceSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.bandFormats && g.bandFormats.length > 0) {
      xml.openNode("c:bandFmts");
      for (const bf of g.bandFormats) {
        xml.openNode("c:bandFmt");
        xml.leafNode("c:idx", { val: String(bf.index) });
        if (bf.spPr) {
          this._renderSpPr(xml, bf.spPr);
        }
        xml.closeNode();
      }
      xml.closeNode();
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderOfPieChart(xml: XmlSink, g: OfPieChartGroup): void {
    xml.leafNode("c:ofPieType", { val: g.ofPieType });
    if (g.varyColors !== undefined) {
      xml.leafNode("c:varyColors", { val: g.varyColors ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderPieSeries(xml, s);
    }
    if (g.dataLabels) {
      this._renderDataLabels(xml, g.dataLabels);
    }
    if (g.gapWidth !== undefined) {
      xml.leafNode("c:gapWidth", { val: String(g.gapWidth) });
    }
    if (g.splitType) {
      xml.leafNode("c:splitType", { val: g.splitType });
    }
    if (g.splitPos !== undefined) {
      xml.leafNode("c:splitPos", { val: String(g.splitPos) });
    }
    if (g.custSplit && g.custSplit.length > 0) {
      xml.openNode("c:custSplit");
      for (const idx of g.custSplit) {
        xml.leafNode("c:secondPiePt", { val: String(idx) });
      }
      xml.closeNode();
    }
    if (g.secondPieSize !== undefined) {
      xml.leafNode("c:secondPieSize", { val: String(g.secondPieSize) });
    }
    if (g.serLines) {
      xml.openNode("c:serLines");
      this._renderSpPr(xml, g.serLines);
      xml.closeNode();
    }
  }

  // ---- Series rendering ----

  private _renderSeriesBase(xml: XmlSink, s: any): void {
    xml.leafNode("c:idx", { val: String(s.index) });
    xml.leafNode("c:order", { val: String(s.order) });
    if (s.tx) {
      this._renderSeriesTx(xml, s.tx);
    }
    if (s.spPr) {
      this._renderSpPr(xml, s.spPr);
    }
  }

  private _renderBarSeries(xml: XmlSink, s: BarSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.invertIfNegative !== undefined) {
      xml.leafNode("c:invertIfNegative", { val: s.invertIfNegative ? "1" : "0" });
    }
    if (s.pictureOptions) {
      this._renderPictureOptions(xml, s.pictureOptions);
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.trendlines) {
      for (const t of s.trendlines) {
        this._renderTrendline(xml, t);
      }
    }
    if (s.errorBars) {
      this._renderErrorBars(xml, s.errorBars);
    }
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.shape) {
      xml.leafNode("c:shape", { val: s.shape });
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderLineSeries(xml: XmlSink, s: LineSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.marker) {
      this._renderMarker(xml, s.marker);
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.trendlines) {
      for (const t of s.trendlines) {
        this._renderTrendline(xml, t);
      }
    }
    if (s.errorBars) {
      this._renderErrorBars(xml, s.errorBars);
    }
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.smooth !== undefined) {
      xml.leafNode("c:smooth", { val: s.smooth ? "1" : "0" });
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderPieSeries(xml: XmlSink, s: PieSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.explosion !== undefined) {
      xml.leafNode("c:explosion", { val: String(s.explosion) });
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderAreaSeries(xml: XmlSink, s: AreaSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.trendlines) {
      for (const t of s.trendlines) {
        this._renderTrendline(xml, t);
      }
    }
    if (s.errorBars) {
      this._renderErrorBars(xml, s.errorBars);
    }
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderScatterSeries(xml: XmlSink, s: ScatterSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.marker) {
      this._renderMarker(xml, s.marker);
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.trendlines) {
      for (const t of s.trendlines) {
        this._renderTrendline(xml, t);
      }
    }
    if (s.errorBars) {
      for (const eb of s.errorBars) {
        this._renderErrorBars(xml, eb);
      }
    }
    if (s.xVal) {
      this._renderCatData(xml, s.xVal, "c:xVal");
    }
    if (s.yVal) {
      this._renderValData(xml, s.yVal, "c:yVal");
    }
    if (s.smooth !== undefined) {
      xml.leafNode("c:smooth", { val: s.smooth ? "1" : "0" });
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderBubbleSeries(xml: XmlSink, s: BubbleSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.invertIfNegative !== undefined) {
      xml.leafNode("c:invertIfNegative", { val: s.invertIfNegative ? "1" : "0" });
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.trendlines) {
      for (const t of s.trendlines) {
        this._renderTrendline(xml, t);
      }
    }
    if (s.errorBars) {
      for (const eb of s.errorBars) {
        this._renderErrorBars(xml, eb);
      }
    }
    if (s.xVal) {
      this._renderCatData(xml, s.xVal, "c:xVal");
    }
    if (s.yVal) {
      this._renderValData(xml, s.yVal, "c:yVal");
    }
    if (s.bubbleSize) {
      this._renderValData(xml, s.bubbleSize, "c:bubbleSize");
    }
    if (s.bubble3D !== undefined) {
      xml.leafNode("c:bubble3D", { val: s.bubble3D ? "1" : "0" });
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderRadarSeries(xml: XmlSink, s: RadarSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.marker) {
      this._renderMarker(xml, s.marker);
    }
    if (s.dataPoints) {
      for (const dp of s.dataPoints) {
        this._renderDataPoint(xml, dp);
      }
    }
    if (s.dataLabels) {
      this._renderDataLabels(xml, s.dataLabels);
    }
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  private _renderSurfaceSeries(xml: XmlSink, s: SurfaceSeries): void {
    xml.openNode("c:ser");
    this._renderSeriesBase(xml, s);
    if (s.cat) {
      this._renderCatData(xml, s.cat);
    }
    if (s.val) {
      this._renderValData(xml, s.val);
    }
    if (s.extLst) {
      xml.writeRaw(s.extLst);
    }
    xml.closeNode();
  }

  // ---- Data references ----

  private _renderSeriesTx(xml: XmlSink, tx: { strRef?: StringReference; value?: string }): void {
    xml.openNode("c:tx");
    if (tx.strRef) {
      this._renderStrRef(xml, tx.strRef);
    } else if (tx.value !== undefined) {
      xml.openNode("c:v");
      xml.writeText(tx.value);
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderCatData(xml: XmlSink, d: AxisDataSource, tag = "c:cat"): void {
    xml.openNode(tag);
    if (d.strRef) {
      this._renderStrRef(xml, d.strRef);
    } else if (d.numRef) {
      this._renderNumRef(xml, d.numRef);
    } else if (d.strLit) {
      this._renderStrLit(xml, d.strLit);
    } else if (d.numLit) {
      this._renderNumLit(xml, d.numLit);
    } else if (d.multiLvlStrRef) {
      this._renderMultiLvlStrRef(xml, d.multiLvlStrRef);
    }
    xml.closeNode();
  }

  private _renderValData(xml: XmlSink, d: NumberDataSource, tag = "c:val"): void {
    xml.openNode(tag);
    if (d.numRef) {
      this._renderNumRef(xml, d.numRef);
    } else if (d.numLit) {
      this._renderNumLit(xml, d.numLit);
    }
    xml.closeNode();
  }

  private _renderNumRef(xml: XmlSink, ref: NumberReference): void {
    xml.openNode("c:numRef");
    xml.openNode("c:f");
    xml.writeText(ref.formula);
    xml.closeNode();
    if (ref.cache) {
      this._renderNumCache(xml, ref.cache);
    }
    xml.closeNode();
  }

  private _renderStrRef(xml: XmlSink, ref: StringReference): void {
    xml.openNode("c:strRef");
    xml.openNode("c:f");
    xml.writeText(ref.formula);
    xml.closeNode();
    if (ref.cache) {
      this._renderStrCache(xml, ref.cache);
    }
    xml.closeNode();
  }

  private _renderNumCache(xml: XmlSink, cache: NumberCache): void {
    xml.openNode("c:numCache");
    if (cache.formatCode) {
      xml.leafNode("c:formatCode", undefined, cache.formatCode);
    }
    if (cache.pointCount !== undefined) {
      xml.leafNode("c:ptCount", { val: String(cache.pointCount) });
    }
    for (const pt of cache.points) {
      if (pt.value !== null) {
        const attrs: any = { idx: String(pt.index) };
        if (pt.formatCode) {
          attrs.formatCode = pt.formatCode;
        }
        xml.openNode("c:pt", attrs);
        xml.openNode("c:v");
        xml.writeText(String(pt.value));
        xml.closeNode();
        xml.closeNode();
      }
    }
    xml.closeNode();
  }

  private _renderStrCache(xml: XmlSink, cache: StringCache): void {
    xml.openNode("c:strCache");
    if (cache.pointCount !== undefined) {
      xml.leafNode("c:ptCount", { val: String(cache.pointCount) });
    }
    for (const pt of cache.points) {
      xml.openNode("c:pt", { idx: String(pt.index) });
      xml.openNode("c:v");
      xml.writeText(pt.value);
      xml.closeNode();
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderNumLit(xml: XmlSink, lit: NumberLiteral): void {
    xml.openNode("c:numLit");
    if (lit.formatCode) {
      xml.leafNode("c:formatCode", undefined, lit.formatCode);
    }
    if (lit.pointCount !== undefined) {
      xml.leafNode("c:ptCount", { val: String(lit.pointCount) });
    }
    for (const pt of lit.points) {
      if (pt.value !== null) {
        xml.openNode("c:pt", { idx: String(pt.index) });
        xml.openNode("c:v");
        xml.writeText(String(pt.value));
        xml.closeNode();
        xml.closeNode();
      }
    }
    xml.closeNode();
  }

  private _renderStrLit(xml: XmlSink, lit: StringLiteral): void {
    xml.openNode("c:strLit");
    if (lit.pointCount !== undefined) {
      xml.leafNode("c:ptCount", { val: String(lit.pointCount) });
    }
    for (const pt of lit.points) {
      xml.openNode("c:pt", { idx: String(pt.index) });
      xml.openNode("c:v");
      xml.writeText(pt.value);
      xml.closeNode();
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderMultiLvlStrRef(xml: XmlSink, ref: MultiLevelStringReference): void {
    xml.openNode("c:multiLvlStrRef");
    xml.openNode("c:f");
    xml.writeText(ref.formula);
    xml.closeNode();
    if (ref.cache) {
      xml.openNode("c:multiLvlStrCache");
      if (ref.cache.pointCount !== undefined) {
        xml.leafNode("c:ptCount", { val: String(ref.cache.pointCount) });
      }
      for (const level of ref.cache.levels) {
        xml.openNode("c:lvl");
        for (const pt of level.points) {
          xml.openNode("c:pt", { idx: String(pt.index) });
          xml.openNode("c:v");
          xml.writeText(pt.value);
          xml.closeNode();
          xml.closeNode();
        }
        xml.closeNode();
      }
      xml.closeNode();
    }
    xml.closeNode();
  }

  // ---- Marker, DataPoint, DataLabels ----

  private _renderMarker(xml: XmlSink, m: ChartMarker): void {
    xml.openNode("c:marker");
    if (m.symbol) {
      xml.leafNode("c:symbol", { val: m.symbol });
    }
    if (m.size !== undefined) {
      xml.leafNode("c:size", { val: String(m.size) });
    }
    if (m.spPr) {
      this._renderSpPr(xml, m.spPr);
    }
    if (m.extLst) {
      xml.writeRaw(m.extLst);
    }
    xml.closeNode();
  }

  private _renderDataPoint(xml: XmlSink, dp: DataPoint): void {
    xml.openNode("c:dPt");
    xml.leafNode("c:idx", { val: String(dp.index) });
    if (dp.invertIfNegative !== undefined) {
      xml.leafNode("c:invertIfNegative", { val: dp.invertIfNegative ? "1" : "0" });
    }
    if (dp.marker) {
      this._renderMarker(xml, dp.marker);
    }
    if (dp.bubble3D !== undefined) {
      xml.leafNode("c:bubble3D", { val: dp.bubble3D ? "1" : "0" });
    }
    if (dp.explosion !== undefined) {
      xml.leafNode("c:explosion", { val: String(dp.explosion) });
    }
    if (dp.spPr) {
      this._renderSpPr(xml, dp.spPr);
    }
    if (dp.pictureOptions) {
      this._renderPictureOptions(xml, dp.pictureOptions);
    }
    if (dp.extLst) {
      xml.writeRaw(dp.extLst);
    }
    xml.closeNode();
  }

  private _renderDataLabels(xml: XmlSink, dl: DataLabels): void {
    xml.openNode("c:dLbls");
    if (dl.entries) {
      for (const e of dl.entries) {
        xml.openNode("c:dLbl");
        xml.leafNode("c:idx", { val: String(e.index) });
        if (e.layout) {
          this._renderLayout(xml, e.layout);
        }
        if (e.rawTx) {
          xml.writeRaw(e.rawTx);
        } else if (e.text) {
          this._renderRichText(xml, e.text, "c:tx");
        }
        if (e.numFmt) {
          xml.leafNode("c:numFmt", {
            formatCode: e.numFmt.formatCode,
            ...(e.numFmt.sourceLinked !== undefined
              ? { sourceLinked: e.numFmt.sourceLinked ? "1" : "0" }
              : {})
          });
        }
        if (e.spPr) {
          this._renderSpPr(xml, e.spPr);
        }
        if (e.txPr) {
          this._renderTxPr(xml, e.txPr);
        }
        if (e.position) {
          xml.leafNode("c:dLblPos", { val: e.position });
        }
        if (e.showLegendKey !== undefined) {
          xml.leafNode("c:showLegendKey", { val: e.showLegendKey ? "1" : "0" });
        }
        if (e.showVal !== undefined) {
          xml.leafNode("c:showVal", { val: e.showVal ? "1" : "0" });
        }
        if (e.showCatName !== undefined) {
          xml.leafNode("c:showCatName", { val: e.showCatName ? "1" : "0" });
        }
        if (e.showSerName !== undefined) {
          xml.leafNode("c:showSerName", { val: e.showSerName ? "1" : "0" });
        }
        if (e.showPercent !== undefined) {
          xml.leafNode("c:showPercent", { val: e.showPercent ? "1" : "0" });
        }
        if (e.showBubbleSize !== undefined) {
          xml.leafNode("c:showBubbleSize", { val: e.showBubbleSize ? "1" : "0" });
        }
        if (e.separator) {
          xml.openNode("c:separator");
          xml.writeText(e.separator);
          xml.closeNode();
        }
        if (e.delete !== undefined) {
          xml.leafNode("c:delete", { val: e.delete ? "1" : "0" });
        }
        if (e.extLst) {
          xml.writeRaw(e.extLst);
        }
        xml.closeNode();
      }
    }
    if (dl.numFmt) {
      xml.leafNode("c:numFmt", {
        formatCode: dl.numFmt.formatCode,
        ...(dl.numFmt.sourceLinked !== undefined
          ? { sourceLinked: dl.numFmt.sourceLinked ? "1" : "0" }
          : {})
      });
    }
    if (dl.spPr) {
      this._renderSpPr(xml, dl.spPr);
    }
    if (dl.txPr) {
      this._renderTxPr(xml, dl.txPr);
    }
    if (dl.position) {
      xml.leafNode("c:dLblPos", { val: dl.position });
    }
    if (dl.showLegendKey !== undefined) {
      xml.leafNode("c:showLegendKey", { val: dl.showLegendKey ? "1" : "0" });
    }
    if (dl.showVal !== undefined) {
      xml.leafNode("c:showVal", { val: dl.showVal ? "1" : "0" });
    }
    if (dl.showCatName !== undefined) {
      xml.leafNode("c:showCatName", { val: dl.showCatName ? "1" : "0" });
    }
    if (dl.showSerName !== undefined) {
      xml.leafNode("c:showSerName", { val: dl.showSerName ? "1" : "0" });
    }
    if (dl.showPercent !== undefined) {
      xml.leafNode("c:showPercent", { val: dl.showPercent ? "1" : "0" });
    }
    if (dl.showBubbleSize !== undefined) {
      xml.leafNode("c:showBubbleSize", { val: dl.showBubbleSize ? "1" : "0" });
    }
    if (dl.showLeaderLines !== undefined) {
      xml.leafNode("c:showLeaderLines", { val: dl.showLeaderLines ? "1" : "0" });
    }
    if (dl.separator) {
      xml.openNode("c:separator");
      xml.writeText(dl.separator);
      xml.closeNode();
    }
    if (dl.extLst) {
      xml.writeRaw(dl.extLst);
    }
    xml.closeNode();
  }

  // ---- Trendline, ErrorBars ----

  private _renderTrendline(xml: XmlSink, t: Trendline): void {
    xml.openNode("c:trendline");
    if (t.name) {
      xml.openNode("c:name");
      xml.writeText(t.name);
      xml.closeNode();
    }
    if (t.spPr) {
      this._renderSpPr(xml, t.spPr);
    }
    xml.leafNode("c:trendlineType", { val: t.type });
    if (t.order !== undefined) {
      xml.leafNode("c:order", { val: String(t.order) });
    }
    if (t.period !== undefined) {
      xml.leafNode("c:period", { val: String(t.period) });
    }
    if (t.forward !== undefined) {
      xml.leafNode("c:forward", { val: String(t.forward) });
    }
    if (t.backward !== undefined) {
      xml.leafNode("c:backward", { val: String(t.backward) });
    }
    if (t.intercept !== undefined) {
      xml.leafNode("c:intercept", { val: String(t.intercept) });
    }
    if (t.displayRSqr !== undefined) {
      xml.leafNode("c:dispRSqr", { val: t.displayRSqr ? "1" : "0" });
    }
    if (t.displayEq !== undefined) {
      xml.leafNode("c:dispEq", { val: t.displayEq ? "1" : "0" });
    }
    if (t.trendlineLbl) {
      xml.openNode("c:trendlineLbl");
      if (t.trendlineLbl.layout) {
        this._renderLayout(xml, t.trendlineLbl.layout);
      }
      if (t.trendlineLbl.rawTx) {
        xml.writeRaw(t.trendlineLbl.rawTx);
      } else if (t.trendlineLbl.text) {
        this._renderRichText(xml, t.trendlineLbl.text, "c:tx");
      }
      if (t.trendlineLbl.numFmt) {
        xml.leafNode("c:numFmt", {
          formatCode: t.trendlineLbl.numFmt.formatCode,
          ...(t.trendlineLbl.numFmt.sourceLinked !== undefined
            ? { sourceLinked: t.trendlineLbl.numFmt.sourceLinked ? "1" : "0" }
            : {})
        });
      }
      if (t.trendlineLbl.spPr) {
        this._renderSpPr(xml, t.trendlineLbl.spPr);
      }
      if (t.trendlineLbl.txPr) {
        this._renderTxPr(xml, t.trendlineLbl.txPr);
      }
      if (t.trendlineLbl.extLst) {
        xml.writeRaw(t.trendlineLbl.extLst);
      }
      xml.closeNode();
    }
    if (t.extLst) {
      xml.writeRaw(t.extLst);
    }
    xml.closeNode();
  }

  private _renderErrorBars(xml: XmlSink, eb: ErrorBars): void {
    xml.openNode("c:errBars");
    if (eb.errDir) {
      xml.leafNode("c:errDir", { val: eb.errDir });
    }
    xml.leafNode("c:errBarType", { val: eb.barDir });
    xml.leafNode("c:errValType", { val: eb.errValType });
    if (eb.noEndCap !== undefined) {
      xml.leafNode("c:noEndCap", { val: eb.noEndCap ? "1" : "0" });
    }
    if (eb.val !== undefined) {
      xml.leafNode("c:val", { val: String(eb.val) });
    }
    if (eb.plus) {
      this._renderValData(xml, eb.plus, "c:plus");
    }
    if (eb.minus) {
      this._renderValData(xml, eb.minus, "c:minus");
    }
    if (eb.spPr) {
      this._renderSpPr(xml, eb.spPr);
    }
    if (eb.extLst) {
      xml.writeRaw(eb.extLst);
    }
    xml.closeNode();
  }

  private _renderPictureOptions(xml: XmlSink, po: PictureOptions): void {
    xml.openNode("c:pictureOptions");
    if (po.applyToFront !== undefined) {
      xml.leafNode("c:applyToFront", { val: po.applyToFront ? "1" : "0" });
    }
    if (po.applyToSides !== undefined) {
      xml.leafNode("c:applyToSides", { val: po.applyToSides ? "1" : "0" });
    }
    if (po.applyToEnd !== undefined) {
      xml.leafNode("c:applyToEnd", { val: po.applyToEnd ? "1" : "0" });
    }
    if (po.pictureFormat) {
      xml.leafNode("c:pictureFormat", { val: po.pictureFormat });
    }
    if (po.pictureStackUnit !== undefined) {
      xml.leafNode("c:pictureStackUnit", { val: String(po.pictureStackUnit) });
    }
    xml.closeNode();
  }

  private _renderUpDownBars(xml: XmlSink, udb: UpDownBars): void {
    xml.openNode("c:upDownBars");
    if (udb.gapWidth !== undefined) {
      xml.leafNode("c:gapWidth", { val: String(udb.gapWidth) });
    }
    if (udb.upBars) {
      xml.openNode("c:upBars");
      this._renderSpPr(xml, udb.upBars);
      xml.closeNode();
    }
    if (udb.downBars) {
      xml.openNode("c:downBars");
      this._renderSpPr(xml, udb.downBars);
      xml.closeNode();
    }
    xml.closeNode();
  }

  // ---- Axes ----

  private _renderAxis(xml: XmlSink, ax: ChartAxis): void {
    const tagMap: Record<string, string> = {
      cat: "c:catAx",
      val: "c:valAx",
      date: "c:dateAx",
      ser: "c:serAx"
    };
    const tag = tagMap[ax.axisType];
    if (!tag) {
      return;
    }

    xml.openNode(tag);
    xml.leafNode("c:axId", { val: String(ax.axId) });

    // Scaling
    xml.openNode("c:scaling");
    if (ax.scaling?.orientation) {
      xml.leafNode("c:orientation", { val: ax.scaling.orientation });
    }
    if (ax.scaling?.max !== undefined) {
      xml.leafNode("c:max", { val: String(ax.scaling.max) });
    }
    if (ax.scaling?.min !== undefined) {
      xml.leafNode("c:min", { val: String(ax.scaling.min) });
    }
    if (ax.scaling?.logBase !== undefined) {
      xml.leafNode("c:logBase", { val: String(ax.scaling.logBase) });
    }
    xml.closeNode();

    if (ax.delete !== undefined) {
      xml.leafNode("c:delete", { val: ax.delete ? "1" : "0" });
    }
    xml.leafNode("c:axPos", { val: ax.axPos });
    if (ax.majorGridlines) {
      xml.openNode("c:majorGridlines");
      this._renderSpPr(xml, ax.majorGridlines);
      xml.closeNode();
    }
    if (ax.minorGridlines) {
      xml.openNode("c:minorGridlines");
      this._renderSpPr(xml, ax.minorGridlines);
      xml.closeNode();
    }
    if (ax.title) {
      this._renderTitle(xml, ax.title, "c:title");
    }
    if (ax.numFmt) {
      xml.leafNode("c:numFmt", {
        formatCode: ax.numFmt.formatCode,
        ...(ax.numFmt.sourceLinked !== undefined
          ? { sourceLinked: ax.numFmt.sourceLinked ? "1" : "0" }
          : {})
      });
    }
    if (ax.majorTickMark) {
      xml.leafNode("c:majorTickMark", { val: ax.majorTickMark });
    }
    if (ax.minorTickMark) {
      xml.leafNode("c:minorTickMark", { val: ax.minorTickMark });
    }
    if (ax.tickLblPos) {
      xml.leafNode("c:tickLblPos", { val: ax.tickLblPos });
    }
    if (ax.spPr) {
      this._renderSpPr(xml, ax.spPr);
    }
    if (ax.txPr) {
      this._renderTxPr(xml, ax.txPr);
    }
    xml.leafNode("c:crossAx", { val: String(ax.crossAx) });
    if (ax.crosses) {
      xml.leafNode("c:crosses", { val: ax.crosses });
    }
    if (ax.crossesAt !== undefined) {
      xml.leafNode("c:crossesAt", { val: String(ax.crossesAt) });
    }

    // Type-specific properties
    if (ax.axisType === "cat") {
      const ca = ax as CategoryAxis;
      if (ca.auto !== undefined) {
        xml.leafNode("c:auto", { val: ca.auto ? "1" : "0" });
      }
      if (ca.lblAlgn) {
        xml.leafNode("c:lblAlgn", { val: ca.lblAlgn });
      }
      if (ca.lblOffset !== undefined) {
        xml.leafNode("c:lblOffset", { val: String(ca.lblOffset) });
      }
      if (ca.tickLblSkip !== undefined) {
        xml.leafNode("c:tickLblSkip", { val: String(ca.tickLblSkip) });
      }
      if (ca.tickMarkSkip !== undefined) {
        xml.leafNode("c:tickMarkSkip", { val: String(ca.tickMarkSkip) });
      }
      if (ca.noMultiLvlLbl !== undefined) {
        xml.leafNode("c:noMultiLvlLbl", { val: ca.noMultiLvlLbl ? "1" : "0" });
      }
    } else if (ax.axisType === "val") {
      const va = ax as ValueAxis;
      if (va.crossBetween) {
        xml.leafNode("c:crossBetween", { val: va.crossBetween });
      }
      if (va.majorUnit !== undefined) {
        xml.leafNode("c:majorUnit", { val: String(va.majorUnit) });
      }
      if (va.minorUnit !== undefined) {
        xml.leafNode("c:minorUnit", { val: String(va.minorUnit) });
      }
      if (va.dispUnits) {
        xml.openNode("c:dispUnits");
        if (va.dispUnits.builtInUnit) {
          xml.leafNode("c:builtInUnit", { val: va.dispUnits.builtInUnit });
        }
        if (va.dispUnits.custUnit !== undefined) {
          xml.leafNode("c:custUnit", { val: String(va.dispUnits.custUnit) });
        }
        if (va.dispUnits.label) {
          this._renderTitle(xml, va.dispUnits.label, "c:dispUnitsLbl");
        }
        xml.closeNode();
      }
    } else if (ax.axisType === "date") {
      const da = ax as DateAxis;
      if (da.auto !== undefined) {
        xml.leafNode("c:auto", { val: da.auto ? "1" : "0" });
      }
      if (da.lblOffset !== undefined) {
        xml.leafNode("c:lblOffset", { val: String(da.lblOffset) });
      }
      if (da.baseTimeUnit) {
        xml.leafNode("c:baseTimeUnit", { val: da.baseTimeUnit });
      }
      if (da.majorUnit !== undefined) {
        xml.leafNode("c:majorUnit", { val: String(da.majorUnit) });
      }
      if (da.majorTimeUnit) {
        xml.leafNode("c:majorTimeUnit", { val: da.majorTimeUnit });
      }
      if (da.minorUnit !== undefined) {
        xml.leafNode("c:minorUnit", { val: String(da.minorUnit) });
      }
      if (da.minorTimeUnit) {
        xml.leafNode("c:minorTimeUnit", { val: da.minorTimeUnit });
      }
    } else if (ax.axisType === "ser") {
      const sa = ax as SeriesAxis;
      if (sa.tickLblSkip !== undefined) {
        xml.leafNode("c:tickLblSkip", { val: String(sa.tickLblSkip) });
      }
      if (sa.tickMarkSkip !== undefined) {
        xml.leafNode("c:tickMarkSkip", { val: String(sa.tickMarkSkip) });
      }
    }

    if (ax.extLst) {
      xml.writeRaw(ax.extLst);
    }

    xml.closeNode();
  }

  // ---- Title, Legend, Layout ----

  private _renderTitle(xml: XmlSink, title: ChartTitle, tag: string): void {
    xml.openNode(tag);
    if (title.rawTx) {
      xml.writeRaw(title.rawTx);
    } else if (title.text) {
      this._renderRichText(xml, title.text, "c:tx");
    } else if (title.strRef) {
      xml.openNode("c:tx");
      this._renderStrRef(xml, title.strRef);
      xml.closeNode();
    }
    if (title.layout) {
      this._renderLayout(xml, title.layout);
    }
    if (title.overlay !== undefined) {
      xml.leafNode("c:overlay", { val: title.overlay ? "1" : "0" });
    }
    if (title.spPr) {
      this._renderSpPr(xml, title.spPr);
    }
    if (title.txPr) {
      this._renderTxPr(xml, title.txPr);
    }
    if (title.extLst) {
      xml.writeRaw(title.extLst);
    }
    xml.closeNode();
  }

  private _renderLegend(xml: XmlSink, legend: ChartLegend): void {
    xml.openNode("c:legend");
    xml.leafNode("c:legendPos", { val: legend.legendPos });
    if (legend.legendEntries) {
      for (const e of legend.legendEntries) {
        xml.openNode("c:legendEntry");
        xml.leafNode("c:idx", { val: String(e.index) });
        if (e.delete !== undefined) {
          xml.leafNode("c:delete", { val: e.delete ? "1" : "0" });
        }
        if (e.txPr) {
          this._renderTxPr(xml, e.txPr);
        }
        if (e.extLst) {
          xml.writeRaw(e.extLst);
        }
        xml.closeNode();
      }
    }
    if (legend.layout) {
      this._renderLayout(xml, legend.layout);
    }
    if (legend.overlay !== undefined) {
      xml.leafNode("c:overlay", { val: legend.overlay ? "1" : "0" });
    }
    if (legend.spPr) {
      this._renderSpPr(xml, legend.spPr);
    }
    if (legend.txPr) {
      this._renderTxPr(xml, legend.txPr);
    }
    if (legend.extLst) {
      xml.writeRaw(legend.extLst);
    }
    xml.closeNode();
  }

  private _renderLayout(xml: XmlSink, layout?: ChartLayout): void {
    if (!layout) {
      xml.leafNode("c:layout");
      return;
    }
    xml.openNode("c:layout");
    if (layout.manualLayout) {
      const ml = layout.manualLayout;
      xml.openNode("c:manualLayout");
      if (ml.layoutTarget) {
        xml.leafNode("c:layoutTarget", { val: ml.layoutTarget });
      }
      if (ml.xMode) {
        xml.leafNode("c:xMode", { val: ml.xMode });
      }
      if (ml.yMode) {
        xml.leafNode("c:yMode", { val: ml.yMode });
      }
      if (ml.wMode) {
        xml.leafNode("c:wMode", { val: ml.wMode });
      }
      if (ml.hMode) {
        xml.leafNode("c:hMode", { val: ml.hMode });
      }
      if (ml.x !== undefined) {
        xml.leafNode("c:x", { val: String(ml.x) });
      }
      if (ml.y !== undefined) {
        xml.leafNode("c:y", { val: String(ml.y) });
      }
      if (ml.w !== undefined) {
        xml.leafNode("c:w", { val: String(ml.w) });
      }
      if (ml.h !== undefined) {
        xml.leafNode("c:h", { val: String(ml.h) });
      }
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderDataTable(xml: XmlSink, dt: DataTable): void {
    xml.openNode("c:dTable");
    if (dt.showHorzBorder !== undefined) {
      xml.leafNode("c:showHorzBorder", { val: dt.showHorzBorder ? "1" : "0" });
    }
    if (dt.showVertBorder !== undefined) {
      xml.leafNode("c:showVertBorder", { val: dt.showVertBorder ? "1" : "0" });
    }
    if (dt.showOutline !== undefined) {
      xml.leafNode("c:showOutline", { val: dt.showOutline ? "1" : "0" });
    }
    if (dt.showKeys !== undefined) {
      xml.leafNode("c:showKeys", { val: dt.showKeys ? "1" : "0" });
    }
    if (dt.spPr) {
      this._renderSpPr(xml, dt.spPr);
    }
    if (dt.txPr) {
      this._renderTxPr(xml, dt.txPr);
    }
    if (dt.extLst) {
      xml.writeRaw(dt.extLst);
    }
    xml.closeNode();
  }

  private _renderRichText(xml: XmlSink, rt: ChartRichText, tag: string): void {
    xml.openNode(tag);
    xml.openNode("c:rich");
    // Render body properties (a:bodyPr) — use structured if present, else default leaf.
    if (rt.bodyProperties) {
      this._renderBodyProperties(xml, rt.bodyProperties);
    } else {
      xml.leafNode("a:bodyPr");
    }
    xml.leafNode("a:lstStyle");
    let paragraphCount = 0;
    for (const p of rt.paragraphs) {
      xml.openNode("a:p");
      if (p.properties || p.runProperties) {
        this._renderParagraphProperties(xml, p.properties, p.runProperties);
      }
      for (const run of p.runs ?? []) {
        xml.openNode("a:r");
        if (run.properties) {
          this._renderRunProperties(xml, run.properties, "a:rPr");
        }
        xml.openNode("a:t");
        xml.writeText(run.text);
        xml.closeNode();
        xml.closeNode();
      }
      if (p.endParaRunProperties) {
        this._renderRunProperties(xml, p.endParaRunProperties, "a:endParaRPr");
      }
      xml.closeNode();
      paragraphCount++;
    }
    if (paragraphCount === 0) {
      xml.openNode("a:p");
      xml.leafNode("a:endParaRPr");
      xml.closeNode();
    }
    xml.closeNode();
    xml.closeNode();
  }

  private _renderBodyProperties(xml: XmlSink, bp: ChartBodyProperties): void {
    const attrs: Record<string, string> = {};
    if (bp.rotation !== undefined) {
      attrs.rot = String(bp.rotation);
    }
    if (bp.horizontalOverflow) {
      attrs.horzOverflow = bp.horizontalOverflow;
    }
    if (bp.vertical) {
      attrs.vert = bp.vertical;
    }
    if (bp.wrap) {
      attrs.wrap = bp.wrap;
    }
    if (bp.anchor) {
      attrs.anchor = bp.anchor;
    }
    xml.leafNode("a:bodyPr", Object.keys(attrs).length > 0 ? attrs : undefined);
  }

  private _renderParagraphProperties(
    xml: XmlSink,
    pPr: ChartParagraphProperties | undefined,
    legacyDefRPr: ChartTextProperties | undefined
  ): void {
    const attrs: Record<string, string> = {};
    if (pPr?.alignment) {
      attrs.algn = pPr.alignment;
    }
    if (pPr?.indent !== undefined) {
      attrs.indent = String(pPr.indent);
    }
    if (pPr?.marginLeft !== undefined) {
      attrs.marL = String(pPr.marginLeft);
    }
    if (pPr?.marginRight !== undefined) {
      attrs.marR = String(pPr.marginRight);
    }
    if (pPr?.level !== undefined) {
      attrs.lvl = String(pPr.level);
    }

    const defRPr = pPr?.defaultRunProperties ?? legacyDefRPr;
    const hasChildren = !!(
      pPr?.lineSpacing ||
      pPr?.spaceBefore ||
      pPr?.spaceAfter ||
      pPr?.bullet ||
      defRPr
    );

    if (!hasChildren) {
      xml.leafNode("a:pPr", Object.keys(attrs).length > 0 ? attrs : undefined);
      return;
    }

    xml.openNode("a:pPr", Object.keys(attrs).length > 0 ? attrs : undefined);
    if (pPr?.lineSpacing) {
      xml.openNode("a:lnSpc");
      this._renderSpacing(xml, pPr.lineSpacing);
      xml.closeNode();
    }
    if (pPr?.spaceBefore) {
      xml.openNode("a:spcBef");
      this._renderSpacing(xml, pPr.spaceBefore);
      xml.closeNode();
    }
    if (pPr?.spaceAfter) {
      xml.openNode("a:spcAft");
      this._renderSpacing(xml, pPr.spaceAfter);
      xml.closeNode();
    }
    if (pPr?.bullet) {
      const b = pPr.bullet;
      if (b.type === "none") {
        xml.leafNode("a:buNone");
      } else if (b.type === "char") {
        xml.leafNode("a:buChar", { char: b.character });
      } else if (b.type === "autoNum") {
        const bAttrs: Record<string, string> = { type: b.scheme };
        if (b.startAt !== undefined) {
          bAttrs.startAt = String(b.startAt);
        }
        xml.leafNode("a:buAutoNum", bAttrs);
      }
    }
    if (defRPr) {
      this._renderRunProperties(xml, defRPr, "a:defRPr");
    }
    xml.closeNode();
  }

  private _renderSpacing(xml: XmlSink, s: ChartLineSpacing): void {
    if (s.type === "percentage") {
      xml.leafNode("a:spcPct", { val: String(s.value) });
    } else {
      xml.leafNode("a:spcPts", { val: String(s.value) });
    }
  }

  private _renderRunProperties(xml: XmlSink, props: ChartTextProperties, tag: string): void {
    if (props._rawXml) {
      xml.writeRaw(props._rawXml);
      return;
    }
    const attrs: Record<string, string> = {};
    if (props.size !== undefined) {
      attrs.sz = String(props.size);
    }
    if (props.bold !== undefined) {
      attrs.b = props.bold ? "1" : "0";
    }
    if (props.italic !== undefined) {
      attrs.i = props.italic ? "1" : "0";
    }
    if (props.underline !== undefined) {
      // Accept boolean shorthand: true -> "sng", false -> "none"
      if (typeof props.underline === "boolean") {
        attrs.u = props.underline ? "sng" : "none";
      } else {
        attrs.u = props.underline;
      }
    }
    if (props.strike) {
      attrs.strike = props.strike;
    }
    if (props.rotation !== undefined) {
      attrs.rot = String(props.rotation);
    }
    if (props.baseline !== undefined) {
      attrs.baseline = String(props.baseline);
    }
    if (props.kern !== undefined) {
      attrs.kern = String(props.kern);
    }
    if (props.spacing !== undefined) {
      attrs.spc = String(props.spacing);
    }
    if (props.cap) {
      attrs.cap = props.cap;
    }
    if (props.lang) {
      attrs.lang = props.lang;
    }

    const hasChildren = !!(
      props.color ||
      props.fontFamily ||
      props.eastAsianFamily ||
      props.complexScriptFamily
    );
    if (!hasChildren) {
      xml.leafNode(tag, Object.keys(attrs).length > 0 ? attrs : undefined);
      return;
    }

    xml.openNode(tag, attrs);
    if (props.color) {
      xml.openNode("a:solidFill");
      this._renderColor(xml, props.color);
      xml.closeNode();
    }
    if (props.fontFamily) {
      xml.leafNode("a:latin", { typeface: props.fontFamily });
    }
    if (props.eastAsianFamily) {
      xml.leafNode("a:ea", { typeface: props.eastAsianFamily });
    }
    if (props.complexScriptFamily) {
      xml.leafNode("a:cs", { typeface: props.complexScriptFamily });
    }
    xml.closeNode();
  }

  // ---- Shape Properties (simplified + raw XML) ----

  private _renderSpPr(xml: XmlSink, spPr: ShapeProperties): void {
    // If we have raw XML from parsing, emit it directly for perfect round-trip
    if (spPr._rawXml) {
      xml.writeRaw(spPr._rawXml);
      return;
    }

    xml.openNode("c:spPr");
    if (spPr.fill) {
      if (spPr.fill.noFill) {
        xml.leafNode("a:noFill");
      } else if (spPr.fill.solid) {
        xml.openNode("a:solidFill");
        this._renderColor(xml, spPr.fill.solid);
        xml.closeNode();
      } else if (spPr.fill.gradient) {
        const g = spPr.fill.gradient;
        if (g.stops.length >= 2) {
          xml.openNode("a:gradFill");
          xml.openNode("a:gsLst");
          for (const stop of g.stops) {
            xml.openNode("a:gs", { pos: String(Math.round(stop.position * 1000)) });
            this._renderColor(xml, stop.color);
            xml.closeNode();
          }
          xml.closeNode();
          if (g.type === "linear" || g.type === undefined) {
            xml.leafNode("a:lin", {
              ang: String((g.angle ?? 0) * 60000),
              scaled: "1"
            });
          } else if (g.type === "circle" || g.type === "rect" || g.type === "shape") {
            xml.openNode("a:path", { path: g.type });
            xml.leafNode("a:fillToRect", { l: "50000", t: "50000", r: "50000", b: "50000" });
            xml.closeNode();
          }
          xml.closeNode();
        }
      } else if (spPr.fill.pattern) {
        const p = spPr.fill.pattern;
        xml.openNode("a:pattFill", { prst: p.preset });
        if (p.foreground) {
          xml.openNode("a:fgClr");
          this._renderColor(xml, p.foreground);
          xml.closeNode();
        }
        if (p.background) {
          xml.openNode("a:bgClr");
          this._renderColor(xml, p.background);
          xml.closeNode();
        }
        xml.closeNode();
      }
    }
    if (spPr.line) {
      const lnAttrs: Record<string, string> = {};
      if (spPr.line.width) {
        lnAttrs.w = String(spPr.line.width);
      }
      if (spPr.line.cap) {
        lnAttrs.cap = spPr.line.cap;
      }
      if (spPr.line.compound) {
        lnAttrs.cmpd = spPr.line.compound;
      }
      xml.openNode("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined);
      if (spPr.line.noFill) {
        xml.leafNode("a:noFill");
      } else if (spPr.line.color) {
        xml.openNode("a:solidFill");
        this._renderColor(xml, spPr.line.color);
        xml.closeNode();
      }
      if (spPr.line.dash) {
        xml.leafNode("a:prstDash", { val: spPr.line.dash });
      }
      if (spPr.line.join) {
        if (spPr.line.join === "round") {
          xml.leafNode("a:round");
        } else if (spPr.line.join === "bevel") {
          xml.leafNode("a:bevel");
        } else if (spPr.line.join === "miter") {
          xml.leafNode("a:miter");
        }
      }
      xml.closeNode();
    }
    if (spPr.effectList) {
      this._renderEffectList(xml, spPr.effectList);
    }
    if (spPr.scene3d) {
      this._renderScene3D(xml, spPr.scene3d);
    }
    if (spPr.sp3d) {
      this._renderSp3D(xml, spPr.sp3d);
    }
    xml.closeNode();
  }

  private _renderEffectList(xml: XmlSink, effects: EffectList): void {
    xml.openNode("a:effectLst");
    if (effects.blur) {
      const attrs: Record<string, string> = {};
      if (effects.blur.radius !== undefined) {
        attrs.rad = String(effects.blur.radius);
      }
      if (effects.blur.grow) {
        attrs.grow = "1";
      }
      xml.leafNode("a:blur", Object.keys(attrs).length > 0 ? attrs : undefined);
    }
    if (effects.outerShadow) {
      this._renderShadowElement(xml, "a:outerShdw", effects.outerShadow);
    }
    if (effects.innerShadow) {
      this._renderShadowElement(xml, "a:innerShdw", effects.innerShadow);
    }
    if (effects.presetShadow) {
      const ps = effects.presetShadow;
      const attrs: Record<string, string> = { prst: ps.preset };
      if (ps.distance !== undefined) {
        attrs.dist = String(ps.distance);
      }
      if (ps.direction !== undefined) {
        attrs.dir = String(ps.direction);
      }
      xml.openNode("a:prstShdw", attrs);
      if (ps.color) {
        this._renderColor(xml, ps.color);
      }
      xml.closeNode();
    }
    if (effects.glow) {
      xml.openNode("a:glow", { rad: String(effects.glow.radius) });
      this._renderColor(xml, effects.glow.color);
      xml.closeNode();
    }
    if (effects.softEdge) {
      xml.leafNode("a:softEdge", { rad: String(effects.softEdge.radius) });
    }
    if (effects.reflection) {
      const r = effects.reflection;
      const attrs: Record<string, string> = {};
      if (r.blurRadius !== undefined) {
        attrs.blurRad = String(r.blurRadius);
      }
      if (r.startOpacity !== undefined) {
        attrs.stA = String(r.startOpacity);
      }
      if (r.startPosition !== undefined) {
        attrs.stPos = String(r.startPosition);
      }
      if (r.endOpacity !== undefined) {
        attrs.endA = String(r.endOpacity);
      }
      if (r.endPosition !== undefined) {
        attrs.endPos = String(r.endPosition);
      }
      if (r.distance !== undefined) {
        attrs.dist = String(r.distance);
      }
      if (r.direction !== undefined) {
        attrs.dir = String(r.direction);
      }
      if (r.fadeDirection !== undefined) {
        attrs.fadeDir = String(r.fadeDirection);
      }
      if (r.scaleHorizontal !== undefined) {
        attrs.sx = String(r.scaleHorizontal);
      }
      if (r.scaleVertical !== undefined) {
        attrs.sy = String(r.scaleVertical);
      }
      if (r.skewHorizontal !== undefined) {
        attrs.kx = String(r.skewHorizontal);
      }
      if (r.skewVertical !== undefined) {
        attrs.ky = String(r.skewVertical);
      }
      if (r.alignment) {
        attrs.algn = r.alignment;
      }
      if (r.rotateWithShape) {
        attrs.rotWithShape = "1";
      }
      xml.leafNode("a:reflection", attrs);
    }
    xml.closeNode();
  }

  private _renderShadowElement(xml: XmlSink, tag: string, shadow: Shadow): void {
    const attrs: Record<string, string> = {};
    if (shadow.blurRadius !== undefined) {
      attrs.blurRad = String(shadow.blurRadius);
    }
    if (shadow.distance !== undefined) {
      attrs.dist = String(shadow.distance);
    }
    if (shadow.direction !== undefined) {
      attrs.dir = String(shadow.direction);
    }
    if (shadow.alignment) {
      attrs.algn = shadow.alignment;
    }
    if (shadow.rotateWithShape) {
      attrs.rotWithShape = "1";
    }
    if (shadow.scaleHorizontal !== undefined) {
      attrs.sx = String(shadow.scaleHorizontal);
    }
    if (shadow.scaleVertical !== undefined) {
      attrs.sy = String(shadow.scaleVertical);
    }
    if (shadow.skewHorizontal !== undefined) {
      attrs.kx = String(shadow.skewHorizontal);
    }
    if (shadow.skewVertical !== undefined) {
      attrs.ky = String(shadow.skewVertical);
    }
    xml.openNode(tag, attrs);
    this._renderColor(xml, shadow.color);
    xml.closeNode();
  }

  private _renderScene3D(xml: XmlSink, scene: Scene3D): void {
    xml.openNode("a:scene3d");
    if (scene.camera) {
      const cam = scene.camera;
      const attrs: Record<string, string> = { prst: cam.preset };
      if (cam.fov !== undefined) {
        attrs.fov = String(cam.fov);
      }
      if (cam.zoom !== undefined) {
        attrs.zoom = String(cam.zoom);
      }
      if (cam.rotation) {
        xml.openNode("a:camera", attrs);
        xml.leafNode("a:rot", {
          lat: String(cam.rotation.lat),
          lon: String(cam.rotation.lon),
          rev: String(cam.rotation.rev)
        });
        xml.closeNode();
      } else {
        xml.leafNode("a:camera", attrs);
      }
    }
    if (scene.lightRig) {
      const lr = scene.lightRig;
      const attrs: Record<string, string> = { rig: lr.rig, dir: lr.direction };
      if (lr.rotation) {
        xml.openNode("a:lightRig", attrs);
        xml.leafNode("a:rot", {
          lat: String(lr.rotation.lat),
          lon: String(lr.rotation.lon),
          rev: String(lr.rotation.rev)
        });
        xml.closeNode();
      } else {
        xml.leafNode("a:lightRig", attrs);
      }
    }
    xml.closeNode();
  }

  private _renderSp3D(xml: XmlSink, sp3d: ShapeProperties3D): void {
    const attrs: Record<string, string> = {};
    if (sp3d.z !== undefined) {
      attrs.z = String(sp3d.z);
    }
    if (sp3d.extrusionHeight !== undefined) {
      attrs.extrusionH = String(sp3d.extrusionHeight);
    }
    if (sp3d.contourWidth !== undefined) {
      attrs.contourW = String(sp3d.contourWidth);
    }
    if (sp3d.material) {
      attrs.prstMaterial = sp3d.material;
    }
    const hasChildren = !!(
      sp3d.bevelTop ||
      sp3d.bevelBottom ||
      sp3d.extrusionColor ||
      sp3d.contourColor
    );
    if (!hasChildren) {
      xml.leafNode("a:sp3d", attrs);
      return;
    }
    xml.openNode("a:sp3d", attrs);
    if (sp3d.bevelTop) {
      const b = sp3d.bevelTop;
      const bAttrs: Record<string, string> = {};
      if (b.width !== undefined) {
        bAttrs.w = String(b.width);
      }
      if (b.height !== undefined) {
        bAttrs.h = String(b.height);
      }
      if (b.preset) {
        bAttrs.prst = b.preset;
      }
      xml.leafNode("a:bevelT", bAttrs);
    }
    if (sp3d.bevelBottom) {
      const b = sp3d.bevelBottom;
      const bAttrs: Record<string, string> = {};
      if (b.width !== undefined) {
        bAttrs.w = String(b.width);
      }
      if (b.height !== undefined) {
        bAttrs.h = String(b.height);
      }
      if (b.preset) {
        bAttrs.prst = b.preset;
      }
      xml.leafNode("a:bevelB", bAttrs);
    }
    if (sp3d.extrusionColor) {
      xml.openNode("a:extrusionClr");
      this._renderColor(xml, sp3d.extrusionColor);
      xml.closeNode();
    }
    if (sp3d.contourColor) {
      xml.openNode("a:contourClr");
      this._renderColor(xml, sp3d.contourColor);
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderTxPr(xml: XmlSink, txPr: ChartTextProperties): void {
    if (txPr._rawXml) {
      xml.writeRaw(txPr._rawXml);
      return;
    }
    // Minimal rendering
    xml.openNode("c:txPr");
    xml.leafNode(
      "a:bodyPr",
      txPr.rotation !== undefined ? { rot: String(txPr.rotation) } : undefined
    );
    xml.leafNode("a:lstStyle");
    xml.openNode("a:p");
    xml.openNode("a:pPr");
    this._renderRunProperties(xml, txPr, "a:defRPr");
    xml.closeNode();
    xml.leafNode("a:endParaRPr");
    xml.closeNode();
    xml.closeNode();
  }

  private _renderColor(xml: XmlSink, color: ChartColor): void {
    if (color.srgb) {
      const attrs: Record<string, string> = { val: color.srgb };
      const hasChildren =
        color.alpha !== undefined ||
        color.tint !== undefined ||
        color.lumMod !== undefined ||
        color.lumOff !== undefined ||
        color.shade !== undefined ||
        color.satMod !== undefined;
      if (!hasChildren) {
        xml.leafNode("a:srgbClr", attrs);
      } else {
        xml.openNode("a:srgbClr", attrs);
        if (color.alpha !== undefined) {
          xml.leafNode("a:alpha", { val: String(color.alpha) });
        }
        if (color.tint !== undefined) {
          xml.leafNode("a:tint", { val: String(Math.round(color.tint * 100000)) });
        }
        if (color.lumMod !== undefined) {
          xml.leafNode("a:lumMod", { val: String(color.lumMod) });
        }
        if (color.lumOff !== undefined) {
          xml.leafNode("a:lumOff", { val: String(color.lumOff) });
        }
        if (color.shade !== undefined) {
          xml.leafNode("a:shade", { val: String(color.shade) });
        }
        if (color.satMod !== undefined) {
          xml.leafNode("a:satMod", { val: String(color.satMod) });
        }
        xml.closeNode();
      }
    } else if (color.theme !== undefined) {
      const themeNames = [
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink"
      ];
      const name = themeNames[color.theme] ?? "dk1";
      const hasChildren =
        color.alpha !== undefined ||
        color.tint !== undefined ||
        color.lumMod !== undefined ||
        color.lumOff !== undefined ||
        color.shade !== undefined ||
        color.satMod !== undefined;
      if (!hasChildren) {
        xml.leafNode("a:schemeClr", { val: name });
      } else {
        xml.openNode("a:schemeClr", { val: name });
        if (color.alpha !== undefined) {
          xml.leafNode("a:alpha", { val: String(color.alpha) });
        }
        let tintEmittedShade = false;
        if (color.tint !== undefined) {
          if (color.tint >= 0) {
            xml.leafNode("a:tint", { val: String(Math.round(color.tint * 100000)) });
          } else {
            xml.leafNode("a:shade", { val: String(Math.round((1 + color.tint) * 100000)) });
            tintEmittedShade = true;
          }
        }
        if (color.lumMod !== undefined) {
          xml.leafNode("a:lumMod", { val: String(color.lumMod) });
        }
        if (color.lumOff !== undefined) {
          xml.leafNode("a:lumOff", { val: String(color.lumOff) });
        }
        if (color.shade !== undefined && !tintEmittedShade) {
          xml.leafNode("a:shade", { val: String(color.shade) });
        }
        if (color.satMod !== undefined) {
          xml.leafNode("a:satMod", { val: String(color.satMod) });
        }
        xml.closeNode();
      }
    } else if (color.sysClr) {
      const hasChildren =
        color.alpha !== undefined ||
        color.tint !== undefined ||
        color.lumMod !== undefined ||
        color.lumOff !== undefined ||
        color.shade !== undefined ||
        color.satMod !== undefined;
      if (!hasChildren) {
        xml.leafNode("a:sysClr", { val: color.sysClr });
      } else {
        xml.openNode("a:sysClr", { val: color.sysClr });
        if (color.alpha !== undefined) {
          xml.leafNode("a:alpha", { val: String(color.alpha) });
        }
        if (color.tint !== undefined) {
          xml.leafNode("a:tint", { val: String(Math.round(color.tint * 100000)) });
        }
        if (color.lumMod !== undefined) {
          xml.leafNode("a:lumMod", { val: String(color.lumMod) });
        }
        if (color.lumOff !== undefined) {
          xml.leafNode("a:lumOff", { val: String(color.lumOff) });
        }
        if (color.shade !== undefined) {
          xml.leafNode("a:shade", { val: String(color.shade) });
        }
        if (color.satMod !== undefined) {
          xml.leafNode("a:satMod", { val: String(color.satMod) });
        }
        xml.closeNode();
      }
    } else if (color.prstClr) {
      const hasChildren =
        color.alpha !== undefined ||
        color.tint !== undefined ||
        color.lumMod !== undefined ||
        color.lumOff !== undefined ||
        color.shade !== undefined ||
        color.satMod !== undefined;
      if (!hasChildren) {
        xml.leafNode("a:prstClr", { val: color.prstClr });
      } else {
        xml.openNode("a:prstClr", { val: color.prstClr });
        if (color.alpha !== undefined) {
          xml.leafNode("a:alpha", { val: String(color.alpha) });
        }
        if (color.tint !== undefined) {
          xml.leafNode("a:tint", { val: String(Math.round(color.tint * 100000)) });
        }
        if (color.lumMod !== undefined) {
          xml.leafNode("a:lumMod", { val: String(color.lumMod) });
        }
        if (color.lumOff !== undefined) {
          xml.leafNode("a:lumOff", { val: String(color.lumOff) });
        }
        if (color.shade !== undefined) {
          xml.leafNode("a:shade", { val: String(color.shade) });
        }
        if (color.satMod !== undefined) {
          xml.leafNode("a:satMod", { val: String(color.satMod) });
        }
        xml.closeNode();
      }
    }
  }

  private _renderPrintSettings(xml: XmlSink, ps: PrintSettings): void {
    xml.openNode("c:printSettings");
    if (typeof ps.headerFooter === "string") {
      xml.writeRaw(ps.headerFooter);
    } else {
      xml.leafNode("c:headerFooter");
    }
    if (ps.pageMargins) {
      xml.leafNode("c:pageMargins", {
        b: String(ps.pageMargins.b),
        l: String(ps.pageMargins.l),
        r: String(ps.pageMargins.r),
        t: String(ps.pageMargins.t),
        header: String(ps.pageMargins.header),
        footer: String(ps.pageMargins.footer)
      });
    }
    if (ps.pageSetup) {
      const attrs: Record<string, string> = {};
      if (ps.pageSetup.orientation) {
        attrs.orientation = ps.pageSetup.orientation;
      }
      if (ps.pageSetup.paperSize !== undefined) {
        attrs.paperSize = String(ps.pageSetup.paperSize);
      }
      xml.leafNode("c:pageSetup", Object.keys(attrs).length > 0 ? attrs : undefined);
    } else {
      xml.leafNode("c:pageSetup");
    }
    xml.closeNode();
  }

  // ============================================================================
  // PARSE
  // ============================================================================

  parseOpen(node: any): boolean {
    // Raw XML capture mode (spPr, txPr)
    if (this.rawCapture) {
      if (node.isSelfClosing) {
        this.rawCapture.handleSelfClose?.(node);
      } else {
        this.rawCapture.handleOpen(node);
      }
      return true;
    }

    const name = node.name;
    const attrs = node.attributes || {};

    // Check if this is a raw XML element we want to capture
    if (
      name === "c:spPr" ||
      name === "c:txPr" ||
      name === "c:pivotSource" ||
      name === "c:clrMapOvr" ||
      name === "c:protection" ||
      (name === "c:headerFooter" && this.currentPrintSettings != null) ||
      name === "mc:AlternateContent" ||
      name === "c:extLst" ||
      (name === "c:tx" && this.currentTitle && !this.currentSeries) ||
      (name === "c:tx" && this.currentDataLabelEntry != null) ||
      (name === "c:tx" && this.currentTrendlineLbl != null) ||
      (name === "c:dLbl" && this.currentPivotFormat && !this.currentDataLabels)
    ) {
      if (node.isSelfClosing) {
        // Empty element — no raw XML needed
        return true;
      }
      this.rawCapture = new RawXmlCapture();
      this.rawCapture.start(node);
      this.rawXmlTarget = name;
      // Determine context for extLst routing
      if (name === "c:extLst") {
        if (this.currentDataLabelEntry) {
          this.rawCaptureContext = "dLblEntry";
        } else if (this.currentDataLabels) {
          this.rawCaptureContext = "dLbls";
        } else if (this.currentTrendlineLbl) {
          this.rawCaptureContext = "trendlineLbl";
        } else if (this.currentTrendline) {
          this.rawCaptureContext = "trendline";
        } else if (this.currentMarker) {
          this.rawCaptureContext = "marker";
        } else if (this.currentDataPoint) {
          this.rawCaptureContext = "dataPoint";
        } else if (this.currentErrorBars) {
          this.rawCaptureContext = "errorBars";
        } else if (this.currentSeries) {
          this.rawCaptureContext = "series";
        } else if (this.currentDataTable) {
          this.rawCaptureContext = "dataTable";
        } else if (this.currentLegendEntry) {
          this.rawCaptureContext = "legendEntry";
        } else if (this.currentLegend) {
          this.rawCaptureContext = "legend";
        } else if (this.currentTitle) {
          this.rawCaptureContext = "title";
        } else if (this.currentAxis) {
          this.rawCaptureContext = "axis";
        } else if (this.currentChartTypeGroup) {
          this.rawCaptureContext = "chartTypeGroup";
        } else if (this.insidePlotArea) {
          this.rawCaptureContext = "plotArea";
        } else if (this.chartData) {
          this.rawCaptureContext = "chart";
        } else {
          this.rawCaptureContext = "chartSpace";
        }
      }
      // c:tx inside title → capture raw XML for round-trip
      if (name === "c:tx" && this.currentTitle) {
        this.rawXmlTarget = "c:tx:title";
      }
      // c:tx inside dLbl → capture raw XML for round-trip
      if (name === "c:tx" && this.currentDataLabelEntry && !this.currentTitle) {
        this.rawXmlTarget = "c:tx:dLblEntry";
      }
      // c:tx inside trendlineLbl → capture raw XML for round-trip
      if (
        name === "c:tx" &&
        this.currentTrendlineLbl &&
        !this.currentTitle &&
        !this.currentDataLabelEntry
      ) {
        this.rawXmlTarget = "c:tx:trendlineLbl";
      }
      // c:dLbl inside pivotFmt (not wrapped in c:dLbls) → capture raw XML
      if (name === "c:dLbl" && this.currentPivotFormat && !this.currentDataLabels) {
        this.rawXmlTarget = "c:dLbl:pivotFmt";
      }
      return true;
    }

    switch (name) {
      case "c:chartSpace":
        this.chartModel = {
          chart: { plotArea: { chartTypes: [], axes: [] } }
        };
        // Preserve extra xmlns attributes for round-trip
        {
          const extraNs: Record<string, string> = {};
          for (const [k, v] of Object.entries(attrs)) {
            if (k.startsWith("xmlns:") && !["xmlns:c", "xmlns:a", "xmlns:r"].includes(k)) {
              extraNs[k] = v as string;
            }
          }
          if (Object.keys(extraNs).length > 0) {
            this.chartModel.extraNamespaces = extraNs;
          }
        }
        break;
      case "c:lang":
        this.chartModel.lang = attrs.val;
        break;
      case "c:roundedCorners":
        this.chartModel.roundedCorners = attrs.val === "1";
        break;
      case "c:date1904":
        this.chartModel.date1904 = attrs.val === "1";
        break;
      case "c:style":
        this.chartModel.style = parseInt(attrs.val, 10);
        break;
      case "c:chart":
        this.chartData = { plotArea: { chartTypes: [], axes: [] } };
        break;
      case "c:title":
        this.currentTitle = {};
        this.stateStack.push({ tag: "c:title", context: this.currentTitle });
        break;
      case "c:autoTitleDeleted":
        if (this.chartData) {
          this.chartData.autoTitleDeleted = attrs.val === "1";
        }
        break;
      case "c:view3D":
        this.currentView3D = {};
        break;
      case "c:rotX":
        if (this.currentView3D) {
          this.currentView3D.rotX = parseInt(attrs.val, 10);
        }
        break;
      case "c:rotY":
        if (this.currentView3D) {
          this.currentView3D.rotY = parseInt(attrs.val, 10);
        }
        break;
      case "c:depthPercent":
        if (this.currentView3D) {
          this.currentView3D.depthPercent = parseInt(attrs.val, 10);
        }
        break;
      case "c:rAngAx":
        if (this.currentView3D) {
          this.currentView3D.rAngAx = attrs.val === "1";
        }
        break;
      case "c:hPercent":
        if (this.currentView3D) {
          this.currentView3D.hPercent = parseInt(attrs.val, 10);
        }
        break;
      case "c:perspective":
        if (this.currentView3D) {
          this.currentView3D.perspective = parseInt(attrs.val, 10);
        }
        break;
      case "c:plotArea":
        this.plotArea = { chartTypes: [], axes: [] };
        this.insidePlotArea = true;
        break;
      case "c:plotVisOnly":
        if (this.chartData) {
          this.chartData.plotVisOnly = attrs.val === "1";
        }
        break;
      case "c:dispBlanksAs":
        if (this.chartData) {
          this.chartData.dispBlanksAs = attrs.val as any;
        }
        break;
      case "c:showDLblsOverMax":
        if (this.chartData) {
          this.chartData.showDLblsOverMax = attrs.val === "1";
        }
        break;
      case "c:legend":
        this.currentLegend = { legendPos: "b" };
        break;
      case "c:legendPos":
        if (this.currentLegend) {
          this.currentLegend.legendPos = attrs.val as any;
        }
        break;
      case "c:overlay":
        if (this.currentLegend) {
          this.currentLegend.overlay = attrs.val === "1";
        } else if (this.currentTitle) {
          this.currentTitle.overlay = attrs.val === "1";
        }
        break;
      case "c:legendEntry":
        this.currentLegendEntry = { index: 0 };
        break;

      // ---- Print settings ----
      case "c:printSettings":
        this.currentPrintSettings = {};
        break;
      case "c:pageMargins":
        if (this.currentPrintSettings) {
          this.currentPrintSettings.pageMargins = {
            b: parseFloat(attrs.b ?? "0.75"),
            l: parseFloat(attrs.l ?? "0.7"),
            r: parseFloat(attrs.r ?? "0.7"),
            t: parseFloat(attrs.t ?? "0.75"),
            header: parseFloat(attrs.header ?? "0.3"),
            footer: parseFloat(attrs.footer ?? "0.3")
          };
        }
        break;
      case "c:pageSetup":
        if (this.currentPrintSettings) {
          this.currentPrintSettings.pageSetup = {};
          if (attrs.orientation) {
            this.currentPrintSettings.pageSetup.orientation = attrs.orientation as any;
          }
          if (attrs.paperSize) {
            this.currentPrintSettings.pageSetup.paperSize = parseInt(attrs.paperSize, 10);
          }
        }
        break;
      case "c:externalData":
        this.chartModel.externalData = { id: attrs["r:id"] ?? "" };
        break;
      case "c:autoUpdate":
        if (this.chartModel.externalData) {
          this.chartModel.externalData.autoUpdate = attrs.val === "1";
        }
        break;

      // ---- Pivot formats (#4) ----
      case "c:pivotFmts":
        if (this.chartData) {
          this.chartData.pivotFormats = [];
        }
        break;
      case "c:pivotFmt":
        this.currentPivotFormat = { index: 0 };
        break;

      // ---- Floor / Side Wall / Back Wall (#5) ----
      case "c:floor":
      case "c:sideWall":
      case "c:backWall":
        this.currentFloorWall = name;
        break;

      // ---- Layout (#3) ----
      case "c:layout":
        this.currentLayout = {};
        break;
      case "c:manualLayout":
        if (this.currentLayout) {
          this.currentLayout.manualLayout = {};
          this.currentManualLayout = this.currentLayout.manualLayout;
        }
        break;
      case "c:layoutTarget":
        if (this.currentManualLayout) {
          this.currentManualLayout.layoutTarget = attrs.val as any;
        }
        break;
      case "c:xMode":
        if (this.currentManualLayout) {
          this.currentManualLayout.xMode = attrs.val as any;
        }
        break;
      case "c:yMode":
        if (this.currentManualLayout) {
          this.currentManualLayout.yMode = attrs.val as any;
        }
        break;
      case "c:wMode":
        if (this.currentManualLayout) {
          this.currentManualLayout.wMode = attrs.val as any;
        }
        break;
      case "c:hMode":
        if (this.currentManualLayout) {
          this.currentManualLayout.hMode = attrs.val as any;
        }
        break;
      case "c:x":
        if (this.currentManualLayout) {
          this.currentManualLayout.x = parseFloat(attrs.val);
        }
        break;
      case "c:y":
        if (this.currentManualLayout) {
          this.currentManualLayout.y = parseFloat(attrs.val);
        }
        break;
      case "c:w":
        if (this.currentManualLayout) {
          this.currentManualLayout.w = parseFloat(attrs.val);
        }
        break;
      case "c:h":
        if (this.currentManualLayout) {
          this.currentManualLayout.h = parseFloat(attrs.val);
        }
        break;

      // ---- Data Table (#8) ----
      case "c:dTable":
        this.currentDataTable = {};
        break;
      case "c:showHorzBorder":
        if (this.currentDataTable) {
          this.currentDataTable.showHorzBorder = attrs.val === "1";
        }
        break;
      case "c:showVertBorder":
        if (this.currentDataTable) {
          this.currentDataTable.showVertBorder = attrs.val === "1";
        }
        break;
      case "c:showOutline":
        if (this.currentDataTable) {
          this.currentDataTable.showOutline = attrs.val === "1";
        }
        break;
      case "c:showKeys":
        if (this.currentDataTable) {
          this.currentDataTable.showKeys = attrs.val === "1";
        }
        break;

      // ---- Display Units (#9) ----
      case "c:dispUnits":
        this.currentDisplayUnits = {};
        break;
      case "c:builtInUnit":
        if (this.currentDisplayUnits) {
          this.currentDisplayUnits.builtInUnit = attrs.val as any;
        }
        break;
      case "c:custUnit":
        if (this.currentDisplayUnits) {
          this.currentDisplayUnits.custUnit = parseFloat(attrs.val);
        }
        break;
      case "c:dispUnitsLbl":
        // dispUnitsLbl is a title-like element
        this.currentTitle = {};
        this.stateStack.push({ tag: "c:title", context: this.currentTitle });
        break;

      default:
        // Chart type groups
        if (CHART_TYPE_TAGS.has(name)) {
          this.currentChartTypeGroup = this._createChartTypeGroup(name);
          return true;
        }
        // Axes
        if (AXIS_TAGS.has(name)) {
          this.currentAxis = this._createAxis(name);
          return true;
        }
        // Series
        if (name === "c:ser") {
          this.currentSeries = { index: 0, order: 0 };
          return true;
        }
        // Series elements
        this._parseSeriesElement(name, attrs);
        // Axis elements
        this._parseAxisElement(name, attrs);
        break;
    }

    this.textBuf = "";
    return true;
  }

  parseText(text: string): void {
    if (this.rawCapture) {
      this.rawCapture.handleText(text);
      return;
    }
    this.textBuf += text;
  }

  parseClose(name: string): boolean {
    // Raw XML capture
    if (this.rawCapture) {
      if (!this.rawCapture.handleClose(name)) {
        const rawXml = this.rawCapture.xml;
        this.rawCapture = null;
        this._attachRawXml(this.rawXmlTarget!, rawXml);
        this.rawXmlTarget = null;
      }
      return true;
    }

    switch (name) {
      case "c:chartSpace":
        this.model = this.chartModel;
        return false; // done

      case "c:chart":
        this.chartModel.chart = this.chartData;
        this.chartData = null as any;
        break;

      case "c:view3D":
        if (this.chartData && this.currentView3D) {
          this.chartData.view3D = this.currentView3D;
          this.currentView3D = null;
        }
        break;

      case "c:plotArea":
        if (this.chartData) {
          this.chartData.plotArea = this.plotArea;
        }
        this.insidePlotArea = false;
        break;

      case "c:legend":
        if (this.chartData && this.currentLegend) {
          this.chartData.legend = this.currentLegend;
          this.currentLegend = null;
        }
        break;

      case "c:legendEntry":
        if (this.currentLegend && this.currentLegendEntry) {
          if (!this.currentLegend.legendEntries) {
            this.currentLegend.legendEntries = [];
          }
          this.currentLegend.legendEntries.push(this.currentLegendEntry);
          this.currentLegendEntry = null;
        }
        break;

      case "c:title":
        if (this.stateStack.length > 0) {
          this.stateStack.pop();
          // Attach title to the correct parent
          if (this.currentAxis) {
            this.currentAxis.title = this.currentTitle;
          } else if (this.chartData && !this.currentAxis && !this.currentChartTypeGroup) {
            this.chartData.title = this.currentTitle;
          }
          this.currentTitle = null;
        }
        break;

      case "c:printSettings":
        if (this.currentPrintSettings) {
          this.chartModel.printSettings = this.currentPrintSettings;
          this.currentPrintSettings = null;
        }
        break;

      // Series close
      case "c:ser":
        if (this.currentChartTypeGroup && this.currentSeries) {
          if (!this.currentChartTypeGroup.series) {
            this.currentChartTypeGroup.series = [];
          }
          this.currentChartTypeGroup.series.push(this.currentSeries);
          this.currentSeries = null;
        }
        break;

      // Data references — handle c:f (formula) and c:v (value)
      case "c:f":
        if (this.currentMultiLvlStrRef && !this.currentNumRef && !this.currentStrRef) {
          this.currentMultiLvlStrRef.formula = this.textBuf;
        } else if (this.currentNumRef) {
          this.currentNumRef.formula = this.textBuf;
        } else if (this.currentStrRef) {
          this.currentStrRef.formula = this.textBuf;
        }
        break;
      case "c:v":
        this._handleValueClose();
        break;
      case "c:numRef":
        this._attachNumRef();
        this.currentNumRef = null;
        break;
      case "c:strRef":
        this._attachStrRef();
        this.currentStrRef = null;
        break;
      case "c:numCache":
        if (this.currentNumRef && this.currentNumCache) {
          this.currentNumRef.cache = this.currentNumCache;
          this.currentNumCache = null;
        }
        break;
      case "c:strCache":
        if (this.currentStrRef && this.currentStrCache) {
          this.currentStrRef.cache = this.currentStrCache;
          this.currentStrCache = null;
        }
        break;

      // formatCode (#1)
      case "c:formatCode":
        if (this.currentNumCache) {
          this.currentNumCache.formatCode = this.textBuf;
        } else if (this.currentNumLit) {
          this.currentNumLit.formatCode = this.textBuf;
        }
        break;

      // Numeric/string literals (#12)
      case "c:numLit":
        this._attachNumLit();
        this.currentNumLit = null;
        break;
      case "c:strLit":
        this._attachStrLit();
        this.currentStrLit = null;
        break;

      // Multi-level string reference (#13)
      case "c:lvl":
        if (this.currentMultiLvlStrRef?.cache && this.currentLvl) {
          this.currentMultiLvlStrRef.cache.levels.push(this.currentLvl);
          this.currentLvl = null;
        }
        break;
      case "c:multiLvlStrCache":
        // cache already attached to ref
        break;
      case "c:multiLvlStrRef":
        this._attachMultiLvlStrRef();
        this.currentMultiLvlStrRef = null;
        break;

      // Data labels
      case "c:dLbls":
        if (this.currentPivotFormat && this.currentDataLabels) {
          this.currentPivotFormat.dataLabels = this.currentDataLabels;
        } else if (this.currentSeries && this.currentDataLabels) {
          this.currentSeries.dataLabels = this.currentDataLabels;
        } else if (this.currentChartTypeGroup && this.currentDataLabels) {
          this.currentChartTypeGroup.dataLabels = this.currentDataLabels;
        }
        this.currentDataLabels = null;
        break;
      case "c:dLbl":
        if (this.currentDataLabels && this.currentDataLabelEntry) {
          if (!this.currentDataLabels.entries) {
            this.currentDataLabels.entries = [];
          }
          this.currentDataLabels.entries.push(this.currentDataLabelEntry);
          this.currentDataLabelEntry = null;
        }
        break;

      // Marker close
      case "c:marker":
        if (this.currentMarker) {
          if (this.currentPivotFormat) {
            this.currentPivotFormat.marker = this.currentMarker;
          } else if (this.currentDataPoint) {
            this.currentDataPoint.marker = this.currentMarker;
          } else if (this.currentSeries) {
            this.currentSeries.marker = this.currentMarker;
          }
          this.currentMarker = null;
        }
        break;

      // Data point
      case "c:dPt":
        if (this.currentSeries && this.currentDataPoint) {
          if (!this.currentSeries.dataPoints) {
            this.currentSeries.dataPoints = [];
          }
          this.currentSeries.dataPoints.push(this.currentDataPoint);
          this.currentDataPoint = null;
        }
        break;

      // Trendline
      case "c:trendline":
        if (this.currentSeries && this.currentTrendline) {
          if (!this.currentSeries.trendlines) {
            this.currentSeries.trendlines = [];
          }
          this.currentSeries.trendlines.push(this.currentTrendline);
          this.currentTrendline = null;
        }
        break;

      // Trendline label (#14)
      case "c:trendlineLbl":
        if (this.currentTrendline && this.currentTrendlineLbl) {
          this.currentTrendline.trendlineLbl = this.currentTrendlineLbl;
          this.currentTrendlineLbl = null;
        }
        break;

      // Trendline name (#20)
      case "c:name":
        if (this.currentTrendline) {
          this.currentTrendline.name = this.textBuf;
        }
        break;

      // Error bars
      case "c:errBars":
        if (this.currentSeries && this.currentErrorBars) {
          // Scatter and bubble series support multiple error bars (x and y)
          const chartType = this.currentChartTypeGroup?.type;
          if (chartType === "scatter" || chartType === "bubble") {
            if (!this.currentSeries.errorBars) {
              this.currentSeries.errorBars = [];
            }
            (this.currentSeries.errorBars as ErrorBars[]).push(this.currentErrorBars);
          } else {
            this.currentSeries.errorBars = this.currentErrorBars;
          }
          this.currentErrorBars = null;
        }
        break;

      // Pivot format (#4)
      case "c:pivotFmt":
        if (this.chartData?.pivotFormats && this.currentPivotFormat) {
          this.chartData.pivotFormats.push(this.currentPivotFormat);
          this.currentPivotFormat = null;
        }
        break;

      // Floor / wall (#5)
      case "c:floor":
      case "c:sideWall":
      case "c:backWall":
        this.currentFloorWall = null;
        break;

      // Layout (#3)
      case "c:layout":
        if (this.currentTitle) {
          this.currentTitle.layout = this.currentLayout ?? undefined;
        } else if (this.currentLegend) {
          this.currentLegend.layout = this.currentLayout ?? undefined;
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.layout = this.currentLayout ?? undefined;
        } else if (this.currentTrendlineLbl) {
          this.currentTrendlineLbl.layout = this.currentLayout ?? undefined;
        } else if (this.insidePlotArea) {
          this.plotArea.layout = this.currentLayout ?? undefined;
        }
        this.currentLayout = null;
        this.currentManualLayout = null;
        break;
      case "c:manualLayout":
        this.currentManualLayout = null;
        break;

      // Data table (#8)
      case "c:dTable":
        if (this.plotArea && this.currentDataTable) {
          this.plotArea.dataTable = this.currentDataTable;
          this.currentDataTable = null;
        }
        break;

      // Display units (#9)
      case "c:dispUnits":
        if (this.currentAxis && this.currentDisplayUnits) {
          (this.currentAxis as any).dispUnits = this.currentDisplayUnits;
          this.currentDisplayUnits = null;
        }
        break;
      case "c:dispUnitsLbl":
        if (this.currentDisplayUnits && this.currentTitle) {
          this.currentDisplayUnits.label = this.currentTitle;
          this.currentTitle = null;
          if (
            this.stateStack.length > 0 &&
            this.stateStack[this.stateStack.length - 1].tag === "c:title"
          ) {
            this.stateStack.pop();
          }
        }
        break;

      // Line contexts (#6)
      case "c:hiLowLines":
        if (this.currentChartTypeGroup && this._pendingLineSpPr) {
          this.currentChartTypeGroup.hiLowLines = this._pendingLineSpPr;
        }
        this.currentLineContext = null;
        this._pendingLineSpPr = null;
        break;
      case "c:dropLines":
        if (this.currentChartTypeGroup && this._pendingLineSpPr) {
          this.currentChartTypeGroup.dropLines = this._pendingLineSpPr;
        }
        this.currentLineContext = null;
        this._pendingLineSpPr = null;
        break;
      case "c:serLines":
        if (this.currentChartTypeGroup && this._pendingLineSpPr) {
          this.currentChartTypeGroup.serLines = this._pendingLineSpPr;
        }
        this.currentLineContext = null;
        this._pendingLineSpPr = null;
        break;

      // Up/down bars (#7)
      case "c:upBars":
      case "c:downBars":
        if (
          this.stateStack.length > 0 &&
          (this.stateStack[this.stateStack.length - 1].tag === "c:upBars" ||
            this.stateStack[this.stateStack.length - 1].tag === "c:downBars")
        ) {
          this.stateStack.pop();
        }
        break;
      case "c:upDownBars":
        if (this.currentChartTypeGroup && this.currentUpDownBars) {
          this.currentChartTypeGroup.upDownBars = this.currentUpDownBars;
          this.currentUpDownBars = null;
        }
        break;

      // Band formats (#10)
      case "c:bandFmt":
        if (this.currentChartTypeGroup?.bandFormats) {
          const state =
            this.stateStack.length > 0 &&
            this.stateStack[this.stateStack.length - 1].tag === "c:bandFmt"
              ? this.stateStack.pop()
              : null;
          if (state) {
            this.currentChartTypeGroup.bandFormats.push(state.context);
          }
        }
        break;

      // Picture options (#15)
      case "c:pictureOptions":
        if (this.currentSeries && this.currentPictureOptions) {
          this.currentSeries.pictureOptions = this.currentPictureOptions;
        } else if (this.currentDataPoint && this.currentPictureOptions) {
          this.currentDataPoint.pictureOptions = this.currentPictureOptions;
        }
        this.currentPictureOptions = null;
        break;

      // Separator (#16)
      case "c:separator":
        if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.separator = this.textBuf;
        } else if (this.currentDataLabels) {
          this.currentDataLabels.separator = this.textBuf;
        }
        break;

      // Stack cleanup — data source wrappers (#23)
      case "c:tx":
      case "c:cat":
      case "c:xVal":
      case "c:yVal":
      case "c:bubbleSize":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === name
        ) {
          this.stateStack.pop();
        }
        break;
      case "c:val":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === "c:val"
        ) {
          this.stateStack.pop();
        }
        break;
      case "c:plus":
      case "c:minus":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === name
        ) {
          this.stateStack.pop();
        }
        break;
      case "c:pt":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === "c:pt"
        ) {
          this.stateStack.pop();
        }
        break;

      // Gridlines (#23)
      case "c:majorGridlines":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === "c:majorGridlines"
        ) {
          const state = this.stateStack.pop();
          if (this.currentAxis && state?.context?.spPr) {
            this.currentAxis.majorGridlines = state.context.spPr;
          }
        }
        break;
      case "c:minorGridlines":
        if (
          this.stateStack.length > 0 &&
          this.stateStack[this.stateStack.length - 1].tag === "c:minorGridlines"
        ) {
          const state = this.stateStack.pop();
          if (this.currentAxis && state?.context?.spPr) {
            this.currentAxis.minorGridlines = state.context.spPr;
          }
        }
        break;

      default:
        // Chart type group close
        if (CHART_TYPE_TAGS.has(name) && this.currentChartTypeGroup) {
          this.plotArea.chartTypes.push(this.currentChartTypeGroup);
          this.currentChartTypeGroup = null;
        }
        // Axis close
        if (AXIS_TAGS.has(name) && this.currentAxis) {
          this.plotArea.axes.push(this.currentAxis);
          this.currentAxis = null;
        }
        break;
    }

    this.textBuf = "";
    return true;
  }

  // ---- Parse helpers ----

  private _createChartTypeGroup(tag: string): any {
    const typeMap: Record<string, string> = {
      "c:barChart": "bar",
      "c:bar3DChart": "bar3D",
      "c:lineChart": "line",
      "c:line3DChart": "line3D",
      "c:pieChart": "pie",
      "c:pie3DChart": "pie3D",
      "c:doughnutChart": "doughnut",
      "c:areaChart": "area",
      "c:area3DChart": "area3D",
      "c:scatterChart": "scatter",
      "c:bubbleChart": "bubble",
      "c:radarChart": "radar",
      "c:stockChart": "stock",
      "c:surfaceChart": "surface",
      "c:surface3DChart": "surface3D",
      "c:ofPieChart": "ofPie"
    };
    return { type: typeMap[tag], series: [], axisIds: [] };
  }

  private _createAxis(tag: string): any {
    const typeMap: Record<string, string> = {
      "c:catAx": "cat",
      "c:valAx": "val",
      "c:dateAx": "date",
      "c:serAx": "ser"
    };
    return { axisType: typeMap[tag], axId: 0, axPos: "b", crossAx: 0 };
  }

  private _parseSeriesElement(name: string, attrs: any): void {
    if (
      !this.currentSeries &&
      !this.currentChartTypeGroup &&
      !this.currentAxis &&
      !this.currentPivotFormat
    ) {
      return;
    }

    switch (name) {
      // Series basic
      case "c:idx":
        if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.index = parseInt(attrs.val, 10);
        } else if (this.currentDataPoint) {
          this.currentDataPoint.index = parseInt(attrs.val, 10);
        } else if (this.currentLegendEntry) {
          this.currentLegendEntry.index = parseInt(attrs.val, 10);
        } else if (this.currentPivotFormat) {
          this.currentPivotFormat.index = parseInt(attrs.val, 10);
        } else {
          // Check for bandFmt context in stateStack
          const bandFmtState = this.stateStack.find(s => s.tag === "c:bandFmt");
          if (bandFmtState) {
            bandFmtState.context.index = parseInt(attrs.val, 10);
          } else if (this.currentSeries) {
            this.currentSeries.index = parseInt(attrs.val, 10);
          }
        }
        break;
      case "c:order":
        if (this.currentTrendline) {
          this.currentTrendline.order = parseInt(attrs.val, 10);
        } else if (this.currentSeries) {
          this.currentSeries.order = parseInt(attrs.val, 10);
        }
        break;

      // Chart type group attributes
      case "c:barDir":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.barDir = attrs.val;
        }
        break;
      case "c:grouping":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.grouping = attrs.val;
        }
        break;
      case "c:varyColors":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.varyColors = attrs.val === "1";
        }
        break;
      case "c:scatterStyle":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.scatterStyle = attrs.val;
        }
        break;
      case "c:radarStyle":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.radarStyle = attrs.val;
        }
        break;
      case "c:ofPieType":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.ofPieType = attrs.val;
        }
        break;
      case "c:gapWidth":
        if (this.currentUpDownBars) {
          this.currentUpDownBars.gapWidth = parseInt(attrs.val, 10);
        } else if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.gapWidth = parseInt(attrs.val, 10);
        }
        break;
      case "c:overlap":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.overlap = parseInt(attrs.val, 10);
        }
        break;
      case "c:firstSliceAng":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.firstSliceAng = parseInt(attrs.val, 10);
        }
        break;
      case "c:holeSize":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.holeSize = parseInt(attrs.val, 10);
        }
        break;
      case "c:bubbleScale":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.bubbleScale = parseInt(attrs.val, 10);
        }
        break;
      case "c:showNegBubbles":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.showNegBubbles = attrs.val === "1";
        }
        break;
      case "c:sizeRepresents":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.sizeRepresents = attrs.val;
        }
        break;
      case "c:wireframe":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.wireframe = attrs.val === "1";
        }
        break;
      case "c:splitType":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.splitType = attrs.val;
        }
        break;
      case "c:splitPos":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.splitPos = parseFloat(attrs.val);
        }
        break;
      case "c:secondPieSize":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.secondPieSize = parseInt(attrs.val, 10);
        }
        break;
      case "c:axId":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.axisIds.push(parseInt(attrs.val, 10));
        }
        break;

      // Series-specific
      case "c:invertIfNegative":
        if (this.currentSeries) {
          this.currentSeries.invertIfNegative = attrs.val === "1";
        }
        break;
      case "c:smooth":
        if (this.currentSeries) {
          this.currentSeries.smooth = attrs.val === "1";
        } else if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.smooth = attrs.val === "1";
        }
        break;
      case "c:explosion":
        if (this.currentDataPoint) {
          this.currentDataPoint.explosion = parseInt(attrs.val, 10);
        } else if (this.currentSeries) {
          this.currentSeries.explosion = parseInt(attrs.val, 10);
        }
        break;
      case "c:bubble3D":
        if (this.currentDataPoint) {
          this.currentDataPoint.bubble3D = attrs.val === "1";
        } else if (this.currentSeries) {
          this.currentSeries.bubble3D = attrs.val === "1";
        }
        break;
      case "c:shape":
        if (this.currentSeries) {
          this.currentSeries.shape = attrs.val;
        } else if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.shape = attrs.val;
        }
        break;

      // Data refs
      case "c:tx":
      case "c:cat":
      case "c:xVal":
      case "c:yVal":
      case "c:bubbleSize":
        this.stateStack.push({ tag: name, context: null });
        break;
      case "c:val":
        // c:val inside errBars is a leaf <c:val val="5"/> with val attribute
        if (this.currentErrorBars && attrs.val !== undefined) {
          this.currentErrorBars.val = parseFloat(attrs.val);
        } else {
          this.stateStack.push({ tag: name, context: null });
        }
        break;
      case "c:plus":
      case "c:minus":
        this.stateStack.push({ tag: name, context: null });
        break;
      case "c:numRef":
        this.currentNumRef = { formula: "", cache: { points: [] } };
        break;
      case "c:strRef":
        this.currentStrRef = { formula: "", cache: { points: [] } };
        break;
      case "c:numCache":
        this.currentNumCache = { points: [] };
        break;
      case "c:strCache":
        this.currentStrCache = { points: [] };
        break;
      case "c:formatCode":
        // Will be captured on text close
        break;
      case "c:ptCount":
        if (this.currentNumCache) {
          this.currentNumCache.pointCount = parseInt(attrs.val, 10);
        } else if (this.currentStrCache) {
          this.currentStrCache.pointCount = parseInt(attrs.val, 10);
        } else if (this.currentNumLit) {
          this.currentNumLit.pointCount = parseInt(attrs.val, 10);
        } else if (this.currentStrLit) {
          this.currentStrLit.pointCount = parseInt(attrs.val, 10);
        } else if (this.currentMultiLvlStrRef?.cache) {
          this.currentMultiLvlStrRef.cache.pointCount = parseInt(attrs.val, 10);
        }
        break;
      case "c:pt":
        this.stateStack.push({ tag: "c:pt", context: { index: parseInt(attrs.idx ?? "0", 10) } });
        break;

      // DataLabels
      case "c:dLbls":
        this.currentDataLabels = {};
        break;
      case "c:dLbl":
        this.currentDataLabelEntry = { index: 0 };
        break;
      case "c:showLegendKey":
        this._setBoolOnLabels("showLegendKey", attrs.val === "1");
        break;
      case "c:showVal":
        this._setBoolOnLabels("showVal", attrs.val === "1");
        break;
      case "c:showCatName":
        this._setBoolOnLabels("showCatName", attrs.val === "1");
        break;
      case "c:showSerName":
        this._setBoolOnLabels("showSerName", attrs.val === "1");
        break;
      case "c:showPercent":
        this._setBoolOnLabels("showPercent", attrs.val === "1");
        break;
      case "c:showBubbleSize":
        this._setBoolOnLabels("showBubbleSize", attrs.val === "1");
        break;
      case "c:showLeaderLines":
        if (this.currentDataLabels) {
          this.currentDataLabels.showLeaderLines = attrs.val === "1";
        }
        break;
      case "c:dLblPos":
        if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.position = attrs.val as any;
        } else if (this.currentDataLabels) {
          this.currentDataLabels.position = attrs.val as any;
        }
        break;
      case "c:numFmt":
        if (this.currentTrendlineLbl) {
          this.currentTrendlineLbl.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: attrs.sourceLinked === "1"
          };
        } else if (this.currentAxis) {
          this.currentAxis.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: attrs.sourceLinked === "1"
          };
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: attrs.sourceLinked === "1"
          };
        } else if (this.currentDataLabels) {
          this.currentDataLabels.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: attrs.sourceLinked === "1"
          };
        }
        break;

      // Marker
      case "c:marker":
        if (attrs.val !== undefined) {
          // Group-level boolean marker (line/stock charts)
          if (this.currentChartTypeGroup) {
            this.currentChartTypeGroup.marker = attrs.val === "1";
          }
        } else {
          // Container marker (series/pivotFmt/dataPoint level)
          this.currentMarker = {};
        }
        break;
      case "c:symbol":
        if (this.currentMarker) {
          this.currentMarker.symbol = attrs.val as any;
        }
        break;
      case "c:size":
        if (this.currentMarker) {
          this.currentMarker.size = parseInt(attrs.val, 10);
        }
        break;

      // Data point
      case "c:dPt":
        this.currentDataPoint = { index: 0 };
        break;

      // Trendline
      case "c:trendline":
        this.currentTrendline = { type: "linear" };
        break;
      case "c:trendlineType":
        if (this.currentTrendline) {
          this.currentTrendline.type = attrs.val as any;
        }
        break;
      case "c:period":
        if (this.currentTrendline) {
          this.currentTrendline.period = parseInt(attrs.val, 10);
        }
        break;
      case "c:forward":
        if (this.currentTrendline) {
          this.currentTrendline.forward = parseFloat(attrs.val);
        }
        break;
      case "c:backward":
        if (this.currentTrendline) {
          this.currentTrendline.backward = parseFloat(attrs.val);
        }
        break;
      case "c:intercept":
        if (this.currentTrendline) {
          this.currentTrendline.intercept = parseFloat(attrs.val);
        }
        break;
      case "c:dispRSqr":
        if (this.currentTrendline) {
          this.currentTrendline.displayRSqr = attrs.val === "1";
        }
        break;
      case "c:dispEq":
        if (this.currentTrendline) {
          this.currentTrendline.displayEq = attrs.val === "1";
        }
        break;

      // Error bars
      case "c:errBars":
        this.currentErrorBars = { barDir: "both", errValType: "fixedVal" };
        break;
      case "c:errBarType":
        if (this.currentErrorBars) {
          this.currentErrorBars.barDir = attrs.val as ErrorBarType;
        }
        break;
      case "c:errValType":
        if (this.currentErrorBars) {
          this.currentErrorBars.errValType = attrs.val as any;
        }
        break;
      case "c:noEndCap":
        if (this.currentErrorBars) {
          this.currentErrorBars.noEndCap = attrs.val === "1";
        }
        break;
      case "c:errDir":
        if (this.currentErrorBars && attrs.val) {
          this.currentErrorBars.errDir = attrs.val as "x" | "y";
        }
        break;

      // Trendline label (#14)
      case "c:trendlineLbl":
        this.currentTrendlineLbl = {};
        break;
      // Trendline name (#20)
      case "c:name":
        // Will be captured on text close
        break;

      // Numeric/string literals (#12)
      case "c:numLit":
        this.currentNumLit = { points: [] };
        break;
      case "c:strLit":
        this.currentStrLit = { points: [] };
        break;

      // Multi-level string reference (#13)
      case "c:multiLvlStrRef":
        this.currentMultiLvlStrRef = { formula: "" };
        break;
      case "c:multiLvlStrCache":
        if (this.currentMultiLvlStrRef) {
          this.currentMultiLvlStrRef.cache = { levels: [] };
        }
        break;
      case "c:lvl":
        this.currentLvl = { points: [] };
        break;

      // Line contexts: hiLowLines, dropLines, serLines (#6)
      case "c:hiLowLines":
      case "c:dropLines":
      case "c:serLines":
        this.currentLineContext = name;
        break;

      // Up/down bars (#7)
      case "c:upDownBars":
        this.currentUpDownBars = {};
        break;
      case "c:upBars":
        if (this.currentUpDownBars) {
          this.stateStack.push({ tag: "c:upBars", context: {} });
        }
        break;
      case "c:downBars":
        if (this.currentUpDownBars) {
          this.stateStack.push({ tag: "c:downBars", context: {} });
        }
        break;

      // Band formats (#10)
      case "c:bandFmts":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.bandFormats = [];
        }
        break;
      case "c:bandFmt":
        this.stateStack.push({ tag: "c:bandFmt", context: { index: 0 } });
        break;

      // Custom split (#11)
      case "c:custSplit":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.custSplit = [];
        }
        break;
      case "c:secondPiePt":
        if (this.currentChartTypeGroup?.custSplit) {
          this.currentChartTypeGroup.custSplit.push(parseInt(attrs.val, 10));
        }
        break;

      // Picture options (#15)
      case "c:pictureOptions":
        this.currentPictureOptions = {};
        break;
      case "c:applyToFront":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToFront = attrs.val === "1";
        }
        break;
      case "c:applyToSides":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToSides = attrs.val === "1";
        }
        break;
      case "c:applyToEnd":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToEnd = attrs.val === "1";
        }
        break;
      case "c:pictureFormat":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.pictureFormat = attrs.val as any;
        }
        break;
      case "c:pictureStackUnit":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.pictureStackUnit = parseFloat(attrs.val);
        }
        break;

      // Separator (#16)
      case "c:separator":
        // Text will be captured in parseClose
        break;

      // delete on legend entry
      case "c:delete":
        if (this.currentLegendEntry) {
          this.currentLegendEntry.delete = attrs.val === "1";
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.delete = attrs.val === "1";
        }
        break;
    }
  }

  private _parseAxisElement(name: string, attrs: any): void {
    if (!this.currentAxis) {
      return;
    }

    switch (name) {
      case "c:axId":
        // axId in axis context (vs chart type group)
        if (!this.currentChartTypeGroup) {
          this.currentAxis.axId = parseInt(attrs.val, 10);
        }
        break;
      case "c:scaling":
        this.currentAxis.scaling = {};
        break;
      case "c:orientation":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.orientation = attrs.val;
        }
        break;
      case "c:max":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.max = parseFloat(attrs.val);
        }
        break;
      case "c:min":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.min = parseFloat(attrs.val);
        }
        break;
      case "c:logBase":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.logBase = parseFloat(attrs.val);
        }
        break;
      case "c:delete":
        if (!this.currentLegendEntry && !this.currentDataLabelEntry) {
          this.currentAxis.delete = attrs.val === "1";
        }
        break;
      case "c:axPos":
        this.currentAxis.axPos = attrs.val;
        break;
      case "c:majorGridlines":
        // Will capture spPr inside
        this.stateStack.push({ tag: "c:majorGridlines", context: {} });
        break;
      case "c:minorGridlines":
        this.stateStack.push({ tag: "c:minorGridlines", context: {} });
        break;
      case "c:majorTickMark":
        this.currentAxis.majorTickMark = attrs.val;
        break;
      case "c:minorTickMark":
        this.currentAxis.minorTickMark = attrs.val;
        break;
      case "c:tickLblPos":
        this.currentAxis.tickLblPos = attrs.val;
        break;
      case "c:crossAx":
        this.currentAxis.crossAx = parseInt(attrs.val, 10);
        break;
      case "c:crosses":
        this.currentAxis.crosses = attrs.val;
        break;
      case "c:crossesAt":
        this.currentAxis.crossesAt = parseFloat(attrs.val);
        break;
      case "c:auto":
        this.currentAxis.auto = attrs.val === "1";
        break;
      case "c:lblAlgn":
        this.currentAxis.lblAlgn = attrs.val;
        break;
      case "c:lblOffset":
        this.currentAxis.lblOffset = parseInt(attrs.val, 10);
        break;
      case "c:tickLblSkip":
        this.currentAxis.tickLblSkip = parseInt(attrs.val, 10);
        break;
      case "c:tickMarkSkip":
        this.currentAxis.tickMarkSkip = parseInt(attrs.val, 10);
        break;
      case "c:noMultiLvlLbl":
        this.currentAxis.noMultiLvlLbl = attrs.val === "1";
        break;
      case "c:crossBetween":
        this.currentAxis.crossBetween = attrs.val;
        break;
      case "c:majorUnit":
        this.currentAxis.majorUnit = parseFloat(attrs.val);
        break;
      case "c:minorUnit":
        this.currentAxis.minorUnit = parseFloat(attrs.val);
        break;
      case "c:baseTimeUnit":
        this.currentAxis.baseTimeUnit = attrs.val;
        break;
      case "c:majorTimeUnit":
        this.currentAxis.majorTimeUnit = attrs.val;
        break;
      case "c:minorTimeUnit":
        this.currentAxis.minorTimeUnit = attrs.val;
        break;
    }
  }

  private _handleValueClose(): void {
    const text = this.textBuf;
    // c:v inside a c:pt
    const ptState = this.stateStack.find(s => s.tag === "c:pt");
    if (ptState) {
      const ptCtx = ptState.context;
      // Check numLit/strLit/lvl before numCache/strCache
      if (this.currentLvl) {
        this.currentLvl.points.push({ index: ptCtx.index, value: text });
      } else if (this.currentNumLit) {
        this.currentNumLit.points.push({
          index: ptCtx.index,
          value: text ? parseFloat(text) : null
        });
      } else if (this.currentStrLit) {
        this.currentStrLit.points.push({ index: ptCtx.index, value: text });
      } else if (this.currentNumCache) {
        this.currentNumCache.points.push({
          index: ptCtx.index,
          value: text ? parseFloat(text) : null
        });
      } else if (this.currentStrCache) {
        this.currentStrCache.points.push({ index: ptCtx.index, value: text });
      }
      return;
    }

    // c:v inside c:tx (series name literal)
    const txState = this.stateStack.find(s => s.tag === "c:tx");
    if (txState && this.currentSeries) {
      if (!this.currentSeries.tx) {
        this.currentSeries.tx = {};
      }
      this.currentSeries.tx.value = text;
    }
  }

  private _attachNumRef(): void {
    if (!this.currentNumRef) {
      return;
    }
    const parent = this._findDataParent();
    if (!parent) {
      return;
    }

    switch (parent.tag) {
      case "c:val":
      case "c:yVal":
      case "c:bubbleSize":
        if (this.currentSeries) {
          const key =
            parent.tag === "c:val" ? "val" : parent.tag === "c:yVal" ? "yVal" : "bubbleSize";
          this.currentSeries[key] = { numRef: this.currentNumRef };
        }
        break;
      case "c:cat":
      case "c:xVal":
        if (this.currentSeries) {
          const key = parent.tag === "c:cat" ? "cat" : "xVal";
          this.currentSeries[key] = { numRef: this.currentNumRef };
        }
        break;
      case "c:plus":
        if (this.currentErrorBars) {
          this.currentErrorBars.plus = { numRef: this.currentNumRef };
        }
        break;
      case "c:minus":
        if (this.currentErrorBars) {
          this.currentErrorBars.minus = { numRef: this.currentNumRef };
        }
        break;
    }
  }

  private _attachStrRef(): void {
    if (!this.currentStrRef) {
      return;
    }
    const parent = this._findDataParent();
    if (!parent) {
      return;
    }

    switch (parent.tag) {
      case "c:tx":
        if (this.currentSeries) {
          if (!this.currentSeries.tx) {
            this.currentSeries.tx = {};
          }
          this.currentSeries.tx.strRef = this.currentStrRef;
        }
        break;
      case "c:cat":
      case "c:xVal":
        if (this.currentSeries) {
          const key = parent.tag === "c:cat" ? "cat" : "xVal";
          this.currentSeries[key] = { strRef: this.currentStrRef };
        }
        break;
    }
  }

  private _attachNumLit(): void {
    if (!this.currentNumLit) {
      return;
    }
    const parent = this._findDataParent();
    if (!parent) {
      return;
    }
    switch (parent.tag) {
      case "c:val":
      case "c:yVal":
      case "c:bubbleSize":
        if (this.currentSeries) {
          const key =
            parent.tag === "c:val" ? "val" : parent.tag === "c:yVal" ? "yVal" : "bubbleSize";
          this.currentSeries[key] = { numLit: this.currentNumLit };
        }
        break;
      case "c:cat":
      case "c:xVal":
        if (this.currentSeries) {
          const key = parent.tag === "c:cat" ? "cat" : "xVal";
          this.currentSeries[key] = { numLit: this.currentNumLit };
        }
        break;
    }
  }

  private _attachStrLit(): void {
    if (!this.currentStrLit) {
      return;
    }
    const parent = this._findDataParent();
    if (!parent) {
      return;
    }
    switch (parent.tag) {
      case "c:cat":
      case "c:xVal":
        if (this.currentSeries) {
          const key = parent.tag === "c:cat" ? "cat" : "xVal";
          this.currentSeries[key] = { strLit: this.currentStrLit };
        }
        break;
    }
  }

  private _attachMultiLvlStrRef(): void {
    if (!this.currentMultiLvlStrRef) {
      return;
    }
    const parent = this._findDataParent();
    if (!parent) {
      return;
    }
    if (parent.tag === "c:cat" || parent.tag === "c:xVal") {
      if (this.currentSeries) {
        const key = parent.tag === "c:cat" ? "cat" : "xVal";
        this.currentSeries[key] = { multiLvlStrRef: this.currentMultiLvlStrRef };
      }
    }
  }

  private _findDataParent(): ParseState | null {
    for (let i = this.stateStack.length - 1; i >= 0; i--) {
      const s = this.stateStack[i];
      if (
        [
          "c:tx",
          "c:cat",
          "c:val",
          "c:xVal",
          "c:yVal",
          "c:bubbleSize",
          "c:plus",
          "c:minus"
        ].includes(s.tag)
      ) {
        return s;
      }
    }
    return null;
  }

  private _attachRawXml(target: string, rawXml: string): void {
    const spPrObj: ShapeProperties = { _rawXml: rawXml };
    const txPrObj: ChartTextProperties = { _rawXml: rawXml };

    if (target === "c:tx:title") {
      // Raw c:tx captured inside a title context (#2)
      if (this.currentTitle) {
        this.currentTitle.rawTx = rawXml;
      }
    } else if (target === "c:tx:dLblEntry") {
      // Raw c:tx captured inside a dLbl entry — store as rawTx for round-trip
      if (this.currentDataLabelEntry) {
        this.currentDataLabelEntry.rawTx = rawXml;
      }
    } else if (target === "c:tx:trendlineLbl") {
      // Raw c:tx captured inside a trendlineLbl — store as rawTx for round-trip
      if (this.currentTrendlineLbl) {
        this.currentTrendlineLbl.rawTx = rawXml;
      }
    } else if (target === "c:dLbl:pivotFmt") {
      // Raw c:dLbl captured inside a pivotFmt context (not wrapped in c:dLbls)
      if (this.currentPivotFormat) {
        this.currentPivotFormat.rawDLbl = rawXml;
      }
    } else if (target === "c:spPr") {
      // Check for line contexts first (#6)
      if (this.currentLineContext) {
        this._pendingLineSpPr = spPrObj;
      } else if (this.currentUpDownBars) {
        // spPr inside upBars/downBars (#7)
        const barState = this.stateStack.find(s => s.tag === "c:upBars" || s.tag === "c:downBars");
        if (barState?.tag === "c:upBars") {
          this.currentUpDownBars.upBars = spPrObj;
        } else if (barState?.tag === "c:downBars") {
          this.currentUpDownBars.downBars = spPrObj;
        }
      } else if (this.currentFloorWall) {
        // spPr inside floor/wall (#5)
        if (this.chartData) {
          if (this.currentFloorWall === "c:floor") {
            this.chartData.floor = spPrObj;
          } else if (this.currentFloorWall === "c:sideWall") {
            this.chartData.sideWall = spPrObj;
          } else if (this.currentFloorWall === "c:backWall") {
            this.chartData.backWall = spPrObj;
          }
        }
      } else {
        // Check for bandFmt context (#10)
        const bandFmtState = this.stateStack.find(s => s.tag === "c:bandFmt");
        // Check for gridlines context (#23)
        const gridlineState = this.stateStack.find(
          s => s.tag === "c:majorGridlines" || s.tag === "c:minorGridlines"
        );

        if (bandFmtState) {
          bandFmtState.context.spPr = spPrObj;
        } else if (this.currentDataPoint) {
          this.currentDataPoint.spPr = spPrObj;
        } else if (this.currentMarker) {
          this.currentMarker.spPr = spPrObj;
        } else if (this.currentTrendlineLbl) {
          this.currentTrendlineLbl.spPr = spPrObj;
        } else if (this.currentTrendline) {
          this.currentTrendline.spPr = spPrObj;
        } else if (this.currentErrorBars) {
          this.currentErrorBars.spPr = spPrObj;
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.spPr = spPrObj;
        } else if (this.currentDataLabels) {
          this.currentDataLabels.spPr = spPrObj;
        } else if (this.currentPivotFormat) {
          this.currentPivotFormat.spPr = spPrObj;
        } else if (this.currentSeries) {
          this.currentSeries.spPr = spPrObj;
        } else if (gridlineState) {
          gridlineState.context.spPr = spPrObj;
        } else if (this.currentAxis) {
          this.currentAxis.spPr = spPrObj;
        } else if (this.currentLegend) {
          this.currentLegend.spPr = spPrObj;
        } else if (this.currentTitle) {
          this.currentTitle.spPr = spPrObj;
        } else if (this.currentDataTable) {
          this.currentDataTable.spPr = spPrObj;
        } else if (this.insidePlotArea && this.plotArea) {
          this.plotArea.spPr = spPrObj;
        } else {
          this.chartModel.spPr = spPrObj;
        }
      }
    } else if (target === "c:txPr") {
      if (this.currentTitle) {
        this.currentTitle.txPr = txPrObj;
      } else if (this.currentTrendlineLbl) {
        this.currentTrendlineLbl.txPr = txPrObj;
      } else if (this.currentDataLabelEntry) {
        this.currentDataLabelEntry.txPr = txPrObj;
      } else if (this.currentDataLabels) {
        this.currentDataLabels.txPr = txPrObj;
      } else if (this.currentDataTable) {
        this.currentDataTable.txPr = txPrObj;
      } else if (this.currentAxis) {
        this.currentAxis.txPr = txPrObj;
      } else if (this.currentLegendEntry) {
        this.currentLegendEntry.txPr = txPrObj;
      } else if (this.currentLegend) {
        this.currentLegend.txPr = txPrObj;
      } else {
        this.chartModel.txPr = txPrObj;
      }
    } else if (target === "c:pivotSource") {
      this.chartModel.pivotSource = rawXml;
    } else if (target === "c:clrMapOvr") {
      this.chartModel.clrMapOvr = rawXml;
    } else if (target === "c:protection") {
      this.chartModel.protection = rawXml;
    } else if (target === "c:headerFooter") {
      if (this.currentPrintSettings) {
        this.currentPrintSettings.headerFooter = rawXml;
      }
    } else if (target === "mc:AlternateContent") {
      this.chartModel.alternateContentStyle = rawXml;
    } else if (target === "c:extLst") {
      switch (this.rawCaptureContext) {
        case "dLblEntry":
          if (this.currentDataLabelEntry) {
            this.currentDataLabelEntry.extLst = rawXml;
          }
          break;
        case "dLbls":
          if (this.currentDataLabels) {
            this.currentDataLabels.extLst = rawXml;
          }
          break;
        case "trendlineLbl":
          if (this.currentTrendlineLbl) {
            this.currentTrendlineLbl.extLst = rawXml;
          }
          break;
        case "trendline":
          if (this.currentTrendline) {
            this.currentTrendline.extLst = rawXml;
          }
          break;
        case "marker":
          if (this.currentMarker) {
            this.currentMarker.extLst = rawXml;
          }
          break;
        case "dataPoint":
          if (this.currentDataPoint) {
            this.currentDataPoint.extLst = rawXml;
          }
          break;
        case "errorBars":
          if (this.currentErrorBars) {
            this.currentErrorBars.extLst = rawXml;
          }
          break;
        case "series":
          if (this.currentSeries) {
            this.currentSeries.extLst = rawXml;
          }
          break;
        case "dataTable":
          if (this.currentDataTable) {
            this.currentDataTable.extLst = rawXml;
          }
          break;
        case "legendEntry":
          if (this.currentLegendEntry) {
            this.currentLegendEntry.extLst = rawXml;
          }
          break;
        case "legend":
          if (this.currentLegend) {
            this.currentLegend.extLst = rawXml;
          }
          break;
        case "title":
          if (this.currentTitle) {
            this.currentTitle.extLst = rawXml;
          }
          break;
        case "axis":
          if (this.currentAxis) {
            this.currentAxis.extLst = rawXml;
          }
          break;
        case "chartTypeGroup":
          if (this.currentChartTypeGroup) {
            this.currentChartTypeGroup.extLst = rawXml;
          }
          break;
        case "plotArea":
          if (this.plotArea) {
            this.plotArea.extLst = rawXml;
          }
          break;
        case "chart":
          if (this.chartData) {
            this.chartData.extLst = rawXml;
          }
          break;
        default:
          this.chartModel.extLst = rawXml;
          break;
      }
      this.rawCaptureContext = null;
    }
  }

  private _setBoolOnLabels(key: string, value: boolean): void {
    if (this.currentDataLabelEntry) {
      (this.currentDataLabelEntry as any)[key] = value;
    } else if (this.currentDataLabels) {
      (this.currentDataLabels as any)[key] = value;
    }
  }

  // ============================================================================
  // Constants
  // ============================================================================

  static CHART_SPACE_ATTRIBUTES = {
    "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  };
}

export { ChartSpaceXform };
