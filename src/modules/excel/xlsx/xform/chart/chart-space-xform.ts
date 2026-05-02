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

import { escapeXml, escapeXmlAttr, themeIndexToName } from "@excel/chart/chart-utils";
import { isRawXmlShape } from "@excel/chart/shape-properties";
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
  ShapeTransform,
  PresetGeometry,
  CustomGeometry,
  CustomGeometryCommand,
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
  PivotChartOptions,
  PivotFormat,
  DataLabelsRange,
  ChartBlipFill,
  UpDownBars,
  DisplayUnits,
  LegendEntry,
  PictureOptions
} from "@excel/chart/types";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import {
  parseXsdBoolean as sharedParseXsdBoolean,
  parseXsdInt as sharedParseXsdInt,
  parseXsdFloat as sharedParseXsdFloat
} from "@excel/xlsx/xform/xsd-values";
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
        // Use the attribute-specific escape so control characters
        // (`\t` / `\n` / `\r`) survive the XML attribute-value
        // normaliser, and both quote styles (`'` and `"`) are
        // escaped defensively.
        s += ` ${k}="${escapeXmlAttr(String(v))}"`;
      }
    }
    s += ">";
    return s;
  }

  private _selfCloseTag(node: any): string {
    let s = `<${node.name}`;
    if (node.attributes) {
      for (const [k, v] of Object.entries(node.attributes)) {
        s += ` ${k}="${escapeXmlAttr(String(v))}"`;
      }
    }
    return s + "/>";
  }
}

// `escapeXml` / `escapeXmlAttr` / `escapeXmlText` are imported from
// `@excel/chart/chart-utils`. They're the single authoritative entry
// points every chart writer uses for text content / attribute values —
// the xform used to carry its own local copies that (1) did not strip
// the C0 / C1 control characters XML 1.0 forbids, letting malformed
// upstream input produce an xlsx no consumer could reopen, and
// (2) drifted from the `chart-renderer` / `chart-ex-renderer` /
// `chart-sidecar` equivalents across releases. `escapeXmlAttr` adds
// `\t\n\r` → numeric-ref encoding so attribute round-trip preserves
// whitespace.

/**
 * Does the string require `xml:space="preserve"` to survive XML
 * whitespace normalisation? The XML default (`xml:space="default"`)
 * collapses leading/trailing whitespace and reduces internal
 * whitespace runs to single spaces; any tab or newline is equivalent
 * to a space. Therefore a run needs `preserve` iff it has:
 *   - leading or trailing whitespace,
 *   - ≥2 consecutive whitespace characters internally, or
 *   - any tab / newline / carriage return character.
 */
function needsXmlSpacePreserve(text: string): boolean {
  if (text === "") {
    return false;
  }
  if (/^\s|\s$/.test(text)) {
    return true;
  }
  if (/[\t\n\r]/.test(text)) {
    return true;
  }
  if (/\s{2,}/.test(text)) {
    return true;
  }
  return false;
}

/**
 * Parse an OOXML `xsd:boolean` attribute value. OOXML allows the four
 * canonical forms: `"0" | "1" | "false" | "true"`. Excel itself always
 * emits `"0"` or `"1"`, but non-Microsoft writers (LibreOffice, custom
 * pipelines, hand-authored XML) use the textual forms — the parser used
 * to silently treat them as `false`, which corrupted round-trip.
 *
 * Returns `true`/`false`/`undefined` (when the attribute is missing or
 * not a recognised boolean).
 *
 * For OOXML elements whose schema default is `true` when the attribute
 * is absent (CT_Boolean variants — e.g. `<c:auto/>`), pass the result
 * through `?? true` at the call site.
 */
/**
 * Parse an OOXML boolean attribute. Re-exported from the shared
 * `xsd-values` helper so every xform shares a single source of truth;
 * chartsheet-xform, worksheet-xform, etc. import from
 * `@excel/xlsx/xform/xsd-values` directly.
 *
 * For OOXML elements whose schema default is `true` when the attribute
 * is absent (CT_Boolean variants — e.g. `<c:auto/>`), pass the result
 * through `?? true` at the call site.
 */
function parseXsdBoolean(value: string | undefined): boolean | undefined {
  return sharedParseXsdBoolean(value);
}

/**
 * Parse an OOXML integer attribute safely. Returns `undefined` when the
 * input is missing, empty or non-numeric rather than producing `NaN`
 * (which then propagates through the model and surfaces as `"NaN"` in
 * the round-trip XML). Call sites that need a default should pipe
 * through `?? fallback`.
 */
function parseXsdInt(value: string | undefined): number | undefined {
  return sharedParseXsdInt(value);
}

/**
 * Parse an OOXML floating-point attribute safely. See
 * {@link parseXsdInt} for the rationale.
 */
function parseXsdFloat(value: string | undefined): number | undefined {
  return sharedParseXsdFloat(value);
}

/**
 * Narrow a parsed attribute value to a known enum tuple. Unknown values
 * are returned as the raw string so that `as any` casts in the parser
 * don't silently coerce garbage into the model — the caller can then
 * decide to drop, log, or surface the anomaly. Returns `undefined` when
 * the input is `undefined` so callers can distinguish "absent" from
 * "present but invalid".
 *
 * Type parameter trickery: when the input matches, the return type is
 * the narrowed union; when it doesn't, we still return a `string`
 * (widened) so the caller sees its typed field populated with a value
 * that may be an illegal enum member. The recommended pattern is to
 * forward this to `unknownElements` diagnostics rather than storing the
 * raw string on the model, but that requires each call site to decide;
 * this helper is the minimum hygienic improvement over a naked `as any`.
 */
function narrowEnumValue<T extends string>(
  value: string | undefined,
  allowed: readonly T[]
): T | undefined {
  if (value === undefined) {
    return undefined;
  }
  return (allowed as readonly string[]).includes(value) ? (value as T) : undefined;
}

/**
 * Map the public API's friendly tick-mark vocabulary (`inside` /
 * `outside`, matching {@link TickMark}) back to the OOXML
 * `ST_TickMark` tokens (`in` / `out`) on the way out. Kept aligned
 * with the inverse map used in {@link parseTickMarkFromOoxml} and
 * with the ChartEx renderer's copy of the same helper.
 */
function tickMarkToOoxml(value: "none" | "inside" | "outside" | "cross"): string {
  if (value === "inside") {
    return "in";
  }
  if (value === "outside") {
    return "out";
  }
  return value;
}

/**
 * Inverse of {@link tickMarkToOoxml} for parse side.
 */
function parseTickMarkFromOoxml(
  value: string | undefined
): "none" | "inside" | "outside" | "cross" | undefined {
  if (value === "in") {
    return "inside";
  }
  if (value === "out") {
    return "outside";
  }
  if (value === "none" || value === "cross") {
    return value;
  }
  return undefined;
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

// MS Office 2010 pivot chart options extension — ECMA-376 MS-XLSX §2.3.11.
// The URI is a literal GUID-ish identifier Excel looks for when routing the
// `<c:ext>` element back to the `c14:pivotOptions` parser. Confirmed by
// inspecting Excel-authored xlsx test fixtures
// (`src/modules/excel/__tests__/data/chart-pivot-sample.xlsx`) — Excel 2019
// emits this exact GUID.
const C14_PIVOT_OPTIONS_EXT_URI = "{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}";
const C14_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
// Office 2014 pivot chart options16 extension — sibling to c14:pivotOptions.
// URI confirmed from an Excel-authored pivot chart sample.
const C16_PIVOT_OPTIONS16_EXT_URI = "{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}";
const C16_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2014/chart";
// Excel 2013+ "Value From Cells" / dataLabelsRange extension. The URI is
// fixed by the MS schema and appears verbatim in every Excel-authored
// file that uses the feature (see `tmp/pivot-sample/xl/charts/chart1.xml`).
const C15_DATA_LABELS_RANGE_EXT_URI = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";
const C15_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2012/chart";

// `escapeXmlText` is a local alias for {@link escapeXml}. Kept so the
// extension builders below read clearly at the call site ("text node
// content" vs. "attribute value"); both paths need identical escaping
// semantics — `&`, `<`, `>`, `"`, `'` entity-encoded, C0/C1 control
// chars and lone surrogates stripped.
const escapeXmlText = escapeXml;

/**
 * Parse the contents of a `<c:ext uri="{781A…}">` element that wraps a
 * `<c14:pivotOptions>` child into a structured {@link PivotChartOptions}.
 *
 * Accepts the full `<c:ext …>…</c:ext>` string (including the outer
 * wrapper). Returns `undefined` when the wrapper does not contain a
 * `c14:pivotOptions` element (defensive — the caller already matched the
 * URI, so this should never happen for well-formed Excel output).
 *
 * Boolean children use OOXML `CT_BooleanFalse` semantics: the `val`
 * attribute is `"0"` / `"1"` / `"true"` / `"false"`; absence of the
 * attribute defaults to `true`. Missing child elements themselves remain
 * undefined in the structured model, not `false`, so the writer does not
 * have to re-emit noise for fields Excel didn't originally include.
 */
function parseC14PivotOptions(extXml: string): PivotChartOptions | undefined {
  const pivotMatch = /<c14:pivotOptions\b[^>]*>([\s\S]*?)<\/c14:pivotOptions>/.exec(extXml);
  if (!pivotMatch) {
    return undefined;
  }
  const body = pivotMatch[1];
  const readBool = (tag: string): boolean | undefined => {
    const re = new RegExp(`<c14:${tag}\\b([^/>]*)(?:/>|>\\s*</c14:${tag}>)`, "i");
    const m = re.exec(body);
    if (!m) {
      return undefined;
    }
    const valMatch = /\bval\s*=\s*"([^"]*)"/.exec(m[1]);
    // CT_BooleanFalse default is true when @val is omitted.
    if (!valMatch) {
      return true;
    }
    return valMatch[1] === "1" || valMatch[1].toLowerCase() === "true";
  };
  const options: PivotChartOptions = {};
  const f = readBool("dropZoneFilter");
  if (f !== undefined) {
    options.dropZoneFilter = f;
  }
  const c = readBool("dropZoneCategories");
  if (c !== undefined) {
    options.dropZoneCategories = c;
  }
  const d = readBool("dropZoneData");
  if (d !== undefined) {
    options.dropZoneData = d;
  }
  const s = readBool("dropZoneSeries");
  if (s !== undefined) {
    options.dropZoneSeries = s;
  }
  const v = readBool("dropZonesVisible");
  if (v !== undefined) {
    options.dropZonesVisible = v;
  }
  // `refreshOnOpen` is not a child of c14:pivotOptions — it lives on the
  // PivotCacheDefinition part. Leaving it undefined here lets the caller
  // merge it from the cache side if needed.
  return Object.keys(options).length > 0 ? options : undefined;
}

/**
 * Parse the Office 2014+ `c16:pivotOptions16` extension into a
 * structured {@link PivotChartOptions} slice. Currently captures the
 * `showExpandCollapseFieldButtons` flag only — the extension has room
 * for more fields but this is the only one Excel emits in our sample
 * corpus.
 */
function parseC16PivotOptions16(extXml: string): PivotChartOptions | undefined {
  const match = /<c16:pivotOptions16\b[^>]*>([\s\S]*?)<\/c16:pivotOptions16>/.exec(extXml);
  if (!match) {
    return undefined;
  }
  const body = match[1];
  const toggleMatch = /<c16:showExpandCollapseFieldButtons\b[^/]*\bval\s*=\s*"([^"]*)"/i.exec(body);
  if (!toggleMatch) {
    return undefined;
  }
  const v = toggleMatch[1];
  return { showExpandCollapseFieldButtons: v === "1" || v.toLowerCase() === "true" };
}

/**
 * Parse the contents of a `<c:ext uri="{CE6537A1…}">` wrapper around a
 * `<c15:datalabelsRange>` child into a structured {@link DataLabelsRange}.
 * Accepts the full wrapper string (including the outer `<c:ext>` tag).
 *
 * Shape expected (matches real Excel output, see
 * `tmp/pivot-sample/xl/charts/chart1.xml`):
 *
 * ```xml
 * <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="…">
 *   <c15:datalabelsRange>
 *     <c15:f>Sheet1!$A$1:$A$5</c15:f>
 *     <c15:dlblRangeCache>
 *       <c15:ptCount val="5"/>
 *       <c15:pt idx="0"><c15:v>Label 1</c15:v></c15:pt>
 *       …
 *     </c15:dlblRangeCache>
 *   </c15:datalabelsRange>
 * </c:ext>
 * ```
 */
function parseC15DataLabelsRange(extXml: string): DataLabelsRange | undefined {
  const rangeMatch = /<c15:datalabelsRange\b[^>]*>([\s\S]*?)<\/c15:datalabelsRange>/.exec(extXml);
  if (!rangeMatch) {
    return undefined;
  }
  const body = rangeMatch[1];
  const fMatch = /<c15:f\b[^>]*>([\s\S]*?)<\/c15:f>/.exec(body);
  if (!fMatch) {
    return undefined;
  }
  const result: DataLabelsRange = { formula: decodeXmlText(fMatch[1]) };

  const cacheMatch = /<c15:dlblRangeCache\b[^>]*>([\s\S]*?)<\/c15:dlblRangeCache>/.exec(body);
  if (cacheMatch) {
    const cacheBody = cacheMatch[1];
    const ptCountMatch = /<c15:ptCount\b[^>]*\bval\s*=\s*"(\d+)"/.exec(cacheBody);
    const pts: Array<{ index: number; value: string }> = [];
    const ptRe =
      /<c15:pt\b[^>]*\bidx\s*=\s*"(\d+)"[^>]*>\s*<c15:v>([\s\S]*?)<\/c15:v>\s*<\/c15:pt>/g;
    let m: RegExpExecArray | null;
    while ((m = ptRe.exec(cacheBody)) !== null) {
      pts.push({ index: parseInt(m[1], 10), value: decodeXmlText(m[2]) });
    }
    if (pts.length > 0 || ptCountMatch) {
      result.cache = {
        pointCount: ptCountMatch ? parseInt(ptCountMatch[1], 10) : undefined,
        points: pts
      };
    }
  }
  return result;
}

function decodeXmlText(value: string): string {
  // Also decode numeric character references (`&#nnn;` decimal and
  // `&#xHH;` hex). Required so that formulas containing non-ASCII
  // characters via numeric escape (e.g. German `é` written as
  // `&#233;`, bullet `&#x2022;`, most Excel-authored CJK) round-trip
  // back to their original code points; the previous implementation
  // only decoded the five named entities and silently left numeric
  // references as literal text (`"&#233;"`) in the formula string.
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hex) => {
      const cp = parseInt(hex, 16);
      return Number.isFinite(cp) && cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "";
    })
    .replace(/&#(\d+);/g, (_, dec) => {
      const cp = parseInt(dec, 10);
      return Number.isFinite(cp) && cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "";
    })
    .replace(/&amp;/g, "&");
}

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
    // Clear every bit of parse state. Previously only `stateStack`,
    // `rawCapture` and `textBuf` were cleared, which meant that if the
    // same xform instance was reused after an exception (or via
    // `BaseXform.parseStreamDirect`), leftover `currentXxx` fields from
    // the previous parse could leak into the next — most visibly
    // causing `currentDataLabelEntry` to carry over to the next
    // series's `<c:dLbls>` and mis-attribute `c:dLblPos` etc. to the
    // wrong chart.
    this.stateStack = [];
    this.rawCapture = null;
    this.rawCaptureContext = null;
    this.textBuf = "";
    this.currentChartTypeGroup = null;
    this.currentSeries = null;
    this.currentAxis = null;
    this.currentTitle = null;
    this.currentLegend = null;
    this.currentDataLabels = null;
    this.currentDataLabelEntry = null;
    this.currentMarker = null;
    this.currentDataPoint = null;
    this.currentTrendline = null;
    this.currentErrorBars = null;
    this.currentNumRef = null;
    this.currentStrRef = null;
    this.currentNumLit = null;
    this.currentStrLit = null;
    this.currentNumCache = null;
    this.currentStrCache = null;
    this.currentLayout = null;
    this.currentManualLayout = null;
    this.currentView3D = null;
    this.currentUpDownBars = null;
    this.currentLegendEntry = null;
    this.currentDataTable = null;
    this.currentDisplayUnits = null;
    this.currentPivotFormat = null;
    this.currentTrendlineLbl = null;
    this.currentPictureOptions = null;
    this.currentPrintSettings = null;
    this.currentMultiLvlStrRef = null;
    this.currentLvl = null;
    this.currentFloorWall = null;
    this.currentLineContext = null;
    this._pendingLineSpPr = null;
    this.insidePlotArea = false;
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
    // ECMA-376 §21.2.2.29 `CT_ChartSpace` child order:
    //   date1904, lang, roundedCorners, (AlternateContent | style),
    //   clrMapOvr, pivotSource, protection, chart,
    //   spPr, txPr, externalData, printSettings, userShapes, extLst.
    // `clrMapOvr` must appear BEFORE `chart` (the chart's colour
    // references may already point at the mapping); previously we
    // emitted it after `chart`, which strict OOXML validators reject.
    if (m.clrMapOvr) {
      xmlStream.writeRaw(m.clrMapOvr);
    }
    if (m.pivotSource) {
      xmlStream.writeRaw(m.pivotSource);
    }
    if (m.protection) {
      xmlStream.writeRaw(m.protection);
    }

    this._renderChart(xmlStream, m.chart);

    if (m.spPr) {
      this._renderSpPr(xmlStream, m.spPr);
    }
    if (m.txPr) {
      this._renderTxPr(xmlStream, m.txPr);
    }
    if (m.externalData && m.externalData.id) {
      // `<c:externalData>` requires a non-empty `r:id` referencing a
      // real relationship in the chart's `.rels` part. Emitting
      // `r:id=""` (which happened when the source XML had been parsed
      // from a chart whose `r:id` was missing or blank) produces a
      // workbook that Excel's strict-open rejects with the "repair"
      // dialog on reopen. Drop the element entirely when we have no
      // valid relationship id to point at.
      xmlStream.openNode("c:externalData", { "r:id": m.externalData.id });
      if (m.externalData.autoUpdate !== undefined) {
        xmlStream.leafNode("c:autoUpdate", { val: m.externalData.autoUpdate ? "1" : "0" });
      }
      xmlStream.closeNode();
    }
    if (m.printSettings) {
      this._renderPrintSettings(xmlStream, m.printSettings);
    }
    // `<c:userShapes r:id="…"/>` — optional reference to a separate
    // drawing part holding user-drawn annotations on the chart. The
    // part itself rides along via the chart rels; here we emit the
    // reference so Excel knows where to find it. Per schema, it is
    // the last structured child before `extLst`.
    if (m.userShapesRelId) {
      xmlStream.leafNode("c:userShapes", { "r:id": m.userShapesRelId });
    }
    if (m.extLst || m.pivotOptions) {
      this._renderChartSpaceExtLst(xmlStream, m);
    }

    xmlStream.closeNode();
  }

  /**
   * Render the chart-space-level `<c:extLst>` element, merging the raw XML
   * captured at parse time (`m.extLst`) with any structured pivot chart
   * metadata (`m.pivotOptions`).
   *
   * The two inputs live side by side for good reason:
   *   - `m.extLst` holds extensions we do not structurally understand
   *     (c15 filtered series, c16r2 markup, c16r5 axis-title rich text…).
   *     These are byte-preserved verbatim for round-trip safety.
   *   - `m.pivotOptions` is the structured view of the MS Office 2010+
   *     `c14:pivotOptions` extension. The parser extracts it out of the
   *     raw extLst precisely so editors can read/mutate it programmatically
   *     without string surgery.
   *
   * Merge strategy: if `m.extLst` is present, splice the new c14 ext
   * element before its closing `</c:extLst>`; otherwise create a fresh
   * `<c:extLst>` wrapper containing just the c14 ext.
   */
  private _renderChartSpaceExtLst(xml: XmlSink, m: ChartModel): void {
    // Combine every chartSpace-level extension we know how to emit. Each
    // piece is optional; we only emit a `<c:extLst>` wrapper when at
    // least one piece is present.
    const extPieces: string[] = [];
    // Emit the pivot extension chain whenever we have structured
    // `pivotOptions`, regardless of whether `pivotSource` is also
    // present. Historically the gate required both because
    // `pivotSource` is what links the chart to its PivotTable, but
    // programmatic builders may legitimately set `pivotOptions` first
    // (e.g. during incremental construction) and have the source
    // linked later — dropping the options silently on write lost the
    // author's work. If no options are set, both helpers return empty
    // strings and no extension is emitted.
    if (m.pivotOptions) {
      const c14 = this._buildC14PivotOptionsExt(m.pivotOptions);
      if (c14) {
        extPieces.push(c14);
      }
      const c16 = this._buildC16PivotOptions16Ext(m.pivotOptions);
      if (c16) {
        extPieces.push(c16);
      }
    }
    const extraExt = extPieces.join("");
    if (m.extLst) {
      if (!extraExt) {
        xml.writeRaw(m.extLst);
        return;
      }
      // Splice the new ext(s) before `</c:extLst>`. We look for the
      // closing tag; the raw XML was produced by our own parser which
      // always emits a closing tag on its own.
      const close = m.extLst.lastIndexOf("</c:extLst");
      if (close >= 0) {
        xml.writeRaw(m.extLst.slice(0, close) + extraExt + m.extLst.slice(close));
      } else {
        // Malformed raw extLst (shouldn't happen from our parser) — emit
        // both independently so we don't lose information.
        xml.writeRaw(m.extLst);
        xml.writeRaw(`<c:extLst>${extraExt}</c:extLst>`);
      }
    } else if (extraExt) {
      xml.writeRaw(`<c:extLst>${extraExt}</c:extLst>`);
    }
  }

  /**
   * Serialise the Office 2014+ `c16:pivotOptions16` extension
   * (`c:chartSpace/c:extLst/c:ext[uri={E28EC0CA-…}]`). Only the
   * `showExpandCollapseFieldButtons` toggle is modelled here — Excel
   * adds this block alongside {@link _buildC14PivotOptionsExt} to
   * control the newer expand/collapse field-button affordances.
   *
   * Returns `""` when nothing needs to be written so the caller can
   * omit the extension entirely.
   */
  private _buildC16PivotOptions16Ext(options: PivotChartOptions): string {
    if (options.showExpandCollapseFieldButtons === undefined) {
      return "";
    }
    return (
      `<c:ext uri="${C16_PIVOT_OPTIONS16_EXT_URI}" xmlns:c16="${C16_CHART_NAMESPACE}">` +
      `<c16:pivotOptions16>` +
      `<c16:showExpandCollapseFieldButtons val="${
        options.showExpandCollapseFieldButtons ? "1" : "0"
      }"/>` +
      `</c16:pivotOptions16>` +
      `</c:ext>`
    );
  }

  /**
   * Serialise a {@link PivotChartOptions} as the `c:ext` wrapper Excel
   * expects:
   *
   * ```xml
   * <c:ext uri="{…}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
   *   <c14:pivotOptions>
   *     <c14:dropZoneFilter val="0"/>
   *     …
   *   </c14:pivotOptions>
   * </c:ext>
   * ```
   *
   * The `uri` attribute is the well-known identifier for the c14 pivot
   * options extension (see MS-XLSX §2.3.11). Excel 2010+ uses this URI as
   * the key when deciding whether to honour the extension.
   */
  private _buildC14PivotOptionsExt(options: PivotChartOptions): string {
    const children: string[] = [];
    // MS-XLSX schema order: filter, categories, data, series, dropZonesVisible.
    if (options.dropZoneFilter !== undefined) {
      children.push(`<c14:dropZoneFilter val="${options.dropZoneFilter ? "1" : "0"}"/>`);
    }
    if (options.dropZoneCategories !== undefined) {
      children.push(`<c14:dropZoneCategories val="${options.dropZoneCategories ? "1" : "0"}"/>`);
    }
    if (options.dropZoneData !== undefined) {
      children.push(`<c14:dropZoneData val="${options.dropZoneData ? "1" : "0"}"/>`);
    }
    if (options.dropZoneSeries !== undefined) {
      children.push(`<c14:dropZoneSeries val="${options.dropZoneSeries ? "1" : "0"}"/>`);
    }
    if (options.dropZonesVisible !== undefined) {
      children.push(`<c14:dropZonesVisible val="${options.dropZonesVisible ? "1" : "0"}"/>`);
    }
    if (children.length === 0) {
      return "";
    }
    return (
      `<c:ext uri="${C14_PIVOT_OPTIONS_EXT_URI}" xmlns:c14="${C14_CHART_NAMESPACE}">` +
      `<c14:pivotOptions>${children.join("")}</c14:pivotOptions>` +
      `</c:ext>`
    );
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
        // CT_PivotFmt child order: `idx, spPr?, txPr?, marker?, dLbl?`.
        // Previously `txPr` was never written out; any pivotFmt-level
        // typography set on the model would vanish on round-trip.
        if (pf.spPr) {
          this._renderSpPr(xml, pf.spPr);
        }
        if (pf.txPr) {
          this._renderTxPr(xml, pf.txPr);
        }
        if (pf.marker) {
          this._renderMarker(xml, pf.marker);
        }
        // `c:pivotFmt` may carry a bare `c:dLbl` (no `c:dLbls` wrapper).
        // Prefer the structured representation when present so callers
        // can mutate label attributes programmatically; fall back to
        // the captured raw bytes for files parsed before the structured
        // slot existed.
        if (pf.dLbl) {
          this._renderDataLabelEntry(xml, pf.dLbl);
        } else if (pf.rawDLbl) {
          xml.writeRaw(pf.rawDLbl);
        } else if (pf.dataLabels) {
          this._renderDataLabels(xml, pf.dataLabels);
        }
        // CT_PivotFmt ends with an optional `extLst`. Preserved as raw
        // bytes so any unrecognised vendor extensions round-trip.
        if (pf.extLst) {
          xml.writeRaw(pf.extLst);
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
    // ECMA-376 §21.2.2.228 CT_View3D child order:
    //   rotX?, hPercent?, rotY?, depthPercent?, rAngAx?, perspective?, extLst?
    // `hPercent` sits between `rotX` and `rotY`; the previous
    // implementation placed it after `rAngAx`, which strict schema
    // validators rejected.
    if (v.rotX !== undefined) {
      xml.leafNode("c:rotX", { val: String(v.rotX) });
    }
    if (v.hPercent !== undefined) {
      xml.leafNode("c:hPercent", { val: String(v.hPercent) });
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
    if (v.perspective !== undefined) {
      xml.leafNode("c:perspective", { val: String(v.perspective) });
    }
    if (v.extLst) {
      xml.writeRaw(v.extLst);
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
        this._renderBarChart(xml, ctg);
        break;
      case "bar3D":
        this._renderBar3DChart(xml, ctg);
        break;
      case "line":
        this._renderLineChart(xml, ctg);
        break;
      case "line3D":
        this._renderLine3DChart(xml, ctg);
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
    // `CT_BarChart` (2D) schema sequence:
    //   barDir, grouping?, varyColors?, ser*, dLbls?, gapWidth?,
    //   overlap?, serLines*, axId{2}.
    //
    // `c:barDir` and `c:grouping` are required children, but the parser
    // may leave them `undefined` if the source XML had an invalid value
    // that `narrowEnumValue` rejected. Emitting `val=""` produces a
    // file Excel rejects outright. Fall back to the schema defaults
    // (`"col"` for barDir, `"clustered"` for grouping) — these match
    // Excel's behaviour when the attributes are absent.
    xml.leafNode("c:barDir", { val: g.barDir ?? "col" });
    xml.leafNode("c:grouping", { val: g.grouping ?? "clustered" });
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
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  /**
   * Render a 3-D bar chart per `CT_Bar3DChart`. The 3-D schema DROPS
   * `overlap` and `serLines` (both 2-D-only) and ADDS `gapDepth` and
   * `shape`. The child order differs from 2-D:
   *
   *   barDir, grouping?, varyColors?, ser*, dLbls?, gapWidth?,
   *   gapDepth?, shape?, axId{3}.
   *
   * Previously the builder shared `_renderBarChart` with bar2D, which
   * emitted `overlap` / `serLines` on bar3D (schema violation) and put
   * `shape` before `gapDepth` (sequence violation — Excel tolerates it
   * on read, strict validators reject).
   */
  private _renderBar3DChart(xml: XmlSink, g: BarChartGroup): void {
    xml.leafNode("c:barDir", { val: g.barDir ?? "col" });
    xml.leafNode("c:grouping", { val: g.grouping ?? "clustered" });
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
    if (g.gapDepth !== undefined) {
      xml.leafNode("c:gapDepth", { val: String(g.gapDepth) });
    }
    if (g.shape) {
      xml.leafNode("c:shape", { val: g.shape });
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderLineChart(xml: XmlSink, g: LineChartGroup): void {
    // `CT_LineChart` (2D) schema sequence:
    //   grouping, varyColors?, ser*, dLbls?, dropLines?, hiLowLines?,
    //   upDownBars?, marker?, smooth?, axId{2}.
    //
    // Guard required `c:grouping` with schema default `"standard"` when
    // parser narrowing rejected the source value; see `_renderBarChart`.
    xml.leafNode("c:grouping", { val: g.grouping ?? "standard" });
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

  /**
   * Render a 3-D line chart per `CT_Line3DChart`. Schema DROPS
   * `hiLowLines`, `upDownBars`, `marker`, `smooth` (all 2-D-only) and
   * ADDS `gapDepth`. Child order:
   *
   *   grouping, varyColors?, ser*, dLbls?, dropLines?, gapDepth?,
   *   axId{3}.
   *
   * Previously the builder shared `_renderLineChart` with line2D,
   * producing `<c:line3DChart>` with illegal children (most Excel-
   * authored tools would still open the file, but strict validators
   * reject it; more importantly, Excel's own read path silently
   * ignored the illegal children, so the line3D chart lost its marker
   * / smooth / up-down bars on round-trip).
   */
  private _renderLine3DChart(xml: XmlSink, g: LineChartGroup): void {
    xml.leafNode("c:grouping", { val: g.grouping ?? "standard" });
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
    if (g.gapDepth !== undefined) {
      xml.leafNode("c:gapDepth", { val: String(g.gapDepth) });
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
    // Guard required `c:grouping` — see `_renderBarChart`.
    xml.leafNode("c:grouping", { val: g.grouping ?? "standard" });
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
    // `c:gapDepth` exists only on `CT_Area3DChart`, not on plain
    // `CT_AreaChart`. The builder's validator rejects `gapDepth` for
    // 2-D area, but a direct model mutation (`group.gapDepth = 50`)
    // would previously sneak the value through and produce an
    // invalid `<c:areaChart><c:gapDepth/>` — every strict validator
    // rejects that. Gate on the group type so only `area3D` emits
    // the element.
    if (g.type === "area3D" && g.gapDepth !== undefined) {
      xml.leafNode("c:gapDepth", { val: String(g.gapDepth) });
    }
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
  }

  private _renderScatterChart(xml: XmlSink, g: ScatterChartGroup): void {
    // `c:scatterStyle` is required; fall back to `"marker"` (Excel's
    // default scatter style) when parser narrowing rejected the source.
    xml.leafNode("c:scatterStyle", { val: g.scatterStyle ?? "marker" });
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
    // `c:radarStyle` is required; default to `"standard"` per schema
    // when parser narrowing rejected the source value.
    xml.leafNode("c:radarStyle", { val: g.radarStyle ?? "standard" });
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
    // `CT_StockChart` child sequence per ECMA-376 §21.2.2.206:
    //   ser+ → dLbls? → dropLines? → hiLowLines? → upDownBars? →
    //   axId{2} → extLst?.
    //
    // The previous emission order was:
    //   varyColors, ser, dLbls, hiLowLines, upDownBars, dropLines, axId
    // which was invalid on two counts: `CT_StockChart` has no
    // `varyColors` attribute (fields are bound by the schema, not the
    // more permissive `CT_LineChart` superset), and `dropLines`
    // precedes `hiLowLines` in the sequence. Strict validators
    // (`ooxml-validate`, LibreOffice strict) rejected both.
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
    for (const id of g.axisIds) {
      xml.leafNode("c:axId", { val: String(id) });
    }
    // NOTE: `extLst` is intentionally omitted here — the caller
    // (`_renderChartTypeGroup`) emits `ctg.extLst` for every group
    // type unconditionally at the very end of the group element.
    // Writing it here too would produce a duplicate `<c:extLst>`
    // inside `<c:stockChart>`, which breaks strict validators.
  }

  private _renderSurfaceChart(xml: XmlSink, g: SurfaceChartGroup): void {
    if (g.wireframe !== undefined) {
      xml.leafNode("c:wireframe", { val: g.wireframe ? "1" : "0" });
    }
    for (const s of g.series) {
      this._renderSurfaceSeries(xml, s);
    }
    // `CT_SurfaceChart` has no `dLbls` child per schema; the builder
    // rejects `opts.dataLabels` for surface charts and the type no
    // longer carries the slot, so we simply skip emitting it here.
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
    // `c:ofPieType` is required; default to `"pie"` (Excel's default
    // bar-of-pie / pie-of-pie style) when parser narrowing rejected.
    xml.leafNode("c:ofPieType", { val: g.ofPieType ?? "pie" });
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
    // OOXML `CT_AreaSer` child sequence: SerShared (index, order, tx,
    // spPr) → pictureOptions? → dPt* → dLbls → trendline* → errBars
    // → cat → val → extLst. Previously `pictureOptions` was parsed
    // (see `_processSeries` at line 4617) but never written back, so
    // a round-trip of a texture-filled area chart lost the
    // pictureFormat (`stretch` / `stack` / …) and the author's scale
    // parameters — Excel fell back to `stretch` as the default.
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
    // ptCount: prefer the explicit value from the model (preserves the
    // "declared N points but M ≤ N have values" OOXML convention for
    // sparse caches); otherwise fall back to `points.length` so
    // hand-authored caches never emit a cache missing the required
    // `c:ptCount` element.
    const ptCount = cache.pointCount ?? cache.points.length;
    xml.leafNode("c:ptCount", { val: String(ptCount) });
    for (const pt of cache.points) {
      // `!= null` catches both `null` (sparse slot sentinel) AND
      // `undefined` (uninitialised points on hand-built caches).
      // Strict `!== null` would let `undefined` through and emit
      // `<c:v>undefined</c:v>` — Excel parses that as a string "undefined"
      // on a numeric cache, corrupting the series on reopen.
      if (pt.value != null) {
        const attrs: Record<string, string> = { idx: String(pt.index) };
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
    const ptCount = cache.pointCount ?? cache.points.length;
    xml.leafNode("c:ptCount", { val: String(ptCount) });
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
    // `CT_NumData` requires `ptCount` per ECMA-376 §21.2.2.163 (same
    // convention as `CT_NumRef/numCache`). Fall back to `points.length`
    // when the caller did not set an explicit `pointCount` so
    // programmatically-built literal lists always emit the required
    // element — the previous `if (lit.pointCount !== undefined)` gate
    // produced schema-invalid XML that both validators and Excel's
    // strict open path rejected.
    const ptCount = lit.pointCount ?? lit.points.length;
    xml.leafNode("c:ptCount", { val: String(ptCount) });
    for (const pt of lit.points) {
      // `!= null` rather than `!== null` so `undefined` also skips
      // emission — see `_renderNumCache` for the rationale.
      if (pt.value != null) {
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
    // `CT_StrData` mandates `ptCount`; mirror the `_renderNumLit`
    // fallback rather than silently omitting the element when
    // `lit.pointCount` is unset.
    const ptCount = lit.pointCount ?? lit.points.length;
    xml.leafNode("c:ptCount", { val: String(ptCount) });
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

  /**
   * Serialise a single {@link DataLabelEntry} as `<c:dLbl>…</c:dLbl>`.
   * Shared between the series-level `<c:dLbls><c:dLbl/>` loop and
   * pivot-format containers that emit a bare `<c:dLbl/>` without the
   * enclosing `<c:dLbls>` wrapper.
   */
  private _renderDataLabelEntry(xml: XmlSink, e: DataLabelEntry): void {
    xml.openNode("c:dLbl");
    xml.leafNode("c:idx", { val: String(e.index) });
    // ECMA-376 CT_DLbl is `idx, choice(delete | (layout, tx, numFmt, spPr,
    // txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName,
    // showPercent, showBubbleSize, separator), extLst)`. When `delete` is
    // set we must emit `<c:delete val="1"/>` alone (choice left branch);
    // all the display-option children belong to the other branch and
    // would make the XML schema-invalid — some Excel builds reject it.
    if (e.delete) {
      xml.leafNode("c:delete", { val: "1" });
      if (e.extLst) {
        xml.writeRaw(e.extLst);
      }
      xml.closeNode();
      return;
    }
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
    if (e.extLst) {
      xml.writeRaw(e.extLst);
    }
    xml.closeNode();
  }

  private _renderDataLabels(xml: XmlSink, dl: DataLabels): void {
    xml.openNode("c:dLbls");
    // ECMA-376 CT_DLbls: `(dLbl* & display-flags) | delete` — the
    // `delete` choice is mutually exclusive with the dLbl list and
    // every display-option child. Emit `delete` alone (+ extLst) when
    // set so strict validators (xsd, Excel strict mode) accept the
    // output. Previously we emitted `entries` before checking
    // `delete`, producing `<c:dLbl>…<c:delete val="1"/>` which violates
    // the choice.
    if (dl.delete) {
      xml.leafNode("c:delete", { val: "1" });
      if (dl.extLst) {
        xml.writeRaw(dl.extLst);
      }
      xml.closeNode();
      return;
    }
    if (dl.entries) {
      for (const e of dl.entries) {
        this._renderDataLabelEntry(xml, e);
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
    // `dataLabelsRange` (Excel 2013+ "Value From Cells") serialises as a
    // `c15:datalabelsRange` element inside `c:dLbls/c:extLst/c:ext`. When
    // both `dataLabelsRange` and `extLst` are set, splice the ext into
    // the existing rawXml so other extensions already captured there are
    // preserved. When only `dataLabelsRange` is set, emit a fresh wrapper.
    const dlRangeExt = dl.dataLabelsRange
      ? this._buildC15DataLabelsRangeExt(dl.dataLabelsRange)
      : "";
    if (dlRangeExt && dl.extLst) {
      const close = dl.extLst.lastIndexOf("</c:extLst");
      if (close >= 0) {
        xml.writeRaw(dl.extLst.slice(0, close) + dlRangeExt + dl.extLst.slice(close));
      } else {
        xml.writeRaw(dl.extLst);
        xml.writeRaw(`<c:extLst>${dlRangeExt}</c:extLst>`);
      }
    } else if (dl.extLst) {
      xml.writeRaw(dl.extLst);
    } else if (dlRangeExt) {
      xml.writeRaw(`<c:extLst>${dlRangeExt}</c:extLst>`);
    }
    xml.closeNode();
  }

  /**
   * Serialise a {@link DataLabelsRange} as the `c:ext` wrapper Excel
   * 2013+ expects:
   *
   * ```xml
   * <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"
   *        xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">
   *   <c15:datalabelsRange>
   *     <c15:f>Sheet1!$A$1:$A$5</c15:f>
   *     <c15:dlblRangeCache>
   *       <c15:ptCount val="5"/>
   *       <c15:pt idx="0"><c15:v>Label 1</c15:v></c15:pt>
   *       …
   *     </c15:dlblRangeCache>
   *   </c15:datalabelsRange>
   * </c:ext>
   * ```
   *
   * The URI and child-element names are exactly what Excel emits —
   * verified against a user-authored xlsx via
   * `tmp/pivot-sample/xl/charts/chart1.xml` (the extension URI also
   * appears as a placeholder inside `c:pivotFmt/c:dLbl`).
   */
  private _buildC15DataLabelsRangeExt(range: DataLabelsRange): string {
    const parts: string[] = [];
    parts.push(`<c:ext uri="${C15_DATA_LABELS_RANGE_EXT_URI}" xmlns:c15="${C15_CHART_NAMESPACE}">`);
    parts.push(`<c15:datalabelsRange>`);
    parts.push(`<c15:f>${escapeXmlText(range.formula)}</c15:f>`);
    // Emit the cache whenever the parser / cache-populator saw ANY
    // resolved cells — an all-blank range has `pointCount > 0` but
    // `points = []`, and without the cache element Excel loses the
    // array length and desynchronises labels from the underlying data
    // points. Previously gated on `points.length > 0`, which dropped
    // the `ptCount` for the blank case.
    if (range.cache && (range.cache.pointCount !== undefined || range.cache.points.length > 0)) {
      parts.push(`<c15:dlblRangeCache>`);
      if (range.cache.pointCount !== undefined) {
        parts.push(`<c15:ptCount val="${range.cache.pointCount}"/>`);
      }
      for (const pt of range.cache.points) {
        parts.push(`<c15:pt idx="${pt.index}"><c15:v>${escapeXmlText(pt.value)}</c15:v></c15:pt>`);
      }
      parts.push(`</c15:dlblRangeCache>`);
    }
    parts.push(`</c15:datalabelsRange>`);
    parts.push(`</c:ext>`);
    return parts.join("");
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
    // ECMA-376 CT_ErrBars child order:
    //   errDir?, errBarType, errValType, noEndCap?, plus?, minus?, val?, spPr?, extLst?
    // `c:val` must come after `plus`/`minus` — we previously emitted
    // `val` before `plus`/`minus` which produced schema-invalid XML
    // that strict OOXML validators (and some Excel versions in strict
    // mode) reject.
    if (eb.errDir) {
      xml.leafNode("c:errDir", { val: eb.errDir });
    }
    xml.leafNode("c:errBarType", { val: eb.barDir });
    xml.leafNode("c:errValType", { val: eb.errValType });
    if (eb.noEndCap !== undefined) {
      xml.leafNode("c:noEndCap", { val: eb.noEndCap ? "1" : "0" });
    }
    if (eb.plus) {
      this._renderValData(xml, eb.plus, "c:plus");
    }
    if (eb.minus) {
      this._renderValData(xml, eb.minus, "c:minus");
    }
    if (eb.val !== undefined) {
      xml.leafNode("c:val", { val: String(eb.val) });
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
    if (udb.extLst) {
      xml.writeRaw(udb.extLst);
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

    // Scaling — per ECMA-376 `CT_Scaling` sequence: `logBase?`,
    // `orientation?`, `min?`, `max?`, `extLst?`. The previous order
    // (`orientation, max, min, logBase`) is a schema violation and
    // strict validators reject the document. Excel tolerates the old
    // order on read, but LibreOffice strict mode and `ooxml-validate`
    // do not; swap to the spec-correct order.
    xml.openNode("c:scaling");
    if (ax.scaling?.logBase !== undefined) {
      xml.leafNode("c:logBase", { val: String(ax.scaling.logBase) });
    }
    if (ax.scaling?.orientation) {
      xml.leafNode("c:orientation", { val: ax.scaling.orientation });
    }
    if (ax.scaling?.min !== undefined) {
      xml.leafNode("c:min", { val: String(ax.scaling.min) });
    }
    if (ax.scaling?.max !== undefined) {
      xml.leafNode("c:max", { val: String(ax.scaling.max) });
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
      xml.leafNode("c:majorTickMark", { val: tickMarkToOoxml(ax.majorTickMark) });
    }
    if (ax.minorTickMark) {
      xml.leafNode("c:minorTickMark", { val: tickMarkToOoxml(ax.minorTickMark) });
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
    // `c:crossAx` is a required child of every axis per ECMA-376
    // `CT_ValAx` / `CT_CatAx`. Guard against `undefined` leaking from
    // a mutated model — emitting `val="undefined"` here produces a
    // file Excel rejects. Skip the element when the field is missing;
    // the xform's axis-id allocator should have filled it in at parse
    // time, so an absent crossAx at write time signals a bug upstream
    // that's better surfaced as a validation error from Excel than
    // silently corrupted XML.
    if (typeof ax.crossAx === "number" && Number.isFinite(ax.crossAx)) {
      xml.leafNode("c:crossAx", { val: String(ax.crossAx) });
    }
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
        // ECMA-376 CT_DispUnits: `(builtInUnit | custUnit)?` — the two
        // are a choice, so emitting both produces schema-invalid XML.
        // `builtInUnit` wins when the model (unusually) carries both.
        if (va.dispUnits.builtInUnit) {
          xml.leafNode("c:builtInUnit", { val: va.dispUnits.builtInUnit });
        } else if (va.dispUnits.custUnit !== undefined) {
          xml.leafNode("c:custUnit", { val: String(va.dispUnits.custUnit) });
        }
        if (va.dispUnits.label) {
          this._renderTitle(xml, va.dispUnits.label, "c:dispUnitsLbl");
        }
        if (va.dispUnits.extLst) {
          xml.writeRaw(va.dispUnits.extLst);
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
    // `c:legendPos` is optional per CT_Legend. Skip when the model
    // didn't set it rather than emit a literal `val="undefined"` which
    // some validators / older Excel versions reject.
    if (legend.legendPos) {
      xml.leafNode("c:legendPos", { val: legend.legendPos });
    }
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
      // Previously this emitted an empty `<c:layout/>` sentinel. That
      // produced a byte-diff on every round-trip of charts whose source
      // XML did not include a `<c:layout>` element (which is the norm:
      // Excel omits it when no manual layout is set). Callers that
      // genuinely need the placeholder now emit `<c:layout/>` themselves.
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
        // Emit `xml:space="preserve"` when the run text contains
        // whitespace that XML normalisation would strip: leading /
        // trailing whitespace, runs of ≥2 spaces, or any tab/newline
        // in the middle. The ChartEx renderer has the equivalent
        // guard; previously the classic `<a:t>` writer silently
        // dropped padding whitespace on round-trip of axis titles and
        // category labels that intentionally carry alignment spaces.
        const text = run.text ?? "";
        if (needsXmlSpacePreserve(text)) {
          xml.openNode("a:t", { "xml:space": "preserve" });
        } else {
          xml.openNode("a:t");
        }
        xml.writeText(text);
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
      // Rewrite the outer element name to match the target tag before
      // emitting the raw bytes. The parser captures run properties
      // under whatever element it found (`<a:defRPr>` on paragraph
      // defaults, `<a:rPr>` inside `<a:r>`, `<a:endParaRPr>` after the
      // final run); the caller passes a target tag that may differ
      // from the captured one (e.g. restyling a paragraph default as a
      // per-run property). Previously the function emitted the bytes
      // verbatim — producing schema-invalid output when the target tag
      // was `a:endParaRPr` but the bytes were `<a:defRPr>…</a:defRPr>`,
      // or a `<a:defRPr>` appearing where the writer asked for `<a:rPr>`.
      const raw = props._rawXml;
      const openRe = /<a:(?:defRPr|rPr|endParaRPr)\b([^>]*)(\/?)>/;
      const closeRe = /<\/a:(?:defRPr|rPr|endParaRPr)>/;
      const openMatch = openRe.exec(raw);
      if (!openMatch) {
        // Captured bytes aren't a recognisable run-properties element.
        // Drop to the structured path below so we don't emit garbage.
      } else {
        const attrs = openMatch[1] ?? "";
        const selfClosing = openMatch[2] === "/";
        const stripNamespace = tag.replace(/^a:/, "");
        const targetTag = `a:${stripNamespace}`;
        if (selfClosing) {
          xml.writeRaw(`<${targetTag}${attrs}/>`);
          return;
        }
        const closeMatch = closeRe.exec(raw);
        if (closeMatch) {
          const inner = raw.slice(openMatch.index + openMatch[0].length, closeMatch.index);
          xml.writeRaw(`<${targetTag}${attrs}>${inner}</${targetTag}>`);
          return;
        }
        // Malformed — fall through to structured path.
      }
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
    // Emit raw XML only when the model is a pure raw-capture — i.e.
    // `_rawXml` is present AND no structured field has been set. If
    // the caller assigned `spPr.fill = {...}` directly (without going
    // through `setSpPrFill`, which clears `_rawXml`), the stale raw
    // bytes are stale and the structured writer must win. Previously
    // this check was `if (spPr._rawXml)` — so any direct mutation of
    // a loaded chart's spPr was silently dropped on save.
    if (isRawXmlShape(spPr)) {
      xml.writeRaw(spPr._rawXml!);
      return;
    }

    xml.openNode("c:spPr");
    // Per DrawingML §20.1.2.2.35 (CT_ShapeProperties), child order is:
    // xfrm? → (prstGeom | custGeom)? → (fill)? → ln? → effectLst|effectDag? →
    // scene3d? → sp3d? → extLst?. We emit in that order so round-trip
    // through `parseSpPr → buildSpPr` re-serialises to valid OOXML.
    if (spPr.transform) {
      this._renderShapeTransform(xml, spPr.transform);
    }
    if (spPr.presetGeometry) {
      this._renderPresetGeometry(xml, spPr.presetGeometry);
    } else if (spPr.customGeometry) {
      this._renderCustomGeometry(xml, spPr.customGeometry);
    }
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
            // OOXML `<a:gs pos>` encodes position as hundredths of a
            // percent (0–100000), NOT thousandths. The previous writer
            // used ×1000 — emitting `pos="1000"` for a 100% stop, which
            // schema-validating readers treat as 1%. Matching fix lives
            // in `chart-ex-renderer.ts` and the gradient parser in
            // `shape-properties.ts`.
            const encoded = Math.max(0, Math.min(100000, Math.round(stop.position * 100000)));
            xml.openNode("a:gs", { pos: String(encoded) });
            this._renderColor(xml, stop.color);
            xml.closeNode();
          }
          xml.closeNode();
          if (g.type === "linear" || g.type === undefined) {
            // `<a:lin ang>` is in 60000ths of a degree. `Math.round`
            // keeps the attribute an integer even for fractional
            // degrees passed through the structured model.
            //
            // `scaled` is only emitted when the author explicitly set
            // it — leaving it absent lets DrawingML apply the implicit
            // default (`scaled="1"`). Unconditionally writing
            // `scaled="1"` previously overwrote a parsed `scaled="0"`
            // on round-trip, silently toggling whether the angle
            // scales with the shape's aspect ratio.
            const angleEmu = Math.round((g.angle ?? 0) * 60000);
            const linAttrs: Record<string, string> = { ang: String(angleEmu) };
            if (g.scaled !== undefined) {
              linAttrs.scaled = g.scaled ? "1" : "0";
            }
            xml.leafNode("a:lin", linAttrs);
          } else if (g.type === "circle" || g.type === "rect" || g.type === "shape") {
            // Path gradient with optional focal rectangle. Preserve
            // parsed `fillToRect` so off-centre radial gradients
            // round-trip; fall back to Excel's centred default.
            // `CT_FillToRectangle` sides are `ST_Percentage` — the
            // full signed range is legal; don't clamp to `[0, 100000]`
            // or off-centre focal points get lost on re-save.
            const rect = g.fillToRect;
            const pct = (v: number | undefined, def: number): number => {
              if (v === undefined) {
                return def;
              }
              return Math.round(v * 100000);
            };
            xml.openNode("a:path", { path: g.type });
            xml.leafNode("a:fillToRect", {
              l: String(pct(rect?.left, 50000)),
              t: String(pct(rect?.top, 50000)),
              r: String(pct(rect?.right, 50000)),
              b: String(pct(rect?.bottom, 50000))
            });
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
      } else if (spPr.fill.blip?.relationshipId) {
        // Picture fill. The rel id was allocated during
        // worksheet._registerChart (see chart-images.ts); by the time
        // the writer runs it is always a concrete `rIdN`. When the id
        // is missing we skip the element entirely rather than emit a
        // broken `<a:blip r:embed="">` — the caller kept their
        // pictureOptions, so the chart still opens correctly.
        this._renderBlipFill(xml, spPr.fill.blip);
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

  /**
   * Emit `<a:xfrm>` with optional `@rot` / `@flipH` / `@flipV` attributes
   * and nested `<a:off>` / `<a:ext>`. Matches what `parseSpPr` captures in
   * `shape-properties.ts`.
   */
  private _renderShapeTransform(xml: XmlSink, transform: ShapeTransform): void {
    const attrs: Record<string, string> = {};
    if (transform.rotation !== undefined && transform.rotation !== 0) {
      attrs.rot = String(transform.rotation);
    }
    if (transform.flipHorizontal) {
      attrs.flipH = "1";
    }
    if (transform.flipVertical) {
      attrs.flipV = "1";
    }
    const hasOff = transform.offsetX !== undefined || transform.offsetY !== undefined;
    const hasExt = transform.width !== undefined || transform.height !== undefined;
    if (!hasOff && !hasExt && Object.keys(attrs).length === 0) {
      // Nothing meaningful to emit.
      return;
    }
    if (!hasOff && !hasExt) {
      xml.leafNode("a:xfrm", attrs);
      return;
    }
    xml.openNode("a:xfrm", Object.keys(attrs).length > 0 ? attrs : undefined);
    if (hasOff) {
      xml.leafNode("a:off", {
        x: String(transform.offsetX ?? 0),
        y: String(transform.offsetY ?? 0)
      });
    }
    if (hasExt) {
      xml.leafNode("a:ext", {
        cx: String(transform.width ?? 0),
        cy: String(transform.height ?? 0)
      });
    }
    xml.closeNode();
  }

  /**
   * Emit `<a:prstGeom prst="…"><a:avLst>…</a:avLst></a:prstGeom>`. The
   * `a:avLst` wrapper is always written (even when empty) to match Excel's
   * own output; adjustments are optional inside.
   */
  private _renderPresetGeometry(xml: XmlSink, geom: PresetGeometry): void {
    xml.openNode("a:prstGeom", { prst: geom.preset });
    xml.openNode("a:avLst");
    for (const adj of geom.adjustments ?? []) {
      xml.leafNode("a:gd", { name: adj.name, fmla: adj.fmla });
    }
    xml.closeNode();
    xml.closeNode();
  }

  /**
   * Emit `<a:custGeom>` with its adjustments list and the `<a:pathLst>`
   * command stream. Commands map 1:1 to the DrawingML path command set
   * (`moveTo` / `lnTo` / `arcTo` / `cubicBezTo` / `quadBezTo` / `close`);
   * `arcTo` carries its parameters as attributes instead of control
   * points, matching the parser.
   */
  private _renderCustomGeometry(xml: XmlSink, geom: CustomGeometry): void {
    xml.openNode("a:custGeom");
    xml.openNode("a:avLst");
    for (const adj of geom.adjustments ?? []) {
      xml.leafNode("a:gd", { name: adj.name, fmla: adj.fmla });
    }
    xml.closeNode();
    // `a:gdLst`, `a:ahLst`, `a:cxnLst`, `a:rect` are omitted here — the
    // parser does not consume them and the round-trip path is dominated
    // by `_rawXml` for geometry-heavy shapes.
    xml.openNode("a:pathLst");
    for (const path of geom.paths ?? []) {
      const pathAttrs: Record<string, string> = {};
      if (path.w !== undefined) {
        pathAttrs.w = String(path.w);
      }
      if (path.h !== undefined) {
        pathAttrs.h = String(path.h);
      }
      if (path.fill !== undefined) {
        pathAttrs.fill = path.fill;
      }
      if (path.stroke !== undefined) {
        pathAttrs.stroke = path.stroke ? "1" : "0";
      }
      xml.openNode("a:path", Object.keys(pathAttrs).length > 0 ? pathAttrs : undefined);
      for (const cmd of path.commands) {
        this._renderCustomGeometryCommand(xml, cmd);
      }
      xml.closeNode();
    }
    xml.closeNode();
    xml.closeNode();
  }

  private _renderCustomGeometryCommand(xml: XmlSink, cmd: CustomGeometryCommand): void {
    if (cmd.type === "close") {
      xml.leafNode("a:close");
      return;
    }
    if (cmd.type === "arcTo") {
      const p = cmd.arcParams;
      if (!p) {
        return;
      }
      xml.leafNode("a:arcTo", {
        wR: String(p.wR),
        hR: String(p.hR),
        stAng: String(p.stAng),
        swAng: String(p.swAng)
      });
      return;
    }
    const tag =
      cmd.type === "moveTo"
        ? "a:moveTo"
        : cmd.type === "lnTo"
          ? "a:lnTo"
          : cmd.type === "cubicBezTo"
            ? "a:cubicBezTo"
            : "a:quadBezTo";
    xml.openNode(tag);
    for (const point of cmd.points ?? []) {
      xml.leafNode("a:pt", { x: String(point.x), y: String(point.y) });
    }
    xml.closeNode();
  }

  /**
   * Render an `<a:blipFill>` element for a picture fill. The element is
   * `<a:blipFill><a:blip r:embed="rIdN"/>[<a:srcRect …/>]<a:stretch …>
   * or <a:tile …></a:blipFill>` matching the OOXML shape-fill subschema
   * (DrawingML §20.1.8.14).
   *
   * Callers are responsible for making sure `blip.relationshipId` is
   * populated — the writer no-ops when it is missing (see the comment
   * at the call site in `_renderSpPr`).
   */
  private _renderBlipFill(xml: XmlSink, blip: ChartBlipFill): void {
    xml.openNode("a:blipFill");
    xml.leafNode("a:blip", { "r:embed": blip.relationshipId! });
    if (blip.sourceRectangle) {
      const sr = blip.sourceRectangle;
      const attrs: Record<string, string> = {};
      if (sr.left !== undefined) {
        attrs.l = String(sr.left);
      }
      if (sr.top !== undefined) {
        attrs.t = String(sr.top);
      }
      if (sr.right !== undefined) {
        attrs.r = String(sr.right);
      }
      if (sr.bottom !== undefined) {
        attrs.b = String(sr.bottom);
      }
      xml.leafNode("a:srcRect", Object.keys(attrs).length > 0 ? attrs : undefined);
    }
    if (blip.fillMode === "tile") {
      const t = blip.tile ?? {};
      const tileAttrs: Record<string, string> = {};
      if (t.tx !== undefined) {
        tileAttrs.tx = String(t.tx);
      }
      if (t.ty !== undefined) {
        tileAttrs.ty = String(t.ty);
      }
      if (t.sx !== undefined) {
        tileAttrs.sx = String(t.sx);
      }
      if (t.sy !== undefined) {
        tileAttrs.sy = String(t.sy);
      }
      if (t.flip) {
        tileAttrs.flip = t.flip;
      }
      if (t.alignment) {
        tileAttrs.algn = t.alignment;
      }
      xml.leafNode("a:tile", Object.keys(tileAttrs).length > 0 ? tileAttrs : undefined);
    } else if (blip.fillMode !== "none") {
      // Default and "stretch" produce a fillRect-stretch.
      xml.openNode("a:stretch");
      xml.leafNode("a:fillRect");
      xml.closeNode();
    }
    xml.closeNode();
  }

  private _renderEffectList(xml: XmlSink, effects: EffectList): void {
    xml.openNode("a:effectLst");
    // DrawingML `CT_EffectList` declares a strict `<xsd:sequence>`:
    //   blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw →
    //   reflection → softEdge.
    // The previous emission order (blur → outerShdw → innerShdw →
    // prstShdw → glow → softEdge → reflection) was accepted by Excel
    // but rejected by strict validators (`ooxml-validate`, LibreOffice
    // strict mode). Keep this ordering in sync with
    // `chart-ex-renderer.ts:renderEffectList`. `fillOverlay` is not
    // modelled yet and its slot is intentionally empty.
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
    // fillOverlay — reserved slot per schema, not currently modelled.
    if (effects.glow) {
      xml.openNode("a:glow", { rad: String(effects.glow.radius) });
      this._renderColor(xml, effects.glow.color);
      xml.closeNode();
    }
    if (effects.innerShadow) {
      this._renderShadowElement(xml, "a:innerShdw", effects.innerShadow);
    }
    if (effects.outerShadow) {
      this._renderShadowElement(xml, "a:outerShdw", effects.outerShadow);
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
    if (effects.softEdge) {
      xml.leafNode("a:softEdge", { rad: String(effects.softEdge.radius) });
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
    // Children shared across all five DrawingML colour-element kinds
    // (srgbClr, schemeClr, sysClr, schemeClr-by-name, prstClr). Rather
    // than duplicating the guard + emission logic in every branch,
    // route through these two helpers.
    //
    // `hasModifiers` avoids the pathological
    // `<a:srgbClr val="…"></a:srgbClr>` output (invalid child count);
    // `emitModifiers` enforces:
    //   1. Fields stored as OOXML integers (`alpha`, `shade`, `satMod`,
    //      `lumMod`, `lumOff`) round to guard against `NaN.toString()`
    //      producing `"NaN"` in attribute values (strict readers reject).
    //   2. `tint` is a fraction in 0..1 per `ChartColor.tint` contract;
    //      multiply by 100000. Negative `tint` on `schemeClr` is a
    //      known Excel-UI convention for "darken by |tint|"; map to
    //      `<a:shade>` exactly as the existing code did.
    const hasModifiers = (c: ChartColor): boolean =>
      c.alpha !== undefined ||
      c.tint !== undefined ||
      c.lumMod !== undefined ||
      c.lumOff !== undefined ||
      c.shade !== undefined ||
      c.satMod !== undefined;
    const emitInt = (tag: string, value: number | undefined): void => {
      if (value === undefined || !Number.isFinite(value)) {
        return;
      }
      xml.leafNode(tag, { val: String(Math.round(value)) });
    };
    const emitModifiers = (c: ChartColor, tintAsShade: boolean = false): void => {
      emitInt("a:alpha", c.alpha);
      let tintEmittedShade = false;
      if (c.tint !== undefined && Number.isFinite(c.tint)) {
        if (tintAsShade && c.tint < 0) {
          // Legacy scheme-colour convention: negative tint → equivalent
          // shade. `(1 + tint) * 100000` lives on the 0..100000 range
          // for tint ∈ [-1, 0].
          xml.leafNode("a:shade", { val: String(Math.round((1 + c.tint) * 100000)) });
          tintEmittedShade = true;
        } else {
          xml.leafNode("a:tint", { val: String(Math.round(c.tint * 100000)) });
        }
      }
      emitInt("a:lumMod", c.lumMod);
      emitInt("a:lumOff", c.lumOff);
      if (!tintEmittedShade) {
        emitInt("a:shade", c.shade);
      }
      emitInt("a:satMod", c.satMod);
    };

    if (color.srgb) {
      const attrs: Record<string, string> = { val: color.srgb };
      if (!hasModifiers(color)) {
        xml.leafNode("a:srgbClr", attrs);
      } else {
        xml.openNode("a:srgbClr", attrs);
        emitModifiers(color);
        xml.closeNode();
      }
    } else if (color.theme !== undefined) {
      const name = themeIndexToName(color.theme);
      if (!hasModifiers(color)) {
        xml.leafNode("a:schemeClr", { val: name });
      } else {
        xml.openNode("a:schemeClr", { val: name });
        // Scheme-colour is the only variant that honours the negative-
        // tint-as-shade legacy mapping; the other elements don't.
        emitModifiers(color, /* tintAsShade */ true);
        xml.closeNode();
      }
    } else if (color.sysClr) {
      if (!hasModifiers(color)) {
        xml.leafNode("a:sysClr", { val: color.sysClr });
      } else {
        xml.openNode("a:sysClr", { val: color.sysClr });
        emitModifiers(color);
        xml.closeNode();
      }
    } else if (color.schemeName) {
      // Scheme-colour tokens that can't be mapped onto a theme index
      // (e.g. `phClr`, vendor extensions). Round-trip as
      // `<a:schemeClr>` rather than silently re-emitting as
      // `<a:sysClr>` (which the old parser did, but which is a
      // semantically-different DrawingML element kind).
      if (!hasModifiers(color)) {
        xml.leafNode("a:schemeClr", { val: color.schemeName });
      } else {
        xml.openNode("a:schemeClr", { val: color.schemeName });
        emitModifiers(color);
        xml.closeNode();
      }
    } else if (color.prstClr) {
      if (!hasModifiers(color)) {
        xml.leafNode("a:prstClr", { val: color.prstClr });
      } else {
        xml.openNode("a:prstClr", { val: color.prstClr });
        emitModifiers(color);
        xml.closeNode();
      }
    }
    // else: unknown / empty color — drop silently (same as before)
  }

  private _renderPrintSettings(xml: XmlSink, ps: PrintSettings): void {
    xml.openNode("c:printSettings");
    // Only emit `c:headerFooter` / `c:pageSetup` when the model
    // actually carries them. Previously we unconditionally wrote empty
    // `<c:headerFooter/>` / `<c:pageSetup/>` sentinels, which broke
    // byte-for-byte round-trip on files that did not originally
    // include them (and added elements Excel would gladly default,
    // inflating diffs / confusing compare tools).
    if (typeof ps.headerFooter === "string") {
      xml.writeRaw(ps.headerFooter);
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
      (name === "c:tx" && this.currentTrendlineLbl != null)
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
        } else if (this.currentPivotFormat) {
          // Must be checked before `series`/`chartTypeGroup` because
          // pivotFmt lives alongside (not inside) the chart type group
          // but is set independently; without this branch a pivotFmt
          // `extLst` would fall through to `chart` and be attached at
          // the wrong level.
          this.rawCaptureContext = "pivotFormat";
        } else if (this.currentDisplayUnits) {
          this.rawCaptureContext = "displayUnits";
        } else if (this.currentUpDownBars) {
          this.rawCaptureContext = "upDownBars";
        } else if (this.currentView3D) {
          this.rawCaptureContext = "view3D";
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
      // c:dLbl inside pivotFmt is now parsed structurally (see the
      // `currentDataLabelEntry` branch below — `c:dLbl` open always
      // starts an entry, and the matching close routes it to the
      // pivotFormat or the data-labels bundle depending on the
      // surrounding state). The previous raw-capture branch that
      // preempted structured parsing has been removed because it
      // made `PivotFormat.dLbl` unreachable from XML — callers could
      // only populate the field by mutating the model directly, and
      // the writer then had to consult both `dLbl` and the
      // deprecated `rawDLbl` fallback. Now the parser always
      // produces `dLbl`; `rawDLbl` remains on the type only for
      // legacy round-trip of files parsed by older versions.
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
        this.chartModel.roundedCorners = parseXsdBoolean(attrs.val) ?? true;
        break;
      case "c:date1904":
        this.chartModel.date1904 = parseXsdBoolean(attrs.val) ?? true;
        break;
      case "c:userShapes":
        // `<c:userShapes r:id="rIdN"/>` — reference to a sibling drawing
        // part holding freehand annotations. The part itself travels
        // with the chart rels, we only preserve the reference.
        if (attrs["r:id"]) {
          this.chartModel.userShapesRelId = attrs["r:id"];
        }
        break;
      case "c:style": {
        const style = parseXsdInt(attrs.val);
        if (style !== undefined) {
          this.chartModel.style = style;
        }
        break;
      }
      case "c:chart":
        this.chartData = { plotArea: { chartTypes: [], axes: [] } };
        break;
      case "c:title":
        this.currentTitle = {};
        this.stateStack.push({ tag: "c:title", context: this.currentTitle });
        break;
      case "c:autoTitleDeleted":
        if (this.chartData) {
          this.chartData.autoTitleDeleted = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:view3D":
        this.currentView3D = {};
        break;
      case "c:rotX":
        if (this.currentView3D) {
          // View3D integer fields default to chart-type-specific values
          // (`rotX=15` for most, `0` for pie3D) rather than a single
          // universal default. Forcing `0` on absent attribute changes
          // rotation semantics AND round-trips as `val="0"` — breaking
          // byte-compare tests. Leave undefined so the writer omits.
          this.currentView3D.rotX = parseXsdInt(attrs.val);
        }
        break;
      case "c:rotY":
        if (this.currentView3D) {
          this.currentView3D.rotY = parseXsdInt(attrs.val);
        }
        break;
      case "c:depthPercent":
        if (this.currentView3D) {
          this.currentView3D.depthPercent = parseXsdInt(attrs.val);
        }
        break;
      case "c:rAngAx":
        if (this.currentView3D) {
          this.currentView3D.rAngAx = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:hPercent":
        if (this.currentView3D) {
          this.currentView3D.hPercent = parseXsdInt(attrs.val);
        }
        break;
      case "c:perspective":
        if (this.currentView3D) {
          this.currentView3D.perspective = parseXsdInt(attrs.val);
        }
        break;
      case "c:plotArea":
        this.plotArea = { chartTypes: [], axes: [] };
        this.insidePlotArea = true;
        break;
      case "c:plotVisOnly":
        if (this.chartData) {
          this.chartData.plotVisOnly = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:dispBlanksAs":
        if (this.chartData) {
          // `ST_DispBlanksAs` values: gap | span | zero.
          this.chartData.dispBlanksAs = narrowEnumValue(attrs.val, [
            "gap",
            "span",
            "zero"
          ] as const);
        }
        break;
      case "c:showDLblsOverMax":
        if (this.chartData) {
          this.chartData.showDLblsOverMax = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:legend":
        // Don't seed `legendPos` to `"b"` up-front — `CT_Legend/legendPos`
        // is optional per ECMA-376 §21.2.2.118 and Excel happily omits
        // it when the user hasn't chosen a specific position. A hard
        // default would pollute every round-tripped legend with an
        // unchanged-but-suddenly-explicit position (`<c:legendPos val="b"/>`),
        // breaking byte-for-byte round-trip and interfering with the
        // "legendPos undefined → don't emit" fix in `_renderLegend`.
        this.currentLegend = {};
        break;
      case "c:legendPos":
        if (this.currentLegend) {
          // `ST_LegendPos` values per ECMA-376 §21.2.3.25.
          this.currentLegend.legendPos = narrowEnumValue(attrs.val, [
            "b",
            "tr",
            "l",
            "r",
            "t"
          ] as const);
        }
        break;
      case "c:overlay":
        // ECMA-376 `CT_Boolean` defaults `val` to `true` when absent,
        // but specific elements override that. `c:overlay` on both
        // `c:title` and `c:legend` documents its default as `false`
        // (the title / legend should NOT overlay the plot area unless
        // explicitly asked). Parsing `<c:overlay/>` as `true` flipped
        // the convention and moved titles/legends on top of charts
        // after any round-trip of files that carried the element with
        // no `val` attribute.
        if (this.currentLegend) {
          this.currentLegend.overlay = parseXsdBoolean(attrs.val) ?? false;
        } else if (this.currentTitle) {
          this.currentTitle.overlay = parseXsdBoolean(attrs.val) ?? false;
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
          const paperSize = parseXsdInt(attrs.paperSize);
          if (paperSize !== undefined) {
            this.currentPrintSettings.pageSetup.paperSize = paperSize;
          }
        }
        break;
      case "c:externalData":
        this.chartModel.externalData = { id: attrs["r:id"] ?? "" };
        break;
      case "c:autoUpdate":
        if (this.chartModel.externalData) {
          this.chartModel.externalData.autoUpdate = parseXsdBoolean(attrs.val) ?? true;
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
          // `ST_LayoutTarget` values: inner | outer (ECMA-376 §21.2.3.21).
          this.currentManualLayout.layoutTarget = narrowEnumValue(attrs.val, [
            "inner",
            "outer"
          ] as const);
        }
        break;
      case "c:xMode":
      case "c:yMode":
      case "c:wMode":
      case "c:hMode":
        if (this.currentManualLayout) {
          // `ST_LayoutMode` values: edge | factor (ECMA-376 §21.2.3.22).
          const mode = narrowEnumValue(attrs.val, ["edge", "factor"] as const);
          if (name === "c:xMode") {
            this.currentManualLayout.xMode = mode;
          } else if (name === "c:yMode") {
            this.currentManualLayout.yMode = mode;
          } else if (name === "c:wMode") {
            this.currentManualLayout.wMode = mode;
          } else {
            this.currentManualLayout.hMode = mode;
          }
        }
        break;
      case "c:x":
        if (this.currentManualLayout) {
          const v = parseXsdFloat(attrs.val);
          if (v !== undefined) {
            this.currentManualLayout.x = v;
          }
        }
        break;
      case "c:y":
        if (this.currentManualLayout) {
          const v = parseXsdFloat(attrs.val);
          if (v !== undefined) {
            this.currentManualLayout.y = v;
          }
        }
        break;
      case "c:w":
        if (this.currentManualLayout) {
          const v = parseXsdFloat(attrs.val);
          if (v !== undefined) {
            this.currentManualLayout.w = v;
          }
        }
        break;
      case "c:h":
        if (this.currentManualLayout) {
          const v = parseXsdFloat(attrs.val);
          if (v !== undefined) {
            this.currentManualLayout.h = v;
          }
        }
        break;

      // ---- Data Table (#8) ----
      case "c:dTable":
        this.currentDataTable = {};
        break;
      case "c:showHorzBorder":
        if (this.currentDataTable) {
          this.currentDataTable.showHorzBorder = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:showVertBorder":
        if (this.currentDataTable) {
          this.currentDataTable.showVertBorder = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:showOutline":
        if (this.currentDataTable) {
          this.currentDataTable.showOutline = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:showKeys":
        if (this.currentDataTable) {
          this.currentDataTable.showKeys = parseXsdBoolean(attrs.val) ?? true;
        }
        break;

      // ---- Display Units (#9) ----
      case "c:dispUnits":
        this.currentDisplayUnits = {};
        break;
      case "c:builtInUnit":
        if (this.currentDisplayUnits) {
          // `ST_BuiltInUnit` enumeration (ECMA-376 §21.2.3.2). Values
          // outside the set fall through to `undefined`, which
          // `narrowEnumValue` signals — caller keeps whatever default
          // the model already has.
          const v = narrowEnumValue(attrs.val, [
            "hundreds",
            "thousands",
            "tenThousands",
            "hundredThousands",
            "millions",
            "tenMillions",
            "hundredMillions",
            "billions",
            "trillions"
          ] as const);
          if (v !== undefined) {
            this.currentDisplayUnits.builtInUnit = v;
          }
        }
        break;
      case "c:custUnit":
        if (this.currentDisplayUnits) {
          this.currentDisplayUnits.custUnit = parseXsdFloat(attrs.val) ?? 0;
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
        // Non-c: namespaced elements at chart-space top level (not already
        // captured via rawXmlTarget and not c:/a:/r: native) are vendor
        // extensions; surface them on the model so `strict` template mode
        // can warn when a structural rebuild would drop them. This
        // mirrors ChartExModel.unknownElements and uses the same shape
        // (ChartExUnknownElement-compatible { name, path }).
        //
        // We also **enter raw-capture** for the unknown element's full
        // subtree. Without this, children named like `c:idx` / `c:val`
        // inside the unknown block would fall through to the core
        // dispatcher and write to the current parent (e.g.
        // `currentSeries.index`), corrupting the structural model. The
        // captured bytes themselves are discarded — we only care that
        // the inner dispatcher cannot touch current* state.
        if (
          this.rawCapture === null &&
          name.includes(":") &&
          !name.startsWith("c:") &&
          !name.startsWith("a:") &&
          !name.startsWith("r:") &&
          !name.startsWith("mc:") &&
          !name.startsWith("xsi:") &&
          !name.startsWith("xml:")
        ) {
          const parent = this.stateStack[this.stateStack.length - 1]?.tag ?? "c:chartSpace";
          if (!this.chartModel.unknownElements) {
            this.chartModel.unknownElements = [];
          }
          this.chartModel.unknownElements.push({ name, path: `${parent}/${name}` });
          if (!node.isSelfClosing) {
            this.rawCapture = new RawXmlCapture();
            this.rawCapture.start(node);
            // Deliberately do NOT set `rawXmlTarget` / `rawCaptureContext`
            // so the matching close handler discards the buffer instead
            // of trying to attach it to a structured model slot.
            this.rawXmlTarget = "__unknown__";
          }
        }
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
        const target = this.rawXmlTarget;
        this.rawXmlTarget = null;
        // `__unknown__` targets are vendor / unknown-namespace subtrees
        // that we captured solely to prevent the inner dispatcher from
        // polluting `current*` state. The captured bytes have no
        // structured sink — drop them. (`unknownElements` was already
        // recorded on the open side.)
        if (target !== null && target !== "__unknown__") {
          this._attachRawXml(target, rawXml);
        }
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
          // `CT_SurfaceChart` is the only chart-type group that has no
          // `dLbls` child per schema. A `c:dLbls` element encountered
          // under `<c:surfaceChart>` / `<c:surface3DChart>` means the
          // author wrote invalid XML — dropping the element here
          // matches the builder's `dataLabels is not supported for
          // surface charts` validation and prevents round-trip from
          // re-emitting the invalid structure. Every other group
          // type honours the attribute.
          const groupType = (this.currentChartTypeGroup as { type?: string }).type;
          if (groupType !== "surface" && groupType !== "surface3D") {
            (this.currentChartTypeGroup as { dataLabels?: unknown }).dataLabels =
              this.currentDataLabels;
          }
        }
        this.currentDataLabels = null;
        // Defensive: if a `<c:dLbl>` closed mid-parse (e.g. an error
        // captured before the `dLbl` pop or an XML file where an `idx`
        // was followed by a `delete` without later elements), the entry
        // may still be dangling. Clear it so the next series's
        // `<c:dLbls>` starts clean — otherwise a stray `c:dLblPos`
        // parsed after this `dLbls` close was mis-attributed to the
        // previous series's dLbl entry.
        this.currentDataLabelEntry = null;
        break;
      case "c:dLbl":
        if (this.currentDataLabels && this.currentDataLabelEntry) {
          if (!this.currentDataLabels.entries) {
            this.currentDataLabels.entries = [];
          }
          this.currentDataLabels.entries.push(this.currentDataLabelEntry);
          this.currentDataLabelEntry = null;
        } else if (this.currentPivotFormat && this.currentDataLabelEntry) {
          // Bare `<c:dLbl>` inside `<c:pivotFmt>` (no `<c:dLbls>`
          // wrapper). Pivot samples from Excel emit a single dLbl
          // here and the structured slot on `PivotFormat.dLbl`
          // captures it. Previously this element was raw-captured
          // into `rawDLbl`, which made the structured field
          // unreachable from on-disk files.
          this.currentPivotFormat.dLbl = this.currentDataLabelEntry;
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
          this.currentDataLabelEntry.index = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentDataPoint) {
          this.currentDataPoint.index = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentLegendEntry) {
          this.currentLegendEntry.index = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentPivotFormat) {
          this.currentPivotFormat.index = parseXsdInt(attrs.val) ?? 0;
        } else {
          // Check for bandFmt context in stateStack
          const bandFmtState = this.stateStack.find(s => s.tag === "c:bandFmt");
          if (bandFmtState) {
            bandFmtState.context.index = parseXsdInt(attrs.val) ?? 0;
          } else if (this.currentSeries) {
            this.currentSeries.index = parseXsdInt(attrs.val) ?? 0;
          }
        }
        break;
      case "c:order":
        if (this.currentTrendline) {
          this.currentTrendline.order = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentSeries) {
          this.currentSeries.order = parseXsdInt(attrs.val) ?? 0;
        }
        break;

      // Chart type group attributes
      case "c:barDir":
        // Skip when `val` is missing or unrecognised so a malformed
        // `<c:barDir/>` doesn't land `undefined` on the model, which
        // would later serialise as the literal string `"undefined"` in
        // writer attributes. `ST_BarDir` admits `"bar"` / `"col"`.
        if (this.currentChartTypeGroup) {
          const v = narrowEnumValue(attrs.val, ["bar", "col"] as const);
          if (v !== undefined) {
            this.currentChartTypeGroup.barDir = v;
          }
        }
        break;
      case "c:grouping":
        if (this.currentChartTypeGroup) {
          // `ST_Grouping` for bar / line / area: standard, stacked,
          // percentStacked. Bar additionally admits `clustered`.
          const v = narrowEnumValue(attrs.val, [
            "standard",
            "clustered",
            "stacked",
            "percentStacked"
          ] as const);
          if (v !== undefined) {
            this.currentChartTypeGroup.grouping = v;
          }
        }
        break;
      case "c:varyColors":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.varyColors = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:scatterStyle":
        if (this.currentChartTypeGroup && attrs.val !== undefined) {
          this.currentChartTypeGroup.scatterStyle = attrs.val;
        }
        break;
      case "c:radarStyle":
        if (this.currentChartTypeGroup && attrs.val !== undefined) {
          this.currentChartTypeGroup.radarStyle = attrs.val;
        }
        break;
      case "c:ofPieType":
        if (this.currentChartTypeGroup && attrs.val !== undefined) {
          this.currentChartTypeGroup.ofPieType = attrs.val;
        }
        break;
      case "c:gapWidth":
        if (this.currentUpDownBars) {
          this.currentUpDownBars.gapWidth = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.gapWidth = parseXsdInt(attrs.val) ?? 0;
        }
        break;
      case "c:gapDepth":
        if (this.currentChartTypeGroup) {
          // Schema default for `c:gapDepth` is 150 (not 0). Coercing a
          // missing-attribute `<c:gapDepth/>` to 0 produced a bar3D
          // with zero inter-series gap — not Excel's actual default.
          // Leave `undefined` so the writer's `!== undefined` guard
          // omits the element and Excel applies its own 150 default.
          (this.currentChartTypeGroup as { gapDepth?: number }).gapDepth = parseXsdInt(attrs.val);
        }
        break;
      case "c:overlap":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.overlap = parseXsdInt(attrs.val) ?? 0;
        }
        break;
      case "c:firstSliceAng":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.firstSliceAng = parseXsdInt(attrs.val) ?? 0;
        }
        break;
      case "c:holeSize":
        if (this.currentChartTypeGroup) {
          // `c:holeSize` has valid range 10–90 (schema default 10 for
          // doughnut). Coercing a missing / blank `<c:holeSize/>` to
          // `0` produced out-of-range output that Excel's strict
          // validators flag and some OOXML consumers reject. Leave
          // undefined so the writer omits the element.
          this.currentChartTypeGroup.holeSize = parseXsdInt(attrs.val);
        }
        break;
      case "c:bubbleScale":
        if (this.currentChartTypeGroup) {
          // Schema default 100 (valid range 0–300). Same rationale as
          // `c:holeSize` — don't write `0` back when the attribute was
          // absent on the wire.
          this.currentChartTypeGroup.bubbleScale = parseXsdInt(attrs.val);
        }
        break;
      case "c:showNegBubbles":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.showNegBubbles = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:sizeRepresents":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.sizeRepresents = attrs.val;
        }
        break;
      case "c:wireframe":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.wireframe = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:splitType":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.splitType = attrs.val;
        }
        break;
      case "c:splitPos":
        if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.splitPos = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:secondPieSize":
        if (this.currentChartTypeGroup) {
          // `c:secondPieSize` valid range 5–200 (schema default 75).
          // 0 is invalid per schema; coercing to 0 triggers Excel's
          // repair dialog on reopen. Leave undefined on missing val.
          this.currentChartTypeGroup.secondPieSize = parseXsdInt(attrs.val);
        }
        break;
      case "c:axId":
        if (this.currentChartTypeGroup) {
          const v = parseXsdInt(attrs.val);
          if (v !== undefined) {
            this.currentChartTypeGroup.axisIds.push(v);
          }
        }
        break;

      // Series-specific
      case "c:invertIfNegative":
        // `c:invertIfNegative` is valid both on `c:ser` (series-level)
        // and on `c:dPt` (per-point override). Check the inner
        // context first so a per-point value doesn't leak onto the
        // whole series — previously this silently overwrote
        // `currentSeries.invertIfNegative` with the last `dPt` value
        // and dropped the per-point flag.
        if (this.currentDataPoint) {
          this.currentDataPoint.invertIfNegative = parseXsdBoolean(attrs.val) ?? true;
        } else if (this.currentSeries) {
          this.currentSeries.invertIfNegative = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:smooth":
        if (this.currentSeries) {
          this.currentSeries.smooth = parseXsdBoolean(attrs.val) ?? true;
        } else if (this.currentChartTypeGroup) {
          this.currentChartTypeGroup.smooth = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:explosion":
        if (this.currentDataPoint) {
          this.currentDataPoint.explosion = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentSeries) {
          this.currentSeries.explosion = parseXsdInt(attrs.val) ?? 0;
        }
        break;
      case "c:bubble3D":
        if (this.currentDataPoint) {
          this.currentDataPoint.bubble3D = parseXsdBoolean(attrs.val) ?? true;
        } else if (this.currentSeries) {
          this.currentSeries.bubble3D = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:shape":
        {
          // `ST_Shape` for bar3D: cone, coneToMax, box, cylinder,
          // pyramid, pyramidToMax. Default is `"box"`.
          const v = narrowEnumValue(attrs.val, [
            "cone",
            "coneToMax",
            "box",
            "cylinder",
            "pyramid",
            "pyramidToMax"
          ] as const);
          if (v !== undefined) {
            if (this.currentSeries) {
              this.currentSeries.shape = v;
            } else if (this.currentChartTypeGroup) {
              this.currentChartTypeGroup.shape = v;
            }
          }
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
          this.currentErrorBars.val = parseXsdFloat(attrs.val) ?? 0;
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
          this.currentNumCache.pointCount = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentStrCache) {
          this.currentStrCache.pointCount = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentNumLit) {
          this.currentNumLit.pointCount = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentStrLit) {
          this.currentStrLit.pointCount = parseXsdInt(attrs.val) ?? 0;
        } else if (this.currentMultiLvlStrRef?.cache) {
          this.currentMultiLvlStrRef.cache.pointCount = parseXsdInt(attrs.val) ?? 0;
        }
        break;
      case "c:pt":
        // Capture the optional per-point `formatCode` attribute. Some
        // Excel versions (and notably mixed-format pivot charts) emit
        // different number formats for each point inside the same
        // `c:numCache` — see the serializer in `_renderNumCache` which
        // already emits `formatCode` on `<c:pt>`. Dropping it on parse
        // made every round-trip flatten back to the cache-level
        // `formatCode`, losing the per-point override.
        this.stateStack.push({
          tag: "c:pt",
          context: { index: parseXsdInt(attrs.idx) ?? 0, formatCode: attrs.formatCode }
        });
        break;

      // DataLabels
      case "c:dLbls":
        this.currentDataLabels = {};
        break;
      case "c:dLbl":
        this.currentDataLabelEntry = { index: 0 };
        break;
      case "c:showLegendKey":
        this._setBoolOnLabels("showLegendKey", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showVal":
        this._setBoolOnLabels("showVal", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showCatName":
        this._setBoolOnLabels("showCatName", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showSerName":
        this._setBoolOnLabels("showSerName", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showPercent":
        this._setBoolOnLabels("showPercent", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showBubbleSize":
        this._setBoolOnLabels("showBubbleSize", parseXsdBoolean(attrs.val) ?? true);
        break;
      case "c:showLeaderLines":
        if (this.currentDataLabels) {
          this.currentDataLabels.showLeaderLines = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:dLblPos": {
        // `ST_DLblPos` enumeration (ECMA-376 §21.2.3.11). Unknown
        // values are dropped so we don't echo `val="invalid"` back
        // into the output on save.
        const v = narrowEnumValue(attrs.val, [
          "ctr",
          "l",
          "r",
          "t",
          "b",
          "bestFit",
          "inBase",
          "inEnd",
          "outEnd"
        ] as const);
        if (v !== undefined) {
          if (this.currentDataLabelEntry) {
            this.currentDataLabelEntry.position = v;
          } else if (this.currentDataLabels) {
            this.currentDataLabels.position = v;
          }
        }
        break;
      }
      case "c:numFmt":
        if (this.currentTrendlineLbl) {
          this.currentTrendlineLbl.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: parseXsdBoolean(attrs.sourceLinked) ?? true
          };
        } else if (this.currentAxis) {
          this.currentAxis.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: parseXsdBoolean(attrs.sourceLinked) ?? true
          };
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: parseXsdBoolean(attrs.sourceLinked) ?? true
          };
        } else if (this.currentDataLabels) {
          this.currentDataLabels.numFmt = {
            formatCode: attrs.formatCode,
            sourceLinked: parseXsdBoolean(attrs.sourceLinked) ?? true
          };
        }
        break;

      // Marker
      case "c:marker":
        if (attrs.val !== undefined) {
          // Group-level boolean marker (line/stock charts)
          if (this.currentChartTypeGroup) {
            this.currentChartTypeGroup.marker = parseXsdBoolean(attrs.val) ?? true;
          }
        } else if (
          this.currentChartTypeGroup &&
          !this.currentSeries &&
          !this.currentDataPoint &&
          !this.currentPivotFormat
        ) {
          // Self-closing `<c:marker/>` on a line / stock chart group is
          // CT_Boolean (`<c:marker val="1"/>` with val omitted → default
          // `true` per ECMA-376). Previously we treated it as the start
          // of a container marker, silently losing the boolean and
          // leaving `currentMarker` dangling.
          this.currentChartTypeGroup.marker = true;
        } else {
          // Container marker (series/pivotFmt/dataPoint level)
          this.currentMarker = {};
        }
        break;
      case "c:symbol":
        if (this.currentMarker) {
          // `ST_MarkerStyle` values per ECMA-376 §21.2.3.30.
          this.currentMarker.symbol = narrowEnumValue(attrs.val, [
            "circle",
            "dash",
            "diamond",
            "dot",
            "none",
            "picture",
            "plus",
            "square",
            "star",
            "triangle",
            "x",
            "auto"
          ] as const);
        }
        break;
      case "c:size":
        if (this.currentMarker) {
          // Marker `c:size` valid range 2–72 (schema default 5).
          // Writing `0` is out-of-range — leave undefined when the
          // attribute is absent so the writer omits the element and
          // Excel applies its own default.
          this.currentMarker.size = parseXsdInt(attrs.val);
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
          // `ST_TrendlineType` values per ECMA-376 §21.2.3.49.
          const type = narrowEnumValue(attrs.val, [
            "exp",
            "linear",
            "log",
            "movingAvg",
            "poly",
            "power"
          ] as const);
          // Fall back to the constructor default `"linear"` so the
          // trendline structure remains renderable; the original attrs
          // value (if unrecognised) is simply dropped.
          this.currentTrendline.type = type ?? "linear";
        }
        break;
      case "c:period":
        if (this.currentTrendline) {
          // Moving-average `c:period` has `minInclusive=2` per schema
          // (default 2); writing `0` is out-of-range and Excel
          // clamps/repairs on reopen. Leave undefined so the writer
          // omits the element.
          this.currentTrendline.period = parseXsdInt(attrs.val);
        }
        break;
      case "c:forward":
        if (this.currentTrendline) {
          this.currentTrendline.forward = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:backward":
        if (this.currentTrendline) {
          this.currentTrendline.backward = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:intercept":
        if (this.currentTrendline) {
          this.currentTrendline.intercept = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:dispRSqr":
        if (this.currentTrendline) {
          this.currentTrendline.displayRSqr = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:dispEq":
        if (this.currentTrendline) {
          this.currentTrendline.displayEq = parseXsdBoolean(attrs.val) ?? true;
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
          // `ST_ErrValType` values per ECMA-376 §21.2.3.15. The model
          // field is required, so fall back to `"fixedVal"` (Excel's
          // default) when the source XML carries an unrecognised or
          // missing value — preserves the error-bar structure rather
          // than throwing at parse time.
          this.currentErrorBars.errValType =
            narrowEnumValue(attrs.val, [
              "cust",
              "fixedVal",
              "percentage",
              "stdDev",
              "stdErr"
            ] as const) ?? "fixedVal";
        }
        break;
      case "c:noEndCap":
        if (this.currentErrorBars) {
          this.currentErrorBars.noEndCap = parseXsdBoolean(attrs.val) ?? true;
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
          const v = parseXsdInt(attrs.val);
          if (v !== undefined) {
            this.currentChartTypeGroup.custSplit.push(v);
          }
        }
        break;

      // Picture options (#15)
      case "c:pictureOptions":
        this.currentPictureOptions = {};
        break;
      case "c:applyToFront":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToFront = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:applyToSides":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToSides = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:applyToEnd":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.applyToEnd = parseXsdBoolean(attrs.val) ?? true;
        }
        break;
      case "c:pictureFormat":
        if (this.currentPictureOptions) {
          // `ST_PictureFormat` enumeration (ECMA-376 §21.2.3.38).
          const v = narrowEnumValue(attrs.val, ["stretch", "stack", "stackScale"] as const);
          if (v !== undefined) {
            this.currentPictureOptions.pictureFormat = v;
          }
        }
        break;
      case "c:pictureStackUnit":
        if (this.currentPictureOptions) {
          this.currentPictureOptions.pictureStackUnit = parseXsdFloat(attrs.val) ?? 0;
        }
        break;

      // Separator (#16)
      case "c:separator":
        // Text will be captured in parseClose
        break;

      // delete on legend entry / data-label entry / data-labels block.
      // `<c:delete/>` with no `val` attribute means `true` (CT_Boolean
      // default), per ECMA-376. `currentDataLabelEntry` is checked
      // before `currentDataLabels` so per-label delete wins over
      // block-level delete when both are in scope (unlikely in real
      // files but defensive).
      case "c:delete":
        if (this.currentLegendEntry) {
          this.currentLegendEntry.delete = parseXsdBoolean(attrs.val) ?? true;
        } else if (this.currentDataLabelEntry) {
          this.currentDataLabelEntry.delete = parseXsdBoolean(attrs.val) ?? true;
        } else if (this.currentDataLabels) {
          this.currentDataLabels.delete = parseXsdBoolean(attrs.val) ?? true;
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
        // axId in axis context (vs chart type group).
        // A malformed / missing `val` leaves `currentAxis.axId = 0`
        // from the axis initialiser; refusing to overwrite with NaN
        // preserves that sentinel rather than poisoning the chart.
        if (!this.currentChartTypeGroup) {
          const v = parseXsdInt(attrs.val);
          if (v !== undefined) {
            this.currentAxis.axId = v;
          }
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
          this.currentAxis.scaling.max = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:min":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.min = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:logBase":
        if (this.currentAxis.scaling) {
          this.currentAxis.scaling.logBase = parseXsdFloat(attrs.val) ?? 0;
        }
        break;
      case "c:delete":
        if (!this.currentLegendEntry && !this.currentDataLabelEntry) {
          this.currentAxis.delete = parseXsdBoolean(attrs.val) ?? true;
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
        this.currentAxis.majorTickMark = parseTickMarkFromOoxml(attrs.val);
        break;
      case "c:minorTickMark":
        this.currentAxis.minorTickMark = parseTickMarkFromOoxml(attrs.val);
        break;
      case "c:tickLblPos":
        this.currentAxis.tickLblPos = attrs.val;
        break;
      case "c:crossAx":
        this.currentAxis.crossAx = parseXsdInt(attrs.val) ?? 0;
        break;
      case "c:crosses":
        this.currentAxis.crosses = attrs.val;
        break;
      case "c:crossesAt":
        // `c:crossesAt` has no schema default; when absent the axis
        // uses Excel's `c:crosses` choice ("autoZero"/"min"/"max")
        // instead. Forcing `0` when `val` was missing changed semantics
        // ("axis crosses at 0") and leaked to round-trip XML as
        // `val="0"`. Leave undefined so the writer omits the element.
        this.currentAxis.crossesAt = parseXsdFloat(attrs.val);
        break;
      case "c:auto":
        // CT_Boolean — absent `val` attribute means `true` per ECMA-376.
        this.currentAxis.auto = parseXsdBoolean(attrs.val) ?? true;
        break;
      case "c:lblAlgn":
        this.currentAxis.lblAlgn = attrs.val;
        break;
      case "c:lblOffset":
        // Schema default is `100` (percent). Leaving undefined so the
        // writer omits the element when the source did; forcing `0`
        // semantically meant "no offset" but round-tripped `val="0"`,
        // which Excel interprets DIFFERENTLY from an omitted element.
        this.currentAxis.lblOffset = parseXsdInt(attrs.val);
        break;
      case "c:tickLblSkip":
        // Schema default is `1` (show every tick). Similar round-trip
        // fidelity reasoning — preserve "attribute was absent".
        this.currentAxis.tickLblSkip = parseXsdInt(attrs.val);
        break;
      case "c:tickMarkSkip":
        this.currentAxis.tickMarkSkip = parseXsdInt(attrs.val);
        break;
      case "c:noMultiLvlLbl":
        this.currentAxis.noMultiLvlLbl = parseXsdBoolean(attrs.val) ?? true;
        break;
      case "c:crossBetween":
        this.currentAxis.crossBetween = attrs.val;
        break;
      case "c:majorUnit":
        // No schema default — omitting the element means "auto unit".
        // Writing `<c:majorUnit val="0"/>` produces a divide-by-zero
        // condition in tick positioning on the consumer side; leave
        // undefined so the writer omits the element.
        this.currentAxis.majorUnit = parseXsdFloat(attrs.val);
        break;
      case "c:minorUnit":
        this.currentAxis.minorUnit = parseXsdFloat(attrs.val);
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
      const perPointFormat: string | undefined = ptCtx.formatCode;
      // Check numLit/strLit/lvl before numCache/strCache
      if (this.currentLvl) {
        this.currentLvl.points.push({ index: ptCtx.index, value: text });
      } else if (this.currentNumLit) {
        // `parseFloat` returns NaN for blank / non-numeric text. Treat
        // anything we can't turn into a finite number as `null`
        // (OOXML's "blank cell" sentinel in a numeric cache) — keeps
        // the model free of NaN that would later round-trip as the
        // literal string "NaN".
        const n = text ? parseFloat(text) : NaN;
        this.currentNumLit.points.push({
          index: ptCtx.index,
          value: Number.isFinite(n) ? n : null
        });
      } else if (this.currentStrLit) {
        this.currentStrLit.points.push({ index: ptCtx.index, value: text });
      } else if (this.currentNumCache) {
        const n = text ? parseFloat(text) : NaN;
        // Propagate per-point `formatCode` captured from the `<c:pt>`
        // open handler. Only set it when present so hand-built caches
        // without per-point overrides stay clean (the default
        // `cache.formatCode` applies downstream).
        const point: NumberCache["points"][number] = {
          index: ptCtx.index,
          value: Number.isFinite(n) ? n : null
        };
        if (perPointFormat) {
          point.formatCode = perPointFormat;
        }
        this.currentNumCache.points.push(point);
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
      // Error-bar custom values can also be numLit (literal list) rather
      // than numRef (sheet reference). `_attachNumRef` handles the
      // reference case; without the mirrored branches here the literal
      // form was silently dropped on parse.
      case "c:plus":
        if (this.currentErrorBars) {
          this.currentErrorBars.plus = { numLit: this.currentNumLit };
        }
        break;
      case "c:minus":
        if (this.currentErrorBars) {
          this.currentErrorBars.minus = { numLit: this.currentNumLit };
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
      // Route the parsed txPr object to the nearest enclosing model
      // slot. Order matters: inner contexts (title, trendlineLbl,
      // dataLabelEntry, dataLabels, dataTable, pivotFmt) win over
      // outer ones (axis, legendEntry, legend, chartSpace). Missing
      // `currentPivotFormat` used to let a pivotFmt-internal `c:txPr`
      // fall through all the way to `chartModel.txPr`, silently
      // stealing chart-space-level typography.
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
      } else if (this.currentPivotFormat) {
        this.currentPivotFormat.txPr = txPrObj;
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
            // Strip the c15:datalabelsRange extension out into a structured
            // slot so editors can mutate "Value From Cells" without string
            // surgery; keep the rest of the extLst raw for round-trip of
            // other extensions.
            this.currentDataLabels.extLst = this._extractDataLabelsExtLst(
              rawXml,
              this.currentDataLabels
            );
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
        case "pivotFormat":
          if (this.currentPivotFormat) {
            this.currentPivotFormat.extLst = rawXml;
          }
          break;
        case "displayUnits":
          if (this.currentDisplayUnits) {
            this.currentDisplayUnits.extLst = rawXml;
          }
          break;
        case "upDownBars":
          if (this.currentUpDownBars) {
            this.currentUpDownBars.extLst = rawXml;
          }
          break;
        case "view3D":
          if (this.currentView3D) {
            this.currentView3D.extLst = rawXml;
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
          // chartSpace-level extLst: Parse and strip the MS `c14:pivotOptions`
          // extension into the structured `pivotOptions` slot so editors can
          // mutate it, and keep the remaining raw XML for everything else
          // (c15/c16 extensions we don't model structurally, vendor ext, …).
          this.chartModel.extLst = this._extractChartSpaceExtLst(rawXml);
          break;
      }
      this.rawCaptureContext = null;
    }
  }

  /**
   * Parse the MS `c14:pivotOptions` extension out of a captured
   * chartSpace-level `<c:extLst>` raw XML blob and populate
   * {@link ChartModel.pivotOptions} with the structured view.
   *
   * Returns the raw XML with the matching `<c:ext uri="{781A…}" …/>`
   * element removed so the writer does not emit it twice when it
   * reconstructs the extLst from the merged (raw + structured) state.
   * If the extension is absent, the input is returned unchanged.
   *
   * This function uses regex rather than a full XML parse because the
   * raw blob is already guaranteed well-formed (it was emitted by our
   * own raw-capture path) and we need to preserve byte exactness of the
   * surviving `<c:ext>` elements for round-trip of extensions we do not
   * understand.
   */
  private _extractChartSpaceExtLst(rawXml: string): string {
    // Match a single `<c:ext ... uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}" ... >…</c:ext>`
    // in a case-insensitive way (Excel has been seen to emit mixed case
    // GUIDs historically). The URI must appear anywhere in the opening
    // tag's attributes, so we anchor the regex to the opening `<c:ext`
    // and look for the URI attribute inside the same tag.
    let stripped = rawXml;
    stripped = this._extractAndStripPivotExt(stripped, C14_PIVOT_OPTIONS_EXT_URI, match => {
      const options = parseC14PivotOptions(match);
      if (options) {
        // Merge into any options already harvested from a sibling ext.
        this.chartModel.pivotOptions = { ...this.chartModel.pivotOptions, ...options };
      }
    });
    stripped = this._extractAndStripPivotExt(stripped, C16_PIVOT_OPTIONS16_EXT_URI, match => {
      const options = parseC16PivotOptions16(match);
      if (options) {
        this.chartModel.pivotOptions = { ...this.chartModel.pivotOptions, ...options };
      }
    });

    // If stripping left an empty `<c:extLst></c:extLst>` or
    // `<c:extLst/>`, drop the entire wrapper so the writer doesn't emit
    // a noise element.
    const emptyExtLst = /^\s*<c:extLst\b[^>]*>\s*<\/c:extLst>\s*$|^\s*<c:extLst\b[^/]*\/>\s*$/;
    return emptyExtLst.test(stripped) ? "" : stripped;
  }

  /**
   * Locate and remove a single `<c:ext uri="{…}">…</c:ext>` element
   * whose `uri` matches `extUri`, passing the captured XML to
   * `handler`. Returns the source string with the ext removed. When the
   * ext is absent the input is returned unchanged.
   */
  private _extractAndStripPivotExt(
    rawXml: string,
    extUri: string,
    handler: (match: string) => void
  ): string {
    const uriEscaped = extUri.replace(/[{}]/g, ch => `\\${ch}`);
    const extRegex = new RegExp(
      `<c:ext\\b[^>]*\\buri="${uriEscaped}"[^>]*>[\\s\\S]*?</c:ext>`,
      "i"
    );
    const match = extRegex.exec(rawXml);
    if (!match) {
      return rawXml;
    }
    handler(match[0]);
    return rawXml.slice(0, match.index) + rawXml.slice(match.index + match[0].length);
  }

  /**
   * Parse and strip the MS `c15:datalabelsRange` extension (Excel 2013+
   * "Value From Cells") from a captured `<c:dLbls>/<c:extLst>` raw XML
   * blob, populating the structured {@link DataLabels.dataLabelsRange}
   * slot on the given DataLabels object.
   *
   * Returns the raw XML with the matching ext element removed so the
   * writer does not emit it twice when it reconstructs the extLst from
   * the merged (raw + structured) state. If the extension is absent the
   * input is returned unchanged.
   */
  private _extractDataLabelsExtLst(rawXml: string, dl: DataLabels): string {
    const uriEscaped = C15_DATA_LABELS_RANGE_EXT_URI.replace(/[{}]/g, ch => `\\${ch}`);
    const extRegex = new RegExp(
      `<c:ext\\b[^>]*\\buri="${uriEscaped}"[^>]*>[\\s\\S]*?</c:ext>`,
      "i"
    );
    const match = extRegex.exec(rawXml);
    if (!match) {
      return rawXml;
    }
    const range = parseC15DataLabelsRange(match[0]);
    if (range) {
      dl.dataLabelsRange = range;
    }
    const stripped = rawXml.slice(0, match.index) + rawXml.slice(match.index + match[0].length);
    const emptyExtLst = /^\s*<c:extLst\b[^>]*>\s*<\/c:extLst>\s*$|^\s*<c:extLst\b[^/]*\/>\s*$/;
    return emptyExtLst.test(stripped) ? "" : stripped;
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
