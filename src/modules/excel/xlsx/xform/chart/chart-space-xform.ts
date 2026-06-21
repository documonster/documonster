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
  NumberLiteral,
  StringLiteral,
  MultiLevelStringReference,
  ShapeProperties,
  ChartTextProperties,
  PrintSettings,
  PivotChartOptions,
  PivotFormat,
  DataLabelsRange,
  UpDownBars,
  DisplayUnits,
  LegendEntry,
  PictureOptions
} from "@excel/chart/model/types";
import { escapeXml, escapeXmlAttr } from "@excel/chart/shared/chart-utils";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import {
  C14_PIVOT_OPTIONS_EXT_URI,
  C15_DATA_LABELS_RANGE_EXT_URI,
  C16_PIVOT_OPTIONS16_EXT_URI
} from "@excel/xlsx/xform/chart/chart-space-constants";
import { renderChartSpace } from "@excel/xlsx/xform/chart/chart-space-render";
import {
  parseXsdBoolean as sharedParseXsdBoolean,
  parseXsdInt as sharedParseXsdInt,
  parseXsdFloat as sharedParseXsdFloat
} from "@excel/xlsx/xform/xsd-values";
import type { ParseOpenTag, XmlSink } from "@xml/types";

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
// `@excel/chart/shared/chart-utils`. They're the single authoritative entry
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
 *
 * Render-side only — moved to `./chart-space-render` alongside the writers.
 */

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
 * Inverse of the render-side `tickMarkToOoxml` (moved to
 * `./chart-space-render`) for the parse side.
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

// MS Office extension URIs (`C14_PIVOT_OPTIONS_EXT_URI`,
// `C16_PIVOT_OPTIONS16_EXT_URI`, `C15_DATA_LABELS_RANGE_EXT_URI`) and their
// namespaces are defined in `./chart-space-render` and shared with the
// renderer; the parse side imports the URIs it needs at the top of this file.

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
    renderChartSpace(xmlStream, m);
  }

  // ============================================================================
  // PARSE
  // ============================================================================

  parseOpen(node: ParseOpenTag): boolean {
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
        // Self-closing raw-preserve elements (e.g. `<c:extLst/>`,
        // `<c:protection/>`) must still be captured and attached so
        // they survive round-trip. Build the self-closing tag inline
        // and route through `_attachRawXml` like the normal path.
        let selfCloseXml = `<${node.name}`;
        if (node.attributes) {
          for (const [k, v] of Object.entries(node.attributes)) {
            selfCloseXml += ` ${k}="${escapeXmlAttr(String(v))}"`;
          }
        }
        selfCloseXml += "/>";
        // Determine the correct rawXmlTarget for routing (mirrors the
        // non-self-closing path below). For `c:tx` we need the
        // context-qualified target name.
        let selfCloseTarget: string = name;
        if (name === "c:tx" && this.currentTitle && !this.currentSeries) {
          selfCloseTarget = "c:tx:title";
        } else if (name === "c:tx" && this.currentDataLabelEntry != null) {
          selfCloseTarget = "c:tx:dLblEntry";
        } else if (name === "c:tx" && this.currentTrendlineLbl != null) {
          selfCloseTarget = "c:tx:trendlineLbl";
        }
        this._attachRawXml(selfCloseTarget, selfCloseXml);
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
      // c:dLbl inside pivotFmt is parsed structurally — see the
      // `currentDataLabelEntry` branch below. `c:dLbl` open always
      // starts an entry, and the matching close routes it to the
      // pivotFormat or the data-labels bundle depending on the
      // surrounding state.
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
            b: parseXsdFloat(attrs.b) ?? 0.75,
            l: parseXsdFloat(attrs.l) ?? 0.7,
            r: parseXsdFloat(attrs.r) ?? 0.7,
            t: parseXsdFloat(attrs.t) ?? 0.75,
            header: parseXsdFloat(attrs.header) ?? 0.3,
            footer: parseXsdFloat(attrs.footer) ?? 0.3
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
          // captures it.
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
}

export { ChartSpaceXform };
