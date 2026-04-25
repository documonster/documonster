/**
 * Sparkline (mini chart) data model and utilities.
 *
 * Sparklines are cell-sized charts stored in the worksheet extLst under
 * `x14:sparklineGroups`. Each group contains:
 *   - display options (type, line weight, markers, axis settings, colors)
 *   - one or more sparklines, each pairing a data reference with an anchor cell.
 *
 * Reference: ECMA-376 §18.18.92 + Office Open XML extension `x14` namespace.
 */

/**
 * Top-level sparkline group — matches `x14:sparklineGroup`.
 */
export interface SparklineGroup {
  /** Chart type: line | column | stacked (win-loss) */
  type?: SparklineType;
  /** Line weight in points (0.25 - 2.25) */
  lineWeight?: number;
  /** Display empty cells as: gap, zero, or span */
  displayEmptyCellsAs?: "gap" | "zero" | "span";
  /** Whether to display markers */
  markers?: boolean;
  /** High point marker */
  high?: boolean;
  /** Low point marker */
  low?: boolean;
  /** First point marker */
  first?: boolean;
  /** Last point marker */
  last?: boolean;
  /** Negative point marker */
  negative?: boolean;
  /** Display X axis */
  displayXAxis?: boolean;
  /** Display hidden cells data */
  displayHidden?: boolean;
  /** Min axis type: individual, group, custom */
  minAxisType?: SparklineAxisType;
  /** Max axis type */
  maxAxisType?: SparklineAxisType;
  /** Manual min (when minAxisType === "custom") */
  manualMin?: number;
  /** Manual max (when maxAxisType === "custom") */
  manualMax?: number;
  /** Right-to-left */
  rightToLeft?: boolean;
  /** Color series */
  colorSeries?: SparklineColor;
  /** Color negative */
  colorNegative?: SparklineColor;
  /** Color axis */
  colorAxis?: SparklineColor;
  /** Color markers */
  colorMarkers?: SparklineColor;
  /** Color first */
  colorFirst?: SparklineColor;
  /** Color last */
  colorLast?: SparklineColor;
  /** Color high */
  colorHigh?: SparklineColor;
  /** Color low */
  colorLow?: SparklineColor;
  /** Date axis source (reference range of dates) */
  dateAxis?: string;
  /** Sparklines in this group */
  sparklines: Sparkline[];
}

export type SparklineType = "line" | "column" | "stacked";

export type SparklineAxisType = "individual" | "group" | "custom";

/**
 * A single sparkline within a group.
 */
export interface Sparkline {
  /** Data reference (e.g. "Sheet1!B2:G2") */
  dataRef: string;
  /** Anchor cell reference (e.g. "H2") */
  cellRef: string;
}

/**
 * Sparkline color — theme reference or sRGB.
 */
export interface SparklineColor {
  /** Theme index (0-11) */
  theme?: number;
  /** sRGB hex */
  rgb?: string;
  /** Tint (-1 to 1) */
  tint?: number;
  /** Auto color */
  auto?: boolean;
}

// ============================================================================
// High-level options for Worksheet.addSparklineGroup
// ============================================================================

/**
 * High-level options for creating a sparkline group.
 */
export interface AddSparklineGroupOptions {
  /** Chart type */
  type: SparklineType;
  /** List of sparklines (data + anchor) */
  sparklines: Sparkline[];
  /** Line weight in points */
  lineWeight?: number;
  /** Show markers (line type) */
  markers?: boolean;
  /** Show high point */
  high?: boolean;
  /** Show low point */
  low?: boolean;
  /** Show first point */
  first?: boolean;
  /** Show last point */
  last?: boolean;
  /** Show negative points */
  negative?: boolean;
  /** Line color */
  lineColor?: string;
  /** Negative bar color */
  negativeColor?: string;
  /** High marker color */
  highColor?: string;
  /** Low marker color */
  lowColor?: string;
  /** First marker color */
  firstColor?: string;
  /** Last marker color */
  lastColor?: string;
  /** Min axis type */
  minAxisType?: SparklineAxisType;
  /** Max axis type */
  maxAxisType?: SparklineAxisType;
  /** Manual min */
  manualMin?: number;
  /** Manual max */
  manualMax?: number;
  /** Show X axis */
  displayXAxis?: boolean;
  /** Right-to-left */
  rightToLeft?: boolean;
  /** Display empty cells as */
  displayEmptyCellsAs?: "gap" | "zero" | "span";
  /** Date axis source */
  dateAxis?: string;
}

/**
 * Build a SparklineGroup from simplified options.
 */
export function buildSparklineGroup(opts: AddSparklineGroupOptions): SparklineGroup {
  const group: SparklineGroup = {
    type: opts.type,
    sparklines: opts.sparklines
  };
  if (opts.lineWeight !== undefined) {
    group.lineWeight = opts.lineWeight;
  }
  if (opts.markers !== undefined) {
    group.markers = opts.markers;
  }
  if (opts.high !== undefined) {
    group.high = opts.high;
  }
  if (opts.low !== undefined) {
    group.low = opts.low;
  }
  if (opts.first !== undefined) {
    group.first = opts.first;
  }
  if (opts.last !== undefined) {
    group.last = opts.last;
  }
  if (opts.negative !== undefined) {
    group.negative = opts.negative;
  }
  if (opts.lineColor) {
    group.colorSeries = hexToSparklineColor(opts.lineColor);
  }
  if (opts.negativeColor) {
    group.colorNegative = hexToSparklineColor(opts.negativeColor);
  }
  if (opts.highColor) {
    group.colorHigh = hexToSparklineColor(opts.highColor);
  }
  if (opts.lowColor) {
    group.colorLow = hexToSparklineColor(opts.lowColor);
  }
  if (opts.firstColor) {
    group.colorFirst = hexToSparklineColor(opts.firstColor);
  }
  if (opts.lastColor) {
    group.colorLast = hexToSparklineColor(opts.lastColor);
  }
  if (opts.minAxisType) {
    group.minAxisType = opts.minAxisType;
  }
  if (opts.maxAxisType) {
    group.maxAxisType = opts.maxAxisType;
  }
  if (opts.manualMin !== undefined) {
    group.manualMin = opts.manualMin;
  }
  if (opts.manualMax !== undefined) {
    group.manualMax = opts.manualMax;
  }
  if (opts.displayXAxis !== undefined) {
    group.displayXAxis = opts.displayXAxis;
  }
  if (opts.rightToLeft !== undefined) {
    group.rightToLeft = opts.rightToLeft;
  }
  if (opts.displayEmptyCellsAs) {
    group.displayEmptyCellsAs = opts.displayEmptyCellsAs;
  }
  if (opts.dateAxis) {
    group.dateAxis = opts.dateAxis;
  }
  return group;
}

function hexToSparklineColor(hex: string): SparklineColor {
  return { rgb: hex.replace(/^#/, "").toUpperCase() };
}

// ============================================================================
// XML rendering
// ============================================================================

/**
 * Render all sparkline groups on a worksheet to an x14:sparklineGroups
 * XML fragment. Returns a string (empty if no groups).
 */
export function renderSparklineGroups(groups: SparklineGroup[]): string {
  if (!groups || groups.length === 0) {
    return "";
  }
  const parts: string[] = [];
  parts.push(
    '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
  );
  for (const g of groups) {
    parts.push(renderSparklineGroup(g));
  }
  parts.push("</x14:sparklineGroups>");
  return parts.join("");
}

function renderSparklineGroup(g: SparklineGroup): string {
  const attrs: string[] = [];
  if (g.type !== undefined && g.type !== "line") {
    attrs.push(`type="${g.type}"`);
  }
  if (g.lineWeight !== undefined) {
    attrs.push(`lineWeight="${g.lineWeight}"`);
  }
  if (g.displayEmptyCellsAs) {
    attrs.push(`displayEmptyCellsAs="${g.displayEmptyCellsAs}"`);
  }
  if (g.markers) {
    attrs.push('markers="1"');
  }
  if (g.high) {
    attrs.push('high="1"');
  }
  if (g.low) {
    attrs.push('low="1"');
  }
  if (g.first) {
    attrs.push('first="1"');
  }
  if (g.last) {
    attrs.push('last="1"');
  }
  if (g.negative) {
    attrs.push('negative="1"');
  }
  if (g.displayXAxis) {
    attrs.push('displayXAxis="1"');
  }
  if (g.displayHidden) {
    attrs.push('displayHidden="1"');
  }
  if (g.minAxisType && g.minAxisType !== "individual") {
    attrs.push(`minAxisType="${g.minAxisType}"`);
  }
  if (g.maxAxisType && g.maxAxisType !== "individual") {
    attrs.push(`maxAxisType="${g.maxAxisType}"`);
  }
  if (g.manualMin !== undefined) {
    attrs.push(`manualMin="${g.manualMin}"`);
  }
  if (g.manualMax !== undefined) {
    attrs.push(`manualMax="${g.manualMax}"`);
  }
  if (g.rightToLeft) {
    attrs.push('rightToLeft="1"');
  }

  const parts: string[] = [];
  parts.push(`<x14:sparklineGroup${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}>`);

  // Colors (emit only those set, in OOXML order)
  if (g.colorSeries) {
    parts.push(`<x14:colorSeries ${sparklineColorAttrs(g.colorSeries)}/>`);
  }
  if (g.colorNegative) {
    parts.push(`<x14:colorNegative ${sparklineColorAttrs(g.colorNegative)}/>`);
  }
  if (g.colorAxis) {
    parts.push(`<x14:colorAxis ${sparklineColorAttrs(g.colorAxis)}/>`);
  }
  if (g.colorMarkers) {
    parts.push(`<x14:colorMarkers ${sparklineColorAttrs(g.colorMarkers)}/>`);
  }
  if (g.colorFirst) {
    parts.push(`<x14:colorFirst ${sparklineColorAttrs(g.colorFirst)}/>`);
  }
  if (g.colorLast) {
    parts.push(`<x14:colorLast ${sparklineColorAttrs(g.colorLast)}/>`);
  }
  if (g.colorHigh) {
    parts.push(`<x14:colorHigh ${sparklineColorAttrs(g.colorHigh)}/>`);
  }
  if (g.colorLow) {
    parts.push(`<x14:colorLow ${sparklineColorAttrs(g.colorLow)}/>`);
  }
  if (g.dateAxis) {
    parts.push(`<xm:f>${escapeXml(g.dateAxis)}</xm:f>`);
  }

  // Sparklines
  parts.push("<x14:sparklines>");
  for (const s of g.sparklines) {
    parts.push("<x14:sparkline>");
    parts.push(`<xm:f>${escapeXml(s.dataRef)}</xm:f>`);
    parts.push(`<xm:sqref>${escapeXml(s.cellRef)}</xm:sqref>`);
    parts.push("</x14:sparkline>");
  }
  parts.push("</x14:sparklines>");
  parts.push("</x14:sparklineGroup>");
  return parts.join("");
}

function sparklineColorAttrs(c: SparklineColor): string {
  const parts: string[] = [];
  if (c.rgb) {
    parts.push(`rgb="${c.rgb}"`);
  }
  if (c.theme !== undefined) {
    parts.push(`theme="${c.theme}"`);
  }
  if (c.tint !== undefined) {
    parts.push(`tint="${c.tint}"`);
  }
  if (c.auto) {
    parts.push('auto="1"');
  }
  return parts.join(" ");
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ============================================================================
// XML parsing (best-effort from raw XML fragment)
// ============================================================================

/**
 * Parse an x14:sparklineGroups XML fragment into structured groups.
 * Best-effort regex-based parser — sufficient for round-trip via rebuild.
 */
export function parseSparklineGroups(xml: string): SparklineGroup[] {
  const groups: SparklineGroup[] = [];
  const groupRe = /<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g;
  let m: RegExpExecArray | null;
  while ((m = groupRe.exec(xml)) !== null) {
    const g = parseGroupBlock(m[1], m[2]);
    groups.push(g);
  }
  return groups;
}

// Pre-compiled regexes for sparkline color tag parsing
const COLOR_TAG_RES = new Map<string, RegExp>([
  ["colorSeries", /<x14:colorSeries\b([^/]*)\/>/],
  ["colorNegative", /<x14:colorNegative\b([^/]*)\/>/],
  ["colorAxis", /<x14:colorAxis\b([^/]*)\/>/],
  ["colorMarkers", /<x14:colorMarkers\b([^/]*)\/>/],
  ["colorFirst", /<x14:colorFirst\b([^/]*)\/>/],
  ["colorLast", /<x14:colorLast\b([^/]*)\/>/],
  ["colorHigh", /<x14:colorHigh\b([^/]*)\/>/],
  ["colorLow", /<x14:colorLow\b([^/]*)\/>/]
]);

function parseGroupBlock(attrXml: string, inner: string): SparklineGroup {
  const g: SparklineGroup = { sparklines: [] };
  const attrRe = /(\w+)="([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(attrXml)) !== null) {
    const [, name, val] = m;
    switch (name) {
      case "type":
        g.type = val as SparklineType;
        break;
      case "lineWeight":
        g.lineWeight = parseFloat(val);
        break;
      case "displayEmptyCellsAs":
        g.displayEmptyCellsAs = val as NonNullable<SparklineGroup["displayEmptyCellsAs"]>;
        break;
      case "markers":
        g.markers = val === "1";
        break;
      case "high":
        g.high = val === "1";
        break;
      case "low":
        g.low = val === "1";
        break;
      case "first":
        g.first = val === "1";
        break;
      case "last":
        g.last = val === "1";
        break;
      case "negative":
        g.negative = val === "1";
        break;
      case "displayXAxis":
        g.displayXAxis = val === "1";
        break;
      case "displayHidden":
        g.displayHidden = val === "1";
        break;
      case "minAxisType":
        g.minAxisType = val as SparklineAxisType;
        break;
      case "maxAxisType":
        g.maxAxisType = val as SparklineAxisType;
        break;
      case "manualMin":
        g.manualMin = parseFloat(val);
        break;
      case "manualMax":
        g.manualMax = parseFloat(val);
        break;
      case "rightToLeft":
        g.rightToLeft = val === "1";
        break;
    }
  }
  // Parse colors
  const colorTags: Array<keyof SparklineGroup> = [
    "colorSeries",
    "colorNegative",
    "colorAxis",
    "colorMarkers",
    "colorFirst",
    "colorLast",
    "colorHigh",
    "colorLow"
  ];
  for (const tag of colorTags) {
    const re = COLOR_TAG_RES.get(tag as string)!;
    const cm = re.exec(inner);
    if (cm) {
      (g as any)[tag] = parseColorAttrs(cm[1]);
    }
  }
  // Parse sparklines
  const sparkRe =
    /<x14:sparkline>\s*<xm:f>([\s\S]*?)<\/xm:f>\s*<xm:sqref>([\s\S]*?)<\/xm:sqref>\s*<\/x14:sparkline>/g;
  let sm: RegExpExecArray | null;
  while ((sm = sparkRe.exec(inner)) !== null) {
    g.sparklines.push({
      dataRef: decodeXml(sm[1]),
      cellRef: decodeXml(sm[2])
    });
  }
  return g;
}

function parseColorAttrs(attrXml: string): SparklineColor {
  const c: SparklineColor = {};
  const attrRe = /(\w+)="([^"]*)"/g;
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(attrXml)) !== null) {
    const [, name, val] = m;
    switch (name) {
      case "rgb":
        c.rgb = val;
        break;
      case "theme":
        c.theme = parseInt(val, 10);
        break;
      case "tint":
        c.tint = parseFloat(val);
        break;
      case "auto":
        c.auto = val === "1";
        break;
    }
  }
  return c;
}

function decodeXml(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
