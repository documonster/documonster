/**
 * Chart - High-level chart object for worksheet embedding.
 *
 * Similar to Image, a Chart belongs to a Worksheet and carries the
 * structural data needed for both the DrawingML anchor and the
 * standalone chart XML part.
 */

import { Anchor, type AnchorModel } from "@excel/anchor";
import type { ChartExModel } from "@excel/chart/chart-ex-types";
import type {
  AddChartRange,
  ChartModel,
  ChartStyleModel,
  ChartColorsModel,
  ChartTitle,
  ChartTypeGroup,
  ChartAxis,
  ChartLegend,
  PlotArea,
  ShapeProperties,
  ChartRichText,
  ChartTextProperties,
  SeriesBase
} from "@excel/chart/types";
import { colCache } from "@excel/utils/col-cache";
import type { Worksheet } from "@excel/worksheet";

import { buildChartModel } from "./chart-builder";

// Pre-compiled regexes for _extractTextFromRawTx
const RAW_TX_AT_RE = /<a:t>([\s\S]*?)<\/a:t>/g;
const RAW_TX_CV_RE = /<c:v>([\s\S]*?)<\/c:v>/g;

/**
 * Internal model stored on the Workbook for each chart.
 * Contains the fully-parsed (or programmatically created) chart data
 * plus ancillary style/colors files.
 */
export interface ChartEntry {
  /** 1-based chart number (matches chart{N}.xml) */
  chartNumber: number;
  /** Full chart model */
  model: ChartModel;
  /** Chart style (styleN.xml) — raw XML for round-trip */
  style?: ChartStyleModel;
  /** Chart colors (colorsN.xml) — raw XML for round-trip */
  colors?: ChartColorsModel;
  /** Chart rels (chart{N}.xml.rels) — raw entries for round-trip */
  rels?: any[];
}

/**
 * Stored entry for a structured ChartEx (Office 2016+ extended chart).
 * When a ChartEx is created programmatically via `addChartEx()`, a structured
 * model is stored here and serialised through the builder/renderer on write.
 * When a ChartEx is round-tripped, raw bytes are used instead (stored under
 * `workbook._chartExEntries`).
 */
export interface ChartExEntry {
  /** 1-based chartEx number (matches chartEx{N}.xml) */
  chartExNumber: number;
  /** Structured model (built from addChartEx options) */
  model: ChartExModel;
  /** ChartEx rels — preserved for round-trip */
  rels?: any[];
}

/**
 * The range a chart occupies on a worksheet.
 */
interface ChartAnchorRange {
  /** Top-left anchor (always present) */
  tl: Anchor;
  /** Bottom-right anchor (only for twoCellAnchor) */
  br?: Anchor;
  /** Absolute position in EMU (only for absoluteAnchor) */
  pos?: { x: number; y: number };
  /** Extent in EMU (for oneCellAnchor and absoluteAnchor) */
  ext?: { cx: number; cy: number };
  /** Anchor behaviour: oneCell, twoCell, or absolute */
  editAs?: string;
}

/**
 * A chart embedded in a worksheet.
 *
 * Charts come in two flavours:
 * - **Classic** (`chartNumber` set): fully parsed c:chart with a ChartModel.
 * - **ChartEx** (`chartExNumber` set): Office 2016+ cx:chart stored as raw bytes.
 *   The high-level accessors (`chartModel`, `chartTypes`, `axes`, etc.) return
 *   `undefined` / empty for chartEx charts because the data is not parsed.
 */
class Chart {
  readonly worksheet: Worksheet;
  /** 1-based chart number for classic c:chart (0 when this is a chartEx) */
  chartNumber: number;
  /** 1-based chartEx number for cx:chart (0 when this is a classic chart) */
  chartExNumber: number;
  range: ChartAnchorRange;

  constructor(
    worksheet: Worksheet,
    ids: { chartNumber?: number; chartExNumber?: number },
    range: AddChartRange | ChartAnchorModel["range"]
  ) {
    this.worksheet = worksheet;
    this.chartNumber = ids.chartNumber ?? 0;
    this.chartExNumber = ids.chartExNumber ?? 0;
    this.range = Chart.parseRange(worksheet, range);
  }

  /** Whether this is an Office 2016+ extended chart (cx:chart) */
  get isChartEx(): boolean {
    return this.chartExNumber > 0;
  }

  private static parseRange(
    worksheet: Worksheet,
    range: AddChartRange | ChartAnchorModel["range"]
  ): ChartAnchorRange {
    if (typeof range === "string") {
      const decoded = colCache.decode(range);
      if ("top" in decoded) {
        return {
          tl: new Anchor(worksheet, { col: decoded.left, row: decoded.top }, -1),
          br: new Anchor(worksheet, { col: decoded.right, row: decoded.bottom }, 0),
          editAs: "twoCell"
        };
      }
      // Single cell — default to 10 cols x 15 rows
      const addr = colCache.decodeAddress(range);
      return {
        tl: new Anchor(worksheet, { col: addr.col, row: addr.row }, -1),
        br: new Anchor(worksheet, { col: addr.col + 10, row: addr.row + 15 }, 0),
        editAs: "twoCell"
      };
    }

    const parseAnchor = (v: { col: number; row: number } | AnchorModel | string): Anchor => {
      if (typeof v === "string") {
        return new Anchor(worksheet, v, -1);
      }
      return new Anchor(worksheet, v, 0);
    };

    // Absolute anchor: { pos, ext }
    if ("pos" in range && range.pos !== undefined) {
      return {
        // Absolute anchors have no meaningful "tl" cell, but the drawing layer
        // still requires one — default to (0, 0) with the provided offsets.
        tl: new Anchor(worksheet, { col: 0, row: 0 }, 0),
        pos: range.pos,
        ext: range.ext,
        editAs: range.editAs ?? "absolute"
      };
    }

    // One-cell anchor: { tl, ext } (has tl, no br)
    if ("ext" in range && range.ext !== undefined && "tl" in range && !("br" in range)) {
      return {
        tl: parseAnchor(range.tl),
        ext: range.ext,
        editAs: range.editAs ?? "oneCell"
      };
    }

    // Two-cell anchor: { tl, br }
    const twoCell = range as {
      tl: { col: number; row: number } | AnchorModel | string;
      br: { col: number; row: number } | AnchorModel | string;
      editAs?: string;
    };
    return {
      tl: parseAnchor(twoCell.tl),
      br: parseAnchor(twoCell.br),
      editAs: twoCell.editAs ?? "twoCell"
    };
  }

  get model(): ChartAnchorModel {
    const base: ChartAnchorModel = {
      chartNumber: this.chartNumber,
      chartExNumber: this.chartExNumber,
      range: {
        tl: this.range.tl.model,
        br: this.range.br ? this.range.br.model : undefined,
        editAs: this.range.editAs,
        pos: this.range.pos,
        ext: this.range.ext
      } as ChartAnchorModel["range"]
    };
    return base;
  }

  // ===========================================================================
  // Chart data access
  // ===========================================================================

  /** Get the full ChartModel from the workbook's chart entries (classic charts only) */
  get chartModel(): ChartModel | undefined {
    if (this.chartNumber <= 0) {
      return undefined;
    }
    return this.worksheet.workbook.getChartEntry(this.chartNumber)?.model;
  }

  /** Get the chart title text. Returns undefined if no title. */
  get title(): string | undefined {
    const titleObj = this.chartModel?.chart?.title;
    if (!titleObj) {
      return undefined;
    }
    if (titleObj.text) {
      return this._extractTextFromRichText(titleObj.text);
    }
    if (titleObj.strRef?.cache?.points) {
      return titleObj.strRef.cache.points.map(p => p.value).join("");
    }
    // Fall back to rawTx — extract text from <a:t> elements
    if (titleObj.rawTx) {
      return this._extractTextFromRawTx(titleObj.rawTx);
    }
    return undefined;
  }

  /**
   * Set the chart title.
   *
   * Accepts:
   * - `undefined` → removes title (sets autoTitleDeleted).
   * - `string` → sets title text. If the existing title had rich-text formatting,
   *   the formatting of the first run is preserved and only the text is replaced.
   * - `ChartRichText` → replaces the title with fully-structured rich text.
   * - `{ formula: string }` → sets the title to a worksheet formula reference.
   */
  set title(value: string | ChartRichText | { formula: string } | undefined) {
    const cm = this.chartModel;
    if (!cm) {
      return;
    }
    if (value === undefined) {
      cm.chart.title = undefined;
      cm.chart.autoTitleDeleted = true;
      return;
    }
    if (typeof value === "object" && "formula" in value && typeof value.formula === "string") {
      cm.chart.title = {
        strRef: { formula: value.formula, cache: { points: [] } },
        overlay: cm.chart.title?.overlay ?? false
      };
      cm.chart.autoTitleDeleted = false;
      return;
    }
    if (typeof value === "object" && "paragraphs" in value) {
      // Rich text
      cm.chart.title = {
        ...(cm.chart.title ?? {}),
        strRef: undefined,
        rawTx: undefined,
        text: value,
        overlay: cm.chart.title?.overlay ?? false
      };
      cm.chart.autoTitleDeleted = false;
      return;
    }
    // Plain string — try to preserve first-run formatting of existing title.
    if (typeof value !== "string") {
      return; // Type-safety guard; other object shapes handled above.
    }
    const existing = cm.chart.title;
    const preservedProps = this._extractFirstRunProperties(existing);
    cm.chart.title = {
      ...(existing ?? {}),
      strRef: undefined,
      rawTx: undefined,
      text: {
        paragraphs: [
          {
            runs: [{ text: value, properties: preservedProps }]
          }
        ]
      },
      overlay: existing?.overlay ?? false
    };
    cm.chart.autoTitleDeleted = false;
  }

  /**
   * Set the chart title using structured rich text.
   * Convenience method equivalent to `chart.title = richText`.
   */
  setTitleRichText(richText: ChartRichText): void {
    this.title = richText;
  }

  /**
   * Get the structured rich text of the chart title.
   * Returns undefined if the title is unset, formula-only, or captured as rawTx.
   */
  get titleRichText(): ChartRichText | undefined {
    return this.chartModel?.chart?.title?.text;
  }

  /** Get the chart type groups in the plot area */
  get chartTypes(): ChartTypeGroup[] {
    return this.chartModel?.chart?.plotArea?.chartTypes ?? [];
  }

  /** Get the chart axes */
  get axes(): ChartAxis[] {
    return this.chartModel?.chart?.plotArea?.axes ?? [];
  }

  /** Get the chart legend */
  get legend(): ChartLegend | undefined {
    return this.chartModel?.chart?.legend;
  }

  /** Set the chart legend */
  set legend(value: ChartLegend | undefined) {
    const cm = this.chartModel;
    if (cm) {
      cm.chart.legend = value;
    }
  }

  /** Get the chart-level shape properties */
  get spPr(): ShapeProperties | undefined {
    return this.chartModel?.spPr;
  }

  /** Set the chart-level shape properties */
  set spPr(value: ShapeProperties | undefined) {
    const cm = this.chartModel;
    if (cm) {
      cm.spPr = value;
    }
  }

  /**
   * Get an axis by its ID.
   */
  getAxis(axId: number): ChartAxis | undefined {
    return this.axes.find(ax => ax.axId === axId);
  }

  /**
   * Get the category (X) axis, if any.
   */
  get categoryAxis(): ChartAxis | undefined {
    return this.axes.find(ax => ax.axisType === "cat" || ax.axisType === "date");
  }

  /**
   * Get the value (Y) axis, if any.
   */
  get valueAxis(): ChartAxis | undefined {
    return this.axes.find(ax => ax.axisType === "val");
  }

  /** Get the plot area for direct manipulation */
  get plotArea(): PlotArea | undefined {
    return this.chartModel?.chart?.plotArea;
  }

  /**
   * Add a series to a chart type group.
   *
   * @param series - The series object matching the expected series type for the chart.
   * @param groupIndex - 0-based index of the chart type group (for combo charts). Defaults to 0.
   */
  addSeries(series: SeriesBase, groupIndex = 0): void {
    const ctg = this.chartTypes[groupIndex];
    if (ctg) {
      ctg.series.push(series as any);
    }
  }

  /**
   * Remove a series from a chart type group by index.
   *
   * @param index - 0-based index of the series within the group.
   * @param groupIndex - 0-based index of the chart type group (for combo charts). Defaults to 0.
   * @returns The removed series, or undefined if out of range.
   */
  removeSeries(index: number, groupIndex = 0): SeriesBase | undefined {
    const ctg = this.chartTypes[groupIndex];
    if (!ctg || index < 0 || index >= ctg.series.length) {
      return undefined;
    }
    return ctg.series.splice(index, 1)[0];
  }

  /**
   * Get a series from a chart type group.
   *
   * @param index - 0-based index of the series within the group.
   * @param groupIndex - 0-based index of the chart type group (for combo charts). Defaults to 0.
   */
  getSeries(index: number, groupIndex = 0): SeriesBase | undefined {
    const ctg = this.chartTypes[groupIndex];
    if (!ctg) {
      return undefined;
    }
    return ctg.series[index];
  }

  /** Get the total number of series across all chart type groups. */
  get totalSeriesCount(): number {
    return this.chartTypes.reduce((sum, ctg) => sum + (ctg.series?.length ?? 0), 0);
  }

  /**
   * Get the number of series in a specific chart type group.
   *
   * @param groupIndex - 0-based index of the chart type group (defaults to 0).
   */
  getSeriesCount(groupIndex = 0): number {
    const ctg = this.chartTypes[groupIndex];
    return ctg?.series?.length ?? 0;
  }

  /**
   * @deprecated Use `getSeriesCount()` instead. Returns the count of series in the first chart type group.
   */
  get seriesCount(): number {
    return this.getSeriesCount(0);
  }

  // ===========================================================================
  // Chart cloning and duplication
  // ===========================================================================

  /**
   * Create a deep copy of this chart and add it to a target worksheet.
   * The new chart receives a fresh chartNumber; the original is untouched.
   *
   * @param targetWs - Worksheet to receive the clone. Defaults to this chart's worksheet.
   * @param range - Anchor range for the clone. Defaults to the original range.
   * @returns The new chartNumber in the target workbook.
   */
  copyTo(targetWs: Worksheet, range?: AddChartRange): number {
    const sourceModel = this.chartModel;
    if (!sourceModel) {
      throw new Error("Cannot copy a chartEx chart via copyTo() — chartEx is stored as raw bytes");
    }
    const clonedModel = deepClone(sourceModel);
    const targetRange = range ?? this._cloneRange();
    // Register with target worksheet
    return (targetWs as any)._registerChart(clonedModel, targetRange);
  }

  /**
   * Create a deep copy of this chart in the same worksheet.
   * @param range - Anchor for the clone (defaults to shifting right of original).
   */
  clone(range?: AddChartRange): number {
    return this.copyTo(this.worksheet, range ?? this._cloneRange());
  }

  private _cloneRange(): AddChartRange {
    // Default clone range: shift right of original by its width
    if (this.range.br) {
      const tl = this.range.tl.model;
      const br = this.range.br.model;
      const width = br.nativeCol - tl.nativeCol;
      return {
        tl: { col: br.nativeCol + 1, row: tl.nativeRow },
        br: { col: br.nativeCol + 1 + width, row: br.nativeRow }
      };
    }
    // Fallback: same location (caller should override)
    return {
      tl: { col: this.range.tl.nativeCol, row: this.range.tl.nativeRow },
      br: { col: this.range.tl.nativeCol + 10, row: this.range.tl.nativeRow + 15 }
    };
  }

  // ===========================================================================
  // Private helpers
  // ===========================================================================

  private _extractTextFromRichText(rt: ChartRichText): string {
    return rt.paragraphs.map(p => (p.runs ?? []).map(r => r.text).join("")).join("\n");
  }

  private _extractTextFromRawTx(rawTx: string): string | undefined {
    // Extract all <a:t>…</a:t> text content from raw c:tx XML.
    // Also check <c:v>…</c:v> for strRef-based titles.
    const parts: string[] = [];
    RAW_TX_AT_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = RAW_TX_AT_RE.exec(rawTx)) !== null) {
      parts.push(this._decodeXmlEntities(m[1]));
    }
    if (parts.length > 0) {
      return parts.join("");
    }
    // strRef inside rawTx: look for <c:v> elements
    RAW_TX_CV_RE.lastIndex = 0;
    while ((m = RAW_TX_CV_RE.exec(rawTx)) !== null) {
      parts.push(this._decodeXmlEntities(m[1]));
    }
    return parts.length > 0 ? parts.join("") : undefined;
  }

  private _decodeXmlEntities(text: string): string {
    return text
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  /**
   * Extract the run properties of the first text run in an existing title,
   * if any. Used to preserve formatting when replacing plain-string title text.
   * If only `rawTx` is available, we attempt to parse the first `<a:rPr>` from it.
   */
  private _extractFirstRunProperties(
    title: ChartTitle | undefined
  ): ChartTextProperties | undefined {
    if (!title) {
      return undefined;
    }
    // Structured path
    const firstRun = title.text?.paragraphs?.[0]?.runs?.[0];
    if (firstRun?.properties) {
      return { ...firstRun.properties };
    }
    // rawTx path — capture the first run's rPr attributes as opaque raw XML.
    if (title.rawTx) {
      const rPrMatch = /<a:rPr\b[^>]*\/?>/.exec(title.rawTx);
      if (rPrMatch) {
        // Build a text-properties object that renders to the same attributes.
        return this._parseRunAttrsToProps(rPrMatch[0]);
      }
    }
    return undefined;
  }

  /** Parse `<a:rPr ...>` attributes into a ChartTextProperties. */
  private _parseRunAttrsToProps(tag: string): ChartTextProperties {
    const p: ChartTextProperties = {};
    const attrMatch = (name: string): string | undefined => {
      const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag);
      return m ? m[1] : undefined;
    };
    const sz = attrMatch("sz");
    if (sz) {
      p.size = parseInt(sz, 10);
    }
    const b = attrMatch("b");
    if (b) {
      p.bold = b === "1";
    }
    const i = attrMatch("i");
    if (i) {
      p.italic = i === "1";
    }
    const u = attrMatch("u");
    if (u && u !== "none") {
      p.underline = u as NonNullable<ChartTextProperties["underline"]>;
    }
    const strike = attrMatch("strike");
    if (strike) {
      p.strike = strike as NonNullable<ChartTextProperties["strike"]>;
    }
    const lang = attrMatch("lang");
    if (lang) {
      p.lang = lang;
    }
    return p;
  }
}

export interface ChartAnchorModel {
  chartNumber: number;
  /** 1-based chartEx number (cx:chart). 0 or absent for classic charts. */
  chartExNumber?: number;
  range: {
    tl: AnchorModel;
    /** Bottom-right (only for twoCellAnchor) */
    br?: AnchorModel;
    /** Absolute position in EMU (only for absoluteAnchor) */
    pos?: { x: number; y: number };
    /** Extent in EMU (for oneCellAnchor and absoluteAnchor) */
    ext?: { cx: number; cy: number };
    editAs?: string;
  };
}

/**
 * Deep-clone a plain data structure (ChartModel is JSON-safe — no cycles, no
 * class instances). Uses `structuredClone` when available, else a JSON fallback.
 */
function deepClone<T>(obj: T): T {
  if (typeof structuredClone === "function") {
    return structuredClone(obj);
  }
  return JSON.parse(JSON.stringify(obj));
}

export { Chart, buildChartModel };
