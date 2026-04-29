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
  SeriesBase,
  AddChartSeriesOptions
} from "@excel/chart/types";
import { ChartOptionsError } from "@excel/errors";
import { colCache } from "@excel/utils/col-cache";
import type { Worksheet } from "@excel/worksheet";

import {
  applyChartSeriesOptionsPatch,
  buildChartModel,
  buildChartSeriesForType
} from "./chart-builder";
import { renderChartExPng, renderChartExSvg } from "./chart-ex-renderer";
import { renderChartPng, renderChartSvg, type ChartRenderOptions } from "./chart-renderer";

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
  /** Original chart XML bytes from a loaded workbook, used for clean round-trip passthrough */
  rawData?: Uint8Array;
  /** JSON snapshot of `model` taken when `rawData` was parsed */
  modelSnapshot?: string;
  /** True once a high-level API mutates the parsed chart model */
  dirty?: boolean;
  /** When true, simple high-level mutations may patch raw XML instead of full re-render. */
  preferRawPatch?: boolean;
  /** When true, writing fails instead of re-rendering if raw XML cannot be safely patched. */
  requireRawPatch?: boolean;
  /** Chart style (styleN.xml) — raw XML for round-trip */
  style?: ChartStyleModel;
  /** Chart colors (colorsN.xml) — raw XML for round-trip */
  colors?: ChartColorsModel;
  /** Chart rels (chart{N}.xml.rels) — raw entries for round-trip */
  rels?: any[];
  /**
   * Raw bytes of the user-shapes drawing part targeted by this chart's
   * `c:userShapes r:id="…"` reference. OOXML stores annotation shapes
   * (arrows, callouts, text boxes the user drew on top of the chart)
   * in a separate `xl/drawings/drawingN.xml` part that uses relative
   * anchors instead of the regular `xdr:twoCellAnchor` schema. The
   * full DrawingML subsystem is out of scope for this library, so the
   * bytes are kept verbatim for round-trip and exposed programmatically
   * via {@link Chart.userShapesXml} so callers can inject / replace the
   * drawing part if they need to.
   */
  userShapesXml?: Uint8Array;
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
  /** Original chartEx XML bytes from a loaded workbook, used for clean round-trip passthrough */
  rawData?: Uint8Array;
  /** JSON snapshot of `model` taken when `rawData` was parsed */
  modelSnapshot?: string;
  /** True once a high-level API mutates the parsed chartEx model */
  dirty?: boolean;
  /** When true, simple high-level mutations may patch raw ChartEx XML instead of full re-render. */
  preferRawPatch?: boolean;
  /** When true, writing fails instead of re-rendering if raw ChartEx XML cannot be safely patched. */
  requireRawPatch?: boolean;
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

  /** Get the structured ChartEx model, when this chart is an Office 2016+ chartEx. */
  get chartExModel(): ChartExModel | undefined {
    if (this.chartExNumber <= 0) {
      return undefined;
    }
    return this.worksheet.workbook.getChartExStructuredEntry(this.chartExNumber)?.model;
  }

  /**
   * Vendor-extension elements the parser observed but could not map to a
   * structured field. Populated when the chart was loaded from an existing
   * `.xlsx` whose author emitted `c15:`/`cx14:` style extension tags (often
   * MSO-internal; the OOXML spec treats them as implementation-specific).
   *
   * The array is **purely informational** in the default `preserve` writer
   * mode — structural rebuilds keep the raw XML pass-through that contains
   * those tags. In **`strictTemplateMode`** the writer surfaces this list
   * in its failure message when a mutation cannot be expressed as a raw
   * patch, so authors can decide between:
   *
   *   1. Relaxing `strictTemplateMode` and accepting that a rebuild will
   *      drop these vendor tags, or
   *   2. Reshaping the mutation to land on a patch-friendly path (for
   *      example editing `title` text rather than replacing the whole
   *      `c:chart` subtree).
   *
   * Returns `undefined` if the chart was freshly created (no raw XML was
   * ever parsed) or if every child was recognised. See
   * {@link XlsxWriteOptions.strictTemplateMode} for the writer surface.
   */
  get unknownElements(): Array<{ name: string; path: string }> | undefined {
    const classic = this.chartModel?.unknownElements;
    if (classic && classic.length > 0) {
      return classic.slice();
    }
    const chartEx = this.chartExModel?.unknownElements;
    if (chartEx && chartEx.length > 0) {
      return chartEx.slice();
    }
    return undefined;
  }

  // ===========================================================================
  // User-shape drawing part (c:userShapes)
  // ===========================================================================

  /**
   * Raw XML bytes of the `c:userShapes` drawing part attached to this
   * chart, or `undefined` when the chart has no user shapes. User
   * shapes are annotation overlays (callouts, arrows, free text) that
   * Excel stores in a separate `xl/drawings/drawingN.xml` part using
   * relative anchors — the OOXML spec keeps their DrawingML schema
   * distinct from worksheet drawings, so the library treats the part
   * as opaque bytes for round-trip and programmatic replacement
   * instead of exposing a full structural API.
   *
   * Classic charts only. Returns `undefined` on chartEx charts (they
   * have their own extension mechanism and do not use `c:userShapes`).
   */
  get userShapesXml(): Uint8Array | undefined {
    if (this.chartNumber <= 0) {
      return undefined;
    }
    return this.worksheet.workbook.getChartEntry(this.chartNumber)?.userShapesXml;
  }

  /**
   * Replace the `c:userShapes` drawing part for this chart with the
   * supplied raw XML bytes. Passing `undefined` or an empty byte
   * array removes the user-shapes reference entirely (equivalent to
   * {@link removeUserShapes}).
   *
   * The XML must be a complete DrawingML document whose root is a
   * `c:userShapes` element containing `c:relSizeAnchor` /
   * `c:absSizeAnchor` children. The library performs a shallow sanity
   * check only — no schema validation.
   *
   * When a chart did not previously have user shapes, this allocates
   * a new `r:id` and adds the drawing-part rel so the reference is
   * discoverable in `chartN.xml.rels`. When a chart already had user
   * shapes, the existing `r:id` is kept and only the bytes are
   * updated.
   *
   * Classic charts only. Throws on chartEx charts.
   */
  setUserShapesXml(xml: Uint8Array | string | undefined): void {
    if (this.chartNumber <= 0) {
      throw new Error("Chart.setUserShapesXml is only supported on classic charts");
    }
    const entry = this.worksheet.workbook.getChartEntry(this.chartNumber);
    if (!entry) {
      throw new Error(`Chart ${this.chartNumber} has no registered chart entry`);
    }
    if (xml === undefined) {
      this.removeUserShapes();
      return;
    }
    const bytes = typeof xml === "string" ? new TextEncoder().encode(xml) : xml;
    if (bytes.length === 0) {
      this.removeUserShapes();
      return;
    }
    // Minimal sanity check — full validation is left to downstream consumers.
    const peek = new TextDecoder().decode(bytes.slice(0, 512));
    if (!peek.includes("userShapes")) {
      throw new Error(
        "Chart.setUserShapesXml expects a DrawingML document whose root is c:userShapes"
      );
    }
    entry.userShapesXml = bytes;
    entry.dirty = true;
    // Allocate an r:id if we didn't have one. The writer fills in the
    // actual rel entry when emitting chartN.xml.rels — we only need to
    // reserve a stable r:id so the chart XML can reference it.
    if (!entry.model.userShapesRelId) {
      const rels = (entry.rels ??= []);
      const existing = new Set(rels.map(r => r?.Id).filter(Boolean));
      let counter = rels.length + 1;
      let rid = `rId${counter}`;
      while (existing.has(rid)) {
        counter += 1;
        rid = `rId${counter}`;
      }
      entry.model.userShapesRelId = rid;
    }
  }

  /**
   * Drop the `c:userShapes` reference and its backing drawing part.
   * Classic charts only. No-op when the chart has no user shapes.
   */
  removeUserShapes(): void {
    if (this.chartNumber <= 0) {
      return;
    }
    const entry = this.worksheet.workbook.getChartEntry(this.chartNumber);
    if (!entry) {
      return;
    }
    if (!entry.userShapesXml && !entry.model.userShapesRelId) {
      return;
    }
    entry.userShapesXml = undefined;
    entry.dirty = true;
    const relId = entry.model.userShapesRelId;
    entry.model.userShapesRelId = undefined;
    if (relId && Array.isArray(entry.rels)) {
      entry.rels = entry.rels.filter(r => r?.Id !== relId);
    }
  }

  /**
   * Render this chart as a **zero-dependency deterministic preview** SVG.
   *
   * The output is suitable for thumbnails, email attachments,
   * server-side report generation, CI smoke tests, and README images.
   * It is **not** an Excel-pixel-perfect compositor — text layout,
   * font metrics, and 3D projection are approximated for a stable
   * preview rather than reproduced from Excel's internal renderer.
   *
   * For production-grade rendering (Excel-identical layout, real 3D
   * for non-bar types, exact font hinting), round-trip the `.xlsx`
   * through headless LibreOffice (`soffice --convert-to pdf`) — the
   * byte-preserving round-trip + `templateMode: "strict"` guarantees
   * in this library make that a safe handoff.
   *
   * See `src/modules/excel/README.md` → "Rendering scope" for the
   * complete boundary list.
   */
  toSVG(options: ChartRenderOptions = {}): string {
    const model = this.chartModel;
    if (model) {
      return renderChartSvg(model, options);
    }
    const chartEx = this.chartExModel;
    if (chartEx) {
      return renderChartExSvg(chartEx, options);
    }
    throw new Error("Cannot render chart because no chart model is available");
  }

  /**
   * Render this chart as a **zero-dependency deterministic preview** PNG.
   *
   * Browsers use a `<canvas>` pipeline; Node.js uses the built-in
   * `BasicRasterCanvas` rasteriser (a pure-JS SVG-subset rasteriser —
   * no native canvas dependency). DrawingML effect filters
   * (shadow/glow/soft-edge/blur/reflection) round-trip through XML and
   * emit as SVG `<filter>`, but the Node PNG rasteriser silently drops
   * them; the browser canvas path renders them natively.
   *
   * See {@link toSVG} for the full scope-boundary note. For pixel-perfect
   * output, convert through LibreOffice.
   */
  toPNG(options: ChartRenderOptions = {}): Promise<Uint8Array> {
    const model = this.chartModel;
    if (model) {
      return renderChartPng(model, options);
    }
    const chartEx = this.chartExModel;
    if (chartEx) {
      return renderChartExPng(chartEx, options);
    }
    throw new Error("Cannot render chart because no chart model is available");
  }

  /** Get the chart title text. Returns undefined if no title. */
  get title(): string | undefined {
    const titleObj = this.chartModel?.chart?.title ?? this.chartExModel?.chartSpace?.chart?.title;
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
    const chart = this.chartModel?.chart ?? this.chartExModel?.chartSpace?.chart;
    if (!chart) {
      return;
    }
    if (value === undefined) {
      this._markDirty(true);
      chart.title = undefined;
      chart.autoTitleDeleted = true;
      return;
    }
    if (typeof value === "object" && "formula" in value && typeof value.formula === "string") {
      this._markDirty();
      chart.title = {
        strRef: { formula: value.formula, cache: { points: [] } },
        overlay: chart.title?.overlay ?? false
      };
      chart.autoTitleDeleted = false;
      return;
    }
    if (typeof value === "object" && "paragraphs" in value) {
      this._markDirty();
      // Rich text
      chart.title = {
        ...(chart.title ?? {}),
        strRef: undefined,
        rawTx: undefined,
        text: value,
        overlay: chart.title?.overlay ?? false
      };
      chart.autoTitleDeleted = false;
      return;
    }
    // Plain string — try to preserve first-run formatting of existing title.
    if (typeof value !== "string") {
      return; // Type-safety guard; other object shapes handled above.
    }
    this._markDirty(true);
    const existing = chart.title;
    const preservedProps = this._extractFirstRunProperties(existing);
    chart.title = {
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
    chart.autoTitleDeleted = false;
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
    return this.chartModel?.chart?.title?.text ?? this.chartExModel?.chartSpace?.chart?.title?.text;
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
    return this.chartModel?.chart?.legend ?? this.chartExModel?.chartSpace?.chart?.legend;
  }

  /** Set the chart legend */
  set legend(value: ChartLegend | undefined) {
    const chart = this.chartModel?.chart ?? this.chartExModel?.chartSpace?.chart;
    if (chart) {
      this._markDirty(true);
      chart.legend = value;
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
      this._markDirty();
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

  mutate(
    mutator: (model: ChartModel) => void,
    options: { preferRawPatch?: boolean; requireRawPatch?: boolean } = {}
  ): this {
    const model = this.chartModel;
    if (!model) {
      throw new Error(
        "Cannot mutate a classic chart model because this chart is ChartEx or missing"
      );
    }
    mutator(model);
    this._markDirty(options.preferRawPatch ?? false, options.requireRawPatch ?? false);
    return this;
  }

  mutateChartEx(
    mutator: (model: ChartExModel) => void,
    options: { preferRawPatch?: boolean; requireRawPatch?: boolean } = {}
  ): this {
    const model = this.chartExModel;
    if (!model) {
      throw new Error("Cannot mutate a ChartEx model because this chart is classic or missing");
    }
    mutator(model);
    // Invalidate the model-level rawXml cache so subsequent calls to
    // `renderChartEx(model)` (e.g. standalone preview, tests) honour the
    // mutation instead of short-circuiting to the bytes captured at parse
    // time. The xlsx writer path has its own change-detection (see
    // `hasChartExEntryChanged` in xlsx.browser.ts) and does not depend on
    // this flag, but direct consumers of `renderChartEx` do.
    //
    // `preferRawPatch` consumers explicitly opt in to surgical byte
    // patching — they keep the rawXml so the patcher in xlsx.browser.ts
    // can reuse it.
    if (!options.preferRawPatch && !options.requireRawPatch) {
      model.rawXml = undefined;
    }
    this._markDirty(options.preferRawPatch ?? false, options.requireRawPatch ?? false);
    return this;
  }

  /**
   * Set the built-in chart style by index.
   *
   * Writes `<c:style val="N"/>` on the classic chart (`chartN.xml`).
   * Valid values are 1–48, matching the legacy Excel 2007/2010 style
   * catalogue and `xlsxwriter`'s `set_style(N)`. This is the lightweight
   * option — it does **not** emit a `styleN.xml` / `colorsN.xml` sidecar;
   * for modern Office-2013-era styling use {@link setChartStyle} instead.
   *
   * ChartEx charts do not honour this field; calling `setStyle` on one
   * throws, which matches OOXML: `c:style` only exists in the classic
   * `c:` namespace.
   *
   * @param style - Integer in the range 1–48.
   * @throws {ChartOptionsError} when `style` is outside 1–48 or the
   *         chart is a ChartEx.
   */
  setStyle(style: number): this {
    if (!Number.isInteger(style) || style < 1 || style > 48) {
      throw new ChartOptionsError(
        `Chart.setStyle: built-in style must be an integer in 1..48, received ${style}`
      );
    }
    const model = this.chartModel;
    if (!model) {
      throw new ChartOptionsError(
        "Chart.setStyle is only valid on classic charts; ChartEx does not honour `c:style`"
      );
    }
    model.style = style;
    this._markDirty();
    return this;
  }

  /**
   * Alias for {@link setStyle} that matches the `xlsxwriter` terminology
   * used by Python/Rust users migrating their chart code. Equivalent in
   * every way — both write the same `<c:style val>` attribute.
   */
  setBuiltInStyle(style: number): this {
    return this.setStyle(style);
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
      this._markDirty();
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
    this._markDirty();
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

  /** Update common fields on an existing classic chart series. */
  updateSeries(index: number, options: Partial<AddChartSeriesOptions>, groupIndex = 0): boolean {
    const series = this.getSeries(index, groupIndex);
    if (!series) {
      return false;
    }
    this._markDirty(true);
    applyChartSeriesOptionsPatch(series, options, this.chartTypes[groupIndex]?.type);
    return true;
  }

  /** Add a series from high-level options, matching the target chart type group. */
  addSeriesFromOptions(options: AddChartSeriesOptions, groupIndex = 0): boolean {
    const ctg = this.chartTypes[groupIndex];
    if (!ctg) {
      return false;
    }
    const index = ctg.series.length;
    this._markDirty();
    ctg.series.push(buildChartSeriesForType(ctg.type, options, index) as any);
    return true;
  }

  /** Update the value range for a series. */
  setSeriesValues(index: number, values: string, groupIndex = 0): boolean {
    return this.updateSeries(index, { values }, groupIndex);
  }

  /** Update the category range for a series. */
  setSeriesCategories(index: number, categories: string, groupIndex = 0): boolean {
    return this.updateSeries(index, { categories }, groupIndex);
  }

  /** Update the series display name or formula reference. */
  setSeriesName(
    index: number,
    name: NonNullable<AddChartSeriesOptions["name"]>,
    groupIndex = 0
  ): boolean {
    return this.updateSeries(index, { name }, groupIndex);
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
    const sourceChartExModel = this.chartExModel;
    if (!sourceModel && !sourceChartExModel) {
      throw new Error("Cannot copy chart because no chart model is available");
    }
    const targetRange = range ?? this._cloneRange();
    if (sourceChartExModel) {
      return (targetWs as any)._registerChartEx(deepClone(sourceChartExModel), targetRange);
    }
    const clonedModel = deepClone(sourceModel!);
    const chartNumber = (targetWs as any)._registerChart(clonedModel, targetRange);
    this.worksheet.workbook.copyChartSidecars?.(this.chartNumber, chartNumber, targetWs.workbook);
    return chartNumber;
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

  private _markDirty(preferRawPatch = false, requireRawPatch = false): void {
    if (this.chartNumber > 0) {
      const entry = this.worksheet.workbook.getChartEntry(this.chartNumber);
      if (entry) {
        entry.dirty = true;
        entry.preferRawPatch = preferRawPatch;
        entry.requireRawPatch = requireRawPatch;
      }
    }
    if (this.chartExNumber > 0) {
      const entry = this.worksheet.workbook.getChartExStructuredEntry(this.chartExNumber);
      if (entry) {
        entry.dirty = true;
        entry.preferRawPatch = preferRawPatch;
        entry.requireRawPatch = requireRawPatch;
      }
    }
  }

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
