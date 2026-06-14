/**
 * Chart Support Registry
 *
 * A tiny indirection layer that lets `Worksheet`, `Workbook` and
 * `Chartsheet` call into chart-building, chart-rendering and chart
 * cache-population code **only if the host application has opted in**
 * by calling `installChartSupport()` from
 * `@cj-tech-master/excelts/chart`.
 *
 * ## Why a registry
 *
 * The chart module ships 17 files covering classic charts, ChartEx
 * (Excel 2016+), SVG/PNG/PDF renderers, preset catalogues, shape
 * properties, TopoJSON region maps and cache populators — roughly
 * 30,000 lines of TypeScript. Empirical measurement shows that making
 * `Worksheet` import from chart directly forces ~1.2 MB of chart code
 * into every consumer bundle, even for consumers who only read and
 * write cells.
 *
 * This file is the single point of indirection. Importing `Worksheet`
 * pulls it in (~1 KB) but does NOT pull in the chart implementation;
 * the implementation only arrives once someone calls
 * `installChartSupport()`, which bundlers resolve to the
 * `@cj-tech-master/excelts/chart` subpath import graph.
 *
 * If chart-manipulating APIs (`worksheet.addChart()`, `workbook.writeFile()`
 * with charts, etc.) are called without `installChartSupport()` having
 * run, a clear error is thrown telling the developer how to fix it.
 *
 * ## Why it lives in `excel/` rather than `excel/chart/`
 *
 * Host-registry pattern convention: the slot lives with the consumer,
 * not with the provider. `worksheet.ts`, `workbook.browser.ts` and
 * `chartsheet.ts` are the consumers and they all live in `excel/`, so
 * the slot sits next to them. The chart module only depends on it when
 * `install()` is called — a one-way arrow from provider to consumer.
 *
 * This differs from `@formula/host-registry` which lives inside the
 * formula module because formula is a separate top-level module (layer
 * 3) that cannot import from excel. Chart is nested under excel
 * (layer 4), so the slot naturally lives in excel.
 *
 * ## Relationship to formula's host-registry
 *
 * The two registries are independent slots for independent optional
 * features. A consumer may install chart without formula, formula
 * without chart, both, or neither.
 */

import type { fillChartCaches, fillChartExCaches } from "./chart/cache-populator";
import type { Chart, ChartAnchorModel, buildChartModel } from "./chart/chart";
import type {
  chartExOptionsFromRows,
  chartExOptionsFromTable,
  chartOptionsFromRows,
  chartOptionsFromTable,
  seriesFromColumns
} from "./chart/chart-api";
import type { buildComboChartModel } from "./chart/chart-builder";
import type { buildChartExModel } from "./chart/chart-ex-builder";
import type { parseChartEx } from "./chart/chart-ex-parser";
import type {
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  renderChartEx,
  renderChartExLegendXml,
  renderChartExPng,
  rewriteChartExDataRefsToDefinedNames
} from "./chart/chart-ex-renderer";
import type { resolvePendingChartImages } from "./chart/chart-images";
import type { applyChartExPreset, applyChartPreset } from "./chart/chart-presets";
import type { drawChartPdf, renderChartPng } from "./chart/chart-renderer";
import type { buildChartColors, buildChartStyle } from "./chart/chart-sidecar";
import type { themeIndexToName } from "./chart/chart-utils";
import type { AddChartRange } from "./chart/types";
import type { Worksheet } from "./worksheet";
import type { ChartSpaceXform } from "./xlsx/xform/chart/chart-space-xform";

/**
 * Runtime chart functions the excel module needs from the chart
 * module. Populated by `installChartSupport()` from
 * `@cj-tech-master/excelts/chart`.
 *
 * Using `typeof <fn>` keeps the signatures in lockstep with the real
 * implementations — add an argument to `buildChartModel` and this
 * interface updates automatically, as do all call sites. Consumers who
 * provide a custom chart-support implementation (test doubles, mock
 * providers) get a type error when signatures drift.
 *
 * The `Chart` class constructor is wrapped in a factory function
 * (`createChart`) rather than exposed as a raw constructor reference:
 * `new (getChartSupport().Chart)(...)` reads poorly at the call site,
 * `getChartSupport().createChart(...)` reads cleanly.
 */
export interface ChartSupport {
  readonly buildChartModel: typeof buildChartModel;
  readonly buildComboChartModel: typeof buildComboChartModel;
  readonly buildChartExModel: typeof buildChartExModel;

  readonly applyChartPreset: typeof applyChartPreset;
  readonly applyChartExPreset: typeof applyChartExPreset;

  readonly fillChartCaches: typeof fillChartCaches;
  readonly fillChartExCaches: typeof fillChartExCaches;
  readonly resolvePendingChartImages: typeof resolvePendingChartImages;

  readonly buildChartStyle: typeof buildChartStyle;
  readonly buildChartColors: typeof buildChartColors;

  /**
   * Factory for a `Chart` instance. Preferred over exposing the raw
   * constructor because `new (getChartSupport().Chart)(...)` is
   * awkward at call sites.
   *
   * Accepts either an `AddChartRange` (the public addChart API
   * surface: A1-notation strings, tl/br pairs, tl+ext, or pos+ext)
   * **or** a `ChartAnchorModel["range"]` (the model-shaped form
   * re-hydrated from an XLSX drawing anchor). The Chart constructor
   * accepts both and routes through the same `parseRange` helper.
   */
  readonly createChart: (
    worksheet: Worksheet,
    ids: { chartNumber?: number; chartExNumber?: number },
    range: AddChartRange | ChartAnchorModel["range"]
  ) => Chart;

  readonly chartOptionsFromTable: typeof chartOptionsFromTable;
  readonly chartOptionsFromRows: typeof chartOptionsFromRows;
  readonly chartExOptionsFromTable: typeof chartExOptionsFromTable;
  readonly chartExOptionsFromRows: typeof chartExOptionsFromRows;
  readonly seriesFromColumns: typeof seriesFromColumns;

  // XLSX serialisation helpers — consumed by `xlsx.browser.ts` during
  // `Workbook.writeXlsx(workbook)` / `readFile()` when charts are present.
  readonly parseChartEx: typeof parseChartEx;
  readonly renderChartEx: typeof renderChartEx;
  readonly renderChartExLegendXml: typeof renderChartExLegendXml;
  readonly rewriteChartExDataRefsToDefinedNames: typeof rewriteChartExDataRefsToDefinedNames;
  readonly themeIndexToName: typeof themeIndexToName;
  /**
   * Factory for the classic-chart XLSX xform (`c:chartSpace` parser +
   * writer). `chart-space-xform.ts` itself statically imports
   * `chart-utils` and `shape-properties` from the chart module, so
   * exposing it through the registry ensures a consumer who never
   * installs chart support also never pulls those chart files into
   * their bundle via the xform graph.
   */
  readonly createChartSpaceXform: () => ChartSpaceXform;

  // PDF export helpers — consumed by `pdf/excel-bridge.ts` when a
  // workbook with charts is converted into a PDF.
  readonly drawChartPdf: typeof drawChartPdf;
  readonly drawChartExPdf: typeof drawChartExPdf;
  readonly renderChartPng: typeof renderChartPng;
  readonly renderChartExPng: typeof renderChartExPng;
  readonly canRenderChartExAsVectorPdf: typeof canRenderChartExAsVectorPdf;
}

let installed: ChartSupport | null = null;

/**
 * Install a chart runtime implementation. Called from
 * `installChartSupport()` in the `@cj-tech-master/excelts/chart`
 * subpath.
 *
 * Re-installing is allowed — the last registration wins. Passing
 * `null` uninstalls chart support, which is useful for tests that
 * exercise the "no chart" error path.
 */
export function registerChartSupport(support: ChartSupport | null): void {
  installed = support;
}

/** Returns true when chart support has been installed. */
export function hasChartSupport(): boolean {
  return installed !== null;
}

/**
 * Retrieve the installed chart support. Throws a descriptive error if
 * no chart support is installed.
 */
export function getChartSupport(): ChartSupport {
  if (!installed) {
    throw new Error(
      "No chart support is installed. " +
        "Call `installChartSupport()` from `@cj-tech-master/excelts/chart` " +
        "once at startup to enable chart creation, loading, and " +
        "serialisation during `Workbook.writeXlsx(workbook)`."
    );
  }
  return installed;
}

/**
 * Retrieve installed chart support, or `null` if none is installed.
 *
 * Used by code paths that can gracefully degrade when chart support is
 * missing — e.g. ChartEx loading, which can preserve raw XML verbatim
 * without a full parse when chart support is absent. Note that classic
 * chart loading currently requires chart support to be installed (the
 * reader calls `getChartSupport()` unconditionally for classic charts).
 */
export function tryGetChartSupport(): ChartSupport | null {
  return installed;
}
