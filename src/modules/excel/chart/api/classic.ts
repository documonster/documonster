/**
 * `ChartClassic` namespace surface — the handle-free engine for the classic chart family
 * (`c:chartSpace` — bar, line, pie, scatter and the rest): build a
 * {@link ChartModel} from options, render it, and read or write the chart sidecar parts.
 *
 * `import { ChartClassic } from "documonster/chart"` →
 *   `ChartClassic.createModel(options)`, `ChartClassic.toSvg(model)`,
 *   `await ChartClassic.toPng(model)`, `ChartClassic.toDrawList(model)`, …
 *
 * A chart goes onto a PDF page the way every display list does:
 * `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.
 *
 * Named for the family it serves, parallel to `ChartEx` for the modern family — not
 * `Chart`, which `documonster/excel` already publishes for the operations on a chart
 * placed in a worksheet, and the two are routinely imported together. It works on models;
 * the members that reach into a workbook are the cache fillers (`fillCaches`, `fillNumRef`,
 * `fillStrRef`) and the option builders that read or write a sheet (`createOptionsFromTable`,
 * `createOptionsFromRows`), which are the ones `Chart` publishes for a placed chart. The
 * DrawingML shape/text property helpers live on `DrawingMl`; the renderer's building
 * blocks on `ChartClassic.Backend`.
 */

import type { DrawList } from "@draw/types";
import {
  checkFillArguments,
  checkObjectArgument,
  checkStringArgument,
  namedCall
} from "@excel/chart/api/call";
import type { ChartFillNumRefOptions, ChartFillRefOptions } from "@excel/chart/api/types";
import {
  fillChartCaches,
  fillNumRef as fillNumRefImpl,
  fillStrRef as fillStrRefImpl
} from "@excel/chart/bridge/excel-chart-host";
import {
  chartOptionsFromRows,
  chartOptionsFromTable,
  seriesFromColumns
} from "@excel/chart/build/chart-api";
import type {
  AddChartFromRowsOptions,
  AddChartFromTableOptions,
  SeriesFromColumnsOptions
} from "@excel/chart/build/chart-api";
import { buildChartModel, buildComboChartModel } from "@excel/chart/build/chart-builder";
import { applyChartPreset } from "@excel/chart/model/chart-presets";
import type { ExcelChartPreset } from "@excel/chart/model/chart-presets";
import type {
  AddChartOptions,
  AddComboChartOptions,
  AddChartSeriesOptions,
  ChartColorsModel,
  ChartModel,
  ChartStyleModel,
  NumberReference,
  StringReference
} from "@excel/chart/model/types";
import { chartDrawList, renderChartPng, renderChartSvg } from "@excel/chart/render/chart-renderer";
import type { ChartRenderOptions } from "@excel/chart/render/chart-renderer";
import {
  buildChartColors,
  buildChartStyle,
  parseChartColors,
  parseChartStyle
} from "@excel/chart/serialize/chart-sidecar";
import type { WorkbookData } from "@excel/core/workbook-core";
import { checkWorksheet, type WorksheetData } from "@excel/core/worksheet-core";
import { argumentChecks } from "@excel/utils/arguments";
import { asPlainBytes } from "@utils/binary";

// Every function here runs through `namedCall`, so an argument error says which call received
// it (`ChartClassic.createModel: options.series must …`) — the builders are shared with
// `Chart.*` in `documonster/excel` and cannot know.

// --- Models ---
/**
 * A classic chart model from options — bar, line, pie, scatter and the rest.
 *
 * @throws {ExcelChartOptionsError} an option has no valid reading (its `field` names it).
 */
export function createModel(options: AddChartOptions): ChartModel {
  return namedCall("ChartClassic.createModel", () => buildChartModel(options));
}

/** A combo chart model: several chart groups sharing a plot area. */
export function createComboModel(options: AddComboChartOptions): ChartModel {
  return namedCall("ChartClassic.createComboModel", () => buildComboChartModel(options));
}

/**
 * A classic chart's options from plain rows, which are written to `worksheet` for the chart to refer to — the same
 * builder `Chart.createOptionsFromRows` uses; everything is checked before the first cell is written.
 */
export function createOptionsFromRows<T extends object>(
  worksheet: WorksheetData,
  rows: readonly T[],
  options: AddChartFromRowsOptions<T>
): AddChartOptions {
  return namedCall("ChartClassic.createOptionsFromRows", () => {
    checkWorksheet(worksheet, "worksheet");
    return chartOptionsFromRows(worksheet, rows, options);
  });
}

/** A classic chart's options from a table of `worksheet` — the same builder `Chart.createOptionsFromTable` uses. */
export function createOptionsFromTable(
  worksheet: WorksheetData,
  table: Parameters<typeof chartOptionsFromTable>[1],
  options: AddChartFromTableOptions
): AddChartOptions {
  return namedCall("ChartClassic.createOptionsFromTable", () => {
    checkWorksheet(worksheet, "worksheet");
    return chartOptionsFromTable(worksheet, table, options);
  });
}

/**
 * A series over ranges of the sheet `sheetName`, each range qualified with it and made absolute. Takes the sheet's
 * name rather than the sheet, so a model can be built without a workbook; `Chart.createSeriesFromColumns` takes the
 * sheet.
 */
export function createSeriesFromColumns(
  sheetName: string,
  options: SeriesFromColumnsOptions
): AddChartSeriesOptions {
  return namedCall("ChartClassic.createSeriesFromColumns", () =>
    seriesFromColumns(sheetName, options)
  );
}

/** Merge a classic preset — the chart type and grouping Excel's gallery names — into `options`. */
export function applyPreset(
  preset: ExcelChartPreset,
  options: Parameters<typeof applyChartPreset>[1]
): AddChartOptions {
  return namedCall("ChartClassic.applyPreset", () => applyChartPreset(preset, options));
}

/**
 * Fill every cached point of a chart model's references from the workbook — what a chart shows before Excel
 * recalculates. `options.worksheet` is the sheet the references are read from when they name none.
 */
export function fillCaches(
  model: ChartModel,
  workbook: WorkbookData,
  options: ChartFillRefOptions = {}
): void {
  checkFillArguments("ChartClassic.fillCaches", workbook, options);
  fillChartCaches(model, workbook, options.worksheet);
}

/**
 * Fill the cached points of one number reference from the workbook — for a reference built by
 * hand rather than through a chart model (`fillCaches` does a whole model).
 */
export function fillNumRef(
  ref: NumberReference,
  workbook: WorkbookData,
  options: ChartFillNumRefOptions = {}
): void {
  checkObjectArgument("ChartClassic.fillNumRef", "ref", ref);
  checkFillArguments("ChartClassic.fillNumRef", workbook, options);
  argumentChecks("ChartClassic.fillNumRef").boolean("options.date1904", options.date1904);
  fillNumRefImpl(ref, workbook, options.date1904, options.worksheet);
}

/** Fill the cached points of one string reference from the workbook. */
export function fillStrRef(
  ref: StringReference,
  workbook: WorkbookData,
  options: ChartFillRefOptions = {}
): void {
  checkObjectArgument("ChartClassic.fillStrRef", "ref", ref);
  checkFillArguments("ChartClassic.fillStrRef", workbook, options);
  fillStrRefImpl(ref, workbook, options.worksheet);
}

// --- Rendering ---
/**
 * Render a classic chart model to PNG bytes.
 *
 * On Node the display list is rasterised by `documonster/draw`, with `options.fonts` and
 * `options.useSystemFonts` as `Draw.toRgba` takes them — `{ fonts, useSystemFonts: false }`
 * gives the same PNG on every machine. In a browser the platform draws the SVG with its own
 * fonts, and those two options do not apply. For the characters no font covered, draw the list
 * yourself: `Draw.toRgba(ChartClassic.toDrawList(model), { fonts }).uncoveredCodePoints`.
 *
 * @throws {ExcelChartOptionsError} `model` is not a classic chart model, or an option has no valid
 *   reading — a `width`, `height`, `scale` or `dpi` that is not a finite number > 0, a title
 *   that is not text, a `background` that is not a colour, or `fonts` / `useSystemFonts` of
 *   the wrong shape.
 */
export async function toPng(
  model: ChartModel,
  options?: ChartRenderOptions
): Promise<Uint8Array<ArrayBuffer>> {
  return asPlainBytes(await renderChartPng(model, options, "ChartClassic.toPng"));
}

/**
 * Render a classic chart model to SVG markup.
 *
 * @throws {ExcelChartOptionsError} as {@link toPng}.
 */
export function toSvg(model: ChartModel, options?: ChartRenderOptions): string {
  return renderChartSvg(model, options, "ChartClassic.toSvg");
}

/**
 * A classic chart model as a display list — the picture `toSvg` serialises and `toPng`
 * rasterises — for drawing it with `documonster/draw`: `Draw.toRgba(list, { fonts,
 * useSystemFonts: false })` for pixels that do not depend on the host's fonts, or a PDF page
 * through `Pdf.createDrawSurface`.
 *
 * @throws {ExcelChartOptionsError} as {@link toPng}.
 */
export function toDrawList(model: ChartModel, options?: ChartRenderOptions): DrawList {
  return chartDrawList(model, options, "ChartClassic.toDrawList");
}

/**
 * The renderer's building blocks, for a caller writing their own backend: the scene the backends draw from, the axis
 * transform, the SVG effect filter and the TopoJSON resolution a region map uses.
 */
export * as Backend from "@excel/chart/api/classic-backend";

// --- Sidecar parts (chart style / chart colours XML) ---
/** The chart colours part (`colorsN.xml`) for a colours model. */
export function toColorsXml(model: ChartColorsModel): string {
  checkObjectArgument("ChartClassic.toColorsXml", "model", model);
  return buildChartColors(model);
}

/** The chart style part (`styleN.xml`) for a style model. */
export function toStyleXml(model: ChartStyleModel): string {
  checkObjectArgument("ChartClassic.toStyleXml", "model", model);
  return buildChartStyle(model);
}

/** Read a chart colours part into a model; what it does not describe is kept as the raw XML. */
export function fromColorsXml(rawXml: string): ChartColorsModel {
  checkStringArgument("ChartClassic.fromColorsXml", "rawXml", rawXml);
  return parseChartColors(rawXml);
}

/** Read a chart style part into a model; what it does not describe is kept as the raw XML. */
export function fromStyleXml(rawXml: string): ChartStyleModel {
  checkStringArgument("ChartClassic.fromStyleXml", "rawXml", rawXml);
  return parseChartStyle(rawXml);
}
