/**
 * `ChartEx` namespace surface — the modern `cx:chartSpace` chart family (histogram,
 * Pareto, box & whisker, waterfall, funnel, treemap, sunburst, region map).
 *
 * `import { ChartEx } from "documonster/chart"` →
 *   `ChartEx.createModel(options)`, `ChartEx.toSvg(model)`, `await ChartEx.toPng(model)`,
 *   `ChartEx.toDrawList(model)`, `ChartEx.fromXml(xml)`, `ChartEx.toXml(model)`.
 *
 * Where an operation exists for both families it has the same name and argument order
 * as on `ChartClassic` (`createModel`, `createOptionsFromRows`, `createOptionsFromTable`,
 * `applyPreset`, `fillCaches`, `toSvg`, `toPng`, `toDrawList`). The rest are specific to
 * one family: this one adds the part XML (`fromXml`, `toXml`); combos, scenes, effect
 * filters and the style/colour sidecars are classic-only.
 *
 * A chart goes onto a PDF page the way every display list does — every layout has one:
 * `Draw.render(ChartEx.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.
 */

import type { DrawList } from "@draw/types";
import { checkFillArguments, namedCall } from "@excel/chart/api/call";
import type { ChartFillRefOptions } from "@excel/chart/api/types";
import { fillChartExCaches } from "@excel/chart/bridge/excel-chart-host";
import { chartExOptionsFromRows, chartExOptionsFromTable } from "@excel/chart/build/chart-api";
import type {
  AddChartExFromRowsOptions,
  AddChartExFromTableOptions
} from "@excel/chart/build/chart-api";
import { buildChartExModel } from "@excel/chart/build/chart-ex-builder";
import type { AddChartExOptions, ChartExModel } from "@excel/chart/model/chart-ex-types";
import { applyChartExPreset } from "@excel/chart/model/chart-presets";
import type { ExcelChartExPreset } from "@excel/chart/model/chart-presets";
import {
  checkChartExModel,
  renderChartExDrawList,
  renderChartExPng,
  renderChartExSvg
} from "@excel/chart/render/chart-ex-renderer";
import type { ChartRenderOptions } from "@excel/chart/render/chart-renderer";
import { renderChartEx } from "@excel/chart/serialize/chart-ex-serialize";
import type { WorkbookData } from "@excel/core/workbook-core";
import { checkWorksheet, type WorksheetData } from "@excel/core/worksheet-core";
import { argumentChecks } from "@excel/utils/arguments";

// Every builder here runs through `namedCall`, so an argument error says which call received
// it (`ChartEx.createModel: options.series must …`) — the builders are shared with `Chart.*`
// in `documonster/excel` and cannot know.

/**
 * An Office 2016 chart model from options — histogram, Pareto, waterfall, funnel, treemap,
 * sunburst, box & whisker or region map.
 *
 * @throws {ExcelChartOptionsError} an option has no valid reading (its `field` names it).
 */
export function createModel(options: AddChartExOptions): ChartExModel {
  return namedCall("ChartEx.createModel", () => buildChartExModel(options));
}

/**
 * An Office 2016 chart's options from plain rows, which are written to `worksheet` for the chart to refer to — the
 * same builder `Chart.createOptionsFromRows` uses for these types.
 */
export function createOptionsFromRows<T extends object>(
  worksheet: WorksheetData,
  rows: readonly T[],
  options: AddChartExFromRowsOptions<T> & {
    type: Exclude<AddChartExOptions["type"], "regionMap">;
  }
): AddChartExOptions {
  return namedCall("ChartEx.createOptionsFromRows", () => {
    checkWorksheet(worksheet, "worksheet");
    return chartExOptionsFromRows(worksheet, rows, options);
  });
}

/** An Office 2016 chart's options from a table of `worksheet`; a region map is not drawn from a table. */
export function createOptionsFromTable(
  worksheet: WorksheetData,
  table: Parameters<typeof chartExOptionsFromTable>[1],
  options: AddChartExFromTableOptions & { type: Exclude<AddChartExOptions["type"], "regionMap"> }
): AddChartExOptions {
  return namedCall("ChartEx.createOptionsFromTable", () => {
    checkWorksheet(worksheet, "worksheet");
    return chartExOptionsFromTable(worksheet, table, options);
  });
}

/** Merge an Office 2016 preset (`"histogram"`, `"treemap"`, …) into `options`. */
export function applyPreset(
  preset: ExcelChartExPreset,
  options: Parameters<typeof applyChartExPreset>[1]
): AddChartExOptions {
  return namedCall("ChartEx.applyPreset", () => applyChartExPreset(preset, options));
}

/**
 * Render an Office 2016 chart model to PNG bytes.
 *
 * On Node the display list is rasterised by `documonster/draw`, with `options.fonts` and
 * `options.useSystemFonts` as `Draw.toRgba` takes them — `{ fonts, useSystemFonts: false }`
 * gives the same PNG on every machine. In a browser the platform draws the SVG with its own
 * fonts, and those two options do not apply. For the characters no font covered, draw the list
 * yourself: `Draw.toRgba(ChartEx.toDrawList(model), { fonts }).uncoveredCodePoints`.
 *
 * @throws {ExcelChartOptionsError} `model` is not an Office 2016 chart model, or an option has no
 *   valid reading — a `width`, `height`, `scale` or `dpi` that is not a finite number > 0, a
 *   title that is not text, a `background` that is not a colour, or `fonts` /
 *   `useSystemFonts` of the wrong shape.
 */
export async function toPng(
  model: ChartExModel,
  options?: ChartRenderOptions
): Promise<Uint8Array> {
  return renderChartExPng(model, options, "ChartEx.toPng");
}

/**
 * Render an Office 2016 chart model to SVG markup.
 *
 * @throws {ExcelChartOptionsError} as {@link toPng}.
 */
export function toSvg(model: ChartExModel, options?: ChartRenderOptions): string {
  return renderChartExSvg(model, options, "ChartEx.toSvg");
}

/**
 * An Office 2016 chart model as a display list — the picture `toPng` rasterises — for drawing
 * it with `documonster/draw`: `Draw.toRgba(list, { fonts, useSystemFonts: false })` for pixels
 * that do not depend on the host's fonts, or a PDF page through `Pdf.createDrawSurface`. (The
 * SVG `toSvg` writes is the same picture; it also marks how a region map was drawn.)
 *
 * @throws {ExcelChartOptionsError} as {@link toPng}.
 */
export function toDrawList(model: ChartExModel, options?: ChartRenderOptions): DrawList {
  return renderChartExDrawList(model, options, "ChartEx.toDrawList");
}

/**
 * Fill every cached point of an Office 2016 chart model's data from the workbook. `options.worksheet` is the sheet
 * the references are read from when they name none.
 */
export function fillCaches(
  model: ChartExModel,
  workbook: WorkbookData,
  options: ChartFillRefOptions = {}
): void {
  checkFillArguments("ChartEx.fillCaches", workbook, options);
  fillChartExCaches(model, workbook, options.worksheet);
}

/** Read a `cx:chartSpace` part into a model; malformed XML throws `XmlParseError`. */
export { parseChartEx as fromXml } from "@excel/chart/serialize/chart-ex-parser";

/**
 * A model as a `cx:chartSpace` part. A model read from a file is written back as it was read unless
 * `options.forceStructural` asks for it to be rebuilt from its fields.
 *
 * @throws {ExcelChartOptionsError} `model` is not an Office 2016 chart model.
 */
export function toXml(model: ChartExModel, options: { forceStructural?: boolean } = {}): string {
  checkChartExModel("ChartEx.toXml", model);
  if (argumentChecks("ChartEx.toXml").record("options", options)) {
    argumentChecks("ChartEx.toXml").boolean("options.forceStructural", options.forceStructural);
  }
  return renderChartEx(model, options);
}
