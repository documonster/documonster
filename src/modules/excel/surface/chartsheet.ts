import type { ChartsheetData, ChartsheetModel } from "@excel/core/chartsheet";
import {
  createChartsheet,
  chartsheetSetName,
  chartsheetSetHeaderFooter,
  chartsheetSetPageMargins,
  chartsheetSetPageSetup,
  chartsheetSetState,
  chartsheetSetTabSelected,
  chartsheetSetWorkbookViewId,
  chartsheetSetZoomScale,
  chartsheetSetZoomToFit,
  chartsheetChart,
  chartsheetChartExModel,
  chartsheetChartExNumber,
  chartsheetChartModel,
  chartsheetChartNumber,
  chartsheetHeaderFooter,
  chartsheetId,
  chartsheetModel,
  chartsheetName,
  chartsheetPageMargins,
  chartsheetPageSetup,
  chartsheetSheetNo,
  chartsheetState,
  chartsheetTabSelected,
  chartsheetWorkbookViewId,
  chartsheetZoomScale,
  chartsheetZoomToFit
} from "@excel/core/chartsheet";
/**
 * `Chartsheet` namespace surface — operations on a chartsheet handle.
 *
 * `import { Chartsheet } from "documonster/excel"` →
 *   `Chartsheet.getName(cs)`, `Chartsheet.getChart(cs)`, `Chartsheet.getModel(cs)`.
 *
 * Chartsheet *creation / management* (`addChartsheet`, `getChartsheet`,
 * `removeChartsheet`) is workbook-centric and lives on the `Workbook`
 * namespace; this namespace holds the per-handle accessors.
 *
 * `isTabSelected` / `isZoomToFit` answer `boolean | undefined`: `undefined`
 * means the chartsheet carries no setting (Excel treats it as `false`), and it is
 * what the matching setter accepts to clear the setting, so a read-then-write round
 * trip leaves the attribute absent rather than writing an explicit default.
 */
import { checkChartsheet } from "@excel/core/feature-handles";
import { refuseLongSheetName } from "@excel/core/workbook-core";
import { ExcelArgumentError } from "@excel/errors";
import { argumentChecks } from "@excel/utils/arguments";
import { checkSheetSettings } from "@excel/utils/sheet-settings-checks";

/** Whether the chartsheet holds an Office 2016 chart (histogram, treemap, …) rather than a classic one. */
export { chartsheetIsChartEx as isChartEx } from "@excel/core/chartsheet";

/**
 * @deprecated Use `Workbook.addChartsheet`. This builds a handle over `model` without adding it to a workbook, so it is
 * not listed and not written. Removed in 1.0.
 */
export const create = createChartsheet;

/**
 * Refuse an edit through the handle of a chartsheet that was removed from its workbook: it changed a model nothing
 * writes any more.
 */
function checkAttached(where: string, cs: ChartsheetData): void {
  const wb = cs?._workbook;
  if (wb && !wb._chartsheets.includes(cs._model)) {
    throw new ExcelArgumentError(
      "chartsheet",
      `${where}: chartsheet "${cs._model.name}" was removed from its workbook; it can no longer be edited.`
    );
  }
}

/**
 * Rename the chartsheet, by the rules a worksheet name follows.
 *
 * @throws {WorksheetNameError} when the name is not one Excel accepts or another sheet has it.
 */
export function setName(cs: ChartsheetData, value: string): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setName", cs);
  // The rule `Workbook.addChartsheet` and `Worksheet.setName` apply: a longer name was cut to 31 characters.
  refuseLongSheetName(value);
  chartsheetSetName(cs, value);
}

// The setters below store what they are given, and the chartsheet writer writes it: each is checked against the same
// vocabulary as a worksheet's settings, so a value Excel would repair the file for is refused here.

/** Set the chartsheet's visibility: `"visible"`, `"hidden"` or `"veryHidden"`. */
export function setState(cs: ChartsheetData, value: ChartsheetModel["state"]): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setState", cs);
  checkSheetSettings("Chartsheet.setState", { state: value }, path =>
    path.replace(/^state/, "value")
  );
  chartsheetSetState(cs, value);
}

const MARGIN_KEYS = ["l", "r", "t", "b", "left", "right", "top", "bottom", "header", "footer"];

/** Set the chartsheet's page margins, in inches. */
export function setPageMargins(cs: ChartsheetData, value: ChartsheetModel["pageMargins"]): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setPageMargins", cs);
  const check = argumentChecks("Chartsheet.setPageMargins");
  if (check.record("value", value)) {
    for (const [side, margin] of Object.entries(value)) {
      // A key the margins do not have was kept and never written.
      if (!MARGIN_KEYS.includes(side)) {
        check.fail(`value.${side}`, `one of the margins ${MARGIN_KEYS.join(", ")}`, side);
      }
      check.number(`value.${side}`, margin);
    }
  }
  chartsheetSetPageMargins(cs, value);
}

/** Set the chartsheet's page setup. */
export function setPageSetup(cs: ChartsheetData, value: ChartsheetModel["pageSetup"]): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setPageSetup", cs);
  checkSheetSettings("Chartsheet.setPageSetup", { pageSetup: value }, path =>
    path.replace(/^pageSetup/, "value")
  );
  const check = argumentChecks("Chartsheet.setPageSetup");
  check.string("value.paperHeight", value?.paperHeight);
  check.string("value.paperWidth", value?.paperWidth);
  chartsheetSetPageSetup(cs, value);
}

/** Set the chartsheet's header and footer. */
export function setHeaderFooter(cs: ChartsheetData, value: ChartsheetModel["headerFooter"]): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setHeaderFooter", cs);
  checkSheetSettings("Chartsheet.setHeaderFooter", { headerFooter: value }, path =>
    path.replace(/^headerFooter/, "value")
  );
  chartsheetSetHeaderFooter(cs, value);
}

/** Set whether the chartsheet's tab is selected. */
export function setTabSelected(cs: ChartsheetData, value: boolean | undefined): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setTabSelected", cs);
  argumentChecks("Chartsheet.setTabSelected").boolean("value", value);
  chartsheetSetTabSelected(cs, value);
}

/** Set the chartsheet's zoom, a whole percentage. */
export function setZoomScale(cs: ChartsheetData, value: number | undefined): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setZoomScale", cs);
  argumentChecks("Chartsheet.setZoomScale").number("value", value, { min: 0, integer: true });
  chartsheetSetZoomScale(cs, value);
}

/** Set which workbook view (0-based) the chartsheet's view belongs to. */
export function setWorkbookViewId(cs: ChartsheetData, value: number | undefined): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setWorkbookViewId", cs);
  argumentChecks("Chartsheet.setWorkbookViewId").number("value", value, {
    min: 0,
    integer: true
  });
  chartsheetSetWorkbookViewId(cs, value);
}

/** Set whether the chart is scaled to fill the window. */
export function setZoomToFit(cs: ChartsheetData, value: boolean | undefined): void {
  checkChartsheet(cs);
  checkAttached("Chartsheet.setZoomToFit", cs);
  argumentChecks("Chartsheet.setZoomToFit").boolean("value", value);
  chartsheetSetZoomToFit(cs, value);
}

/** The chartsheet's sheet id, unique among the workbook's sheets. */
export { chartsheetId as getId };
/** The chartsheet's position among the workbook's sheets, as the file numbers it. */
export { chartsheetSheetNo as getSheetNo };
/** The chartsheet's name. */
export { chartsheetName as getName };
/** The chartsheet's visibility, or `undefined` when it carries none (visible). */
export { chartsheetState as getState };
/** The part number of the classic chart it holds (`chart3.xml` → `3`), or `undefined`. */
export { chartsheetChartNumber as getChartNumber };
/** The part number of the Office 2016 chart it holds, or `undefined`. */
export { chartsheetChartExNumber as getChartExNumber };
/** The model of the classic chart it holds — the live model, as `Chart.getModel` returns it — or `undefined`. */
export { chartsheetChartModel as getChartModel };
/** The model of the Office 2016 chart it holds — the live model — or `undefined`. */
export { chartsheetChartExModel as getChartExModel };
/** A handle on the chart it holds, for the `Chart` namespace, or `undefined`. */
export { chartsheetChart as getChart };
/** Whether the chartsheet's tab is selected; `undefined` when it carries no setting. */
export { chartsheetTabSelected as isTabSelected };
/** The chartsheet's zoom, a whole percentage, or `undefined`. */
export { chartsheetZoomScale as getZoomScale };
/** The workbook view (0-based) the chartsheet's view belongs to, or `undefined`. */
export { chartsheetWorkbookViewId as getWorkbookViewId };
/** Whether the chart is scaled to fill the window; `undefined` when it carries no setting. */
export { chartsheetZoomToFit as isZoomToFit };

/**
 * A copy of the chartsheet's model. Change the chartsheet with the setters; the deprecated `Chartsheet.model` returns
 * the live model, through which a rename skipped the sheet-name rules.
 */
export function getModel(cs: ChartsheetData): ChartsheetModel {
  checkChartsheet(cs);
  return structuredClone(chartsheetModel(cs));
}

/** A copy of the chartsheet's page margins, or `undefined`. */
export function getPageMargins(cs: ChartsheetData): ChartsheetModel["pageMargins"] {
  checkChartsheet(cs);
  return structuredClone(chartsheetPageMargins(cs));
}

/** A copy of the chartsheet's page setup, or `undefined`. */
export function getPageSetup(cs: ChartsheetData): ChartsheetModel["pageSetup"] {
  checkChartsheet(cs);
  return structuredClone(chartsheetPageSetup(cs));
}

/** A copy of the chartsheet's header and footer, or `undefined`. */
export function getHeaderFooter(cs: ChartsheetData): ChartsheetModel["headerFooter"] {
  checkChartsheet(cs);
  return structuredClone(chartsheetHeaderFooter(cs));
}

/** A chartsheet handle. */
export type { ChartsheetData as Handle } from "@excel/core/chartsheet";

// --- deprecated names ---

/** @deprecated Use {@link getId}. Removed in 1.0. */
export const id = chartsheetId;
/** @deprecated Use {@link getSheetNo}. Removed in 1.0. */
export const sheetNo = chartsheetSheetNo;
/** @deprecated Use {@link getName}. Removed in 1.0. */
export const name = chartsheetName;
/** @deprecated Use {@link getState}. Removed in 1.0. */
export const state = chartsheetState;
/** @deprecated Use {@link getChartNumber}. Removed in 1.0. */
export const chartNumber = chartsheetChartNumber;
/** @deprecated Use {@link getChartExNumber}. Removed in 1.0. */
export const chartExNumber = chartsheetChartExNumber;
/** @deprecated Use {@link getChartModel}. Removed in 1.0. */
export const chartModel = chartsheetChartModel;
/** @deprecated Use {@link getChartExModel}. Removed in 1.0. */
export const chartExModel = chartsheetChartExModel;
/** @deprecated Use {@link getChart}. Removed in 1.0. */
export const chart = chartsheetChart;
/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const model = chartsheetModel;
/** @deprecated Use {@link getPageMargins}. Removed in 1.0. */
export const pageMargins = chartsheetPageMargins;
/** @deprecated Use {@link getPageSetup}. Removed in 1.0. */
export const pageSetup = chartsheetPageSetup;
/** @deprecated Use {@link getHeaderFooter}. Removed in 1.0. */
export const headerFooter = chartsheetHeaderFooter;
/** @deprecated Use {@link isTabSelected}. Removed in 1.0. */
export const tabSelected = chartsheetTabSelected;
/** @deprecated Use {@link getZoomScale}. Removed in 1.0. */
export const zoomScale = chartsheetZoomScale;
/** @deprecated Use {@link getWorkbookViewId}. Removed in 1.0. */
export const workbookViewId = chartsheetWorkbookViewId;
/** @deprecated Use {@link isZoomToFit}. Removed in 1.0. */
export const zoomToFit = chartsheetZoomToFit;
