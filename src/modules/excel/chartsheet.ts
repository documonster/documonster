import { getChartSupport } from "@excel/chart-host-registry";
import type { Chart } from "@excel/chart/chart";
import type { AddChartExOptions, ChartExModel } from "@excel/chart/chart-ex-types";
import type { AddChartOptions, AddComboChartOptions, ChartModel } from "@excel/chart/types";
import type { Worksheet } from "@excel/worksheet";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";

/**
 * Nominal anchor range for the full-sheet chart in a chartsheet.
 * Chartsheet charts fill the entire sheet viewport; the anchor value is
 * never used for positioning but is required by the Chart constructor to
 * initialise its internal two-cell-anchor model.
 */
const CHARTSHEET_ANCHOR_RANGE = "A1:K15";

export interface ChartsheetViewOptions {
  /** Whether the chartsheet tab is selected in the workbook view. */
  tabSelected?: boolean;
  /** Zoom scale percentage. */
  zoomScale?: number;
  /**
   * 0-based index into the workbook's `<bookViews>` list that this
   * chartsheet view is bound to. Defaults to 0 (the primary view).
   */
  workbookViewId?: number;
  /**
   * When true the chartsheet scales to fill the window. Matches the
   * OOXML `<sheetView zoomToFit="1"/>` attribute.
   */
  zoomToFit?: boolean;
}

export interface ChartsheetOptions extends ChartsheetViewOptions {
  /** Sheet visibility state. */
  state?: "visible" | "hidden" | "veryHidden";
  /** Page margins for the chartsheet. */
  pageMargins?: ChartsheetModel["pageMargins"];
  /** Page setup for the chartsheet (CT_CsPageSetup). */
  pageSetup?: ChartsheetModel["pageSetup"];
}

export interface AddChartsheetOptions extends ChartsheetOptions {
  /** Chart to place on the chartsheet. */
  chart: AddChartOptions | AddComboChartOptions | AddChartExOptions;
}

export interface AddPivotChartsheetOptions extends ChartsheetOptions {
  /** Classic chart or combo chart to place on the chartsheet as a pivot chart. */
  chart: AddChartOptions | AddComboChartOptions;
}

/**
 * Minimal `Worksheet`-shaped proxy used exclusively to back the `Chart`
 * instance returned from {@link Chartsheet.chart}.
 *
 * A chartsheet is a top-level sheet that hosts a single chart without a
 * surrounding grid — there are no rows, columns, tables, merges, or
 * drawings underneath it. `Chart`, however, assumes it lives inside a
 * `Worksheet` so it can walk up to the workbook, copy relationship
 * sidecars, and compute anchor geometry.
 *
 * Previous versions fabricated an empty `{ workbook, name } as any` to
 * satisfy that type. That escape hatch was a correctness trap: calling
 * {@link Chart.copyTo} or anything that invoked `worksheet.getRow` on a
 * chartsheet-backed Chart would throw because the underlying object had
 * none of the real methods.
 *
 * `ChartsheetChartHost` replaces the cast with a first-class proxy that:
 *   - forwards the `workbook` pointer so `Chart` can reach its model,
 *     sidecars, and copyChartSidecars helper;
 *   - reports a synthetic `id`/`name` so diagnostics stay sensible;
 *   - throws a clear, descriptive error from every grid-centric method so
 *     any accidental reach through fails loudly instead of silently
 *     corrupting state (`getRow`, `addRow`, `getColumn`, `addTable`, …).
 *
 * Callers should prefer {@link Chartsheet.chartModel} /
 * {@link Chartsheet.chartExModel} to reach the underlying data directly.
 */
class ChartsheetChartHost {
  readonly id: number;
  readonly name: string;
  readonly workbook: ChartsheetWorkbook;

  constructor(workbook: ChartsheetWorkbook, chartsheetId: number, chartsheetName: string) {
    this.id = chartsheetId;
    this.name = chartsheetName;
    this.workbook = workbook;
  }

  private _unsupported(method: string): never {
    throw new Error(
      `${method}() is not supported on a Chart attached to a chartsheet. ` +
        `Use the Chartsheet APIs (chartsheet.chartModel, chartsheet.replaceChart, …) instead.`
    );
  }

  // Grid operations — explicitly fail fast rather than silently returning
  // undefined-ish values that could feed bad data into anchor math.
  getRow(_rowNumber: number): never {
    this._unsupported("Worksheet.getRow");
  }
  addRow(): never {
    this._unsupported("Worksheet.addRow");
  }
  addRows(): never {
    this._unsupported("Worksheet.addRows");
  }
  getColumn(_colNumber: number): never {
    this._unsupported("Worksheet.getColumn");
  }
  getCell(): never {
    this._unsupported("Worksheet.getCell");
  }
  addTable(): never {
    this._unsupported("Worksheet.addTable");
  }
  getTables(): never {
    this._unsupported("Worksheet.getTables");
  }
}

class Chartsheet {
  private readonly _model: ChartsheetModel;
  private readonly _workbook?: ChartsheetWorkbook;

  constructor(model: ChartsheetModel, workbook?: ChartsheetWorkbook) {
    this._model = model;
    this._workbook = workbook;
  }

  get id(): number {
    return this._model.id;
  }

  get sheetNo(): number {
    return this._model.sheetNo;
  }

  get name(): string {
    return this._model.name;
  }

  set name(value: string) {
    // Go through the workbook's unified validator so a chartsheet can
    // never silently land on a name that collides with a worksheet (or
    // another chartsheet), includes illegal characters, or exceeds
    // Excel's 31-char limit. Previously this setter just wrote
    // `this._model.name = value` verbatim, letting callers corrupt the
    // model into a state Excel would reject on reopen.
    if (this._workbook) {
      this._model.name = this._workbook.validateSheetName(value, this._model);
    } else {
      this._model.name = value;
    }
  }

  get state(): ChartsheetModel["state"] {
    return this._model.state;
  }

  set state(value: ChartsheetModel["state"]) {
    this._model.state = value;
  }

  get chartNumber(): number | undefined {
    return this._model.chartNumber;
  }

  get chartExNumber(): number | undefined {
    return this._model.chartExNumber;
  }

  get chartModel(): ChartModel | undefined {
    return this._model.chartNumber
      ? this._workbook?.getChartEntry(this._model.chartNumber)?.model
      : undefined;
  }

  get chartExModel(): ChartExModel | undefined {
    return this._model.chartExNumber
      ? this._workbook?.getChartExStructuredEntry(this._model.chartExNumber)?.model
      : undefined;
  }

  /**
   * Return the `Chart` wrapper for this chartsheet's single chart.
   *
   * The wrapper's `.worksheet` field is a {@link ChartsheetChartHost}
   * rather than a real `Worksheet` — grid operations on it will throw a
   * descriptive error. For anything grid-related, use the underlying
   * {@link chartModel}/{@link chartExModel} directly or go through the
   * `Chartsheet` methods ({@link replaceChart}, {@link rename}, …).
   */
  get chart(): Chart | undefined {
    if (!this._workbook || (!this._model.chartNumber && !this._model.chartExNumber)) {
      return undefined;
    }
    const host = new ChartsheetChartHost(this._workbook, this._model.id, this._model.name);
    // The `Chart` constructor types its first argument as `Worksheet`, but
    // internally it only touches `worksheet.workbook`, `worksheet.id`, and
    // `worksheet.name` from that position (plus anchor helpers that we've
    // stubbed above). Casting is safe here because the host deliberately
    // implements that exact contract and loudly rejects anything else.
    return getChartSupport().createChart(
      host as unknown as Worksheet,
      { chartNumber: this._model.chartNumber, chartExNumber: this._model.chartExNumber },
      CHARTSHEET_ANCHOR_RANGE
    );
  }

  get isChartEx(): boolean {
    return !!this._model.chartExNumber;
  }

  get model(): ChartsheetModel {
    return this._model;
  }

  get pageMargins(): ChartsheetModel["pageMargins"] {
    return this._model.pageMargins;
  }

  set pageMargins(value: ChartsheetModel["pageMargins"]) {
    this._model.pageMargins = value;
  }

  get pageSetup(): ChartsheetModel["pageSetup"] {
    return this._model.pageSetup;
  }

  set pageSetup(value: ChartsheetModel["pageSetup"]) {
    this._model.pageSetup = value;
  }

  get tabSelected(): boolean | undefined {
    return this._model.tabSelected;
  }

  set tabSelected(value: boolean | undefined) {
    this._model.tabSelected = value;
  }

  get zoomScale(): number | undefined {
    return this._model.zoomScale;
  }

  set zoomScale(value: number | undefined) {
    this._model.zoomScale = value;
  }

  get workbookViewId(): number | undefined {
    return this._model.workbookViewId;
  }

  set workbookViewId(value: number | undefined) {
    this._model.workbookViewId = value;
  }

  get zoomToFit(): boolean | undefined {
    return this._model.zoomToFit;
  }

  set zoomToFit(value: boolean | undefined) {
    this._model.zoomToFit = value;
  }

  rename(name: string): boolean {
    return this._workbook?.renameChartsheet(this.name, name) ?? false;
  }

  remove(): boolean {
    return this._workbook?.removeChartsheet(this.name) ?? false;
  }

  copy(name?: string): Chartsheet | undefined {
    return this._workbook?.copyChartsheet(this.name, name);
  }

  replaceChart(chart: AddChartsheetOptions["chart"]): boolean {
    return this._workbook?.replaceChartsheetChart(this.name, chart) ?? false;
  }
}

interface ChartsheetWorkbook {
  getChartEntry(chartNumber: number): { model: ChartModel } | undefined;
  getChartExStructuredEntry(chartExNumber: number): { model: ChartExModel } | undefined;
  renameChartsheet(nameOrIndex: string | number, name: string): boolean;
  removeChartsheet(nameOrIndex: string | number): boolean;
  copyChartsheet(nameOrIndex: string | number, name?: string): Chartsheet | undefined;
  replaceChartsheetChart(
    nameOrIndex: string | number,
    chart: AddChartsheetOptions["chart"]
  ): boolean;
  /**
   * Validate a sheet name against Excel's unified namespace (both
   * worksheets and chartsheets). See `Workbook.validateSheetName` for
   * the full contract. Declared here so `Chartsheet.name` setter can
   * route through it without widening the narrow host interface to
   * the full `Workbook` type.
   */
  validateSheetName(name: string, existing?: { name: string }): string;
}

export { Chartsheet };
