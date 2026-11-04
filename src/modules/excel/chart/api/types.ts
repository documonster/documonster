import type { WorksheetData } from "@excel/core/worksheet-core";

/** Options for `ChartClassic.fillStrRef`. */
export interface ChartFillRefOptions {
  /** The sheet a formula is evaluated from: it scopes sheet-local defined names. */
  readonly worksheet?: WorksheetData;
}

/** Options for `ChartClassic.fillNumRef`. */
export interface ChartFillNumRefOptions extends ChartFillRefOptions {
  /** Read dates against the 1904 epoch; defaults to the workbook's own setting. */
  readonly date1904?: boolean;
}
