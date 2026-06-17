/**
 * `Workbook` namespace surface — browser entry.
 *
 * Same as the Node `surface/workbook.ts` but IO comes from
 * `@excel/xlsx-io.browser` (cross-platform `toXlsxBuffer` / `loadXlsx` /
 * streaming only — no Node file-path `readXlsxFile` / `writeXlsx`).
 */
export {
  createWorkbook as create,
  addWorksheet,
  getWorksheet,
  getWorksheets,
  removeWorksheet,
  addChartsheet,
  addPivotChartsheet,
  getChartsheet,
  getChartsheets,
  removeChartsheet,
  renameChartsheet,
  copyChartsheet,
  replaceChartsheetChart,
  getDefinedNames,
  eachSheet,
  importSheet,
  protectWorkbook as protect,
  unprotectWorkbook as unprotect,
  addExternalLink,
  getExternalLink,
  registerPerson,
  registerFunction,
  unregisterFunction,
  getWorkbookModel as getModel,
  setWorkbookModel as setModel
} from "@excel/workbook.browser";

export {
  toXlsxBuffer,
  loadXlsx,
  readXlsxStream,
  writeXlsxStream,
  getXlsxIo
} from "@excel/xlsx-io.browser";

/** A workbook handle (opaque to consumers). */
export type { WorkbookData as Handle } from "@excel/workbook-core";
