/**
 * `Workbook` namespace surface — Node entry.
 *
 * `import { Workbook } from "documonster/excel"` → `Workbook.create()`,
 * `Workbook.addWorksheet(wb, name)`, `Workbook.toXlsxBuffer(wb)`, …
 *
 * Re-exports the workbook management functions (de-prefixed) plus the
 * Node xlsx IO (includes file-path `readXlsxFile` / `writeXlsx`).
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
  readXlsxFile,
  writeXlsx,
  readXlsxStream,
  writeXlsxStream,
  createStreamWriter,
  createStreamReader,
  getXlsxIo
} from "@excel/workbook";

/** A workbook handle (opaque to consumers). */
export type { WorkbookData as Handle } from "@excel/workbook-core";
