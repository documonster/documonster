/**
 * `Workbook` namespace surface — Node entry.
 *
 * `import { Workbook } from "@cj-tech-master/excelts/excel"` → `Workbook.create()`,
 * `Workbook.addWorksheet(wb, name)`, `Workbook.toBuffer(wb)`, `Workbook.read(wb, bytes)`,
 * `Workbook.readFile(wb, path)`, `Workbook.writeFile(wb, path)`, …
 *
 * Re-exports the workbook management functions (de-prefixed) plus the
 * Node xlsx IO (includes file-path `readFile` / `writeFile`).
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
  toBuffer,
  read,
  readFile,
  writeFile,
  readStream,
  writeStream,
  createStreamWriter,
  createStreamReader,
  getXlsxIo
} from "@excel/workbook";

/** A workbook handle (opaque to consumers). */
export type { WorkbookData as Handle } from "@excel/workbook-core";
