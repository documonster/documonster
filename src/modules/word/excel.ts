/**
 * DOCX Module - Excel Bridge (Subpath Entry)
 *
 * Re-exports the Excel bridge API at `excelts/word/excel`. This file is
 * referenced by `package.json#exports["./word/excel"]`; it forwards to the
 * implementation under `./bridge/excel`.
 *
 * @example
 * ```ts
 * import { excelToDocx } from "excelts/word/excel";
 * ```
 */

export {
  excelToDocx,
  extractTablesToExcel,
  renderWordChartSvg,
  buildWordChartExXml,
  wordChartToChartModel,
  generateChartEmbeddedXlsx
} from "./bridge/excel";
export type { ExcelToDocxOptions, WordChartExOptions } from "./bridge/excel";
