/**
 * DOCX Module - Excel Bridge (Subpath Entry)
 *
 * Re-exports the Excel bridge API at `documonster/word/excel`. This file is
 * referenced by `package.json#exports["./word/excel"]`; it forwards to the
 * implementation under `./bridge/excel-bridge`.
 *
 * @example
 * ```ts
 * import { excelToDocx } from "documonster/word/excel";
 * ```
 */

export {
  excelToDocx,
  extractTablesToExcel,
  renderWordChartSvg,
  buildWordChartExXml,
  wordChartToChartModel,
  generateChartEmbeddedXlsx
} from "@word/bridge/excel-bridge";
export type { ExcelToDocxOptions, WordChartExOptions } from "@word/bridge/excel-bridge";
