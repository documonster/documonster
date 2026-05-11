/**
 * DOCX Module - Excel Bridge (Subpath Export)
 *
 * Import separately to avoid pulling Excel dependencies into the bundle
 * when only core Word building is needed.
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
} from "./excel-bridge";
export type { ExcelToDocxOptions, WordChartExOptions } from "./excel-bridge";
