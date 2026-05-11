/** @module Sub-path entry point for direct imports: `import { ... } from "excelts/word/bridge"` */

export {
  excelToDocx,
  extractTablesToExcel,
  renderWordChartSvg,
  buildWordChartExXml,
  generateChartEmbeddedXlsx,
  wordChartToChartModel
} from "./excel-bridge";
export type { ExcelToDocxOptions, WordChartExOptions } from "./excel-bridge";
