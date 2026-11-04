/**
 * DOCX Module - Excel Bridge (Subpath Entry)
 *
 * Public API at `documonster/word/excel`. Referenced by
 * `package.json#exports["./word/excel"]`.
 *
 * @example
 * ```ts
 * import { WordExcel } from "documonster/word/excel";
 *
 * const doc = WordExcel.toDocx(workbook);
 * ```
 */
import {
  buildWordChartExXml as buildWordChartExXmlImpl,
  excelToDocx as excelToDocxImpl,
  extractTablesToExcel as extractTablesToExcelImpl,
  generateChartEmbeddedXlsx as generateChartEmbeddedXlsxImpl,
  renderWordChartSvg as renderWordChartSvgImpl,
  wordChartToChartModel as wordChartToChartModelImpl
} from "@word/bridge/excel-bridge";

export * as WordExcel from "@word/surface/excel";

// The errors this entry's functions throw (also exported from `documonster/word`, the same
// class objects): a wrong argument → `DocxArgumentError`.
export { DocxArgumentError, DocxError, isDocxError } from "@word/errors";
// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
export type {
  ExcelToDocxOptions,
  WordChartExOptions,
  WordChartSvgOptions,
  ChartThemeColors
} from "@word/bridge/excel-bridge";

/** @deprecated Use {@link WordExcel.toDocx}. Removed in 1.0. */
export const excelToDocx = excelToDocxImpl;
/** @deprecated Use {@link WordExcel.extractTables}. Removed in 1.0. */
export const extractTablesToExcel = extractTablesToExcelImpl;
/** @deprecated Use {@link WordExcel.toChartSvg}. Removed in 1.0. */
export const renderWordChartSvg = renderWordChartSvgImpl;
/** @deprecated Use {@link WordExcel.toChartExXml}. Removed in 1.0. */
export const buildWordChartExXml = buildWordChartExXmlImpl;
/** @deprecated Use {@link WordExcel.toChartModel}. Removed in 1.0. */
export const wordChartToChartModel = wordChartToChartModelImpl;
/** @deprecated Use {@link WordExcel.toChartXlsx}. Removed in 1.0. */
export const generateChartEmbeddedXlsx = generateChartEmbeddedXlsxImpl;
