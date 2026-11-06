/**
 * `WordExcel` namespace surface — Excel ↔ DOCX, at `documonster/word/excel`.
 *
 * `WordExcel.toDocx(workbook)` lays a workbook out as a document;
 * `WordExcel.extractTables(doc)` returns a document's tables as named rows of cell values, ready for `Row.addMany`;
 * `WordExcel.toChartSvg(chart)` renders a Word chart with the Excel chart engine;
 * `WordExcel.toChartModel(chart)` converts a Word chart to the Excel chart model;
 * `WordExcel.toChartExXml(options)` builds a `cx:chartSpace` part;
 * `WordExcel.toChartXlsx(series)` builds the embedded `.xlsx` (bytes) a chart part links to.
 */
export {
  excelToDocx as toDocx,
  extractTablesToExcel as extractTables,
  renderWordChartSvg as toChartSvg,
  wordChartToChartModel as toChartModel,
  buildWordChartExXml as toChartExXml,
  generateChartEmbeddedXlsx as toChartXlsx
} from "@word/bridge/excel-bridge";
