/**
 * `Csv` namespace surface — CSV parsing, formatting, detection, streaming,
 * row utilities, dynamic typing, generation.
 *
 * `import { Csv } from "@cj-tech-master/excelts/csv"` →
 *   `Csv.parse(text)`, `Csv.parseAsync(...)`, `Csv.format(rows)`,
 *   `Csv.detectDelimiter(text)`, `new Csv.ParserStream()`, …
 *
 * Single flat namespace (csv is a single-purpose module). Re-exported via
 * `export * as Csv`, tree-shaken per-member on rolldown / rspack.
 */

// Type guards / helpers
export { isFormattedValue, quoted, unquoted } from "@csv/types";

// Parsing
export { parseCsv as parse } from "@csv/parse/sync";
export {
  parseCsvAsync as parseAsync,
  parseCsvRows as parseRows,
  parseCsvWithProgress as parseWithProgress
} from "@csv/parse/async";

// Formatting
export { formatCsv as format } from "@csv/format/index";

// Streaming
export {
  CsvParserStream as ParserStream,
  CsvFormatterStream as FormatterStream,
  createCsvParserStream as createParserStream,
  createCsvFormatterStream as createFormatterStream
} from "@csv/stream/index";

// Detection
export { detectDelimiter, detectLinebreak, stripBom } from "@csv/utils/detect";

// Row utilities
export {
  isRowHashArray,
  rowHashArrayToValues,
  rowHashArrayToHeaders,
  rowHashArrayMapByHeaders,
  processColumns,
  deduplicateHeaders,
  deduplicateHeadersWithRenames
} from "@csv/utils/row";

// Dynamic typing
export { applyDynamicTyping, applyDynamicTypingToRow } from "@csv/utils/dynamic-typing";

// Generation
export {
  csvGenerate as generate,
  csvGenerateRows as generateRows,
  csvGenerateAsync as generateAsync,
  csvGenerateData as generateData,
  createCsvGenerator as createGenerator
} from "@csv/utils/generate";

// Number formatting
export {
  formatNumberForCsv as formatNumber,
  parseNumberFromCsv as parseNumber
} from "@csv/utils/number";
