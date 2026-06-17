/**
 * `Csv` namespace surface — CSV parsing, formatting, detection, streaming,
 * row utilities, dynamic typing, generation.
 *
 * `import { Csv } from "documonster/csv"` →
 *   `Csv.parse(text)`, `Csv.parseAsync(...)`, `Csv.format(rows)`,
 *   `Csv.detectDelimiter(text)`, `new Csv.ParserStream()`, …
 *
 * Single flat namespace (csv is a single-purpose module). Re-exported via
 * `export * as Csv`, tree-shaken per-member on rolldown / rspack.
 */

// Type guards / helpers
export { isFormattedValue, quoted, unquoted } from "../types";

// Parsing
export { parseCsv as parse } from "../parse/sync";
export {
  parseCsvAsync as parseAsync,
  parseCsvRows as parseRows,
  parseCsvWithProgress as parseWithProgress
} from "../parse/async";

// Formatting
export { formatCsv as format } from "../format/index";

// Streaming
export {
  CsvParserStream as ParserStream,
  CsvFormatterStream as FormatterStream,
  createCsvParserStream as createParserStream,
  createCsvFormatterStream as createFormatterStream
} from "../stream/index";

// Detection
export { detectDelimiter, detectLinebreak, stripBom } from "../utils/detect";

// Row utilities
export {
  isRowHashArray,
  rowHashArrayToValues,
  rowHashArrayToHeaders,
  rowHashArrayMapByHeaders,
  processColumns,
  deduplicateHeaders,
  deduplicateHeadersWithRenames
} from "../utils/row";

// Dynamic typing
export { applyDynamicTyping, applyDynamicTypingToRow } from "../utils/dynamic-typing";

// Generation
export {
  csvGenerate as generate,
  csvGenerateRows as generateRows,
  csvGenerateAsync as generateAsync,
  csvGenerateData as generateData,
  createCsvGenerator as createGenerator
} from "../utils/generate";

// Number formatting
export {
  formatNumberForCsv as formatNumber,
  parseNumberFromCsv as parseNumber
} from "../utils/number";
