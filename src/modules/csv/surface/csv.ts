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

import { parseCsvRows, parseCsvWithProgress } from "@csv/parse/async";
import {
  CsvFormatterStream as CsvFormatterStreamImpl,
  CsvParserStream as CsvParserStreamImpl
} from "@csv/stream/index";
import type { CsvFormatOptions, CsvParseOptions } from "@csv/types";
import { quoted as quotedImpl, unquoted as unquotedImpl } from "@csv/types";
import { csvGenerateAsync } from "@csv/utils/generate";
import {
  rowHashArrayMapByHeaders as rowHashArrayMapByHeadersImpl,
  rowHashArrayToHeaders as rowHashArrayToHeadersImpl,
  rowHashArrayToValues as rowHashArrayToValuesImpl
} from "@csv/utils/row";

// Quoting control for a formatted value
export { isFormattedValue, quoted as forceQuote, unquoted as preventQuote } from "@csv/types";

// Parsing
export { parseCsv as parse } from "@csv/parse/sync";
export { parseCsvAsync as parseAsync, parseCsvRows as parseStream } from "@csv/parse/async";

// Formatting
export { formatCsv as format } from "@csv/format/index";

// Detection
export { detectDelimiter, detectLinebreak, stripBom } from "@csv/utils/detect";

// Row utilities
export {
  isRowHashArray,
  getRowValues,
  getRowHeaders,
  getRowValuesByHeaders,
  processColumnsChecked as processColumns,
  deduplicateHeadersChecked as deduplicateHeaders,
  deduplicateHeadersWithRenamesChecked as deduplicateHeadersWithRenames
} from "@csv/utils/row";

// Dynamic typing
export { applyDynamicTyping, applyDynamicTypingToRow } from "@csv/utils/dynamic-typing";

// Generation
export {
  csvGenerate as generate,
  csvGenerateRows as generateRows,
  csvGenerateAsync as generateRowsAsync,
  csvGenerateData as generateData,
  createCsvGenerator as createGenerator
} from "@csv/utils/generate";

// Number formatting
export { formatNumber, parseNumber } from "@csv/utils/number";

/**
 * The streaming CSV parser: `new Csv.ParserStream(options)` — a Transform that takes CSV
 * text or bytes and emits rows.
 */
export const ParserStream = CsvParserStreamImpl;
/** The streaming CSV parser. */
export type ParserStream = CsvParserStreamImpl;
/**
 * The streaming CSV formatter: `new Csv.FormatterStream(options)` — a Transform that takes
 * rows and emits CSV text.
 */
export const FormatterStream = CsvFormatterStreamImpl;
/** The streaming CSV formatter. */
export type FormatterStream = CsvFormatterStreamImpl;
/** @deprecated Use `new Csv.ParserStream(options)`. Removed in 1.0. */
export function createParserStream(options: CsvParseOptions = {}): ParserStream {
  return new CsvParserStreamImpl(options);
}
/** @deprecated Use `new Csv.FormatterStream(options)`. Removed in 1.0. */
export function createFormatterStream(options: CsvFormatOptions = {}): FormatterStream {
  return new CsvFormatterStreamImpl(options);
}
/**
 * @deprecated Use {@link generateRowsAsync}. `Async` marks the asynchronous sibling of a
 * synchronous form, so this read as the sibling of {@link generate}, which returns the whole
 * CSV; what it does is yield rows, as {@link generateRows} does. Removed in 1.0.
 */
export const generateAsync = csvGenerateAsync;
/**
 * @deprecated Use {@link parseAsync} with `{ onProgress }`. Removed in 1.0.
 */
export const parseWithProgress = parseCsvWithProgress;
/**
 * @deprecated Use {@link parseStream}. An incremental reader takes the `Stream` suffix, as
 * `Xml.parseSaxStream` and `Workbook.readStream` do. Removed in 1.0.
 */
export const parseRows = parseCsvRows;
/** @deprecated Use {@link forceQuote}. Removed in 1.0. */
export const quoted = quotedImpl;
/** @deprecated Use {@link preventQuote}. Removed in 1.0. */
export const unquoted = unquotedImpl;
/** @deprecated Use {@link getRowValues}. Removed in 1.0. */
export const rowHashArrayToValues = rowHashArrayToValuesImpl;
/** @deprecated Use {@link getRowHeaders}. Removed in 1.0. */
export const rowHashArrayToHeaders = rowHashArrayToHeadersImpl;
/** @deprecated Use {@link getRowValuesByHeaders}. Removed in 1.0. */
export const rowHashArrayMapByHeaders = rowHashArrayMapByHeadersImpl;
