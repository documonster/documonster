/**
 * `Csv` namespace surface — CSV parsing, formatting, detection, streaming,
 * row utilities, dynamic typing, generation.
 *
 * `import { Csv } from "documonster/csv"` →
 *   `Csv.parse(text)`, `Csv.parseAsync(...)`, `Csv.format(rows)`,
 *   `Csv.detectDelimiter(text)`, `Csv.createParserStream()`, …
 *
 * Single flat namespace (csv is a single-purpose module). Re-exported via
 * `export * as Csv`, tree-shaken per-member on rolldown / rspack.
 */

import { parseCsvWithProgress } from "@csv/parse/async";
import {
  CsvFormatterStream as CsvFormatterStreamImpl,
  CsvParserStream as CsvParserStreamImpl
} from "@csv/stream/index";
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
export { parseCsvAsync as parseAsync, parseCsvRows as parseRows } from "@csv/parse/async";

// Formatting
export { formatCsv as format } from "@csv/format/index";

// Streaming
export {
  createCsvParserStream as createParserStream,
  createCsvFormatterStream as createFormatterStream
} from "@csv/stream/index";

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
 * The streaming parser class. {@link createParserStream} is the usual way to build one;
 * the class is public for `instanceof` checks and subclassing.
 */
export const ParserStream = CsvParserStreamImpl;
/** The streaming parser {@link createParserStream} returns. */
export type ParserStream = CsvParserStreamImpl;
/**
 * The streaming formatter class. {@link createFormatterStream} is the usual way to build
 * one; the class is public for `instanceof` checks and subclassing.
 */
export const FormatterStream = CsvFormatterStreamImpl;
/** The streaming formatter {@link createFormatterStream} returns. */
export type FormatterStream = CsvFormatterStreamImpl;
/**
 * @deprecated Use {@link generateRowsAsync}: it yields rows like {@link generateRows}, and
 * `*Async` alone names a Promise elsewhere in this namespace (`parseAsync`). Removed in 1.0.
 */
export const generateAsync = csvGenerateAsync;
/**
 * @deprecated Use {@link parseAsync} with `{ onProgress }`. Removed in 1.0.
 */
export const parseWithProgress = parseCsvWithProgress;
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
