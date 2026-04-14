/**
 * CSV Format Configuration
 *
 * Pre-compiled configuration for CSV formatting performance.
 * Creates regex patterns, quote lookup functions, and format settings.
 */

import { CsvError } from "../errors";
import type { CsvFormatOptions, TypeTransformMap } from "../types";
import { escapeRegex, normalizeQuoteOption, normalizeEscapeOption } from "../utils/detect";
import type { DecimalSeparator } from "../utils/number";

/**
 * Escape a string for use inside a regex character class [...].
 * Handles all special characters: \ ] ^ -
 */
function escapeForCharClass(str: string): string {
  return str.replace(/[\\\]^-]/g, "\\$&");
}

// =============================================================================
// Types
// =============================================================================

/**
 * Configuration for quoting specific columns
 */
export type QuoteColumnConfig = boolean | boolean[] | Record<string, boolean>;

/**
 * Pre-compiled regex patterns for CSV formatting performance
 */
export interface CsvFormatRegex {
  /** Regex to check if a field needs quoting (contains delimiter, quote, or newline) */
  needsQuoteRegex: RegExp | null;
  /** Regex to find quote/escape characters for escaping */
  escapeQuoteRegex: RegExp | null;
  /** The escaped quote sequence (escape + quote) */
  escapedQuote: string;
  /** Whether quoting is enabled */
  quoteEnabled: boolean;
  /** The quote character */
  quote: string;
  /** The escape character */
  escape: string;
  /** The delimiter character (for fast path) */
  delimiter: string;
  /** Whether we can use fast string.includes() check */
  useFastCheck: boolean;
}

/**
 * Options for creating format regex patterns
 */
export interface FormatRegexOptions {
  /** The quote character (false/null to disable quoting) */
  quote: string | false | null;
  /** The delimiter character */
  delimiter: string;
  /** The escape character (defaults to quote if not provided) */
  escape?: string | false | null;
}

/**
 * Context for formatting a single field
 */
export interface FormatFieldContext {
  /** Column index */
  index: number;
  /** Header name for this column (if known) */
  header?: string;
  /** Whether this is a header row */
  isHeader: boolean;
  /** Current output row index (for transform context) */
  outputRowIndex: number;
  /** Force quote this field */
  forceQuote: boolean;
  /** Quote all fields (when quoteColumns: true) */
  quoteAll: boolean;
  /** Escape formulae (CSV injection protection) */
  escapeFormulae: boolean;
  /** Decimal separator for number formatting */
  decimalSeparator: DecimalSeparator;
  /** Type transform map */
  transform?: TypeTransformMap;
}

/**
 * Options for formatting a row
 */
export interface FormatRowOptions {
  /** Pre-computed quote lookup function */
  quoteLookup: QuoteLookupFn;
  /** Field delimiter */
  delimiter: string;
  /** Header names for columns (used for transform context) */
  headers?: string[];
  /** Whether this row is a header row */
  isHeader: boolean;
  /** Current output row index (0-based) */
  outputRowIndex: number;
  /** Quote all fields (when quoteColumns: true) */
  quoteAll: boolean;
  /** Escape formulae (CSV injection protection) */
  escapeFormulae: boolean;
  /** Decimal separator for number formatting */
  decimalSeparator: DecimalSeparator;
  /** Type transform map */
  transform?: TypeTransformMap;
}

/**
 * Complete format configuration (shared by batch formatter and CsvFormatterStream)
 */
export interface FormatConfig {
  delimiter: string;
  lineEnding: string;
  quoteAll: boolean;
  escapeFormulae: boolean;
  decimalSeparator: DecimalSeparator;
  writeHeaders: boolean;
  bom: boolean;
  trailingNewline: boolean;
  typeTransform?: TypeTransformMap;
  regex: CsvFormatRegex;
  shouldQuoteColumn: QuoteLookupFn;
  shouldQuoteHeader: QuoteLookupFn;
}

// =============================================================================
// Regex Factory
// =============================================================================

/**
 * Create pre-compiled regex patterns for CSV formatting
 */
export function createFormatRegex(options: FormatRegexOptions): CsvFormatRegex {
  const { quote: quoteOption, delimiter, escape: escapeOption } = options;

  // Use centralized normalization utilities
  const { enabled: quoteEnabled, char: quote } = normalizeQuoteOption(quoteOption);
  const escapeNormalized = normalizeEscapeOption(escapeOption, quote);

  if (!quoteEnabled) {
    return {
      needsQuoteRegex: null,
      escapeQuoteRegex: null,
      escapedQuote: "",
      quoteEnabled: false,
      quote: "",
      escape: "",
      delimiter,
      useFastCheck: false
    };
  }

  // When quoting is enabled, we must have a valid escape character to produce valid CSV.
  // If escape was explicitly disabled (escape: false/null), fall back to quote char (RFC 4180 standard).
  // This ensures internal quotes are always properly escaped as "" rather than producing invalid CSV.
  const escape = escapeNormalized.char || quote;

  // Use fast string.includes() check for single-char delimiter, quote, and escape
  const useFastCheck = delimiter.length === 1 && quote.length === 1 && escape.length === 1;

  return {
    needsQuoteRegex: useFastCheck
      ? null // Will use fast check instead
      : (() => {
          // Build character class content using dedicated char-class escaping
          const classContent = `${escapeForCharClass(delimiter)}${escapeForCharClass(quote)}${escape !== quote ? escapeForCharClass(escape) : ""}\r\n`;
          return new RegExp(`[${classContent}]`);
        })(),
    escapeQuoteRegex:
      escape !== quote
        ? new RegExp(`${escapeRegex(quote)}|${escapeRegex(escape)}`, "g")
        : new RegExp(escapeRegex(quote), "g"),
    escapedQuote: escape + quote,
    quoteEnabled: true,
    quote,
    escape,
    delimiter,
    useFastCheck
  };
}

// =============================================================================
// Quote Lookup
// =============================================================================

/**
 * Pre-compute a quote lookup function for better performance.
 * Avoids repeated type checks on quoteConfig during formatting.
 */
export type QuoteLookupFn = (index: number, header: string | undefined) => boolean;

export function createQuoteLookup(quoteConfig: QuoteColumnConfig | undefined): QuoteLookupFn {
  if (quoteConfig === true) {
    return () => true;
  }
  if (quoteConfig === false || quoteConfig === undefined) {
    return () => false;
  }
  if (Array.isArray(quoteConfig)) {
    return (index: number) => !!quoteConfig[index];
  }
  // Record<string, boolean>
  return (_index: number, header: string | undefined) => (header ? !!quoteConfig[header] : false);
}

// =============================================================================
// Config Factory
// =============================================================================

/**
 * Create complete format configuration from options
 */
export function createFormatConfig(options: CsvFormatOptions): FormatConfig {
  const {
    delimiter = ",",
    lineEnding = "\n",
    quote: quoteOption = '"',
    escape: escapeOption,
    quoteColumns = false,
    quoteHeaders = false,
    writeHeaders: writeHeadersOption,
    bom = false,
    trailingNewline = false,
    escapeFormulae = false,
    decimalSeparator = ".",
    typeTransform
  } = options;

  // Validate decimalSeparator - only "." or "," are valid
  if (decimalSeparator !== "." && decimalSeparator !== ",") {
    throw new CsvError(`Invalid decimalSeparator: "${decimalSeparator}". Must be "." or ",".`);
  }

  // Prevent silent data corruption when decimalSeparator matches delimiter
  if (decimalSeparator === delimiter) {
    throw new CsvError("decimalSeparator cannot be the same as delimiter");
  }

  const regex = createFormatRegex({
    quote: quoteOption,
    delimiter,
    escape: escapeOption
  });

  const quoteAll = quoteColumns === true;

  return {
    delimiter,
    lineEnding,
    quoteAll,
    escapeFormulae,
    decimalSeparator,
    writeHeaders: writeHeadersOption ?? true,
    bom,
    trailingNewline,
    typeTransform,
    regex,
    shouldQuoteColumn: createQuoteLookup(quoteColumns),
    shouldQuoteHeader: createQuoteLookup(quoteHeaders)
  };
}
