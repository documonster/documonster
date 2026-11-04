/**
 * CSV Detection Utilities
 *
 * Auto-detection of CSV characteristics:
 * - Delimiter detection (comma, tab, semicolon, pipe, etc.)
 * - Line ending detection (LF, CRLF, CR)
 * - Quote character normalization
 *
 * This module is part of the csv/utils subsystem:
 * - detect.ts: Auto-detection of CSV format
 * - row.ts: Row format conversions (RowHashArray, headers)
 * - dynamic-typing.ts: Type coercion (string -> number/boolean/date)
 * - number.ts: Number parsing utilities
 * - generate.ts: Test data generation
 */

import { CsvArgumentError } from "@csv/errors";
import { detectDelimiterFor } from "@csv/parse/delimiter-detector";
import { checkDetectDelimiterOptions, checkTextInput } from "@csv/utils/check-options";
import { isRecordArgument } from "@utils/object";

// =============================================================================
// Utility Functions
// =============================================================================

/**
 * Escape special regex characters
 */
export function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Normalize quote option to { enabled, char } form.
 * Centralizes the quote/false/null handling logic.
 */
export function normalizeQuoteOption(option: string | false | null | undefined): {
  enabled: boolean;
  char: string;
} {
  if (option === false || option === null) {
    return { enabled: false, char: "" };
  }
  return { enabled: true, char: option ?? '"' };
}

/**
 * Normalize escape option to { enabled, char } form.
 * Consistent with normalizeQuoteOption API design.
 *
 * @param escapeOption - User's escape option (string, false, null, or undefined)
 * @param quoteChar - The quote character (used as default when escape is undefined)
 * @returns { enabled: boolean, char: string }
 *   - enabled=false, char="" when explicitly disabled (false/null)
 *   - enabled=true, char=quoteChar when undefined (default behavior)
 *   - enabled=true, char=escapeOption when string provided
 */
export function normalizeEscapeOption(
  escapeOption: string | false | null | undefined,
  quoteChar: string
): { enabled: boolean; char: string } {
  if (escapeOption === false || escapeOption === null) {
    return { enabled: false, char: "" };
  }
  return { enabled: true, char: escapeOption ?? quoteChar };
}

// =============================================================================
// Constants
// =============================================================================

/**
 * Characters that trigger formula escaping (CSV injection prevention).
 * Per OWASP recommendations, these characters at the start of a field
 * could be interpreted as formulas by spreadsheet applications.
 *
 * @see https://owasp.org/www-community/attacks/CSV_Injection
 */
const FORMULA_ESCAPE_CHARS = new Set([
  "=", // Equals - formula prefix
  "+", // Plus - formula prefix
  "-", // Minus - formula prefix
  "@", // At - formula prefix
  "\t", // Tab (0x09)
  "\r", // Carriage return (0x0D)
  "\n", // Line feed (0x0A)
  "\uFF1D", // ＝ (full-width equals)
  "\uFF0B", // ＋ (full-width plus)
  "\uFF0D", // － (full-width minus)
  "\uFF20" // ＠ (full-width at)
]);

// =============================================================================
// BOM and Formula Detection
// =============================================================================

/**
 * Apply a `beforeFirstChunk` hook and strip any BOM, in that order.
 *
 * Shared by `Csv.parse` and `CsvParserStream` so the hook cannot be validated on one path and
 * ignored on the other, and so a hook is free to introduce or remove a BOM either way.
 *
 * @throws CsvError when the hook returns something that is neither a string nor nullish.
 */
export function applyFirstChunkPreprocessing(
  input: string,
  beforeFirstChunk: ((chunk: string) => string | void) | undefined,
  where = "Csv.parse"
): string {
  let text = input;
  if (beforeFirstChunk) {
    const result = beforeFirstChunk(text);
    if (typeof result === "string") {
      text = result;
    } else if (result !== undefined && result !== null) {
      throw new CsvArgumentError(
        "options.beforeFirstChunk",
        `${where}: beforeFirstChunk must return a string or undefined, got ${typeof result}`
      );
    }
  }
  return stripBom(text);
}

/**
 * Strip UTF-8 BOM (Byte Order Mark) from start of string if present.
 * Excel exports UTF-8 CSV files with BOM (\ufeff).
 *
 * @param input - String to process
 * @returns String without BOM
 */
export function stripBom(input: string): string {
  checkTextInput("Csv.stripBom", input);
  return input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
}

/**
 * Check if a string starts with a formula escape character.
 * Used for CSV injection prevention.
 */
export function startsWithFormulaChar(str: string): boolean {
  return str.length > 0 && FORMULA_ESCAPE_CHARS.has(str[0]);
}

// =============================================================================
// Line Break Detection
// =============================================================================

/**
 * Detect the line terminator used in a string.
 * Uses quote-aware detection to avoid detecting newlines inside quoted fields.
 *
 * @param input - String to analyze
 * @param options - `quote`: quote character (default `'"'`)
 * @returns Detected line terminator or '\n' as default
 *
 * @example
 * detectLinebreak('a,b\r\nc,d') // '\r\n'
 * detectLinebreak('a,b\nc,d') // '\n'
 * detectLinebreak('a,b\rc,d') // '\r'
 * detectLinebreak('a,b,c') // '\n' (default)
 * detectLinebreak('"a\nb",c\r\nd') // '\r\n' (ignores newline in quotes)
 */
export function detectLinebreak(input: string, options?: CsvDetectLinebreakOptions): string;
/** @deprecated Pass `{ quote }`. The positional form is removed in 1.0. */
export function detectLinebreak(input: string, quote: string): string;
export function detectLinebreak(
  input: string,
  optionsOrQuote?: CsvDetectLinebreakOptions | string
): string {
  checkTextInput("Csv.detectLinebreak", input);
  // Only `undefined` defaults the positional quote: a `null` placeholder there has always meant
  // "no quote character", so line breaks inside quotes are not skipped.
  const quote = isRecordArgument(optionsOrQuote)
    ? (optionsOrQuote.quote ?? '"')
    : optionsOrQuote === undefined
      ? '"'
      : optionsOrQuote;
  let inQuote = false;

  for (let i = 0; i < input.length; i++) {
    const char = input[i];

    // Handle quote toggle (including escaped quotes "")
    if (char === quote) {
      // Check for escaped quote (two consecutive quotes)
      if (inQuote && input[i + 1] === quote) {
        i++; // Skip the escaped quote
        continue;
      }
      inQuote = !inQuote;
      continue;
    }

    // Skip characters inside quotes
    if (inQuote) {
      continue;
    }

    // Detect line ending outside of quotes
    if (char === "\r") {
      return input[i + 1] === "\n" ? "\r\n" : "\r";
    }
    if (char === "\n") {
      return "\n";
    }
  }

  // No line ending found outside quotes, default to \n
  return "\n";
}

// =============================================================================
// Delimiter Detection
// =============================================================================

/** Options for `Csv.detectDelimiter`. */
export interface CsvDetectDelimiterOptions {
  /** Quote character. Default `'"'`. */
  quote?: string;
  /** Candidates, in tie-break order. Default `[",", ";", "\t", "|"]`. */
  delimitersToGuess?: readonly string[];
  /** Comment prefix; commented lines are not scored. */
  comment?: string;
}

/** Options for `Csv.detectLinebreak`. */
export interface CsvDetectLinebreakOptions {
  /** Quote character; line breaks inside quotes are ignored. Default `'"'`. */
  quote?: string;
}

// =============================================================================

/**
 * Auto-detect the delimiter used in a CSV string
 *
 * Algorithm:
 * 1. Sample the first few lines (up to 10) for analysis
 * 2. For each candidate delimiter:
 *    - Count occurrences per line (respecting quotes)
 *    - Check consistency: all lines should have the same count
 *    - Higher count = more fields = better delimiter candidate
 * 3. Choose the delimiter with highest consistent field count
 *
 * Tie-breaking rules (in priority order):
 * 1. Lowest delta (variance) wins - more consistent field counts across lines
 * 2. On delta tie, highest avgFieldCount wins - more fields per row
 * 3. On complete tie, array order wins - first delimiter in delimitersToGuess
 *    (default order: comma, semicolon, tab, pipe)
 *
 * @param input - CSV string to analyze
 * @param options - `quote` (default `'"'`), `delimitersToGuess` (default
 *   `[",", ";", "\t", "|"]`) and `comment`
 * @returns Detected delimiter or first delimiter in list
 *
 * @example
 * detectDelimiter('a,b,c\n1,2,3') // ','
 * detectDelimiter('a;b;c\n1;2;3') // ';'
 * detectDelimiter('a\tb\tc\n1\t2\t3') // '\t'
 * detectDelimiter('a:b:c\n1:2:3', { delimitersToGuess: [':'] }) // ':'
 */
export function detectDelimiter(input: string, options?: CsvDetectDelimiterOptions): string;
/**
 * @deprecated Pass `{ quote, delimitersToGuess, comment }`. The positional form is removed
 * in 1.0.
 * @param skipEmptyLines - Ignored. Kept so existing five-argument calls still compile; an
 *   empty or whitespace-only record is never scored, whatever this says.
 */
export function detectDelimiter(
  input: string,
  quote?: string,
  delimitersToGuess?: string[],
  comment?: string,
  skipEmptyLines?: boolean | "greedy"
): string;
export function detectDelimiter(
  input: string,
  optionsOrQuote?: CsvDetectDelimiterOptions | string,
  delimitersToGuess?: string[],
  comment?: string
  // The deprecated overload's fifth `skipEmptyLines` argument is not read: a record that is
  // empty or only whitespace cannot indicate a delimiter and is never scored. See
  // isScorableDetectionRecord.
): string {
  checkTextInput("Csv.detectDelimiter", input);
  if (isRecordArgument(optionsOrQuote)) {
    checkDetectDelimiterOptions(optionsOrQuote as Record<string, unknown>);
    return detectDelimiterFor(input, {
      quote: optionsOrQuote.quote ?? '"',
      comment: optionsOrQuote.comment,
      delimitersToGuess: optionsOrQuote.delimitersToGuess
    });
  }
  // As in `detectLinebreak`, only `undefined` defaults the positional quote.
  return detectDelimiterFor(input, {
    quote: optionsOrQuote === undefined ? '"' : optionsOrQuote,
    comment,
    delimitersToGuess
  });
}
