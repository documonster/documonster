/**
 * CSV Parse Configuration
 *
 * Defines the ParseConfig type and factory function for creating
 * normalized parsing configuration from user options.
 */

import { DEFAULT_LINEBREAK_REGEX } from "@csv/constants";
import { CsvArgumentError } from "@csv/errors";
import { detectDelimiterFor } from "@csv/parse/delimiter-detector";
import { createOnSkipHandler } from "@csv/parse/helpers";
import type { ScannerConfig } from "@csv/parse/scanner";
import type { CsvParseOptions } from "@csv/types";
import { checkParseOptions } from "@csv/utils/check-options";
import {
  applyFirstChunkPreprocessing,
  normalizeQuoteOption,
  normalizeEscapeOption,
  detectLinebreak
} from "@csv/utils/detect";
import { describeValue as describeInput } from "@utils/describe-value";

// =============================================================================
// Types
// =============================================================================

/**
 * Resolved parsing configuration (after option normalization)
 */
export interface ParseConfig {
  delimiter: string;
  linebreak: string;
  /** Pre-compiled regex for line splitting (used in fast mode) */
  linebreakRegex: RegExp | string;
  quote: string;
  escape: string;
  quoteEnabled: boolean;
  trimField: (s: string) => string;
  /** Whether trimField is an identity function (no actual trimming) - cached for performance */
  trimFieldIsIdentity: boolean;
  shouldSkipEmpty: boolean | "greedy";
  skipLines: number;
  skipRows: number;
  maxRows?: number;
  toLine?: number;
  maxRowBytes?: number;
  comment?: string;
  fastMode: boolean;
  relaxQuotes: boolean;
  columnLess: "error" | "pad";
  columnMore: "error" | "truncate" | "keep";
  groupColumnsByName: boolean;
  skipRecordsWithError: boolean;
  skipRecordsWithEmptyValues: boolean;
  infoOption: boolean;
  rawOption: boolean;
  dynamicTyping: CsvParseOptions["dynamicTyping"];
  castDate: CsvParseOptions["castDate"];
  decimalSeparator: "." | ",";
  invokeOnSkip: ReturnType<typeof createOnSkipHandler>;
  headers: CsvParseOptions["headers"];
}

/**
 * Options for creating ParseConfig.
 * - For batch parsing: provide `input` for auto-detection and BOM stripping
 * - For streaming: omit `input` (will use defaults, detection handled separately)
 */
export interface CreateParseConfigOptions {
  /** Raw input string (for batch parsing with auto-detection) */
  input?: string;
  /** CSV parse options */
  options: CsvParseOptions;
  /** Override delimiter (for streaming after detection) */
  detectedDelimiter?: string;
  /** The public call an argument error names (`Csv.parseStream`). Defaults to `Csv.parse`. */
  where?: string;
}

/**
 * Result of createParseConfig
 */
export interface ParseConfigResult {
  /** Resolved parse configuration */
  config: ParseConfig;
  /** Processed input with BOM stripped and beforeFirstChunk applied (if input was provided) */
  processedInput?: string;
}

// =============================================================================
// Configuration Factory
// =============================================================================

/**
 * Create a normalized ParseConfig from options.
 * This is the single source of truth for configuration normalization,
 * used by both sync and streaming parsers.
 *
 * @example Batch parsing
 * ```ts
 * const { config, processedInput } = createParseConfig({ input: csvString, options });
 * ```
 *
 * @example Streaming parsing
 * ```ts
 * const { config } = createParseConfig({ options });
 * // Later, after delimiter detection:
 * config.delimiter = detectedDelimiter;
 * ```
 */
export function createParseConfig(opts: CreateParseConfigOptions): ParseConfigResult {
  const { input, options, detectedDelimiter, where = "Csv.parse" } = opts;
  checkParseOptions(options, where);
  const {
    delimiter: delimiterOption = ",",
    delimitersToGuess,
    lineEnding: lineEndingOption = "",
    quote: quoteOption = '"',
    escape: escapeOption,
    skipEmptyLines = false,
    trim = false,
    ltrim = false,
    rtrim = false,
    headers = false,
    comment,
    maxRows,
    toLine,
    skipLines = 0,
    skipRows = 0,
    columnMismatch,
    groupColumnsByName = false,
    fastMode = false,
    dynamicTyping,
    castDate,
    decimalSeparator = ".",
    beforeFirstChunk,
    info: infoOption = false,
    raw: rawOption = false,
    relaxQuotes = false,
    skipRecordsWithError = false,
    skipRecordsWithEmptyValues = false,
    onSkip,
    maxRowBytes
  } = options;

  // Column mismatch defaults to strict (error on any mismatch)
  const columnLess = columnMismatch?.less ?? "error";
  const columnMore = columnMismatch?.more ?? "error";

  // Process input if provided (batch mode)
  let processedInput: string | undefined;
  if (input !== undefined) {
    processedInput = input;

    processedInput = applyFirstChunkPreprocessing(processedInput, beforeFirstChunk, where);
  }

  const shouldSkipEmpty = skipEmptyLines;

  // Normalize quote/escape
  const { enabled: quoteEnabled, char: quote } = normalizeQuoteOption(quoteOption);
  const escapeNormalized = normalizeEscapeOption(escapeOption, quote);
  const escape = escapeNormalized.enabled ? escapeNormalized.char || quote : "";

  // Determine linebreak
  const linebreak =
    lineEndingOption || (processedInput !== undefined ? detectLinebreak(processedInput) : "\n");

  // Pre-compile linebreak regex for fast mode
  const linebreakRegex =
    linebreak && linebreak !== "\n" && linebreak !== "\r\n" && linebreak !== "\r"
      ? linebreak
      : DEFAULT_LINEBREAK_REGEX;

  // Determine delimiter
  let delimiter: string;
  if (detectedDelimiter !== undefined) {
    // Use externally detected delimiter (streaming mode)
    delimiter = detectedDelimiter;
  } else if (delimiterOption === "" && processedInput !== undefined) {
    // Auto-detect from input (batch mode). The same detector a stream uses, fed in one go, so
    // the two cannot weigh the candidates differently.
    delimiter = detectDelimiterFor(processedInput, {
      // "" when quoting is disabled, so detection scores by the grammar the parse will use.
      quote: quoteEnabled ? quote : "",
      escape,
      relaxQuotes,
      comment,
      delimitersToGuess,
      // Only fastMode actually ends records at a configured separator; standard mode always
      // ends them at CR/LF, so sampling there must not use it.
      lineEnding: fastMode && linebreakRegex !== DEFAULT_LINEBREAK_REGEX ? linebreak : undefined
    });
  } else if (delimiterOption === "") {
    // Streaming mode with auto-detect - use default, will be updated later
    delimiter = ",";
  } else {
    delimiter = delimiterOption;
  }

  const config: ParseConfig = {
    delimiter,
    linebreak,
    linebreakRegex,
    quote,
    escape,
    quoteEnabled,
    trimField: makeTrimField(trim, ltrim, rtrim),
    // Cache whether trimField is identity to avoid per-row checking
    trimFieldIsIdentity: !trim && !ltrim && !rtrim,
    shouldSkipEmpty,
    skipLines,
    skipRows,
    maxRows,
    toLine,
    maxRowBytes,
    comment,
    fastMode,
    relaxQuotes,
    columnLess,
    columnMore,
    groupColumnsByName,
    skipRecordsWithError,
    skipRecordsWithEmptyValues,
    infoOption,
    rawOption,
    dynamicTyping,
    castDate,
    decimalSeparator,
    invokeOnSkip: createOnSkipHandler(onSkip),
    headers
  };

  return { config, processedInput };
}

export function resolveParseConfig(
  input: string,
  options: CsvParseOptions,
  where = "Csv.parse"
): { config: ParseConfig; processedInput: string } {
  // Otherwise a number or `null` failed inside BOM stripping with an untyped `TypeError`.
  if (typeof input !== "string") {
    throw new CsvArgumentError(
      "input",
      `${where}: input must be a string, got ${describeInput(input)}`
    );
  }
  const result = createParseConfig({ input, options, where });
  return {
    config: result.config,
    processedInput: result.processedInput!
  };
}

// =============================================================================
// Scanner Config Helper
// =============================================================================

/**
 * Convert ParseConfig to ScannerConfig
 */
export function toScannerConfig(config: ParseConfig): ScannerConfig {
  return {
    delimiter: config.delimiter,
    quote: config.quote,
    escape: config.escape,
    quoteEnabled: config.quoteEnabled,
    relaxQuotes: config.relaxQuotes
  };
}

// =============================================================================
// Trim Function Factory
// =============================================================================

/**
 * Create a trim function based on options
 */
export function makeTrimField(
  trim: boolean,
  ltrim: boolean,
  rtrim: boolean
): (s: string) => string {
  if (trim || (ltrim && rtrim)) {
    return (s: string) => s.trim();
  }
  if (ltrim) {
    return (s: string) => s.trimStart();
  }
  if (rtrim) {
    return (s: string) => s.trimEnd();
  }
  return (s: string) => s;
}
