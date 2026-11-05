/**
 * Markdown Types
 *
 * Centralized type definitions for the Markdown module.
 */

// =============================================================================
// Alignment Types
// =============================================================================

/**
 * Column alignment in a Markdown table.
 * Determined by colon placement in the separator row:
 * - `:---`  = left
 * - `:---:` = center
 * - `---:`  = right
 * - `---`   = none (defaults to left in most renderers)
 */
export type MarkdownAlignment = "left" | "center" | "right" | "none";

// =============================================================================
// Parse Types
// =============================================================================

/**
 * Result of parsing a Markdown table.
 */
export interface MarkdownParseResult {
  /** Parsed header row (column names) */
  headers: string[];

  /** Parsed data rows (each row is an array of cell values) */
  rows: string[][];

  /** Column alignments extracted from the separator row */
  alignments: MarkdownAlignment[];
}

/**
 * Markdown parsing options.
 */
export interface MarkdownParseOptions {
  /**
   * Trim whitespace from cell values.
   * @default true
   */
  trim?: boolean;

  /**
   * Unescape the escapes `Markdown.format` writes: `\|` -> `|`, `\\` -> `\`, `\<` -> `<`.
   * Other Markdown escapes are left as written.
   * @default true
   */
  unescape?: boolean;

  /**
   * Skip empty rows in the output.
   * @default true
   */
  skipEmptyRows?: boolean;

  /**
   * Maximum number of data rows to parse (excludes header).
   * Useful for previewing large tables.
   */
  maxRows?: number;

  /**
   * Convert `<br>`, `<br/>`, `<br />` tags in cell content to newline characters. Set it to
   * read back what `Markdown.format` writes: it writes a newline as `<br>` and a literal
   * `<br>` as `\<br>`, which is not converted.
   * @default false
   */
  convertBr?: boolean;
}

// =============================================================================
// Format Types
// =============================================================================

/**
 * Column configuration for Markdown formatting.
 */
export interface MarkdownColumnConfig {
  /** Column header text */
  header: string;

  /**
   * Column alignment.
   * @default "left"
   */
  alignment?: MarkdownAlignment;

  /**
   * Minimum column width (in characters, excluding padding).
   * The actual width will be the maximum of this value and the widest cell content.
   * @default 3 (minimum for separator `---`)
   */
  minWidth?: number;
}

/**
 * Markdown formatting options.
 */
export interface MarkdownFormatOptions {
  /**
   * Column configuration. When provided, its headers and alignments replace `headers` column
   * by column; the table has as many columns as the longer of the two. Can be an array of
   * strings (header names) or MarkdownColumnConfig objects.
   */
  columns?: readonly (string | MarkdownColumnConfig)[];

  /**
   * Default alignment for columns without explicit alignment.
   * @default "left"
   */
  alignment?: MarkdownAlignment;

  /**
   * Align columns to equal width by padding cell content with spaces.
   * When false, disables width-alignment padding but retains the single space
   * around cell content required by most Markdown renderers (`| value |`).
   * @default true
   */
  padding?: boolean;

  /**
   * Include a trailing newline at the end of the output.
   * @default true
   */
  trailingNewline?: boolean;

  /**
   * Escape pipe characters (`|`) and backslashes (`\`) in cell content.
   * @default true
   */
  escapeContent?: boolean;

  /**
   * Custom value-to-string converter.
   * Called for each cell value before formatting.
   * Return value is used as the cell text.
   */
  stringify?: (value: unknown) => string;
}

// =============================================================================
// Workbook Integration Types
// =============================================================================

/**
 * Unified Markdown options for Workbook integration.
 * Combines parse and format options with worksheet-specific settings.
 *
 * Published from `documonster/markdown` only as a deprecated name; the Excel bridge's
 * deprecated flat forms still take it. Its successors are the bridge's `ExcelMarkdownReadOptions`
 * and `ExcelMarkdownWriteOptions`.
 */
export interface MarkdownOptions extends MarkdownParseOptions, MarkdownFormatOptions {
  // === Worksheet ===
  /** Name of the worksheet to read from or write to */
  sheetName?: string;
  /** ID of the worksheet to read from or write to */
  sheetId?: number;

  // === Value mapping ===
  /** Custom value mapper for parsing (MD string -> cell value) */
  map?: (value: string, column: number) => unknown;

  // === Write options ===
  /** Date format string for formatting Date values */
  dateFormat?: string;
  /**
   * Take a date's fields from UTC — the calendar value an Excel date cell holds — rather than from the host's
   * timezone. Default `true`; `false` writes local time.
   */
  dateUTC?: boolean;
  /** Include empty rows in output */
  includeEmptyRows?: boolean;
}
