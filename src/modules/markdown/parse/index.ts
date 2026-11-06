/**
 * Markdown Table Parser
 *
 * Parses Markdown tables into structured data.
 *
 * Supports:
 * - Standard GFM (GitHub Flavored Markdown) table syntax
 * - Column alignment detection via separator row
 * - Escaped pipes (`\|`) in cell content
 * - Tables with or without leading/trailing pipes
 * - Tolerant parsing (mismatched column counts, extra whitespace)
 * - Multiline cell content via `<br>` / `<br/>` / `<br />` tags
 *
 * @example
 * ```ts
 * const result = Markdown.parse("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
 * // result.headers = ["Name", "Age"]
 * // result.rows = [["Alice", "30"]]
 * // result.alignments = ["none", "none"]
 * ```
 */

import { checkParseArguments } from "@markdown/arguments";
import { CELL_ESCAPE_OR_BR_REGEX, LINEBREAK_REGEX } from "@markdown/constants";
import { MarkdownParseError } from "@markdown/errors";
import type { MarkdownAlignment, MarkdownParseOptions, MarkdownParseResult } from "@markdown/types";

// =============================================================================
// Character Codes (avoid repeated charCodeAt comparisons with magic numbers)
// =============================================================================

const CH_PIPE = 0x7c; // |
const CH_BACKSLASH = 0x5c; // \
const CH_COLON = 0x3a; // :
const CH_DASH = 0x2d; // -
const CH_SPACE = 0x20; // space
const CH_TAB = 0x09; // tab

// =============================================================================
// Resolved Options (shared between parseMarkdown and parseMarkdownAll)
// =============================================================================

/** Internal resolved options — avoids re-destructuring in every call. */
interface ResolvedParseOpts {
  readonly trim: boolean;
  readonly unescape: boolean;
  readonly skipEmpty: boolean;
  readonly maxRows: number | undefined;
  readonly convertBr: boolean;
}

function resolveParseOpts(
  where: string,
  input: unknown,
  options: MarkdownParseOptions
): ResolvedParseOpts {
  checkParseArguments(where, input, options);
  return {
    trim: options.trim !== false,
    unescape: options.unescape !== false,
    skipEmpty: options.skipEmptyRows !== false,
    maxRows: options.maxRows,
    convertBr: options.convertBr === true
  };
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Split a Markdown table row into cell values.
 * Handles escaped pipes (`\|`) correctly by scanning character by character.
 *
 * Optimized: uses start/end index tracking with `slice()` instead of
 * character-by-character string concatenation to avoid O(n²) worst case.
 */
function splitRow(line: string): string[] {
  const cells: string[] = [];
  const len = line.length;

  // Skip leading pipe
  let i = len > 0 && line.charCodeAt(0) === CH_PIPE ? 1 : 0;

  // Check for trailing pipe (to exclude it from the last cell)
  // Must count consecutive backslashes before the pipe:
  // even count (0, 2, 4...) → backslashes are all escaped, pipe is real
  // odd count (1, 3, 5...) → last backslash escapes the pipe
  let end = len;
  if (len > 1 && line.charCodeAt(len - 1) === CH_PIPE) {
    let backslashCount = 0;
    let k = len - 2;
    while (k >= 0 && line.charCodeAt(k) === CH_BACKSLASH) {
      backslashCount++;
      k--;
    }
    if (backslashCount % 2 === 0) {
      end = len - 1;
    }
  }

  // Track the start of the current cell segment
  // We collect segments (between escape sequences) to minimize allocations
  let segStart = i;
  let current = "";
  let hasEscape = false;

  while (i < end) {
    const ch = line.charCodeAt(i);

    if (ch === CH_BACKSLASH && i + 1 < end) {
      // Escape sequence: flush preceding segment, add escape pair
      hasEscape = true;
      if (i > segStart) {
        current += line.slice(segStart, i);
      }
      current += line.slice(i, i + 2);
      i += 2;
      segStart = i;
    } else if (ch === CH_PIPE) {
      // Cell boundary: flush and push
      if (hasEscape) {
        if (i > segStart) {
          current += line.slice(segStart, i);
        }
        cells.push(current);
        current = "";
        hasEscape = false;
      } else {
        cells.push(line.slice(segStart, i));
      }
      i++;
      segStart = i;
    } else {
      i++;
    }
  }

  // Push the last cell
  if (hasEscape) {
    if (end > segStart) {
      current += line.slice(segStart, end);
    }
    cells.push(current);
  } else {
    cells.push(line.slice(segStart, end));
  }

  return cells;
}

/**
 * Determine column alignment from a separator cell.
 *
 * - `:---:` → center
 * - `:---`  → left
 * - `---:`  → right
 * - `---`   → none
 */
function parseAlignment(cell: string): MarkdownAlignment {
  const trimmed = cell.trim();
  const tLen = trimmed.length;
  if (tLen === 0) {
    return "none";
  }
  const leftColon = trimmed.charCodeAt(0) === CH_COLON;
  const rightColon = trimmed.charCodeAt(tLen - 1) === CH_COLON;

  if (leftColon && rightColon) {
    return "center";
  }
  if (leftColon) {
    return "left";
  }
  if (rightColon) {
    return "right";
  }
  return "none";
}

/**
 * Check if a cell string is a valid separator cell.
 * Hand-rolled check replacing regex for better performance.
 * Pattern: optional whitespace, optional colon, one or more dashes, optional colon, optional whitespace.
 */
function isSeparatorCell(cell: string): boolean {
  const len = cell.length;
  let i = 0;

  // Skip leading whitespace
  while (i < len) {
    const ch = cell.charCodeAt(i);
    if (ch !== CH_SPACE && ch !== CH_TAB) {
      break;
    }
    i++;
  }

  // Optional leading colon
  if (i < len && cell.charCodeAt(i) === CH_COLON) {
    i++;
  }

  // At least one dash required
  const dashStart = i;
  while (i < len && cell.charCodeAt(i) === CH_DASH) {
    i++;
  }
  if (i === dashStart) {
    return false;
  }

  // Optional trailing colon
  if (i < len && cell.charCodeAt(i) === CH_COLON) {
    i++;
  }

  // Skip trailing whitespace
  while (i < len) {
    const ch = cell.charCodeAt(i);
    if (ch !== CH_SPACE && ch !== CH_TAB) {
      return false;
    }
    i++;
  }

  return true;
}

/**
 * Check if a line is a valid separator row.
 * A separator row consists entirely of cells matching the pattern `:?-+:?`.
 */
function isSeparatorRow(cells: string[]): boolean {
  if (cells.length === 0) {
    return false;
  }
  for (let i = 0; i < cells.length; i++) {
    if (!isSeparatorCell(cells[i])) {
      return false;
    }
  }
  return true;
}

/**
 * Process cell content: trim, optionally unescape, and optionally convert `<br>` to newlines.
 */
function processCell(value: string, opts: ResolvedParseOpts): string {
  const result = opts.trim ? value.trim() : value;
  if (!opts.unescape && !opts.convertBr) {
    return result;
  }
  return result.replace(CELL_ESCAPE_OR_BR_REGEX, (match: string, escaped: string | undefined) =>
    escaped !== undefined ? (opts.unescape ? escaped : match) : opts.convertBr ? "\n" : match
  );
}

/**
 * Normalize a row to the expected column count.
 * - If row has fewer cells, pad with empty strings
 * - If row has more cells, truncate
 */
function normalizeRow(cells: string[], columnCount: number, opts: ResolvedParseOpts): string[] {
  const row: string[] = new Array(columnCount);
  for (let i = 0; i < columnCount; i++) {
    row[i] = i < cells.length ? processCell(cells[i], opts) : "";
  }
  return row;
}

/**
 * Check if a row is empty (all cells are empty strings).
 */
function isEmptyRow(row: string[]): boolean {
  for (let i = 0; i < row.length; i++) {
    if (row[i] !== "") {
      return false;
    }
  }
  return true;
}

/**
 * Check if a line could be part of a table (contains a pipe character).
 */
function isTableLine(line: string): boolean {
  return line.indexOf("|") !== -1;
}

/**
 * Check if a line starts with a pipe (after optional leading whitespace).
 * Used to determine the table's "piped" style for data row validation.
 */
function startsWithPipe(line: string): boolean {
  const len = line.length;
  let i = 0;
  while (i < len) {
    const ch = line.charCodeAt(i);
    if (ch !== CH_SPACE && ch !== CH_TAB) {
      return ch === CH_PIPE;
    }
    i++;
  }
  return false;
}

/**
 * Check if a line could be a separator candidate (contains a dash).
 */
function hasDash(line: string): boolean {
  return line.indexOf("-") !== -1;
}

// =============================================================================
// Code blocks
// =============================================================================

/** An opening code fence (CommonMark §4.5): up to three spaces, then three or more ` or ~. */
const FENCE_OPEN = /^ {0,3}(`{3,}|~{3,})/;

/**
 * The line after the fenced code block opening at `start`, or `start` when no fence opens
 * there. An unclosed fence runs to the end of the document, as CommonMark reads it. A table
 * inside a code block is code, not a table.
 */
function skipFencedCode(lines: string[], start: number): number {
  const open = FENCE_OPEN.exec(lines[start]);
  if (!open) {
    return start;
  }
  const fence = open[1];
  if (fence[0] === "`" && lines[start].slice(open[0].length).includes("`")) {
    return start; // a backtick fence's info string cannot hold a backtick
  }
  const close = new RegExp(`^ {0,3}${fence[0] === "`" ? "`" : "~"}{${fence.length},}[ \\t]*$`);
  for (let i = start + 1; i < lines.length; i++) {
    if (close.test(lines[i])) {
      return i + 1;
    }
  }
  return lines.length;
}

/** Whether a line is indented four columns or more — an indented code block, not a table. */
function isIndentedCode(line: string): boolean {
  let columns = 0;
  for (let i = 0; i < line.length && columns < 4; i++) {
    const ch = line.charCodeAt(i);
    if (ch === CH_SPACE) {
      columns++;
    } else if (ch === CH_TAB) {
      columns += 4 - (columns % 4);
    } else {
      return false;
    }
  }
  return columns >= 4;
}

// =============================================================================
// Core Table Parser (shared between parseMarkdown and parseMarkdownAll)
// =============================================================================

/**
 * Attempt to parse a table starting at line index `startLine`.
 *
 * Returns `{ result, endLine }` if a valid table starts here, or `null` otherwise.
 */
function parseTableAt(
  lines: string[],
  startLine: number,
  lineCount: number,
  opts: ResolvedParseOpts
): { result: MarkdownParseResult; endLine: number } | null {
  if (startLine >= lineCount - 1) {
    return null;
  }

  if (isIndentedCode(lines[startLine])) {
    return null;
  }
  const line = lines[startLine].trim();

  // Skip empty lines and non-table content
  if (line === "" || !isTableLine(line)) {
    return null;
  }

  // Candidate header row
  const headerCells = splitRow(line);
  if (headerCells.length < 1) {
    return null;
  }

  // Check if the next line is a valid separator row
  const separatorLine = lines[startLine + 1].trim();
  if (separatorLine === "" || !hasDash(separatorLine)) {
    return null;
  }

  const separatorCells = splitRow(separatorLine);
  if (!isSeparatorRow(separatorCells)) {
    return null;
  }

  // Valid table found — extract headers and alignments
  const columnCount = headerCells.length;
  const headers: string[] = new Array(columnCount);
  const alignments: MarkdownAlignment[] = new Array(columnCount);

  for (let c = 0; c < columnCount; c++) {
    headers[c] = processCell(headerCells[c], opts);
    alignments[c] = c < separatorCells.length ? parseAlignment(separatorCells[c]) : "none";
  }

  // Determine if this is a "piped" table (header starts with `|`).
  // When the header has a leading pipe, data rows must also start with `|`.
  // This prevents prose like "This has a | pipe" from being swallowed as data.
  const piped = startsWithPipe(line);

  // Parse data rows
  const rows: string[][] = [];
  let j = startLine + 2;

  for (; j < lineCount; j++) {
    const dataLine = lines[j].trim();

    // Stop at empty line or non-table content (end of table)
    if (dataLine === "" || !isTableLine(dataLine)) {
      break;
    }

    // For piped tables, data rows must also start with `|`
    if (piped && !startsWithPipe(dataLine)) {
      break;
    }

    // Check maxRows limit
    if (opts.maxRows !== undefined && rows.length >= opts.maxRows) {
      // Skip remaining table rows for parseMarkdownAll to correctly advance
      while (j < lineCount) {
        const remaining = lines[j].trim();
        if (remaining === "" || !isTableLine(remaining)) {
          break;
        }
        if (piped && !startsWithPipe(remaining)) {
          break;
        }
        j++;
      }
      break;
    }

    const dataCells = splitRow(dataLine);
    const row = normalizeRow(dataCells, columnCount, opts);

    if (opts.skipEmpty && isEmptyRow(row)) {
      continue;
    }

    rows.push(row);
  }

  return { result: { headers, rows, alignments }, endLine: j };
}

// =============================================================================
// Main Parser
// =============================================================================

/**
 * Parse a Markdown table string into structured data.
 *
 * The parser looks for the GFM table pattern:
 * 1. A header row (pipe-delimited cells)
 * 2. A separator row (dashes with optional colons for alignment)
 * 3. Zero or more data rows
 *
 * Non-table content before and after the table is ignored, and so is a table inside a
 * fenced (```` ``` ````, `~~~`) or indented code block: GFM renders that as code.
 *
 * @param input - Markdown string containing a table
 * @param options - Parse options
 * @returns Parsed table data with headers, rows, and alignments
 *
 * @throws {MarkdownParseError} When no valid table is found in the input
 *
 * @example
 * ```ts
 * // Basic table
 * const result = Markdown.parse("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
 *
 * // With alignment
 * const result = Markdown.parse("| Left | Center | Right |\n|:---|:---:|---:|\n|a|b|c|");
 * // result.alignments = ["left", "center", "right"]
 *
 * // From a larger Markdown document
 * const result = Markdown.parse(markdownDoc); // Finds the first table
 *
 * // With options
 * const result = Markdown.parse(input, { trim: false, maxRows: 100 });
 * ```
 */
export function parseMarkdown(
  input: string,
  options: MarkdownParseOptions = {}
): MarkdownParseResult {
  const opts = resolveParseOpts("Markdown.parse", input, options);
  const lines = input.split(LINEBREAK_REGEX);
  const lineCount = lines.length;

  for (let i = 0; i < lineCount - 1; i++) {
    const afterCode = skipFencedCode(lines, i);
    if (afterCode !== i) {
      i = afterCode - 1;
      continue;
    }
    const parsed = parseTableAt(lines, i, lineCount, opts);
    if (parsed) {
      return parsed.result;
    }
  }

  throw new MarkdownParseError(
    "No valid Markdown table found in input",
    lineCount > 0 ? lineCount : 1
  );
}

/**
 * Parse all Markdown tables from a document.
 *
 * @param input - Markdown string containing one or more tables
 * @param options - Parse options (maxRows applies per table)
 * @returns Array of parsed tables
 *
 * @example
 * ```ts
 * const tables = Markdown.parseAll(markdownDoc);
 * console.log(`Found ${tables.length} tables`);
 * tables.forEach((t, i) => console.log(`Table ${i}: ${t.headers.join(", ")}`));
 * ```
 */
export function parseMarkdownAll(
  input: string,
  options: MarkdownParseOptions = {}
): MarkdownParseResult[] {
  const opts = resolveParseOpts("Markdown.parseAll", input, options);
  const lines = input.split(LINEBREAK_REGEX);
  const lineCount = lines.length;
  const tables: MarkdownParseResult[] = [];
  let i = 0;

  while (i < lineCount - 1) {
    const afterCode = skipFencedCode(lines, i);
    if (afterCode !== i) {
      i = afterCode;
      continue;
    }
    const parsed = parseTableAt(lines, i, lineCount, opts);
    if (parsed) {
      tables.push(parsed.result);
      i = parsed.endLine;
    } else {
      i++;
    }
  }

  return tables;
}
