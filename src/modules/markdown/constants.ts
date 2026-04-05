/**
 * Markdown Module Constants
 *
 * Shared constants used across the Markdown module.
 */

/**
 * Pre-compiled regex for line splitting (matches CR, LF, or CRLF)
 */
export const LINEBREAK_REGEX = /\r\n|\r|\n/;

/**
 * Characters that need escaping in Markdown table cells.
 * Also matches CRLF/CR/LF so that escaping + newline conversion
 * can be done in a single `replace()` call.
 * Note: `\r\n` must come before `\r` to match CRLF as a single unit.
 */
export const ESCAPE_AND_NEWLINE = /\r\n|[|\\\r\n]/g;

/**
 * Regex to unescape Markdown table cell content (`\|` → `|`, `\\` → `\`)
 */
export const UNESCAPE_REGEX = /\\([|\\])/g;

/**
 * Regex to match `<br>`, `<br/>`, or `<br />` tags (case-insensitive).
 * Used to convert multiline cell representations back to newlines during parsing.
 */
export const BR_TAG_REGEX = /<br\s*\/?>/gi;

/**
 * Regex to match literal newlines (CR, LF, or CRLF) in cell content.
 * Used when escaping is disabled but newline conversion is still needed.
 */
export const NEWLINE_IN_CELL = /\r\n|\r|\n/g;
