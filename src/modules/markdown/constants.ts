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
 *
 * Also matches the `<` of a literal `<br>` tag, which is written as `\<br>`: newlines are
 * written as `<br>`, so a cell holding the text `<br>` would otherwise come back as a newline.
 */
export const ESCAPE_AND_NEWLINE = /\r\n|[|\\\r\n]|<(?=br\s*\/?>)/gi;

/**
 * Escapes (`\|`, `\\`, `\<`) and `<br>`, `<br/>`, `<br />` tags (case-insensitive), read in one
 * pass so an escaped `\<br>` is never taken for a line break.
 */
export const CELL_ESCAPE_OR_BR_REGEX = /\\([|\\<])|<br\s*\/?>/gi;

/**
 * Regex to match literal newlines (CR, LF, or CRLF) in cell content.
 * Used when escaping is disabled but newline conversion is still needed.
 */
export const NEWLINE_IN_CELL = /\r\n|\r|\n/g;
