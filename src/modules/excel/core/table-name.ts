/**
 * Table names, as Excel accepts them.
 *
 * Kept apart from `table.ts` because a worksheet looks a table up by the name it was *requested* under,
 * and `worksheet-core` cannot import `table.ts` without a cycle.
 */

/**
 * Maximum length for an Excel defined name (and therefore table name).
 */
const MAX_TABLE_NAME_LENGTH = 255;

/**
 * Matches an A1-style cell reference pattern like A1, Z99, XFD1048576.
 * Excel rejects table names that match this pattern.
 */
const CELL_REF_PATTERN = /^[A-Za-z]{1,3}\d+$/;

/** An R1C1-style cell reference with both numbers, e.g. R1C1, R100C200. A bare `RC` is not one. */
const R1C1_PATTERN = /^[Rr]\d+[Cc]\d+$/;

/**
 * Single-character names that Excel reserves for row/column navigation.
 * Per Microsoft docs: "You cannot use the uppercase and lowercase characters
 * 'C', 'c', 'R', or 'r' as a defined name."
 */
const RESERVED_SINGLE_CHARS = new Set(["C", "c", "R", "r"]);

/** Regex patterns used by sanitizeTableName, hoisted to module scope to avoid recompilation. */
const WHITESPACE_RE = /\s/g;
const INVALID_CHARS_RE = /[^\p{L}\p{N}_.]/gu;
const VALID_FIRST_CHAR_RE = /^[\p{L}_\\]/u;

const NAME_REST_RE = /^[\p{L}\p{N}_.\\?]*$/u;
const R1C1_REFERENCE_RE = /^[Rr]\d*[Cc]\d*$/;

/**
 * Why `name` is not a name Excel accepts for a defined name or a table, or `undefined` when it is.
 *
 * The same rules {@link sanitizeTableName} repairs, checked instead: a name a caller chose is refused rather than
 * silently changed. A name that reads as a cell reference (`A1`, `XFD1048576`, `R1C1`, `R`, `C`) is the case Excel
 * repairs a file for. Built-in names (`_xlnm.Print_Area`) are accepted as they are.
 */
export function definedNameProblem(name: unknown): string | undefined {
  if (typeof name !== "string" || name.length === 0) {
    return "a name is a non-empty string";
  }
  if (name.length > MAX_TABLE_NAME_LENGTH) {
    return `a name has at most ${MAX_TABLE_NAME_LENGTH} characters`;
  }
  if (name.startsWith("_xlnm.")) {
    return undefined;
  }
  if (
    !VALID_FIRST_CHAR_RE.test(name) ||
    !NAME_REST_RE.test(name.slice(name.codePointAt(0)! > 0xffff ? 2 : 1))
  ) {
    return "a name starts with a letter, an underscore or a backslash, and continues with letters, digits, underscores, periods, backslashes or question marks";
  }
  if (
    RESERVED_SINGLE_CHARS.has(name) ||
    R1C1_REFERENCE_RE.test(name) ||
    isGridCellReference(name)
  ) {
    return "a name cannot read as a cell reference";
  }
  return undefined;
}

/** `A1` … `XFD1048576`: letters naming a column of the grid followed by a row number. */
function isGridCellReference(name: string): boolean {
  if (!CELL_REF_PATTERN.test(name)) {
    return false;
  }
  const letters = /^[A-Za-z]+/.exec(name)![0].toUpperCase();
  let column = 0;
  for (const letter of letters) {
    column = column * 26 + letter.charCodeAt(0) - 64;
  }
  return column <= 16384;
}

/** A table name's first character: a letter, an underscore or a backslash. */
const TABLE_FIRST_RE = /^[\p{L}_\\]/u;
/** The rest of a table name: letters, digits, underscores and periods. A defined name may also use `\` and `?`. */
const TABLE_REST_RE = /^[\p{L}\p{N}_.]*$/u;

/**
 * Whether a table name reads as a cell reference: `R` / `C` on their own, one to three letters and a row number
 * (`A1`, `XFE1`), or an R1C1 reference with both numbers (`R1C1` — a bare `RC` is a name). The rules
 * {@link sanitizeTableName} has always applied, so a name it keeps is one this accepts.
 */
function readsAsCellReference(name: string): boolean {
  return RESERVED_SINGLE_CHARS.has(name) || CELL_REF_PATTERN.test(name) || R1C1_PATTERN.test(name);
}

/**
 * Why `name` is not one Excel accepts for a table — with the name that would be — or `undefined` when it is.
 *
 * A name a caller chose is refused rather than changed: `Table.add` used to store `"T1"` as `"_T1"` and `"My Table"`
 * as `"My_Table"`, so the table did not carry the name the caller would use for it in a formula. A name read from a
 * file is still corrected by {@link sanitizeTableName}, since the file was not the caller's to fix. The two agree
 * exactly: a name this accepts is one `sanitizeTableName` returns unchanged.
 */
export function tableNameProblem(name: unknown): string | undefined {
  if (typeof name !== "string" || name.length === 0) {
    return "a table name is a non-empty string";
  }
  const reason =
    name.length > MAX_TABLE_NAME_LENGTH
      ? `a table name has at most ${MAX_TABLE_NAME_LENGTH} characters`
      : /\s/u.test(name)
        ? "a table name has no spaces"
        : !TABLE_FIRST_RE.test(name)
          ? "a table name starts with a letter, an underscore or a backslash"
          : !TABLE_REST_RE.test(name.slice(name.codePointAt(0)! > 0xffff ? 2 : 1))
            ? "a table name continues with letters, digits, underscores and periods only"
            : readsAsCellReference(name)
              ? "a table name cannot read as a cell reference (A1, R1C1, R, C)"
              : undefined;
  return reason === undefined
    ? undefined
    : `${reason}; ${JSON.stringify(sanitizeTableName(name))} is one Excel accepts`;
}

/**
 * Sanitize a table name to comply with OOXML defined name rules
 * (ECMA-376, 4th edition, Part 1, §18.5.1.2).
 *
 * Rules enforced (per Microsoft documentation):
 * - First character must be a letter (any script), underscore (_), or backslash (\)
 * - Subsequent characters may be letters, digits, underscores, or periods (.)
 * - Backslash is only valid as the first character
 * - Spaces are replaced with underscores
 * - Other invalid characters are stripped
 * - Single-character names "C", "c", "R", "r" are prefixed with _
 * - Names that look like cell references (e.g. A1, R1C1) are prefixed with _
 * - Maximum 255 characters
 * - Empty result falls back to "_Table"
 *
 * This library applies these rules automatically so that generated files
 * always comply with the OOXML schema, avoiding Excel "repair" dialogs.
 */
export function sanitizeTableName(name: string): string {
  // Replace all whitespace characters (space, tab, newline, etc.) with underscores
  let sanitized = name.replace(WHITESPACE_RE, "_");

  // Preserve a leading backslash (valid only as first character per spec)
  let leadingBackslash = false;
  if (sanitized.startsWith("\\")) {
    leadingBackslash = true;
    sanitized = sanitized.slice(1);
  }

  // Strip characters not valid in defined names.
  // Subsequent characters: Unicode letters, digits, underscore, period.
  // Backslash is NOT valid in subsequent positions.
  sanitized = sanitized.replace(INVALID_CHARS_RE, "");

  // Re-attach leading backslash
  if (leadingBackslash) {
    sanitized = `\\${sanitized}`;
  }

  // Ensure the first character is valid (letter, underscore, or backslash).
  // Test the whole string start rather than sanitized[0] to correctly handle
  // supplementary Unicode characters (surrogate pairs).
  if (sanitized.length > 0 && !VALID_FIRST_CHAR_RE.test(sanitized)) {
    sanitized = `_${sanitized}`;
  }

  // Fallback if empty after sanitization
  if (sanitized.length === 0) {
    return "_Table";
  }

  // Avoid names that read as cell references — the same test `tableNameProblem` applies, so the two never disagree.
  if (readsAsCellReference(sanitized)) {
    sanitized = `_${sanitized}`;
  }

  // Truncate to max length last, after all prefix additions,
  // to guarantee the result never exceeds 255 characters.
  if (sanitized.length > MAX_TABLE_NAME_LENGTH) {
    sanitized = sanitized.slice(0, MAX_TABLE_NAME_LENGTH);
  }

  return sanitized;
}
