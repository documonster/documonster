/**
 * Utilities for manipulating the external-workbook prefix in formula strings.
 *
 * Excel formula strings containing external workbook references have the
 * shape `[<workbook>]Sheet!Ref` where `<workbook>` is either
 *
 *   - a 1-based numeric index  — `[1]Sheet1!A1`  (the canonical on-disk form
 *     stored inside `<f>` elements of worksheet XML), or
 *   - a filename / relative path — `[测试.xlsx]Sheet1!A1`  (what Excel
 *     displays in the formula bar; not part of the OOXML storage contract,
 *     but produced by hand-written formulas and some older tools).
 *
 * When writing, excelts always emits the numeric form — indices map
 * positionally into the workbook's `<externalReferences>` list. When a
 * formula arrives with the filename form, the writer assigns (or reuses) an
 * ExternalLinkModel with that filename as its `target` and rewrites the
 * formula to the numeric form. This matches how Excel itself stores formulas
 * and makes them round-trippable.
 *
 * The quoted variant `'[file.xlsx]Sheet with space'!A1` is handled too — Excel
 * wraps the `[name]Sheet` segment in single quotes when the sheet name needs
 * quoting. The matching logic here recognises both the unquoted and quoted
 * forms, rewriting inside the quotes when needed.
 *
 * Edge cases we explicitly *do not* treat as external refs:
 *   - `[@Column]`, `[#Headers]`, `[Column Name]` — table structured refs
 *     (no `]Sheet!` tail). The regex requires the `]<sheet>!` follow-up,
 *     which structured refs never have.
 *   - Array literals `{1,2;3,4}` use `{}`, not `[]`.
 *   - String literals `"[Book]Sheet!A1"` — handled by scanning only outside
 *     string literal regions.
 */

// Matches an external-ref prefix in a formula:
//   Group 1 captures the workbook token inside [...], which is either
//   digits-only (numeric form) or a filename that may include path separators
//   when the whole prefix is single-quoted.
//
// Two variants:
//   1. Unquoted: [Book]Sheet!A1       — Windows-filename-safe chars only in
//      workbook token, then identifier sheet name, then !
//   2. Quoted:   '[path/to/Book]Sheet name'!A1    — quoted segment allows
//      paths / spaces / most punctuation inside the brackets; the sheet
//      name inside the quotes can contain anything except `'` (escaped as '')
//
// The regex matches through and including the trailing `!` so callers don't
// have to re-parse the A1/range part.

// Unquoted form: workbook token must not contain characters that would
// make the formula string ambiguous ( ] / \ space). This matches what Excel
// itself writes for the bare-filename case — anything more exotic is
// written in the quoted form by Excel.
const UNQUOTED_EXTERNAL_REF =
  /\[([^\]\\/:*?"<>|\s]+)\]([A-Za-z_\u00A1-\uFFFF][A-Za-z0-9_\u00A1-\uFFFF.]*)!/g;

// Quoted form: everything inside [] is permissive (any char except `]`),
// since the outer `'..'` quotes absorb the surrounding formula
// delimiters. Sheet name inside quotes is likewise permissive (any char
// except `'`, with `''` representing an escaped quote).
const QUOTED_EXTERNAL_REF = /'\[([^\]]+)\]((?:''|[^'])+)'!/g;

/**
 * A single match of an external reference inside a formula string. The
 * writer uses `workbook` to find/create an ExternalLinkModel and `sheet`
 * for reporting / sheet-name upsert. `replacement` is the substring that
 * should replace `match` in the final formula (with the workbook token
 * rewritten to a numeric index).
 */
export interface ExternalRefMatch {
  /** Full matched prefix including trailing `!`, e.g. `[测试.xlsx]Sheet1!`. */
  match: string;
  /** The workbook token inside `[]` — either numeric or a filename/path. */
  workbook: string;
  /** Whether the workbook token was already a numeric index. */
  numeric: boolean;
  /** The 1-based numeric index parsed from the workbook token, if numeric. */
  index: number | null;
  /** The sheet name (unquoted). */
  sheet: string;
  /** Whether the match came from the quoted variant `'[..]..'!`. */
  quoted: boolean;
  /** Start offset in the source formula. */
  start: number;
  /** End offset (exclusive) in the source formula. */
  end: number;
}

/**
 * Scan a formula string for all external-workbook references. String
 * literals (inside `"..."`) are skipped so that a string value like
 * `"[Book]Sheet!A1"` is not misidentified as a ref.
 *
 * The returned matches are in source order. If a formula contains no
 * external refs, the array is empty.
 */
export function findExternalRefs(formula: string): ExternalRefMatch[] {
  // Fast-fail: external refs always include a `[`. Skipping the
  // string-literal scan + two regex passes for plain cell refs (A1+B1,
  // SUM(A1:A5), 99% of real workbooks) saves noticeable time on large
  // sheets with thousands of formulas.
  if (formula.indexOf("[") === -1) {
    return [];
  }

  const matches: ExternalRefMatch[] = [];
  const safeRegions = stringLiteralRegions(formula);

  const addMatch = (
    start: number,
    end: number,
    workbook: string,
    sheet: string,
    quoted: boolean,
    match: string
  ): void => {
    if (insideAnyRegion(start, safeRegions)) {
      return;
    }
    const numeric = /^\d+$/.test(workbook);
    matches.push({
      match,
      workbook,
      numeric,
      index: numeric ? parseInt(workbook, 10) : null,
      sheet: unquoteSheetName(sheet),
      quoted,
      start,
      end
    });
  };

  UNQUOTED_EXTERNAL_REF.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = UNQUOTED_EXTERNAL_REF.exec(formula)) !== null) {
    addMatch(m.index, m.index + m[0].length, m[1], m[2], false, m[0]);
  }

  QUOTED_EXTERNAL_REF.lastIndex = 0;
  while ((m = QUOTED_EXTERNAL_REF.exec(formula)) !== null) {
    addMatch(m.index, m.index + m[0].length, m[1], m[2], true, m[0]);
  }

  // Sort matches by start offset so callers can process them in source
  // order (needed when rewriting with offset bookkeeping).
  matches.sort((a, b) => a.start - b.start);

  // Deduplicate overlapping matches (the unquoted regex can technically
  // match a prefix of a quoted sheet name in pathological inputs). Take
  // the first of any overlap.
  const out: ExternalRefMatch[] = [];
  let lastEnd = -1;
  for (const ref of matches) {
    if (ref.start >= lastEnd) {
      out.push(ref);
      lastEnd = ref.end;
    }
  }
  return out;
}

/**
 * Replace every external-workbook token in `formula` using the supplied
 * resolver. The resolver is called once per match and returns the numeric
 * index to substitute; returning `null` leaves the match unchanged (useful
 * when the caller cannot resolve a particular filename).
 *
 * Returns the rewritten formula. Offsets inside the original formula are
 * adjusted correctly even when multiple rewrites change the total length.
 */
export function rewriteExternalRefs(
  formula: string,
  resolve: (match: ExternalRefMatch) => number | null
): string {
  const refs = findExternalRefs(formula);
  if (refs.length === 0) {
    return formula;
  }

  let out = "";
  let cursor = 0;
  for (const ref of refs) {
    const index = resolve(ref);
    if (index === null) {
      continue; // leave this ref alone; cursor stays put
    }
    out += formula.slice(cursor, ref.start);
    // Construct the replacement: keep the quoted/unquoted shape, swap the
    // workbook token for [N], preserve the sheet segment exactly.
    if (ref.quoted) {
      // The quoted variant surrounds `[Book]Sheet` with single quotes,
      // followed by `!`. We swap the [Book] part for [N] and keep the
      // rest (including the closing `'!`) unchanged so sheet-name
      // quoting is preserved exactly.
      const inner = formula.slice(ref.start + 1, ref.end - 2); // between quotes, excluding trailing '!
      out += "'" + inner.replace(/^\[[^\]]*\]/, `[${index}]`) + "'!";
    } else {
      out += `[${index}]${ref.match.slice(ref.match.indexOf("]") + 1)}`;
    }
    cursor = ref.end;
  }
  out += formula.slice(cursor);
  return out;
}

// ===========================================================================
// Internal helpers
// ===========================================================================

/**
 * Return the spans of string literal regions in a formula, as half-open
 * [start, end) intervals (exclusive of the surrounding quotes themselves).
 * Used to skip external-ref matches that fall inside a string value.
 */
function stringLiteralRegions(formula: string): Array<[number, number]> {
  // Fast-fail when the formula has no double-quote at all — very common
  // for pure arithmetic / reference formulas.
  if (formula.indexOf('"') === -1) {
    return [];
  }
  const regions: Array<[number, number]> = [];
  const len = formula.length;
  let i = 0;
  while (i < len) {
    if (formula[i] === '"') {
      const start = i;
      i++;
      while (i < len) {
        if (formula[i] === '"') {
          if (i + 1 < len && formula[i + 1] === '"') {
            i += 2; // escaped quote inside string — keep scanning
            continue;
          }
          i++;
          break;
        }
        i++;
      }
      regions.push([start, i]);
    } else {
      i++;
    }
  }
  return regions;
}

function insideAnyRegion(pos: number, regions: Array<[number, number]>): boolean {
  for (const [a, b] of regions) {
    if (pos >= a && pos < b) {
      return true;
    }
  }
  return false;
}

function unquoteSheetName(sheet: string): string {
  // The quoted form captured inside `'..'` may have doubled single quotes
  // that represent a single quote in the logical sheet name.
  return sheet.replace(/''/g, "'");
}
