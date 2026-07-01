/**
 * Lookup / Reference Functions — Native RuntimeValue Implementation
 */

import {
  excelWildcardToRegex,
  getCell,
  hasUnescapedWildcard,
  unescapeExcelWildcard
} from "@formula/functions/_shared";
import type { RuntimeValue, ScalarValue, ArrayValue } from "@formula/runtime/values";
import {
  RVKind,
  ERRORS,
  BLANK,
  rvNumber,
  rvString,
  rvArray,
  toNumberRV,
  toBooleanRV,
  toStringRV,
  topLeft,
  scalarEquals,
  compareScalarsSameKind,
  isError,
  isArray
} from "@formula/runtime/values";

// ============================================================================
// Helpers
// ============================================================================

/** Compare two scalar values for same-type ordering. */
function sameType(a: ScalarValue, b: ScalarValue): boolean {
  return a.kind === b.kind;
}

function scalarIsString(
  v: ScalarValue
): v is { readonly kind: RVKind.String; readonly value: string } {
  return v.kind === RVKind.String;
}

// ============================================================================
// Functions
// ============================================================================

export function fnROW(args: RuntimeValue[]): RuntimeValue {
  // The reference-aware path is handled in evaluator.ts via tryEvaluateRefFunction.
  // This fallback is only reached when the argument is not a reference (e.g.
  // ROW(INDIRECT("A5")) or ROW({1,2,3})), which Excel rejects as #VALUE!.
  return ERRORS.VALUE;
}

export function fnCOLUMN(args: RuntimeValue[]): RuntimeValue {
  // See fnROW. Non-reference argument → #VALUE!.
  return ERRORS.VALUE;
}

export function fnROWS(args: RuntimeValue[]): RuntimeValue {
  const a = args[0];
  if (a.kind === RVKind.Array) {
    return rvNumber(a.height);
  }
  return rvNumber(1);
}

export function fnCOLUMNS(args: RuntimeValue[]): RuntimeValue {
  const a = args[0];
  if (a.kind === RVKind.Array) {
    return rvNumber(a.width);
  }
  return rvNumber(1);
}

export function fnINDEX(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    return topLeft(args[0]);
  }
  const arr = args[0] as ArrayValue;
  const rowNumV = args.length > 1 ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(rowNumV)) {
    return rowNumV;
  }
  // Excel truncates fractional indices toward zero before bounds checks.
  // Without this, `INDEX(a, 1.5, 1)` would index into `arr.rows[0.5]`, which
  // in V8 silently returns `undefined` and corrupts downstream values.
  const rowNum = Math.trunc(rowNumV.value);
  const colNumV = args.length > 2 ? toNumberRV(topLeft(args[2])) : rvNumber(0);
  if (isError(colNumV)) {
    return colNumV;
  }
  const colNum = Math.trunc(colNumV.value);

  if (rowNum < 0 || colNum < 0) {
    return ERRORS.VALUE;
  }

  if (rowNum === 0 && colNum === 0) {
    return arr;
  }

  // rowNum=0: return entire column as array
  if (rowNum === 0) {
    const c = colNum - 1;
    if (c < 0 || c >= arr.width) {
      return ERRORS.REF;
    }
    // Single-row source: a whole-column extract collapses to the one cell.
    if (arr.height === 1) {
      return arr.rows[0][c];
    }
    const rows: ScalarValue[][] = [];
    for (let r = 0; r < arr.height; r++) {
      rows.push([getCell(arr, r, c)]);
    }
    return rvArray(rows);
  }

  // colNum=0: return entire row as array
  if (colNum === 0) {
    const r = rowNum - 1;
    if (r < 0 || r >= arr.height) {
      return ERRORS.REF;
    }
    // Single-column source: a whole-row extract collapses to the one cell.
    // Matches Excel's convention — `INDEX(A1:A5, 2)` yields the scalar A2,
    // not a 1×1 array that downstream arithmetic has to implicit-intersect.
    if (arr.width === 1) {
      return arr.rows[r][0];
    }
    return rvArray([[...arr.rows[r]]]);
  }

  // Single cell
  const r = rowNum - 1;
  const c = colNum - 1;
  if (r < 0 || r >= arr.height || c < 0 || c >= arr.width) {
    return ERRORS.REF;
  }
  return arr.rows[r][c];
}

export function fnMATCH(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const lookupArr = args[1] as ArrayValue;
  // Blank `match_type` → Excel default 1 (largest value ≤ lookup, ascending sort).
  // Previously a blank coerced to 0 via toNumberRV and silently flipped the
  // function to exact-match mode — a behaviour gap vs. Excel's documented
  // default.
  const matchTypeV =
    args.length > 2 && args[2].kind !== RVKind.Blank ? toNumberRV(topLeft(args[2])) : rvNumber(1);
  if (isError(matchTypeV)) {
    return matchTypeV;
  }
  const matchType = matchTypeV.value;

  // Flatten to 1D
  const flat: ScalarValue[] = [];
  for (let r = 0; r < lookupArr.height; r++) {
    for (let c = 0; c < lookupArr.width; c++) {
      flat.push(getCell(lookupArr, r, c));
    }
  }

  if (matchType === 0) {
    // Exact match (with wildcard support for string lookups). The shared
    // `excelWildcardToRegex` converter applies the same `~*`, `~?`, `~~`
    // escape rules used by SEARCH, XLOOKUP, and SUMIF so behaviour is
    // consistent across the engine.
    const lookupStr = scalarIsString(lookupValue) ? lookupValue.value : null;
    const hasWildcard = lookupStr !== null && hasUnescapedWildcard(lookupStr);
    let wildcardRe: RegExp | null = null;
    if (hasWildcard && lookupStr !== null) {
      try {
        wildcardRe = new RegExp("^" + excelWildcardToRegex(lookupStr) + "$", "i");
      } catch {
        wildcardRe = null;
      }
    }
    // Pre-compute the literal (unescaped + lowercased) lookup string
    // once. The old code called `unescapeExcelWildcard(lookupValue.value).toLowerCase()`
    // inside the hot per-cell loop, paying O(n) per cell for what is
    // a constant expression. (For ranges without strings the literal
    // is never consulted — `scalarIsString(fi)` short-circuits first.)
    const lookupLiteralLc =
      lookupStr !== null && !hasWildcard ? unescapeExcelWildcard(lookupStr).toLowerCase() : null;
    for (let i = 0; i < flat.length; i++) {
      if (scalarEquals(flat[i], lookupValue)) {
        return rvNumber(i + 1);
      }
      const fi = flat[i];
      if (scalarIsString(fi) && scalarIsString(lookupValue)) {
        if (wildcardRe) {
          if (wildcardRe.test(fi.value)) {
            return rvNumber(i + 1);
          }
        } else if (lookupLiteralLc !== null) {
          // No unescaped wildcard — but the pattern may still contain
          // `~*` / `~?` / `~~` escape sequences that should reduce to
          // their literal character before comparison. Calling
          // `unescapeExcelWildcard` here matches the treatment that
          // SEARCH and the criteria predicate use; without it,
          // `MATCH("a~*b", ...)` would literally look for `"a~*b"`
          // instead of `"a*b"`.
          if (fi.value.toLowerCase() === lookupLiteralLc) {
            return rvNumber(i + 1);
          }
        }
      }
    }
    return ERRORS.NA;
  }

  if (matchType === 1 || matchType > 0) {
    // Sorted ascending. Find largest value <= lookupValue.
    let bestIdx = -1;
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i];
      if (v.kind !== lookupValue.kind) {
        continue;
      }
      const cmp = compareScalarsSameKind(v, lookupValue);
      if (!Number.isFinite(cmp)) {
        continue;
      }
      if (cmp <= 0) {
        bestIdx = i;
      } else {
        break;
      }
    }
    return bestIdx >= 0 ? rvNumber(bestIdx + 1) : ERRORS.NA;
  }

  // matchType === -1: Sorted descending. Find smallest value >= lookupValue.
  let bestIdx = -1;
  for (let i = 0; i < flat.length; i++) {
    const v = flat[i];
    if (v.kind !== lookupValue.kind) {
      continue;
    }
    const cmp = compareScalarsSameKind(v, lookupValue);
    if (!Number.isFinite(cmp)) {
      continue;
    }
    if (cmp >= 0) {
      bestIdx = i;
    } else {
      break;
    }
  }
  return bestIdx >= 0 ? rvNumber(bestIdx + 1) : ERRORS.NA;
}

export function fnVLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const table = args[1] as ArrayValue;
  const colIndexV = toNumberRV(topLeft(args[2]));
  if (isError(colIndexV)) {
    return colIndexV;
  }
  // VLOOKUP truncates the column index toward zero before bounds checks.
  const colIndex = Math.trunc(colIndexV.value);
  // Blank `range_lookup` → Excel default TRUE. A blank coerces through
  // `toBooleanRV` to FALSE which silently flips to exact match —
  // opposite of Excel's documented default.
  const rangeLookupV =
    args.length > 3 && args[3].kind !== RVKind.Blank
      ? toBooleanRV(topLeft(args[3]))
      : { kind: RVKind.Boolean as const, value: true };
  if (isError(rangeLookupV)) {
    return rangeLookupV;
  }
  const rangeLookup = rangeLookupV.value;

  if (colIndex < 1 || colIndex > table.width) {
    return ERRORS.REF;
  }

  if (!rangeLookup) {
    // Exact match — `scalarEquals` already handles case-insensitive string
    // comparison (see runtime/values.ts), so the earlier `scalarStringEquals`
    // fallback was dead code.
    //
    // Excel's VLOOKUP supports wildcards (`*`, `?`, `~*`, `~?`, `~~`) in
    // the exact-match mode (range_lookup=FALSE). Three paths, chosen by
    // what the lookup string contains:
    //   - unescaped wildcard (`*` or `?`) → regex match
    //   - only escape sequences (`~*`, `~?`, `~~`) → unescape then
    //     literal case-insensitive compare (so `"a~*b"` matches `"a*b"`)
    //   - neither → plain `scalarEquals`
    const lookupStr = lookupValue.kind === RVKind.String ? lookupValue.value : null;
    let wildcardRe: RegExp | null = null;
    let literalLc: string | null = null;
    if (lookupStr !== null && hasUnescapedWildcard(lookupStr)) {
      try {
        wildcardRe = new RegExp("^" + excelWildcardToRegex(lookupStr) + "$", "i");
      } catch {
        wildcardRe = null;
      }
    } else if (lookupStr !== null && /~[*?~]/.test(lookupStr)) {
      literalLc = unescapeExcelWildcard(lookupStr).toLowerCase();
    }
    for (let r = 0; r < table.height; r++) {
      const cell = getCell(table, r, 0);
      if (scalarEquals(cell, lookupValue)) {
        return getCell(table, r, colIndex - 1);
      }
      if (wildcardRe && cell.kind === RVKind.String && wildcardRe.test(cell.value)) {
        return getCell(table, r, colIndex - 1);
      }
      if (
        literalLc !== null &&
        cell.kind === RVKind.String &&
        cell.value.toLowerCase() === literalLc
      ) {
        return getCell(table, r, colIndex - 1);
      }
    }
    return ERRORS.NA;
  }

  // Approximate match: sorted ascending by first column. Binary-search
  // style isn't safe here (Excel allows mixed-type entries which break
  // monotonicity), so walk until we overshoot.
  let bestRow = -1;
  for (let r = 0; r < table.height; r++) {
    const v = getCell(table, r, 0);
    if (v.kind !== lookupValue.kind) {
      continue;
    }
    const cmp = compareScalarsSameKind(v, lookupValue);
    if (!Number.isFinite(cmp)) {
      continue;
    }
    if (cmp <= 0) {
      bestRow = r;
    } else {
      break;
    }
  }
  return bestRow >= 0 ? getCell(table, bestRow, colIndex - 1) : ERRORS.NA;
}

export function fnHLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const table = args[1] as ArrayValue;
  const rowIndexV = toNumberRV(topLeft(args[2]));
  if (isError(rowIndexV)) {
    return rowIndexV;
  }
  // HLOOKUP truncates the row index toward zero before bounds checks.
  const rowIndex = Math.trunc(rowIndexV.value);
  // Blank `range_lookup` → Excel default TRUE (see VLOOKUP rationale).
  const rangeLookupV =
    args.length > 3 && args[3].kind !== RVKind.Blank
      ? toBooleanRV(topLeft(args[3]))
      : { kind: RVKind.Boolean as const, value: true };
  if (isError(rangeLookupV)) {
    return rangeLookupV;
  }
  const rangeLookup = rangeLookupV.value;

  if (rowIndex < 1 || rowIndex > table.height) {
    return ERRORS.REF;
  }

  if (!rangeLookup) {
    // Exact match — supports wildcards on string lookups (see VLOOKUP
    // for full rationale; the paths mirror one another).
    const lookupStr = lookupValue.kind === RVKind.String ? lookupValue.value : null;
    let wildcardRe: RegExp | null = null;
    let literalLc: string | null = null;
    if (lookupStr !== null && hasUnescapedWildcard(lookupStr)) {
      try {
        wildcardRe = new RegExp("^" + excelWildcardToRegex(lookupStr) + "$", "i");
      } catch {
        wildcardRe = null;
      }
    } else if (lookupStr !== null && /~[*?~]/.test(lookupStr)) {
      literalLc = unescapeExcelWildcard(lookupStr).toLowerCase();
    }
    for (let c = 0; c < table.width; c++) {
      const cell = getCell(table, 0, c);
      if (scalarEquals(cell, lookupValue)) {
        return getCell(table, rowIndex - 1, c);
      }
      if (wildcardRe && cell.kind === RVKind.String && wildcardRe.test(cell.value)) {
        return getCell(table, rowIndex - 1, c);
      }
      if (
        literalLc !== null &&
        cell.kind === RVKind.String &&
        cell.value.toLowerCase() === literalLc
      ) {
        return getCell(table, rowIndex - 1, c);
      }
    }
    return ERRORS.NA;
  }

  let bestCol = -1;
  for (let c = 0; c < table.width; c++) {
    const hv = getCell(table, 0, c);
    if (hv.kind !== lookupValue.kind) {
      continue;
    }
    const cmp = compareScalarsSameKind(hv, lookupValue);
    if (!Number.isFinite(cmp)) {
      continue;
    }
    // Approximate match: pick the largest value <= lookupValue. We don't
    // break early when we overshoot because HLOOKUP's legacy behaviour
    // scans the whole row even for unsorted data.
    if (cmp <= 0) {
      bestCol = c;
    }
  }
  return bestCol >= 0 ? getCell(table, rowIndex - 1, bestCol) : ERRORS.NA;
}

export function fnXLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.VALUE;
  }
  const lookupArr = args[1] as ArrayValue;
  if (!isArray(args[2])) {
    return ERRORS.VALUE;
  }
  const returnArr = args[2] as ArrayValue;
  // Blank `if_not_found` → treat as omitted (default #N/A). Without
  // this guard, an explicitly-blank fourth slot would produce BLANK as
  // the fallback, which differs from Excel's "omitted → #N/A" default
  // and from what a user intuitively expects.
  const ifNotFound = args.length > 3 && args[3].kind !== RVKind.Blank ? topLeft(args[3]) : null;
  // Blank `match_mode` → 0 (exact); any non-{-1, 0, 1, 2} is rejected.
  const matchModeV =
    args.length > 4 && args[4].kind !== RVKind.Blank ? toNumberRV(topLeft(args[4])) : rvNumber(0);
  if (isError(matchModeV)) {
    return matchModeV;
  }
  const matchMode = matchModeV.value;
  if (matchMode !== 0 && matchMode !== -1 && matchMode !== 1 && matchMode !== 2) {
    return ERRORS.VALUE;
  }
  // Blank `search_mode` → 1 (first-to-last). Previously a blank coerced
  // to 0 which silently passed through but is not a valid search mode.
  const searchModeV =
    args.length > 5 && args[5].kind !== RVKind.Blank ? toNumberRV(topLeft(args[5])) : rvNumber(1);
  if (isError(searchModeV)) {
    return searchModeV;
  }
  const searchMode = searchModeV.value;
  if (searchMode !== 1 && searchMode !== -1 && searchMode !== 2 && searchMode !== -2) {
    return ERRORS.VALUE;
  }

  // Flatten lookup array to 1D
  const flat: ScalarValue[] = [];
  const isRow = lookupArr.height === 1;
  if (isRow) {
    for (let c = 0; c < lookupArr.width; c++) {
      flat.push(getCell(lookupArr, 0, c));
    }
  } else {
    for (let r = 0; r < lookupArr.height; r++) {
      flat.push(getCell(lookupArr, r, 0));
    }
  }

  let foundIdx = -1;

  const doCompare = (a: ScalarValue, b: ScalarValue): number => {
    // Use `compareScalarsSameKind` (shared with sorting / linear match)
    // instead of `localeCompare`. `localeCompare` is locale-sensitive and
    // can disagree with the `toLowerCase+===` equality check used on the
    // linear-search path — binary search would then skip the exact cell
    // that linear search would find (e.g. Turkish dotless I, ß→ss, etc.).
    // See R6-P1-11.
    const cmp = compareScalarsSameKind(a, b);
    return Number.isFinite(cmp) ? cmp : 0;
  };

  // ── Binary search for sorted data (searchMode = ±2) ──
  // Excel assumes the data is sorted ascending (2) or descending (-2). The
  // array must contain values of a single type compatible with `lookupValue`;
  // cells of an incompatible type make the sort invalid and binary search
  // cannot produce a meaningful result, so we fall back to #N/A in that
  // scenario (matching Excel's behaviour when the data is "not sorted").
  //
  // Supports matchMode 0 (exact), -1 (exact or next smaller), 1 (exact or
  // next larger). Wildcard matchMode (2) is incompatible with binary search
  // by definition — Excel silently downgrades to linear scan, which we do
  // by leaving `searchMode` as 1 below.
  const isBinary = (searchMode === 2 || searchMode === -2) && matchMode !== 2;
  if (isBinary) {
    const ascending = searchMode === 2;
    let lo = 0;
    let hi = flat.length - 1;
    let exact = -1;
    let nextSmaller = -1; // largest index with value < lookupValue (ascending)
    let nextLarger = -1; // smallest index with value > lookupValue (ascending)
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      const v = flat[mid];
      if (!sameType(v, lookupValue)) {
        // Heterogeneous array — binary search preconditions violated.
        exact = -1;
        nextSmaller = -1;
        nextLarger = -1;
        break;
      }
      const cmp = doCompare(v, lookupValue);
      if (cmp === 0) {
        exact = mid;
        break;
      }
      // In descending order, the ordering is inverted: treat `cmp > 0` on
      // the left half as "still greater than target" → search right.
      const goLeft = ascending ? cmp > 0 : cmp < 0;
      if (goLeft) {
        // mid is larger (ascending) or smaller (descending) than target
        if (ascending) {
          nextLarger = nextLarger === -1 || mid < nextLarger ? mid : nextLarger;
        } else {
          nextSmaller = nextSmaller === -1 || mid < nextSmaller ? mid : nextSmaller;
        }
        hi = mid - 1;
      } else {
        if (ascending) {
          nextSmaller = nextSmaller === -1 || mid > nextSmaller ? mid : nextSmaller;
        } else {
          nextLarger = nextLarger === -1 || mid > nextLarger ? mid : nextLarger;
        }
        lo = mid + 1;
      }
    }
    if (exact !== -1) {
      foundIdx = exact;
    } else if (matchMode === -1) {
      foundIdx = nextSmaller;
    } else if (matchMode === 1) {
      foundIdx = nextLarger;
    } else {
      foundIdx = -1;
    }
  } else if (matchMode === 0) {
    // Exact match
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      if (scalarEquals(flat[i], lookupValue)) {
        foundIdx = i;
        break;
      }
    }
  } else if (matchMode === -1) {
    // Exact match or next smaller
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (sameType(flat[i], lookupValue)) {
        const cmp = doCompare(flat[i], lookupValue);
        if (cmp === 0) {
          best = i;
          break;
        }
        if (cmp < 0 && (best === -1 || doCompare(flat[i], flat[best]) > 0)) {
          best = i;
        }
      }
    }
    foundIdx = best;
  } else if (matchMode === 1) {
    // Exact match or next larger
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      if (sameType(flat[i], lookupValue)) {
        const cmp = doCompare(flat[i], lookupValue);
        if (cmp === 0) {
          best = i;
          break;
        }
        if (cmp > 0 && (best === -1 || doCompare(flat[i], flat[best]) < 0)) {
          best = i;
        }
      }
    }
    foundIdx = best;
  } else if (matchMode === 2) {
    // Wildcard match — uses the shared Excel wildcard converter so SEARCH,
    // MATCH, XLOOKUP, and SUMIF/COUNTIF agree on `~*`, `~?`, `~~` escaping.
    const lookupStr = toStringRV(lookupValue);
    const pattern = excelWildcardToRegex(lookupStr);
    try {
      const re = new RegExp("^" + pattern + "$", "i");
      for (let i = 0; i < flat.length; i++) {
        if (re.test(toStringRV(flat[i]))) {
          foundIdx = i;
          break;
        }
      }
    } catch {
      for (let i = 0; i < flat.length; i++) {
        if (toStringRV(flat[i]).toLowerCase() === lookupStr.toLowerCase()) {
          foundIdx = i;
          break;
        }
      }
    }
  }

  if (foundIdx === -1) {
    return ifNotFound !== null ? ifNotFound : ERRORS.NA;
  }

  // Return from return array
  if (isRow) {
    // Horizontal lookup: the lookup axis is the column axis of the return
    // array, so `foundIdx` selects a column. A single-row return array
    // yields a scalar; a multi-row return array yields a column vector.
    if (foundIdx >= returnArr.width) {
      return BLANK;
    }
    if (returnArr.height === 1) {
      return getCell(returnArr, 0, foundIdx);
    }
    const col: ScalarValue[][] = [];
    for (let r = 0; r < returnArr.height; r++) {
      col.push([getCell(returnArr, r, foundIdx)]);
    }
    return rvArray(col);
  }
  // Vertical lookup: `foundIdx` selects a row; a single-column return
  // array yields a scalar; a multi-column array yields a row vector.
  if (foundIdx < returnArr.height) {
    if (returnArr.width === 1) {
      return getCell(returnArr, foundIdx, 0);
    }
    const row: ScalarValue[] = [];
    for (let c = 0; c < returnArr.width; c++) {
      row.push(getCell(returnArr, foundIdx, c));
    }
    return rvArray([row]);
  }
  return BLANK;
}

export function fnXMATCH(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.VALUE;
  }
  const lookupArr = args[1] as ArrayValue;
  // Blank `match_mode` → 0 (exact). Same validation as XLOOKUP.
  const matchModeV =
    args.length > 2 && args[2].kind !== RVKind.Blank ? toNumberRV(topLeft(args[2])) : rvNumber(0);
  if (isError(matchModeV)) {
    return matchModeV;
  }
  const matchMode = matchModeV.value;
  if (matchMode !== 0 && matchMode !== -1 && matchMode !== 1 && matchMode !== 2) {
    return ERRORS.VALUE;
  }
  const searchModeV =
    args.length > 3 && args[3].kind !== RVKind.Blank ? toNumberRV(topLeft(args[3])) : rvNumber(1);
  if (isError(searchModeV)) {
    return searchModeV;
  }
  const searchMode = searchModeV.value;
  if (searchMode !== 1 && searchMode !== -1 && searchMode !== 2 && searchMode !== -2) {
    return ERRORS.VALUE;
  }

  const flat: ScalarValue[] = [];
  if (lookupArr.height === 1) {
    for (let c = 0; c < lookupArr.width; c++) {
      flat.push(getCell(lookupArr, 0, c));
    }
  } else {
    for (let r = 0; r < lookupArr.height; r++) {
      flat.push(getCell(lookupArr, r, 0));
    }
  }

  if (matchMode === 0) {
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      if (scalarEquals(flat[i], lookupValue)) {
        return rvNumber(i + 1);
      }
    }
    return ERRORS.NA;
  }

  if (matchMode === 2) {
    // Wildcard matching — `*` and `?` (with `~` escape). Only meaningful
    // when the lookup value is a string; for non-string lookup values
    // Excel falls back to plain comparison.
    if (lookupValue.kind !== RVKind.String) {
      // Fall through to exact-match semantics for non-string lookups.
      const start = searchMode === -1 ? flat.length - 1 : 0;
      const end = searchMode === -1 ? -1 : flat.length;
      const step = searchMode === -1 ? -1 : 1;
      for (let i = start; i !== end; i += step) {
        if (scalarEquals(flat[i], lookupValue)) {
          return rvNumber(i + 1);
        }
      }
      return ERRORS.NA;
    }
    const pattern = lookupValue.value;
    const matcher = hasUnescapedWildcard(pattern)
      ? new RegExp(`^${excelWildcardToRegex(pattern)}$`, "iu")
      : null;
    const literal = matcher ? null : unescapeExcelWildcard(pattern).toLowerCase();
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      const cell = flat[i];
      if (cell.kind !== RVKind.String) {
        continue;
      }
      if (matcher) {
        if (matcher.test(cell.value)) {
          return rvNumber(i + 1);
        }
      } else if (cell.value.toLowerCase() === literal) {
        return rvNumber(i + 1);
      }
    }
    return ERRORS.NA;
  }

  if (matchMode === -1) {
    // Next-smaller-or-equal: largest item <= lookupValue.
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      const cmp = compareScalarsSameKind(flat[i], lookupValue);
      if (Number.isNaN(cmp)) {
        continue;
      }
      if (cmp <= 0 && (best === -1 || compareScalarsSameKind(flat[i], flat[best]) > 0)) {
        best = i;
      }
    }
    return best >= 0 ? rvNumber(best + 1) : ERRORS.NA;
  }

  if (matchMode === 1) {
    // Next-larger-or-equal: smallest item >= lookupValue.
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      const cmp = compareScalarsSameKind(flat[i], lookupValue);
      if (Number.isNaN(cmp)) {
        continue;
      }
      if (cmp >= 0 && (best === -1 || compareScalarsSameKind(flat[i], flat[best]) < 0)) {
        best = i;
      }
    }
    return best >= 0 ? rvNumber(best + 1) : ERRORS.NA;
  }

  return ERRORS.NA;
}

export function fnADDRESS(args: RuntimeValue[]): RuntimeValue {
  const rowNumV = toNumberRV(topLeft(args[0]));
  if (isError(rowNumV)) {
    return rowNumV;
  }
  const rowNum = Math.trunc(rowNumV.value);
  const colNumV = toNumberRV(topLeft(args[1]));
  if (isError(colNumV)) {
    return colNumV;
  }
  const colNum = Math.trunc(colNumV.value);
  // ADDRESS rejects non-positive row/col with #VALUE!. Without this guard
  // ADDRESS(0, 1) would silently return "$A$0" and ADDRESS(1, 0) would
  // produce "$$1" (no column letter) — neither is a legal cell reference.
  if (!Number.isFinite(rowNum) || !Number.isFinite(colNum) || rowNum < 1 || colNum < 1) {
    return ERRORS.VALUE;
  }
  // Blank `abs_num` → Excel default 1 (fully absolute). Without the
  // blank guard, `toNumberRV(BLANK)` coerces to 0 which falls outside
  // the 1..4 range and surfaces a spurious #VALUE!.
  const absNumV =
    args.length > 2 && args[2].kind !== RVKind.Blank ? toNumberRV(topLeft(args[2])) : rvNumber(1);
  if (isError(absNumV)) {
    return absNumV;
  }
  const absNum = Math.trunc(absNumV.value);
  // Excel only accepts abs_num ∈ {1, 2, 3, 4}; anything else is #VALUE!.
  if (absNum < 1 || absNum > 4) {
    return ERRORS.VALUE;
  }
  // a1 style (true/default) vs r1c1 (false)
  const a1Arg = args.length > 3 ? topLeft(args[3]) : { kind: RVKind.Boolean as const, value: true };
  const a1 = a1Arg.kind === RVKind.Boolean ? a1Arg.value : true;
  const sheetText = args.length > 4 ? toStringRV(topLeft(args[4])) : "";

  if (!a1) {
    // R1C1 style
    const rPart = absNum === 1 || absNum === 2 ? `R${rowNum}` : `R[${rowNum}]`;
    const cPart = absNum === 1 || absNum === 3 ? `C${colNum}` : `C[${colNum}]`;
    const prefix = sheetText ? `${renderSheetPrefix(sheetText)}!` : "";
    return rvString(prefix + rPart + cPart);
  }

  // Convert column number to letters
  let col = "";
  let cv = colNum;
  while (cv > 0) {
    cv--;
    col = String.fromCharCode(65 + (cv % 26)) + col;
    cv = Math.floor(cv / 26);
  }

  let result: string;
  switch (absNum) {
    case 1:
      result = "$" + col + "$" + rowNum;
      break;
    case 2:
      result = col + "$" + rowNum;
      break;
    case 3:
      result = "$" + col + rowNum;
      break;
    case 4:
      result = col + rowNum;
      break;
    default:
      // Unreachable — `absNum` is already validated to {1, 2, 3, 4} above.
      result = "$" + col + "$" + rowNum;
  }

  if (sheetText) {
    result = renderSheetPrefix(sheetText) + "!" + result;
  }
  return rvString(result);
}

/**
 * Quote a sheet name for use in a reference prefix the way Excel does:
 *   - plain `Name` (letters, digits, underscore, leading non-digit) → as-is
 *   - anything else → wrapped in single quotes with embedded `'` doubled
 */
function renderSheetPrefix(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}

export function fnLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (lookupValue.kind === RVKind.Error) {
    return lookupValue;
  }
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const lookupArr = args[1] as ArrayValue;

  if (args.length > 2 && isArray(args[2])) {
    const resultArr = args[2] as ArrayValue;
    const flat: ScalarValue[] = [];
    const isRow = lookupArr.height === 1;
    if (isRow) {
      for (let c = 0; c < lookupArr.width; c++) {
        flat.push(getCell(lookupArr, 0, c));
      }
    } else {
      for (let r = 0; r < lookupArr.height; r++) {
        flat.push(getCell(lookupArr, r, 0));
      }
    }
    const bestIdx = findLastLessEqual(flat, lookupValue);
    if (bestIdx === -1) {
      return ERRORS.NA;
    }
    if (isRow) {
      return resultArr.height === 1
        ? bestIdx < resultArr.width
          ? getCell(resultArr, 0, bestIdx)
          : BLANK
        : bestIdx < resultArr.height
          ? getCell(resultArr, bestIdx, 0)
          : BLANK;
    }
    return bestIdx < resultArr.height ? getCell(resultArr, bestIdx, 0) : BLANK;
  }

  const rows = lookupArr.height;
  const cols = lookupArr.width;
  if (cols === 0) {
    return ERRORS.NA;
  }
  if (cols >= rows) {
    // Array-form LOOKUP, horizontal orientation: lookup runs along first
    // row; result is pulled from the last row of the same column.
    const firstRow: ScalarValue[] = [];
    for (let c = 0; c < cols; c++) {
      firstRow.push(getCell(lookupArr, 0, c));
    }
    const bestIdx = findLastLessEqual(firstRow, lookupValue);
    return bestIdx >= 0 ? getCell(lookupArr, rows - 1, bestIdx) : ERRORS.NA;
  }
  // Array-form LOOKUP, vertical orientation: lookup runs down first column;
  // result is pulled from the last column of the same row.
  const firstCol: ScalarValue[] = [];
  for (let r = 0; r < rows; r++) {
    firstCol.push(getCell(lookupArr, r, 0));
  }
  const bestIdx = findLastLessEqual(firstCol, lookupValue);
  return bestIdx >= 0 ? getCell(lookupArr, bestIdx, cols - 1) : ERRORS.NA;
}

/**
 * Linear scan for the largest same-kind value that is `<= target`.
 * Returns the flat-index of that value, or `-1` when no same-kind
 * value qualifies. Uses `compareScalarsSameKind` so numbers compare by
 * value, strings case-insensitively — the same ordering Excel uses
 * for legacy LOOKUP / VLOOKUP / HLOOKUP approximate matches.
 */
function findLastLessEqual(flat: readonly ScalarValue[], target: ScalarValue): number {
  let bestIdx = -1;
  for (let i = 0; i < flat.length; i++) {
    const v = flat[i];
    if (v.kind !== target.kind) {
      continue;
    }
    const cmp = compareScalarsSameKind(v, target);
    if (!Number.isFinite(cmp)) {
      continue;
    }
    if (cmp <= 0) {
      bestIdx = i;
    }
  }
  return bestIdx;
}

export function fnTRANSPOSE(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    const sv = topLeft(args[0]);
    // Excel propagates errors through TRANSPOSE rather than burying them
    // inside a 1×1 array — callers that then aggregate the result (e.g.
    // `SUM(TRANSPOSE(#N/A))`) expect the error to surface. R8-P1 fix.
    if (sv.kind === RVKind.Error) {
      return sv;
    }
    return rvArray([[sv]]);
  }
  const arr = args[0] as ArrayValue;
  const rows = arr.height;
  const cols = arr.width;
  const result: ScalarValue[][] = [];
  for (let c = 0; c < cols; c++) {
    const row: ScalarValue[] = [];
    for (let r = 0; r < rows; r++) {
      row.push(getCell(arr, r, c));
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnAREAS(args: RuntimeValue[]): RuntimeValue {
  if (args.length === 0) {
    return ERRORS.VALUE;
  }
  // Normally unreachable — the evaluator's reference-aware path in
  // `evaluateCall` intercepts AREAS before eager dereference happens,
  // so by the time this fallback runs the reference has already been
  // flattened into a dereferenced array (losing the area count).
  // Keep the fallback behaviour aligned with the intercept path:
  // arrays and scalars that reach here are not references and should
  // surface as `#VALUE!`.
  const a = args[0];
  if (a.kind === RVKind.Error) {
    return a;
  }
  if (a.kind === RVKind.Reference) {
    return rvNumber(a.areas.length);
  }
  return ERRORS.VALUE;
}
