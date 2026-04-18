/**
 * Conditional Aggregate Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue, ErrorValue } from "../runtime/values";
import { RVKind, ERRORS, rvNumber, toStringRV, topLeft, isArray, isError } from "../runtime/values";
import {
  asArray,
  excelWildcardToRegex,
  getCell,
  hasUnescapedWildcard,
  unescapeExcelWildcard
} from "./_shared";

// ============================================================================
// Criteria Predicate Builder (RuntimeValue version)
// ============================================================================

/**
 * Build a criteria predicate from a ScalarValue.
 * Matches Excel SUMIF/COUNTIF criteria semantics:
 * - number → exact numeric match
 * - boolean → exact boolean match
 * - string with operator prefix (">5", "<=10", "<>abc") → comparison
 * - string with wildcards (* ?) → pattern match
 * - plain string → case-insensitive exact match (or numeric if parseable)
 */
export function buildCriteriaPredicateRV(criteria: ScalarValue): (v: ScalarValue) => boolean {
  if (criteria.kind === RVKind.Number) {
    const n = criteria.value;
    return v => v.kind === RVKind.Number && v.value === n;
  }
  if (criteria.kind === RVKind.Boolean) {
    const b = criteria.value;
    return v => v.kind === RVKind.Boolean && v.value === b;
  }
  if (criteria.kind === RVKind.Error) {
    return () => false;
  }
  // Blank criteria: match blank/empty-string
  if (criteria.kind === RVKind.Blank) {
    return v => v.kind === RVKind.Blank || (v.kind === RVKind.String && v.value === "");
  }

  // String criteria
  const s = criteria.kind === RVKind.String ? criteria.value : "";

  // Operator-prefixed criteria. Order the regex alternatives longest-first
  // so that `<>abc` matches the `<>` branch (not `<` with `>abc` as the
  // value). Without the explicit ordering, `/^[<>]=?/` would greedily
  // consume just `<` and silently mis-route every not-equal criterion to
  // the `<` branch.
  const opMatch = /^(<>|<=|>=|<|>|=)(.*)$/.exec(s);
  if (opMatch) {
    const [, op, valStr] = opMatch;
    const numVal = Number(valStr);
    const isNum = !isNaN(numVal) && valStr.trim() !== "";
    // For numeric comparisons Excel coerces booleans (TRUE→1, FALSE→0) and
    // blank (→0); numeric strings are NOT coerced (COUNTIF stays textual
    // for those). Only real Number / Boolean / Blank cells participate in
    // numeric comparisons; everything else falls back to string compare.
    const numericOf = (v: ScalarValue): number => {
      if (v.kind === RVKind.Number) {
        return v.value;
      }
      if (v.kind === RVKind.Boolean) {
        return v.value ? 1 : 0;
      }
      if (v.kind === RVKind.Blank) {
        return 0;
      }
      return Number.NaN;
    };
    return (v: ScalarValue) => {
      const vn = numericOf(v);
      const vs = toStringRV(v).toLowerCase();
      const cs = valStr.toLowerCase();
      switch (op) {
        case "=":
          return isNum ? vn === numVal : vs === cs;
        case "<>":
          return isNum ? vn !== numVal : vs !== cs;
        case ">":
          return isNum ? vn > numVal : vs > cs;
        case "<":
          return isNum ? vn < numVal : vs < cs;
        case ">=":
          return isNum ? vn >= numVal : vs >= cs;
        case "<=":
          return isNum ? vn <= numVal : vs <= cs;
        default:
          return false;
      }
    };
  }

  // Wildcard match (case-insensitive). Excel treats `~*`, `~?`, `~~` as
  // literal `*`, `?`, `~` and everything else as a regex special character
  // that must be escaped. Only an unescaped `*` or `?` triggers the wildcard
  // path; a pattern like `~*` matches a literal asterisk.
  if (hasUnescapedWildcard(s)) {
    try {
      const re = new RegExp("^" + excelWildcardToRegex(s) + "$", "i");
      return v => re.test(toStringRV(v));
    } catch {
      const literal = unescapeExcelWildcard(s).toLowerCase();
      return v => toStringRV(v).toLowerCase() === literal;
    }
  }
  // No wildcards: strip any `~` escapes and do a literal case-insensitive compare.
  const literal = unescapeExcelWildcard(s);

  // Exact match (case-insensitive for strings, numeric for numbers)
  const numCriteria = Number(literal);
  if (!isNaN(numCriteria) && literal.trim() !== "") {
    return v => v.kind === RVKind.Number && v.value === numCriteria;
  }
  const literalLc = literal.toLowerCase();
  return v => toStringRV(v).toLowerCase() === literalLc;
}

/**
 * Scan a criteria string for an unescaped `*` or `?`. A backslash-style
 * escape in Excel is `~`; so `~*` and `~?` are literals, while `*` and `?`
 * on their own or at a position not preceded by `~` count as wildcards.
 */
// ── Wildcard helpers live in `_shared.ts` and are re-used by SEARCH / MATCH /
//    XLOOKUP / SUMIF / COUNTIF so every function agrees on the same escape
//    semantics (`~*`, `~?`, `~~`). See `excelWildcardToRegex`,
//    `hasUnescapedWildcard`, and `unescapeExcelWildcard`.

// ============================================================================
// Functions
// ============================================================================

export function fnSUMIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const criteriaScalar = topLeft(args[1]);
  if (isError(criteriaScalar)) {
    return criteriaScalar;
  }
  const pred = buildCriteriaPredicateRV(criteriaScalar);
  const sumArr = args.length > 2 ? (asArray(args[2]) ?? rangeArr) : rangeArr;
  let sum = 0;
  for (let r = 0; r < rangeArr.height; r++) {
    for (let c = 0; c < rangeArr.width; c++) {
      if (pred(getCell(rangeArr, r, c))) {
        const sv = getCell(sumArr, r, c);
        if (sv.kind === RVKind.Number) {
          sum += sv.value;
        }
      }
    }
  }
  return rvNumber(sum);
}

// ============================================================================
// Multi-criteria helpers (SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS)
// ============================================================================

interface CriteriaPair {
  arr: ArrayValue;
  pred: (v: ScalarValue) => boolean;
}

/**
 * Scan successive `(range, criteria)` argument pairs starting at
 * `startIdx`. Returns a list of `{ arr, pred }` ready for iteration, or
 * an `{ error }` sentinel when a range argument is missing, a criteria
 * value is itself an error, or a criteria range's shape does not match
 * `target`'s shape (Excel's SUMIFS/COUNTIFS require identical shapes —
 * silently zero-extending with `BLANK` produced wrong counts).
 *
 * Used by SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS which share the
 * "target array + N criteria pairs" shape.
 */
function collectCriteriaPairs(
  args: RuntimeValue[],
  startIdx: number,
  target: ArrayValue
): { pairs: CriteriaPair[] } | { error: ErrorValue } {
  const pairs: CriteriaPair[] = [];
  for (let i = startIdx; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return { error: ERRORS.VALUE };
    }
    if (critRange.height !== target.height || critRange.width !== target.width) {
      return { error: ERRORS.VALUE };
    }
    const cs = topLeft(args[i + 1]);
    if (isError(cs)) {
      return { error: cs };
    }
    pairs.push({ arr: critRange, pred: buildCriteriaPredicateRV(cs) });
  }
  return { pairs };
}

/**
 * Walk every (row, col) position in `target` and invoke `onMatch(r, c)`
 * whenever all criteria predicates evaluate true at that position. The
 * criteria arrays are expected to share `target`'s shape — Excel returns
 * `#VALUE!` at the call site if they don't, which callers can detect by
 * checking `target.height` / `target.width` before invoking this helper.
 */
function iterateMultiCriteria(
  target: ArrayValue,
  pairs: readonly CriteriaPair[],
  onMatch: (r: number, c: number) => void
): void {
  const rows = target.height;
  const cols = target.width;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(getCell(p.arr, r, c))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        onMatch(r, c);
      }
    }
  }
}

export function fnSUMIFS(args: RuntimeValue[]): RuntimeValue {
  const sumArr = asArray(args[0]);
  if (!sumArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs = collectCriteriaPairs(args, 1, sumArr);
  if ("error" in pairs) {
    return pairs.error;
  }
  let sum = 0;
  iterateMultiCriteria(sumArr, pairs.pairs, (r, c) => {
    const sv = getCell(sumArr, r, c);
    if (sv.kind === RVKind.Number) {
      sum += sv.value;
    }
  });
  return rvNumber(sum);
}

export function fnCOUNTIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const cs = topLeft(args[1]);
  if (isError(cs)) {
    return cs;
  }
  const pred = buildCriteriaPredicateRV(cs);
  let count = 0;
  for (let r = 0; r < rangeArr.height; r++) {
    for (let c = 0; c < rangeArr.width; c++) {
      if (pred(getCell(rangeArr, r, c))) {
        count++;
      }
    }
  }
  return rvNumber(count);
}

export function fnCOUNTIFS(args: RuntimeValue[]): RuntimeValue {
  if (args.length < 2 || !isArray(args[0])) {
    return ERRORS.VALUE;
  }
  // The first range defines the target shape; every subsequent criteria
  // range must match. `collectCriteriaPairs` also validates args[0]
  // against itself (trivially passes) as a nice-to-have for consistency.
  const target = args[0];
  const pairs = collectCriteriaPairs(args, 0, target);
  if ("error" in pairs) {
    return pairs.error;
  }
  let count = 0;
  iterateMultiCriteria(target, pairs.pairs, () => {
    count++;
  });
  return rvNumber(count);
}

export function fnAVERAGEIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const cs = topLeft(args[1]);
  if (isError(cs)) {
    return cs;
  }
  const pred = buildCriteriaPredicateRV(cs);
  const avgArr = args.length > 2 ? (asArray(args[2]) ?? rangeArr) : rangeArr;
  let sum = 0;
  let count = 0;
  for (let r = 0; r < rangeArr.height; r++) {
    for (let c = 0; c < rangeArr.width; c++) {
      if (pred(getCell(rangeArr, r, c))) {
        const sv = getCell(avgArr, r, c);
        if (sv.kind === RVKind.Number) {
          sum += sv.value;
          count++;
        }
      }
    }
  }
  return count === 0 ? ERRORS.DIV0 : rvNumber(sum / count);
}

export function fnAVERAGEIFS(args: RuntimeValue[]): RuntimeValue {
  const avgArr = asArray(args[0]);
  if (!avgArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs = collectCriteriaPairs(args, 1, avgArr);
  if ("error" in pairs) {
    return pairs.error;
  }
  let sum = 0;
  let count = 0;
  iterateMultiCriteria(avgArr, pairs.pairs, (r, c) => {
    const sv = getCell(avgArr, r, c);
    if (sv.kind === RVKind.Number) {
      sum += sv.value;
      count++;
    }
  });
  return count === 0 ? ERRORS.DIV0 : rvNumber(sum / count);
}

export function fnMAXIFS(args: RuntimeValue[]): RuntimeValue {
  const maxArr = asArray(args[0]);
  if (!maxArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs = collectCriteriaPairs(args, 1, maxArr);
  if ("error" in pairs) {
    return pairs.error;
  }
  let result = -Infinity;
  let found = false;
  iterateMultiCriteria(maxArr, pairs.pairs, (r, c) => {
    const sv = getCell(maxArr, r, c);
    if (sv.kind === RVKind.Number) {
      if (sv.value > result) {
        result = sv.value;
      }
      found = true;
    }
  });
  return rvNumber(found ? result : 0);
}

export function fnMINIFS(args: RuntimeValue[]): RuntimeValue {
  const minArr = asArray(args[0]);
  if (!minArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs = collectCriteriaPairs(args, 1, minArr);
  if ("error" in pairs) {
    return pairs.error;
  }
  let result = Infinity;
  let found = false;
  iterateMultiCriteria(minArr, pairs.pairs, (r, c) => {
    const sv = getCell(minArr, r, c);
    if (sv.kind === RVKind.Number) {
      if (sv.value < result) {
        result = sv.value;
      }
      found = true;
    }
  });
  return rvNumber(found ? result : 0);
}
