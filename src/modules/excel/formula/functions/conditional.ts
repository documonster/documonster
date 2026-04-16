/**
 * Conditional Aggregate Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
import { RVKind, ERRORS, BLANK, rvNumber, toStringRV, topLeft, isArray } from "../runtime/values";

// ============================================================================
// Criteria Predicate Builder (RuntimeValue version)
// ============================================================================

function scalarToString(v: ScalarValue): string {
  return toStringRV(v);
}

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

  // Operator-prefixed criteria
  const opMatch = /^([<>]=?|[<>]|=|<>)(.*)$/.exec(s);
  if (opMatch) {
    const [, op, valStr] = opMatch;
    const numVal = Number(valStr);
    const isNum = !isNaN(numVal) && valStr.trim() !== "";
    return (v: ScalarValue) => {
      const vn = v.kind === RVKind.Number ? v.value : NaN;
      const vs = scalarToString(v).toLowerCase();
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

  // Wildcard match (case-insensitive)
  if (s.includes("*") || s.includes("?")) {
    const pattern = s
      .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
      .replace(/\*/g, ".*")
      .replace(/\?/g, ".");
    try {
      const re = new RegExp("^" + pattern + "$", "i");
      return v => re.test(scalarToString(v));
    } catch {
      return v => scalarToString(v).toLowerCase() === s.toLowerCase();
    }
  }

  // Exact match (case-insensitive for strings, numeric for numbers)
  const numCriteria = Number(s);
  if (!isNaN(numCriteria) && s.trim() !== "") {
    return v => v.kind === RVKind.Number && v.value === numCriteria;
  }
  return v => scalarToString(v).toLowerCase() === s.toLowerCase();
}

// ============================================================================
// Helper: extract array from RuntimeValue
// ============================================================================

function asArray(v: RuntimeValue): ArrayValue | null {
  return v.kind === RVKind.Array ? v : null;
}

function getCell(arr: ArrayValue, r: number, c: number): ScalarValue {
  if (r < arr.height && c < arr.width) {
    return arr.rows[r][c];
  }
  return BLANK;
}

// ============================================================================
// Functions
// ============================================================================

export function fnSUMIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const criteriaScalar = topLeft(args[1]);
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

export function fnSUMIFS(args: RuntimeValue[]): RuntimeValue {
  const sumArr = asArray(args[0]);
  if (!sumArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs: { arr: ArrayValue; pred: (v: ScalarValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return ERRORS.VALUE;
    }
    pairs.push({
      arr: critRange,
      pred: buildCriteriaPredicateRV(topLeft(args[i + 1]))
    });
  }
  let sum = 0;
  for (let r = 0; r < sumArr.height; r++) {
    for (let c = 0; c < sumArr.width; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(getCell(p.arr, r, c))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = getCell(sumArr, r, c);
        if (sv.kind === RVKind.Number) {
          sum += sv.value;
        }
      }
    }
  }
  return rvNumber(sum);
}

export function fnCOUNTIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const pred = buildCriteriaPredicateRV(topLeft(args[1]));
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
  const pairs: { arr: ArrayValue; pred: (v: ScalarValue) => boolean }[] = [];
  for (let i = 0; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return ERRORS.VALUE;
    }
    pairs.push({
      arr: critRange,
      pred: buildCriteriaPredicateRV(topLeft(args[i + 1]))
    });
  }
  const rows = pairs[0].arr.height;
  const cols = pairs[0].arr.width;
  let count = 0;
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
        count++;
      }
    }
  }
  return rvNumber(count);
}

export function fnAVERAGEIF(args: RuntimeValue[]): RuntimeValue {
  const rangeArr = asArray(args[0]);
  if (!rangeArr) {
    return ERRORS.VALUE;
  }
  const pred = buildCriteriaPredicateRV(topLeft(args[1]));
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
  const pairs: { arr: ArrayValue; pred: (v: ScalarValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return ERRORS.VALUE;
    }
    pairs.push({
      arr: critRange,
      pred: buildCriteriaPredicateRV(topLeft(args[i + 1]))
    });
  }
  let sum = 0;
  let count = 0;
  for (let r = 0; r < avgArr.height; r++) {
    for (let c = 0; c < avgArr.width; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(getCell(p.arr, r, c))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
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

export function fnMAXIFS(args: RuntimeValue[]): RuntimeValue {
  const maxArr = asArray(args[0]);
  if (!maxArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs: { arr: ArrayValue; pred: (v: ScalarValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return ERRORS.VALUE;
    }
    pairs.push({
      arr: critRange,
      pred: buildCriteriaPredicateRV(topLeft(args[i + 1]))
    });
  }
  let result = -Infinity;
  let found = false;
  for (let r = 0; r < maxArr.height; r++) {
    for (let c = 0; c < maxArr.width; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(getCell(p.arr, r, c))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = getCell(maxArr, r, c);
        if (sv.kind === RVKind.Number) {
          if (sv.value > result) {
            result = sv.value;
          }
          found = true;
        }
      }
    }
  }
  return rvNumber(found ? result : 0);
}

export function fnMINIFS(args: RuntimeValue[]): RuntimeValue {
  const minArr = asArray(args[0]);
  if (!minArr || args.length < 3) {
    return ERRORS.VALUE;
  }
  const pairs: { arr: ArrayValue; pred: (v: ScalarValue) => boolean }[] = [];
  for (let i = 1; i < args.length - 1; i += 2) {
    const critRange = asArray(args[i]);
    if (!critRange) {
      return ERRORS.VALUE;
    }
    pairs.push({
      arr: critRange,
      pred: buildCriteriaPredicateRV(topLeft(args[i + 1]))
    });
  }
  let result = Infinity;
  let found = false;
  for (let r = 0; r < minArr.height; r++) {
    for (let c = 0; c < minArr.width; c++) {
      let allMatch = true;
      for (const p of pairs) {
        if (!p.pred(getCell(p.arr, r, c))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        const sv = getCell(minArr, r, c);
        if (sv.kind === RVKind.Number) {
          if (sv.value < result) {
            result = sv.value;
          }
          found = true;
        }
      }
    }
  }
  return rvNumber(found ? result : 0);
}
