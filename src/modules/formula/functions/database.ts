/**
 * Database Functions — Native RuntimeValue Implementation
 */

import { asArray, getCell } from "@formula/functions/_shared";
import { buildCriteriaPredicateRV } from "@formula/functions/conditional";
import type { RuntimeValue, ScalarValue } from "@formula/runtime/values";
import { RVKind, ERRORS, rvNumber, toStringRV, topLeft } from "@formula/runtime/values";

/**
 * Collect the field-column values from every database row that satisfies the
 * criteria range. The returned list preserves each matching cell's original
 * `ScalarValue` (numbers, strings, booleans, blanks, errors) so that callers
 * can decide how to aggregate — numeric aggregators filter to numbers while
 * DGET inspects the raw value.
 *
 * Returns an error RuntimeValue when the inputs are malformed (invalid array
 * shapes, unknown field name, etc.). On success, returns `ScalarValue[]`.
 */
function collectDatabaseMatches(args: RuntimeValue[]): RuntimeValue | ScalarValue[] {
  const dbArr = asArray(args[0]);
  const critArr = asArray(args[2]);
  if (!dbArr || !critArr) {
    return ERRORS.VALUE;
  }
  const fieldArg = topLeft(args[1]);

  if (dbArr.height < 2 || critArr.height < 2) {
    return ERRORS.VALUE;
  }

  // Determine field column index.
  //
  // Excel accepts the field argument as:
  //   - a 1-based integer column index (or TRUE/FALSE coerced to 1/0)
  //   - a string that matches one of the header cells (case-insensitive,
  //     and trimmed so imported data with stray whitespace still matches)
  let fieldIdx = -1;
  if (fieldArg.kind === RVKind.Number) {
    fieldIdx = Math.trunc(fieldArg.value) - 1;
  } else if (fieldArg.kind === RVKind.Boolean) {
    // TRUE → 1-based column 1 (index 0); FALSE → 0 → invalid (Excel
    // rejects FALSE with #VALUE!). Match that routing explicitly.
    fieldIdx = fieldArg.value ? 0 : -1;
  } else if (fieldArg.kind === RVKind.String) {
    const want = fieldArg.value.trim().toLowerCase();
    for (let c = 0; c < dbArr.width; c++) {
      if (
        toStringRV(getCell(dbArr, 0, c))
          .trim()
          .toLowerCase() === want
      ) {
        fieldIdx = c;
        break;
      }
    }
  }
  if (fieldIdx < 0 || fieldIdx >= dbArr.width) {
    return ERRORS.VALUE;
  }

  // Parse criteria: columns map to database header columns (trimmed,
  // case-insensitive, same as the field-name path above).
  const critColIndices: number[] = [];
  for (let cc = 0; cc < critArr.width; cc++) {
    const name = toStringRV(getCell(critArr, 0, cc))
      .trim()
      .toLowerCase();
    let idx = -1;
    for (let hc = 0; hc < dbArr.width; hc++) {
      if (
        toStringRV(getCell(dbArr, 0, hc))
          .trim()
          .toLowerCase() === name
      ) {
        idx = hc;
        break;
      }
    }
    critColIndices.push(idx);
  }

  // Collect matching rows' field values
  const matches: ScalarValue[] = [];
  for (let r = 1; r < dbArr.height; r++) {
    let matchesAnyCritRow = false;
    for (let cr = 1; cr < critArr.height; cr++) {
      let allMatch = true;
      for (let cc = 0; cc < critArr.width; cc++) {
        const critVal = getCell(critArr, cr, cc);
        if (
          critVal.kind === RVKind.Blank ||
          (critVal.kind === RVKind.String && critVal.value === "")
        ) {
          continue;
        }
        const dbCol = critColIndices[cc];
        if (dbCol < 0) {
          allMatch = false;
          break;
        }
        const pred = buildCriteriaPredicateRV(critVal);
        if (!pred(getCell(dbArr, r, dbCol))) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) {
        matchesAnyCritRow = true;
        break;
      }
    }
    if (matchesAnyCritRow) {
      matches.push(getCell(dbArr, r, fieldIdx));
    }
  }

  return matches;
}

/**
 * Shared wrapper for numeric aggregators (DSUM, DAVERAGE, DCOUNT, DMAX, DMIN,
 * DPRODUCT). Collects matches, filters to numeric values, and delegates to the
 * caller-supplied reducer.
 */
function databaseNumericAggregate(
  args: RuntimeValue[],
  aggregator: (values: number[]) => RuntimeValue
): RuntimeValue {
  const matches = collectDatabaseMatches(args);
  if (!Array.isArray(matches)) {
    return matches;
  }
  const values: number[] = [];
  for (const v of matches) {
    if (v.kind === RVKind.Number) {
      values.push(v.value);
    }
  }
  return aggregator(values);
}

export function fnDSUM(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => rvNumber(vals.reduce((a, b) => a + b, 0)));
}

export function fnDAVERAGE(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals =>
    vals.length === 0 ? ERRORS.DIV0 : rvNumber(vals.reduce((a, b) => a + b, 0) / vals.length)
  );
}

export function fnDCOUNT(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => rvNumber(vals.length));
}

export function fnDMAX(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length === 0) {
      return rvNumber(0);
    }
    // Avoid `Math.max(...vals)` — spreading a large numeric array onto the
    // call stack throws `RangeError: Maximum call stack size exceeded` on
    // databases with more than ~65k matching rows.
    let m = vals[0];
    for (let i = 1; i < vals.length; i++) {
      if (vals[i] > m) {
        m = vals[i];
      }
    }
    return rvNumber(m);
  });
}

export function fnDMIN(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length === 0) {
      return rvNumber(0);
    }
    let m = vals[0];
    for (let i = 1; i < vals.length; i++) {
      if (vals[i] < m) {
        m = vals[i];
      }
    }
    return rvNumber(m);
  });
}

export function fnDPRODUCT(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals =>
    rvNumber(vals.length === 0 ? 0 : vals.reduce((a, b) => a * b, 1))
  );
}

export function fnDGET(args: RuntimeValue[]): RuntimeValue {
  const matches = collectDatabaseMatches(args);
  if (!Array.isArray(matches)) {
    return matches;
  }
  // DGET aggregator semantics: 0 matches → #VALUE!, exactly 1 → the value,
  // >1 → #NUM!.
  if (matches.length === 0) {
    return ERRORS.VALUE;
  }
  if (matches.length > 1) {
    return ERRORS.NUM;
  }
  return matches[0];
}

/**
 * DCOUNTA(database, field, criteria) — count non-empty cells that match
 * the criteria, in the specified field. Unlike DCOUNT (numeric-only),
 * DCOUNTA counts any non-blank cell including text and booleans.
 */
export function fnDCOUNTA(args: RuntimeValue[]): RuntimeValue {
  const matches = collectDatabaseMatches(args);
  if (!Array.isArray(matches)) {
    return matches;
  }
  let count = 0;
  for (const v of matches) {
    // Count anything that is not Blank and not an empty string.
    if (v.kind === RVKind.Blank) {
      continue;
    }
    if (v.kind === RVKind.String && v.value === "") {
      continue;
    }
    count++;
  }
  return rvNumber(count);
}

/**
 * DSTDEV(database, field, criteria) — sample standard deviation of
 * numeric cells matching criteria in the specified field.
 */
export function fnDSTDEV(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length < 2) {
      return ERRORS.DIV0;
    }
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    let ss = 0;
    for (const v of vals) {
      ss += (v - mean) * (v - mean);
    }
    return rvNumber(Math.sqrt(ss / (vals.length - 1)));
  });
}

/**
 * DSTDEVP(database, field, criteria) — population standard deviation
 * of numeric cells matching criteria.
 */
export function fnDSTDEVP(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length === 0) {
      return ERRORS.DIV0;
    }
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    let ss = 0;
    for (const v of vals) {
      ss += (v - mean) * (v - mean);
    }
    return rvNumber(Math.sqrt(ss / vals.length));
  });
}

/**
 * DVAR(database, field, criteria) — sample variance of matching numeric
 * cells.
 */
export function fnDVAR(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length < 2) {
      return ERRORS.DIV0;
    }
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    let ss = 0;
    for (const v of vals) {
      ss += (v - mean) * (v - mean);
    }
    return rvNumber(ss / (vals.length - 1));
  });
}

/**
 * DVARP(database, field, criteria) — population variance of matching
 * numeric cells.
 */
export function fnDVARP(args: RuntimeValue[]): RuntimeValue {
  return databaseNumericAggregate(args, vals => {
    if (vals.length === 0) {
      return ERRORS.DIV0;
    }
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    let ss = 0;
    for (const v of vals) {
      ss += (v - mean) * (v - mean);
    }
    return rvNumber(ss / vals.length);
  });
}
