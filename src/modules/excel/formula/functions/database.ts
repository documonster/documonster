/**
 * Database Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
import { RVKind, ERRORS, BLANK, rvNumber, toStringRV, topLeft } from "../runtime/values";
import { buildCriteriaPredicateRV } from "./conditional";

function asArray(v: RuntimeValue): ArrayValue | null {
  return v.kind === RVKind.Array ? v : null;
}

function getCell(arr: ArrayValue, r: number, c: number): ScalarValue {
  return r < arr.height && c < arr.width ? arr.rows[r][c] : BLANK;
}

function scalarToString(v: ScalarValue): string {
  return toStringRV(v);
}

function databaseHelper(
  args: RuntimeValue[],
  aggregator: (values: number[]) => RuntimeValue
): RuntimeValue {
  const dbArr = asArray(args[0]);
  const critArr = asArray(args[2]);
  if (!dbArr || !critArr) {
    return ERRORS.VALUE;
  }
  const fieldArg = topLeft(args[1]);

  if (dbArr.height < 2 || critArr.height < 2) {
    return ERRORS.VALUE;
  }

  // Determine field column index
  let fieldIdx = -1;
  if (fieldArg.kind === RVKind.Number) {
    fieldIdx = fieldArg.value - 1;
  } else if (fieldArg.kind === RVKind.String) {
    for (let c = 0; c < dbArr.width; c++) {
      if (scalarToString(getCell(dbArr, 0, c)).toLowerCase() === fieldArg.value.toLowerCase()) {
        fieldIdx = c;
        break;
      }
    }
  }
  if (fieldIdx < 0 || fieldIdx >= dbArr.width) {
    return ERRORS.VALUE;
  }

  // Parse criteria: columns map to database header columns
  const critColIndices: number[] = [];
  for (let cc = 0; cc < critArr.width; cc++) {
    const name = scalarToString(getCell(critArr, 0, cc)).toLowerCase();
    let idx = -1;
    for (let hc = 0; hc < dbArr.width; hc++) {
      if (scalarToString(getCell(dbArr, 0, hc)).toLowerCase() === name) {
        idx = hc;
        break;
      }
    }
    critColIndices.push(idx);
  }

  // Collect matching rows
  const values: number[] = [];
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
      const v = getCell(dbArr, r, fieldIdx);
      if (v.kind === RVKind.Number) {
        values.push(v.value);
      }
    }
  }

  return aggregator(values);
}

export function fnDSUM(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals => rvNumber(vals.reduce((a, b) => a + b, 0)));
}

export function fnDAVERAGE(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals =>
    vals.length === 0 ? ERRORS.DIV0 : rvNumber(vals.reduce((a, b) => a + b, 0) / vals.length)
  );
}

export function fnDCOUNT(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals => rvNumber(vals.length));
}

export function fnDMAX(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals => rvNumber(vals.length === 0 ? 0 : Math.max(...vals)));
}

export function fnDMIN(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals => rvNumber(vals.length === 0 ? 0 : Math.min(...vals)));
}

export function fnDPRODUCT(args: RuntimeValue[]): RuntimeValue {
  return databaseHelper(args, vals =>
    rvNumber(vals.length === 0 ? 0 : vals.reduce((a, b) => a * b, 1))
  );
}

export function fnDGET(args: RuntimeValue[]): RuntimeValue {
  const dbArr = asArray(args[0]);
  const critArr = asArray(args[2]);
  if (!dbArr || !critArr) {
    return ERRORS.VALUE;
  }
  const fieldArg = topLeft(args[1]);

  if (dbArr.height < 2 || critArr.height < 2) {
    return ERRORS.VALUE;
  }

  let fieldIdx = -1;
  if (fieldArg.kind === RVKind.Number) {
    fieldIdx = fieldArg.value - 1;
  } else if (fieldArg.kind === RVKind.String) {
    for (let c = 0; c < dbArr.width; c++) {
      if (scalarToString(getCell(dbArr, 0, c)).toLowerCase() === fieldArg.value.toLowerCase()) {
        fieldIdx = c;
        break;
      }
    }
  }
  if (fieldIdx < 0 || fieldIdx >= dbArr.width) {
    return ERRORS.VALUE;
  }

  const critColIndices: number[] = [];
  for (let cc = 0; cc < critArr.width; cc++) {
    const name = scalarToString(getCell(critArr, 0, cc)).toLowerCase();
    let idx = -1;
    for (let hc = 0; hc < dbArr.width; hc++) {
      if (scalarToString(getCell(dbArr, 0, hc)).toLowerCase() === name) {
        idx = hc;
        break;
      }
    }
    critColIndices.push(idx);
  }

  let found: ScalarValue = BLANK;
  let count = 0;
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
      found = getCell(dbArr, r, fieldIdx);
      count++;
      if (count > 1) {
        return ERRORS.NUM;
      }
    }
  }
  return count === 0 ? ERRORS.VALUE : found;
}
