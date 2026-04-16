/**
 * Lookup / Reference Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
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
  isError,
  isArray
} from "../runtime/values";

// ============================================================================
// Helpers
// ============================================================================

function getCell(arr: ArrayValue, r: number, c: number): ScalarValue {
  if (r < arr.height && c < arr.width) {
    return arr.rows[r][c];
  }
  return BLANK;
}

function scalarToString(v: ScalarValue): string {
  return toStringRV(v);
}

/** Compare two scalar values for same-type ordering. */
function sameType(a: ScalarValue, b: ScalarValue): boolean {
  return a.kind === b.kind;
}

function scalarIsNumber(
  v: ScalarValue
): v is { readonly kind: RVKind.Number; readonly value: number } {
  return v.kind === RVKind.Number;
}

function scalarIsString(
  v: ScalarValue
): v is { readonly kind: RVKind.String; readonly value: string } {
  return v.kind === RVKind.String;
}

function scalarEquals(a: ScalarValue, b: ScalarValue): boolean {
  if (a.kind !== b.kind) {
    // Case-insensitive string comparison for string+string
    return false;
  }
  if (scalarIsNumber(a) && scalarIsNumber(b)) {
    return a.value === b.value;
  }
  if (scalarIsString(a) && scalarIsString(b)) {
    return a.value.toLowerCase() === b.value.toLowerCase();
  }
  if (a.kind === RVKind.Boolean && b.kind === RVKind.Boolean) {
    return a.value === b.value;
  }
  if (a.kind === RVKind.Blank && b.kind === RVKind.Blank) {
    return true;
  }
  return false;
}

function scalarStringEquals(a: ScalarValue, b: ScalarValue): boolean {
  return scalarIsString(a) && scalarIsString(b) && a.value.toLowerCase() === b.value.toLowerCase();
}

// ============================================================================
// Functions
// ============================================================================

export function fnROW(args: RuntimeValue[]): RuntimeValue {
  if (args.length === 0) {
    return ERRORS.VALUE;
  }
  const v = topLeft(args[0]);
  return v.kind === RVKind.Number ? v : ERRORS.VALUE;
}

export function fnCOLUMN(args: RuntimeValue[]): RuntimeValue {
  if (args.length === 0) {
    return ERRORS.VALUE;
  }
  const v = topLeft(args[0]);
  return v.kind === RVKind.Number ? v : ERRORS.VALUE;
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
  const rowNumV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(rowNumV)) {
    return rowNumV;
  }
  const rowNum = rowNumV.value;
  const colNumV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(0);
  if (isError(colNumV)) {
    return colNumV;
  }
  const colNum = colNumV.value;

  if (rowNum === 0 && colNum === 0) {
    return arr;
  }

  // rowNum=0: return entire column as array
  if (rowNum === 0) {
    const c = colNum - 1;
    if (c < 0 || c >= arr.width) {
      return ERRORS.REF;
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
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const lookupArr = args[1] as ArrayValue;
  const matchTypeV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(1);
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
    // Exact match (with wildcard support for string lookups)
    const lookupStr = scalarIsString(lookupValue) ? lookupValue.value : null;
    const hasWildcard = lookupStr !== null && (lookupStr.includes("*") || lookupStr.includes("?"));
    let wildcardRe: RegExp | null = null;
    if (hasWildcard) {
      const pattern = lookupStr
        .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
        .replace(/\*/g, ".*")
        .replace(/\?/g, ".");
      try {
        wildcardRe = new RegExp("^" + pattern + "$", "i");
      } catch {
        wildcardRe = null;
      }
    }
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
        } else if (fi.value.toLowerCase() === lookupValue.value.toLowerCase()) {
          return rvNumber(i + 1);
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
      if (sameType(v, lookupValue)) {
        if (scalarIsNumber(v) && scalarIsNumber(lookupValue)) {
          if (v.value <= lookupValue.value) {
            bestIdx = i;
          } else {
            break;
          }
        } else if (scalarIsString(v) && scalarIsString(lookupValue)) {
          if (v.value.toLowerCase() <= lookupValue.value.toLowerCase()) {
            bestIdx = i;
          } else {
            break;
          }
        }
      }
    }
    return bestIdx >= 0 ? rvNumber(bestIdx + 1) : ERRORS.NA;
  }

  // matchType === -1: Sorted descending. Find smallest value >= lookupValue.
  let bestIdx = -1;
  for (let i = 0; i < flat.length; i++) {
    const v = flat[i];
    if (sameType(v, lookupValue)) {
      if (scalarIsNumber(v) && scalarIsNumber(lookupValue)) {
        if (v.value >= lookupValue.value) {
          bestIdx = i;
        } else {
          break;
        }
      } else if (scalarIsString(v) && scalarIsString(lookupValue)) {
        if (v.value.toLowerCase() >= lookupValue.value.toLowerCase()) {
          bestIdx = i;
        } else {
          break;
        }
      }
    }
  }
  return bestIdx >= 0 ? rvNumber(bestIdx + 1) : ERRORS.NA;
}

export function fnVLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const table = args[1] as ArrayValue;
  const colIndexV = toNumberRV(args[2]);
  if (isError(colIndexV)) {
    return colIndexV;
  }
  const colIndex = colIndexV.value;
  const rangeLookupV =
    args.length > 3 ? toBooleanRV(args[3]) : { kind: RVKind.Boolean as const, value: true };
  if (isError(rangeLookupV)) {
    return rangeLookupV;
  }
  const rangeLookup = rangeLookupV.value;

  if (colIndex < 1 || colIndex > table.width) {
    return ERRORS.REF;
  }

  if (!rangeLookup) {
    // Exact match
    for (let r = 0; r < table.height; r++) {
      const cell = getCell(table, r, 0);
      if (scalarEquals(cell, lookupValue)) {
        return getCell(table, r, colIndex - 1);
      }
      if (scalarStringEquals(cell, lookupValue)) {
        return getCell(table, r, colIndex - 1);
      }
    }
    return ERRORS.NA;
  }

  // Approximate match: sorted ascending by first column.
  let bestRow = -1;
  for (let r = 0; r < table.height; r++) {
    const v = getCell(table, r, 0);
    if (sameType(v, lookupValue)) {
      if (scalarIsNumber(v) && scalarIsNumber(lookupValue)) {
        if (v.value <= lookupValue.value) {
          bestRow = r;
        } else {
          break;
        }
      } else if (scalarIsString(v) && scalarIsString(lookupValue)) {
        if (v.value.toLowerCase() <= lookupValue.value.toLowerCase()) {
          bestRow = r;
        } else {
          break;
        }
      }
    }
  }
  return bestRow >= 0 ? getCell(table, bestRow, colIndex - 1) : ERRORS.NA;
}

export function fnHLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (!isArray(args[1])) {
    return ERRORS.NA;
  }
  const table = args[1] as ArrayValue;
  const rowIndexV = toNumberRV(args[2]);
  if (isError(rowIndexV)) {
    return rowIndexV;
  }
  const rowIndex = rowIndexV.value;
  const rangeLookupV =
    args.length > 3 ? toBooleanRV(args[3]) : { kind: RVKind.Boolean as const, value: true };
  if (isError(rangeLookupV)) {
    return rangeLookupV;
  }
  const rangeLookup = rangeLookupV.value;

  if (rowIndex < 1 || rowIndex > table.height) {
    return ERRORS.REF;
  }

  if (!rangeLookup) {
    for (let c = 0; c < table.width; c++) {
      if (scalarEquals(getCell(table, 0, c), lookupValue)) {
        return getCell(table, rowIndex - 1, c);
      }
    }
    return ERRORS.NA;
  }

  let bestCol = -1;
  for (let c = 0; c < table.width; c++) {
    const hv = getCell(table, 0, c);
    if (sameType(hv, lookupValue)) {
      // For approximate match, find largest <= lookupValue
      if (scalarIsNumber(hv) && scalarIsNumber(lookupValue)) {
        if (hv.value <= lookupValue.value) {
          if (bestCol === -1 || getCell(table, 0, hv.value as never) !== undefined) {
            bestCol = c;
          }
        }
      } else if (scalarIsString(hv) && scalarIsString(lookupValue)) {
        if (hv.value.toLowerCase() <= lookupValue.value.toLowerCase()) {
          bestCol = c;
        }
      }
    }
  }
  return bestCol >= 0 ? getCell(table, rowIndex - 1, bestCol) : ERRORS.NA;
}

export function fnXLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
  if (!isArray(args[1])) {
    return ERRORS.VALUE;
  }
  const lookupArr = args[1] as ArrayValue;
  if (!isArray(args[2])) {
    return ERRORS.VALUE;
  }
  const returnArr = args[2] as ArrayValue;
  const ifNotFound = args.length > 3 ? topLeft(args[3]) : null;
  const matchModeV = args.length > 4 ? toNumberRV(args[4]) : rvNumber(0);
  if (isError(matchModeV)) {
    return matchModeV;
  }
  const matchMode = matchModeV.value;
  const searchModeV = args.length > 5 ? toNumberRV(args[5]) : rvNumber(1);
  if (isError(searchModeV)) {
    return searchModeV;
  }
  const searchMode = searchModeV.value;

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
    if (scalarIsNumber(a) && scalarIsNumber(b)) {
      return a.value - b.value;
    }
    if (scalarIsString(a) && scalarIsString(b)) {
      return a.value.toLowerCase().localeCompare(b.value.toLowerCase());
    }
    return 0;
  };

  if (matchMode === 0) {
    // Exact match
    const start = searchMode === -1 ? flat.length - 1 : 0;
    const end = searchMode === -1 ? -1 : flat.length;
    const step = searchMode === -1 ? -1 : 1;
    for (let i = start; i !== end; i += step) {
      if (scalarEquals(flat[i], lookupValue)) {
        foundIdx = i;
        break;
      }
      if (scalarStringEquals(flat[i], lookupValue)) {
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
    // Wildcard match
    const lookupStr = scalarToString(lookupValue);
    const pattern = lookupStr
      .replace(/[.*+^${}()|[\]\\]/g, m => (m === "*" || m === "?" ? m : "\\" + m))
      .replace(/\*/g, ".*")
      .replace(/\?/g, ".");
    try {
      const re = new RegExp("^" + pattern + "$", "i");
      for (let i = 0; i < flat.length; i++) {
        if (re.test(scalarToString(flat[i]))) {
          foundIdx = i;
          break;
        }
      }
    } catch {
      for (let i = 0; i < flat.length; i++) {
        if (scalarToString(flat[i]).toLowerCase() === lookupStr.toLowerCase()) {
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
    // Return array is also row-oriented
    if (returnArr.height === 1) {
      return foundIdx < returnArr.width ? getCell(returnArr, 0, foundIdx) : BLANK;
    }
    // Multiple rows in return — return column
    return foundIdx < returnArr.height ? getCell(returnArr, foundIdx, 0) : BLANK;
  }
  // Column lookup — return from same row index
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
  if (!isArray(args[1])) {
    return ERRORS.VALUE;
  }
  const lookupArr = args[1] as ArrayValue;
  const matchModeV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(0);
  if (isError(matchModeV)) {
    return matchModeV;
  }
  const matchMode = matchModeV.value;
  const searchModeV = args.length > 3 ? toNumberRV(args[3]) : rvNumber(1);
  if (isError(searchModeV)) {
    return searchModeV;
  }
  const searchMode = searchModeV.value;

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
      if (scalarStringEquals(flat[i], lookupValue)) {
        return rvNumber(i + 1);
      }
    }
    return ERRORS.NA;
  }

  if (matchMode === -1) {
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      const fi = flat[i];
      if (scalarIsNumber(fi) && scalarIsNumber(lookupValue)) {
        if (fi.value <= lookupValue.value) {
          if (best === -1 || (fi.value as number) > (flat[best] as { value: number }).value) {
            best = i;
          }
        }
      }
    }
    return best >= 0 ? rvNumber(best + 1) : ERRORS.NA;
  }

  if (matchMode === 1) {
    let best = -1;
    for (let i = 0; i < flat.length; i++) {
      const fi = flat[i];
      if (scalarIsNumber(fi) && scalarIsNumber(lookupValue)) {
        if (fi.value >= lookupValue.value) {
          if (best === -1 || (fi.value as number) < (flat[best] as { value: number }).value) {
            best = i;
          }
        }
      }
    }
    return best >= 0 ? rvNumber(best + 1) : ERRORS.NA;
  }

  return ERRORS.NA;
}

export function fnADDRESS(args: RuntimeValue[]): RuntimeValue {
  const rowNumV = toNumberRV(args[0]);
  if (isError(rowNumV)) {
    return rowNumV;
  }
  const rowNum = rowNumV.value;
  const colNumV = toNumberRV(args[1]);
  if (isError(colNumV)) {
    return colNumV;
  }
  const colNum = colNumV.value;
  const absNumV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(1);
  if (isError(absNumV)) {
    return absNumV;
  }
  const absNum = absNumV.value;
  // a1 style (true/default) vs r1c1 (false)
  const a1Arg = args.length > 3 ? topLeft(args[3]) : { kind: RVKind.Boolean as const, value: true };
  const a1 = a1Arg.kind === RVKind.Boolean ? a1Arg.value : true;
  const sheetText = args.length > 4 ? toStringRV(args[4]) : "";

  if (!a1) {
    // R1C1 style
    const rPart = absNum === 1 || absNum === 2 ? `R${rowNum}` : `R[${rowNum}]`;
    const cPart = absNum === 1 || absNum === 3 ? `C${colNum}` : `C[${colNum}]`;
    const prefix = sheetText ? `${sheetText}!` : "";
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

  let result = "";
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
      result = "$" + col + "$" + rowNum;
  }

  if (sheetText) {
    const needsQuote = /\s/.test(sheetText);
    result = (needsQuote ? `'${sheetText}'` : sheetText) + "!" + result;
  }
  return rvString(result);
}

export function fnLOOKUP(args: RuntimeValue[]): RuntimeValue {
  const lookupValue = topLeft(args[0]);
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
    let bestIdx = -1;
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i];
      if (sameType(v, lookupValue)) {
        if (scalarIsNumber(v) && scalarIsNumber(lookupValue) && v.value <= lookupValue.value) {
          bestIdx = i;
        } else if (
          scalarIsString(v) &&
          scalarIsString(lookupValue) &&
          v.value.toLowerCase() <= lookupValue.value.toLowerCase()
        ) {
          bestIdx = i;
        }
      }
    }
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
    let bestIdx = -1;
    for (let c = 0; c < cols; c++) {
      const v = getCell(lookupArr, 0, c);
      if (sameType(v, lookupValue)) {
        if (scalarIsNumber(v) && scalarIsNumber(lookupValue) && v.value <= lookupValue.value) {
          bestIdx = c;
        } else if (
          scalarIsString(v) &&
          scalarIsString(lookupValue) &&
          v.value.toLowerCase() <= lookupValue.value.toLowerCase()
        ) {
          bestIdx = c;
        }
      }
    }
    return bestIdx >= 0 ? getCell(lookupArr, rows - 1, bestIdx) : ERRORS.NA;
  }
  let bestIdx = -1;
  for (let r = 0; r < rows; r++) {
    const v = getCell(lookupArr, r, 0);
    if (sameType(v, lookupValue)) {
      if (scalarIsNumber(v) && scalarIsNumber(lookupValue) && v.value <= lookupValue.value) {
        bestIdx = r;
      } else if (
        scalarIsString(v) &&
        scalarIsString(lookupValue) &&
        v.value.toLowerCase() <= lookupValue.value.toLowerCase()
      ) {
        bestIdx = r;
      }
    }
  }
  return bestIdx >= 0 ? getCell(lookupArr, bestIdx, cols - 1) : ERRORS.NA;
}

export function fnTRANSPOSE(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    const sv = topLeft(args[0]);
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
  return args.length > 0 ? rvNumber(1) : ERRORS.VALUE;
}
