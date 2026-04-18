/**
 * Dynamic Array Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
import {
  RVKind,
  ERRORS,
  compareScalarsSameKind,
  rvNumber,
  rvArray,
  toNumberRV,
  toBooleanRV,
  topLeft,
  isError,
  isArray
} from "../runtime/values";
import {
  asArray,
  getCell,
  stripErrorCells,
  stripHiddenRowCells,
  stripSubtotalMaskedCells
} from "./_shared";
import { fnSUM, fnAVERAGE, fnMIN, fnMAX, fnCOUNT, fnCOUNTA, fnPRODUCT } from "./math";
import {
  fnSTDEV,
  fnSTDEVP,
  fnVAR,
  fnVARP,
  fnMEDIAN,
  fnLARGE,
  fnSMALL,
  fnMODE,
  fnPERCENTILE,
  fnPERCENTILEEXC,
  fnQUARTILE,
  fnQUARTILEEXC
} from "./statistical";

function isScalarError(v: ScalarValue): boolean {
  return v.kind === RVKind.Error;
}
function isScalarBlankOrEmpty(v: ScalarValue): boolean {
  return v.kind === RVKind.Blank || (v.kind === RVKind.String && v.value === "");
}

/**
 * Kind-priority ordering for cross-type sorting.
 *
 * `compareScalarsSameKind` returns `NaN` when the two operands have
 * different kinds, so SORT/SORTBY need a deterministic tiebreak. Excel's
 * workbook-grade sort groups values by kind — Numbers before Strings,
 * Strings before Booleans, Booleans before Blanks/Errors — which is the
 * priority we encode here.
 */
function kindPriority(k: RVKind): number {
  switch (k) {
    case RVKind.Number:
      return 0;
    case RVKind.String:
      return 1;
    case RVKind.Boolean:
      return 2;
    case RVKind.Blank:
      return 3;
    case RVKind.Error:
      return 4;
    default:
      return 5;
  }
}

/**
 * Locale-independent comparator for SORT / SORTBY.
 *
 * Uses `compareScalarsSameKind` for same-kind ordering (strings compared
 * case-insensitively, not via `localeCompare`, so results are stable
 * across machines) and falls back to `kindPriority` when the operands
 * differ in kind.
 */
function compareForSort(a: ScalarValue, b: ScalarValue): number {
  if (a.kind !== b.kind) {
    return kindPriority(a.kind) - kindPriority(b.kind);
  }
  const cmp = compareScalarsSameKind(a, b);
  return Number.isNaN(cmp) ? 0 : cmp;
}

export function fnFILTER(args: RuntimeValue[]): RuntimeValue {
  const dataArr = asArray(args[0]);
  const includeArr = asArray(args[1]);
  if (!dataArr || !includeArr) {
    return ERRORS.VALUE;
  }
  // Excel requires `include` to be a 1-column vector matching data's
  // height (the common row-filter shape). A mismatched shape previously
  // let rows slip through silently because `getCell` out-of-bounds
  // returns BLANK, which reads as FALSE.
  if (includeArr.width !== 1 || includeArr.height !== dataArr.height) {
    return ERRORS.VALUE;
  }
  const ifEmpty = args.length > 2 ? topLeft(args[2]) : null;
  const resultRows: ScalarValue[][] = [];
  for (let r = 0; r < dataArr.height; r++) {
    const inc = getCell(includeArr, r, 0);
    if (inc.kind === RVKind.Error) {
      return inc;
    }
    if (
      (inc.kind === RVKind.Boolean && inc.value) ||
      (inc.kind === RVKind.Number && inc.value !== 0)
    ) {
      const row: ScalarValue[] = [];
      for (let c = 0; c < dataArr.width; c++) {
        row.push(getCell(dataArr, r, c));
      }
      resultRows.push(row);
    }
  }
  if (resultRows.length === 0) {
    return ifEmpty !== null ? rvArray([[ifEmpty]]) : ERRORS.CALC;
  }
  return rvArray(resultRows);
}

export function fnSORT(args: RuntimeValue[]): RuntimeValue {
  const dataArr = asArray(args[0]);
  if (!dataArr) {
    return ERRORS.VALUE;
  }
  const rows: ScalarValue[][] = [];
  for (let r = 0; r < dataArr.height; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < dataArr.width; c++) {
      row.push(getCell(dataArr, r, c));
    }
    rows.push(row);
  }
  const sortIndexV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(1);
  if (isError(sortIndexV)) {
    return sortIndexV;
  }
  const sortOrderV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(1);
  if (isError(sortOrderV)) {
    return sortOrderV;
  }
  const byColV =
    args.length > 3 ? toBooleanRV(args[3]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(byColV)) {
    return byColV;
  }
  // Truncate the index so fractional inputs don't silently become NaN
  // when used as array subscripts. Then bounds-check against the axis
  // SORT will actually walk — this turns `SORT(arr, 0)` / out-of-range
  // values into the documented #VALUE! instead of a soft-fail.
  const sortIndex = Math.trunc(sortIndexV.value);
  if (byColV.value) {
    if (sortIndex < 1 || sortIndex > dataArr.height) {
      return ERRORS.VALUE;
    }
    const colIndices = Array.from({ length: dataArr.width }, (_, i) => i);
    const rowIdx = sortIndex - 1;
    colIndices.sort((a, b) => {
      const va = getCell(dataArr, rowIdx, a);
      const vb = getCell(dataArr, rowIdx, b);
      return compareForSort(va, vb) * sortOrderV.value;
    });
    return rvArray(rows.map(row => colIndices.map(c => row[c])));
  }
  if (sortIndex < 1 || sortIndex > dataArr.width) {
    return ERRORS.VALUE;
  }
  const col = sortIndex - 1;
  rows.sort((a, b) => compareForSort(a[col], b[col]) * sortOrderV.value);
  return rvArray(rows);
}

export function fnUNIQUE(args: RuntimeValue[]): RuntimeValue {
  const dataArr = asArray(args[0]);
  if (!dataArr) {
    return ERRORS.VALUE;
  }
  const byColV =
    args.length > 1 ? toBooleanRV(args[1]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(byColV)) {
    return byColV;
  }
  const exactlyOnceV =
    args.length > 2 ? toBooleanRV(args[2]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(exactlyOnceV)) {
    return exactlyOnceV;
  }
  if (byColV.value) {
    const transposed: ScalarValue[][] = [];
    for (let c = 0; c < dataArr.width; c++) {
      const col: ScalarValue[] = [];
      for (let r = 0; r < dataArr.height; r++) {
        col.push(getCell(dataArr, r, c));
      }
      transposed.push(col);
    }
    const unique = applyUnique(transposed, exactlyOnceV.value);
    if (unique.length === 0) {
      return ERRORS.VALUE;
    }
    const numRows = unique[0].length;
    const result: ScalarValue[][] = [];
    for (let r = 0; r < numRows; r++) {
      result.push(unique.map(ci => ci[r]));
    }
    return rvArray(result);
  }
  const dataRows: ScalarValue[][] = [];
  for (let r = 0; r < dataArr.height; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < dataArr.width; c++) {
      row.push(getCell(dataArr, r, c));
    }
    dataRows.push(row);
  }
  const result = applyUnique(dataRows, exactlyOnceV.value);
  return result.length > 0 ? rvArray(result) : ERRORS.VALUE;
}

function applyUnique(rows: ScalarValue[][], exactlyOnce: boolean): ScalarValue[][] {
  const keyCount = new Map<string, number>();
  const keyToRows = new Map<string, ScalarValue[]>();
  const order: string[] = [];
  for (const row of rows) {
    // Key the row by (kind, textual form) per cell so that UNIQUE preserves
    // Excel's type-aware equality: the number `1` and the string `"1"` are
    // distinct entries. Without the kind prefix they would collide because
    // `toStringRV` renders both as `"1"`.
    const key = row.map(c => `${c.kind}\u0001${scalarUniqueKey(c)}`).join("\u0002");
    if (!keyCount.has(key)) {
      order.push(key);
      keyToRows.set(key, row);
    }
    keyCount.set(key, (keyCount.get(key) ?? 0) + 1);
  }
  const result: ScalarValue[][] = [];
  for (const key of order) {
    if (exactlyOnce && (keyCount.get(key) ?? 0) > 1) {
      continue;
    }
    result.push(keyToRows.get(key)!);
  }
  return result;
}

/**
 * Produce a string that fully identifies a scalar for equality purposes,
 * distinguishing kinds that stringify identically but must not merge.
 */
function scalarUniqueKey(v: ScalarValue): string {
  if (v.kind === RVKind.Number) {
    // Using `toString` instead of localised formatting keeps NaN/Infinity
    // / -0 distinguishable from `0`, which matters for exact duplication
    // detection.
    return Object.is(v.value, -0) ? "-0" : String(v.value);
  }
  if (v.kind === RVKind.Boolean) {
    return v.value ? "1" : "0";
  }
  if (v.kind === RVKind.Error) {
    return v.code;
  }
  if (v.kind === RVKind.Blank) {
    return "";
  }
  return v.value.toLowerCase();
}

export function fnSORTBY(args: RuntimeValue[]): RuntimeValue {
  const dataArr = asArray(args[0]);
  if (!dataArr || args.length < 2) {
    return ERRORS.VALUE;
  }
  const data: { row: ScalarValue[]; idx: number }[] = [];
  for (let r = 0; r < dataArr.height; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < dataArr.width; c++) {
      row.push(getCell(dataArr, r, c));
    }
    data.push({ row, idx: r });
  }
  const sortKeys: { arr: ArrayValue; order: number }[] = [];
  for (let i = 1; i < args.length; i += 2) {
    const keyArr = asArray(args[i]);
    if (!keyArr) {
      return ERRORS.VALUE;
    }
    const orderV = i + 1 < args.length ? toNumberRV(args[i + 1]) : rvNumber(1);
    if (isError(orderV)) {
      return orderV;
    }
    sortKeys.push({ arr: keyArr, order: orderV.value });
  }
  data.sort((a, b) => {
    for (const sk of sortKeys) {
      const va = getCell(sk.arr, a.idx, 0);
      const vb = getCell(sk.arr, b.idx, 0);
      const cmp = compareForSort(va, vb);
      if (cmp !== 0) {
        return cmp * sk.order;
      }
    }
    return 0;
  });
  return rvArray(data.map(d => d.row));
}

export function fnSUBTOTAL(args: RuntimeValue[]): RuntimeValue {
  const funcNumV = toNumberRV(args[0]);
  if (isError(funcNumV)) {
    return funcNumV;
  }
  // Excel skips cells whose own formula is SUBTOTAL/AGGREGATE to avoid
  // double-counting when an outer aggregate spans a totals row or another
  // subtotal cell. `buildRangeArray` marks those cells via
  // `subtotalMask`; we strip them here before handing off to the
  // underlying aggregator.
  let dataArgs = stripSubtotalMaskedCells(args.slice(1));
  // Truncate toward zero: Excel rejects non-integer function codes, so
  // `SUBTOTAL(9.5, …)` should behave as `SUBTOTAL(9, …)` rather than
  // slipping through the switch and returning #VALUE!. (R6-P1-9)
  const rawFn = Math.trunc(funcNumV.value);
  const fn = rawFn > 100 ? rawFn - 100 : rawFn;
  // The 1xx variants (101-111) additionally skip hidden rows. Excel's
  // plain 1-11 codes only skip filter-hidden rows — but our worksheet
  // model carries a single `row.hidden` boolean that conflates the two
  // states, so we apply hidden-row stripping only to the 1xx variants
  // (matching manual-hide semantics; filter-hide remains a limitation
  // at the adapter layer, not here).
  if (rawFn > 100) {
    dataArgs = stripHiddenRowCells(dataArgs);
  }
  switch (fn) {
    case 1:
      return fnAVERAGE(dataArgs);
    case 2:
      return fnCOUNT(dataArgs);
    case 3:
      return fnCOUNTA(dataArgs);
    case 4:
      return fnMAX(dataArgs);
    case 5:
      return fnMIN(dataArgs);
    case 6:
      return fnPRODUCT(dataArgs);
    case 7:
      return fnSTDEV(dataArgs);
    case 8:
      return fnSTDEVP(dataArgs);
    case 9:
      return fnSUM(dataArgs);
    case 10:
      return fnVAR(dataArgs);
    case 11:
      return fnVARP(dataArgs);
    default:
      return ERRORS.VALUE;
  }
}

export function fnAGGREGATE(args: RuntimeValue[]): RuntimeValue {
  const funcNumV = toNumberRV(args[0]);
  if (isError(funcNumV)) {
    return funcNumV;
  }
  // Resolve the "options" parameter (arg[1]). Excel's option codes
  // (0–7) control which cells to skip:
  //   0 or omitted → ignore nested SUBTOTAL/AGGREGATE
  //   1 → ignore hidden rows + nested
  //   2 → ignore errors + nested
  //   3 → ignore hidden rows + errors + nested
  //   4 → ignore nothing
  //   5 → ignore hidden rows (keep nested)
  //   6 → ignore errors (keep nested)
  //   7 → ignore hidden rows + errors (keep nested)
  const optV = args[1] !== undefined ? toNumberRV(args[1]) : undefined;
  if (optV && isError(optV)) {
    return optV;
  }
  const option = optV ? Math.trunc(optV.value) : 0;
  const skipNested = option <= 3; // 0,1,2,3 skip nested SUBTOTAL/AGGREGATE
  const skipHidden = option === 1 || option === 3 || option === 5 || option === 7;
  const skipErrors = option === 2 || option === 3 || option === 6 || option === 7;
  let dataArgs = args.slice(2);
  if (skipNested) {
    dataArgs = stripSubtotalMaskedCells(dataArgs);
  }
  if (skipHidden) {
    dataArgs = stripHiddenRowCells(dataArgs);
  }
  if (skipErrors) {
    dataArgs = stripErrorCells(dataArgs);
  }
  // Excel rejects non-integer function codes; truncate toward zero
  // so AGGREGATE(9.5, …) behaves as AGGREGATE(9, …), matching SUBTOTAL.
  const fnCode = Math.trunc(funcNumV.value);
  switch (fnCode) {
    case 1:
      return fnAVERAGE(dataArgs);
    case 2:
      return fnCOUNT(dataArgs);
    case 3:
      return fnCOUNTA(dataArgs);
    case 4:
      return fnMAX(dataArgs);
    case 5:
      return fnMIN(dataArgs);
    case 6:
      return fnPRODUCT(dataArgs);
    case 7:
      return fnSTDEV(dataArgs);
    case 8:
      return fnSTDEVP(dataArgs);
    case 9:
      return fnSUM(dataArgs);
    case 10:
      return fnVAR(dataArgs);
    case 11:
      return fnVARP(dataArgs);
    case 12:
      return fnMEDIAN(dataArgs);
    case 13:
      return fnMODE(dataArgs);
    case 14:
      return fnLARGE(dataArgs);
    case 15:
      return fnSMALL(dataArgs);
    case 16:
      return fnPERCENTILE(dataArgs);
    case 17:
      return fnQUARTILE(dataArgs);
    case 18:
      return fnPERCENTILEEXC(dataArgs);
    case 19:
      return fnQUARTILEEXC(dataArgs);
    default:
      return ERRORS.VALUE;
  }
}

export function fnSEQUENCE(args: RuntimeValue[]): RuntimeValue {
  const rowsV = toNumberRV(args[0]);
  if (isError(rowsV)) {
    return rowsV;
  }
  const colsV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(1);
  if (isError(colsV)) {
    return colsV;
  }
  const startV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(1);
  if (isError(startV)) {
    return startV;
  }
  const stepV = args.length > 3 ? toNumberRV(args[3]) : rvNumber(1);
  if (isError(stepV)) {
    return stepV;
  }
  // Excel truncates row/col counts toward zero and rejects anything
  // below 1. Without this guard `SEQUENCE(-3)` would produce an empty
  // ArrayValue which breaks downstream rectangularisation, and
  // `SEQUENCE(2.5, 2)` would generate 3 rows instead of Excel's 2.
  const rowCount = Math.trunc(rowsV.value);
  const colCount = Math.trunc(colsV.value);
  if (!Number.isFinite(rowCount) || !Number.isFinite(colCount) || rowCount < 1 || colCount < 1) {
    return ERRORS.NUM;
  }
  // Bound the output to the same 10M-cell budget the rest of the array
  // pipeline uses. SEQUENCE(1e9, 1e9) would otherwise try to allocate
  // 1e18 scalars before OOM-ing (R6-P1-4).
  if (rowCount * colCount > 10_000_000) {
    return ERRORS.NUM;
  }
  const result: ScalarValue[][] = [];
  let val = startV.value;
  for (let r = 0; r < rowCount; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colCount; c++) {
      row.push(rvNumber(val));
      val += stepV.value;
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnRANDARRAY(args: RuntimeValue[]): RuntimeValue {
  const rowsV = args.length > 0 ? toNumberRV(args[0]) : rvNumber(1);
  if (isError(rowsV)) {
    return rowsV;
  }
  const colsV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(1);
  if (isError(colsV)) {
    return colsV;
  }
  const minV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(0);
  if (isError(minV)) {
    return minV;
  }
  const maxV = args.length > 3 ? toNumberRV(args[3]) : rvNumber(1);
  if (isError(maxV)) {
    return maxV;
  }
  const wholeV =
    args.length > 4 ? toBooleanRV(args[4]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(wholeV)) {
    return wholeV;
  }
  const rowCount = Math.trunc(rowsV.value);
  const colCount = Math.trunc(colsV.value);
  // Excel rejects rows/cols < 1 and min > max with #VALUE!, and when the
  // `whole` flag is TRUE it additionally requires integer min/max.
  if (
    !Number.isFinite(rowCount) ||
    !Number.isFinite(colCount) ||
    rowCount < 1 ||
    colCount < 1 ||
    minV.value > maxV.value
  ) {
    return ERRORS.VALUE;
  }
  if (wholeV.value && (!Number.isInteger(minV.value) || !Number.isInteger(maxV.value))) {
    return ERRORS.VALUE;
  }
  // Same 10M-cell budget as SEQUENCE / MAKEARRAY / broadcastBinaryOp —
  // otherwise `RANDARRAY(1e6, 1e6)` OOMs the host. (R6-P1-4)
  if (rowCount * colCount > 10_000_000) {
    return ERRORS.NUM;
  }
  const result: ScalarValue[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colCount; c++) {
      if (wholeV.value) {
        // Draw a uniform integer in [min, max] inclusive. The old code
        // did `Math.floor(min + random() * (max - min))` which excluded
        // `max` almost always.
        const span = maxV.value - minV.value + 1;
        row.push(rvNumber(minV.value + Math.floor(Math.random() * span)));
      } else {
        row.push(rvNumber(minV.value + Math.random() * (maxV.value - minV.value)));
      }
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnTOCOL(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    return rvArray([[topLeft(args[0])]]);
  }
  const arr = args[0] as ArrayValue;
  const ignoreV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(ignoreV)) {
    return ignoreV;
  }
  const scanV =
    args.length > 2 ? toBooleanRV(args[2]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(scanV)) {
    return scanV;
  }
  const result: ScalarValue[][] = [];
  const addCell = (v: ScalarValue) => {
    if (ignoreV.value === 1 && isScalarBlankOrEmpty(v)) {
      return;
    }
    if (ignoreV.value === 2 && isScalarError(v)) {
      return;
    }
    if (ignoreV.value === 3 && (isScalarBlankOrEmpty(v) || isScalarError(v))) {
      return;
    }
    result.push([v]);
  };
  if (scanV.value) {
    for (let c = 0; c < arr.width; c++) {
      for (let r = 0; r < arr.height; r++) {
        addCell(getCell(arr, r, c));
      }
    }
  } else {
    for (let r = 0; r < arr.height; r++) {
      for (let c = 0; c < arr.width; c++) {
        addCell(getCell(arr, r, c));
      }
    }
  }
  return result.length > 0 ? rvArray(result) : ERRORS.CALC;
}

export function fnTOROW(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    return rvArray([[topLeft(args[0])]]);
  }
  const arr = args[0] as ArrayValue;
  const ignoreV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(ignoreV)) {
    return ignoreV;
  }
  const scanV =
    args.length > 2 ? toBooleanRV(args[2]) : { kind: RVKind.Boolean as const, value: false };
  if (isError(scanV)) {
    return scanV;
  }
  const result: ScalarValue[] = [];
  const addCell = (v: ScalarValue) => {
    if (ignoreV.value === 1 && isScalarBlankOrEmpty(v)) {
      return;
    }
    if (ignoreV.value === 2 && isScalarError(v)) {
      return;
    }
    if (ignoreV.value === 3 && (isScalarBlankOrEmpty(v) || isScalarError(v))) {
      return;
    }
    result.push(v);
  };
  if (scanV.value) {
    for (let c = 0; c < arr.width; c++) {
      for (let r = 0; r < arr.height; r++) {
        addCell(getCell(arr, r, c));
      }
    }
  } else {
    for (let r = 0; r < arr.height; r++) {
      for (let c = 0; c < arr.width; c++) {
        addCell(getCell(arr, r, c));
      }
    }
  }
  return result.length > 0 ? rvArray([result]) : ERRORS.CALC;
}

export function fnCHOOSEROWS(args: RuntimeValue[]): RuntimeValue {
  const d = asArray(args[0]);
  if (!d) {
    return ERRORS.VALUE;
  }
  const result: ScalarValue[][] = [];
  for (let i = 1; i < args.length; i++) {
    const nV = toNumberRV(args[i]);
    if (isError(nV)) {
      return nV;
    }
    const idx = nV.value > 0 ? nV.value - 1 : d.height + nV.value;
    if (idx < 0 || idx >= d.height) {
      return ERRORS.VALUE;
    }
    const row: ScalarValue[] = [];
    for (let c = 0; c < d.width; c++) {
      row.push(getCell(d, idx, c));
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnCHOOSECOLS(args: RuntimeValue[]): RuntimeValue {
  const d = asArray(args[0]);
  if (!d) {
    return ERRORS.VALUE;
  }
  const ci: number[] = [];
  for (let i = 1; i < args.length; i++) {
    const nV = toNumberRV(args[i]);
    if (isError(nV)) {
      return nV;
    }
    const idx = nV.value > 0 ? nV.value - 1 : d.width + nV.value;
    if (idx < 0 || idx >= d.width) {
      return ERRORS.VALUE;
    }
    ci.push(idx);
  }
  const result: ScalarValue[][] = [];
  for (let r = 0; r < d.height; r++) {
    result.push(ci.map(c => getCell(d, r, c)));
  }
  return rvArray(result);
}

export function fnVSTACK(args: RuntimeValue[]): RuntimeValue {
  const rows: ScalarValue[][] = [];
  let maxWidth = 0;
  for (const a of args) {
    if (a.kind === RVKind.Array) {
      for (let r = 0; r < a.height; r++) {
        const row: ScalarValue[] = [];
        for (let c = 0; c < a.width; c++) {
          row.push(a.rows[r][c]);
        }
        rows.push(row);
        if (a.width > maxWidth) {
          maxWidth = a.width;
        }
      }
    } else {
      rows.push([topLeft(a)]);
      if (maxWidth < 1) {
        maxWidth = 1;
      }
    }
  }
  if (rows.length === 0) {
    return ERRORS.VALUE;
  }
  // Pad with #N/A (not BLANK) to match Excel, and to stay symmetric with
  // HSTACK's existing behaviour. The previous code left short rows to be
  // rectangularised by `rvArray`, which used BLANK — visually wrong and
  // inconsistent with how HSTACK reports "missing" cells.
  for (const row of rows) {
    while (row.length < maxWidth) {
      row.push(ERRORS.NA);
    }
  }
  return rvArray(rows);
}

export function fnHSTACK(args: RuntimeValue[]): RuntimeValue {
  let maxRows = 0;
  const arrays: ArrayValue[] = [];
  for (const a of args) {
    if (a.kind === RVKind.Array) {
      arrays.push(a);
      if (a.height > maxRows) {
        maxRows = a.height;
      }
    } else {
      arrays.push(rvArray([[topLeft(a)]]));
      if (maxRows < 1) {
        maxRows = 1;
      }
    }
  }
  const result: ScalarValue[][] = [];
  for (let r = 0; r < maxRows; r++) {
    const row: ScalarValue[] = [];
    for (const arr of arrays) {
      if (r < arr.height) {
        for (let c = 0; c < arr.width; c++) {
          row.push(arr.rows[r][c]);
        }
      } else {
        for (let c = 0; c < arr.width; c++) {
          row.push(ERRORS.NA);
        }
      }
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnTAKE(args: RuntimeValue[]): RuntimeValue {
  const d = asArray(args[0]);
  if (!d) {
    return ERRORS.VALUE;
  }
  const rowsV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(d.height);
  if (isError(rowsV)) {
    return rowsV;
  }
  const colsV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(d.width);
  if (isError(colsV)) {
    return colsV;
  }
  const rS = rowsV.value >= 0 ? 0 : Math.max(0, d.height + rowsV.value);
  const rE = rowsV.value >= 0 ? Math.min(rowsV.value, d.height) : d.height;
  const cS = colsV.value >= 0 ? 0 : Math.max(0, d.width + colsV.value);
  const cE = colsV.value >= 0 ? Math.min(colsV.value, d.width) : d.width;
  const result: ScalarValue[][] = [];
  for (let r = rS; r < rE; r++) {
    const row: ScalarValue[] = [];
    for (let c = cS; c < cE; c++) {
      row.push(getCell(d, r, c));
    }
    result.push(row);
  }
  return result.length > 0 ? rvArray(result) : ERRORS.CALC;
}

export function fnDROP(args: RuntimeValue[]): RuntimeValue {
  const d = asArray(args[0]);
  if (!d) {
    return ERRORS.VALUE;
  }
  const rowsV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(0);
  if (isError(rowsV)) {
    return rowsV;
  }
  const colsV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(0);
  if (isError(colsV)) {
    return colsV;
  }
  const rS = rowsV.value >= 0 ? rowsV.value : 0;
  const rE = rowsV.value >= 0 ? d.height : d.height + rowsV.value;
  const cS = colsV.value >= 0 ? colsV.value : 0;
  const cE = colsV.value >= 0 ? d.width : d.width + colsV.value;
  const result: ScalarValue[][] = [];
  for (let r = rS; r < rE; r++) {
    const row: ScalarValue[] = [];
    for (let c = cS; c < cE; c++) {
      row.push(getCell(d, r, c));
    }
    if (row.length > 0) {
      result.push(row);
    }
  }
  return result.length > 0 ? rvArray(result) : ERRORS.CALC;
}

export function fnWRAPROWS(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    return ERRORS.VALUE;
  }
  const arr = args[0] as ArrayValue;
  const flat: ScalarValue[] = [];
  for (let r = 0; r < arr.height; r++) {
    for (let c = 0; c < arr.width; c++) {
      flat.push(getCell(arr, r, c));
    }
  }
  const wcV = toNumberRV(args[1]);
  if (isError(wcV)) {
    return wcV;
  }
  if (wcV.value < 1) {
    return ERRORS.VALUE;
  }
  const pad: ScalarValue = args.length > 2 ? topLeft(args[2]) : ERRORS.NA;
  const result: ScalarValue[][] = [];
  for (let i = 0; i < flat.length; i += wcV.value) {
    const row = flat.slice(i, i + wcV.value);
    while (row.length < wcV.value) {
      row.push(pad);
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnWRAPCOLS(args: RuntimeValue[]): RuntimeValue {
  if (!isArray(args[0])) {
    return ERRORS.VALUE;
  }
  const arr = args[0] as ArrayValue;
  const flat: ScalarValue[] = [];
  for (let r = 0; r < arr.height; r++) {
    for (let c = 0; c < arr.width; c++) {
      flat.push(getCell(arr, r, c));
    }
  }
  const wcV = toNumberRV(args[1]);
  if (isError(wcV)) {
    return wcV;
  }
  if (wcV.value < 1) {
    return ERRORS.VALUE;
  }
  const pad: ScalarValue = args.length > 2 ? topLeft(args[2]) : ERRORS.NA;
  const numCols = Math.ceil(flat.length / wcV.value);
  const result: ScalarValue[][] = [];
  for (let r = 0; r < wcV.value; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < numCols; c++) {
      const idx = c * wcV.value + r;
      row.push(idx < flat.length ? flat[idx] : pad);
    }
    result.push(row);
  }
  return rvArray(result);
}

export function fnEXPAND(args: RuntimeValue[]): RuntimeValue {
  const d = asArray(args[0]);
  if (!d) {
    return ERRORS.VALUE;
  }
  const rowsV = args.length > 1 ? toNumberRV(args[1]) : rvNumber(d.height);
  if (isError(rowsV)) {
    return rowsV;
  }
  const colsV = args.length > 2 ? toNumberRV(args[2]) : rvNumber(d.width);
  if (isError(colsV)) {
    return colsV;
  }
  // Excel requires the target to be at least as large as the source, and
  // both dimensions to be positive. Without Math.trunc the inner loop
  // runs `c < rowsV.value` comparisons against a float, which produces
  // an off-by-one depending on the fractional part (R6-P1-5).
  const rowCount = Math.trunc(rowsV.value);
  const colCount = Math.trunc(colsV.value);
  if (
    !Number.isFinite(rowCount) ||
    !Number.isFinite(colCount) ||
    rowCount < d.height ||
    colCount < d.width
  ) {
    return ERRORS.VALUE;
  }
  // Cap the output size to the same 10M-cell budget the other dynamic
  // array producers use (see MAKEARRAY, broadcastBinaryOp).
  if (rowCount * colCount > 10_000_000) {
    return ERRORS.NUM;
  }
  const pad: ScalarValue = args.length > 3 ? topLeft(args[3]) : ERRORS.NA;
  const result: ScalarValue[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colCount; c++) {
      row.push(r < d.height && c < d.width ? getCell(d, r, c) : pad);
    }
    result.push(row);
  }
  return rvArray(result);
}
