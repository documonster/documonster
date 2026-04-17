/**
 * Dynamic Array Functions — Native RuntimeValue Implementation
 */

import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
import {
  RVKind,
  ERRORS,
  BLANK,
  rvNumber,
  rvArray,
  toNumberRV,
  toBooleanRV,
  toStringRV,
  topLeft,
  isError,
  isArray
} from "../runtime/values";
import { fnSUM, fnAVERAGE, fnMIN, fnMAX, fnCOUNT, fnCOUNTA, fnPRODUCT } from "./math";
import { fnSTDEV, fnSTDEVP, fnVAR, fnVARP, fnMEDIAN, fnLARGE, fnSMALL } from "./statistical";

function asArray(v: RuntimeValue): ArrayValue | null {
  return v.kind === RVKind.Array ? v : null;
}
function getCell(arr: ArrayValue, r: number, c: number): ScalarValue {
  return r < arr.height && c < arr.width ? arr.rows[r][c] : BLANK;
}
function scalarToString(v: ScalarValue): string {
  return toStringRV(v);
}
function isScalarError(v: ScalarValue): boolean {
  return v.kind === RVKind.Error;
}
function isScalarBlankOrEmpty(v: ScalarValue): boolean {
  return v.kind === RVKind.Blank || (v.kind === RVKind.String && v.value === "");
}

export function fnFILTER(args: RuntimeValue[]): RuntimeValue {
  const dataArr = asArray(args[0]);
  const includeArr = asArray(args[1]);
  if (!dataArr || !includeArr) {
    return ERRORS.VALUE;
  }
  const ifEmpty = args.length > 2 ? topLeft(args[2]) : null;
  const resultRows: ScalarValue[][] = [];
  for (let r = 0; r < dataArr.height; r++) {
    const inc = r < includeArr.height ? getCell(includeArr, r, 0) : BLANK;
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
    return ifEmpty !== null ? rvArray([[ifEmpty]]) : ERRORS.VALUE;
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
  if (byColV.value) {
    const colIndices = Array.from({ length: dataArr.width }, (_, i) => i);
    const rowIdx = sortIndexV.value - 1;
    colIndices.sort((a, b) => {
      const va = getCell(dataArr, rowIdx, a);
      const vb = getCell(dataArr, rowIdx, b);
      if (va.kind === RVKind.Number && vb.kind === RVKind.Number) {
        return (va.value - vb.value) * sortOrderV.value;
      }
      return scalarToString(va).localeCompare(scalarToString(vb)) * sortOrderV.value;
    });
    return rvArray(rows.map(row => colIndices.map(c => row[c])));
  }
  const col = sortIndexV.value - 1;
  rows.sort((a, b) => {
    const va = a[col];
    const vb = b[col];
    if (va.kind === RVKind.Number && vb.kind === RVKind.Number) {
      return (va.value - vb.value) * sortOrderV.value;
    }
    return scalarToString(va).localeCompare(scalarToString(vb)) * sortOrderV.value;
  });
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
    const key = row.map(c => scalarToString(c)).join("\0");
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
      let cmp: number;
      if (va.kind === RVKind.Number && vb.kind === RVKind.Number) {
        cmp = va.value - vb.value;
      } else {
        cmp = scalarToString(va).localeCompare(scalarToString(vb));
      }
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
  const dataArgs = args.slice(1);
  const fn = funcNumV.value > 100 ? funcNumV.value - 100 : funcNumV.value;
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
  const dataArgs = args.slice(2);
  switch (funcNumV.value) {
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
    case 14:
      return fnLARGE(dataArgs);
    case 15:
      return fnSMALL(dataArgs);
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
  const result: ScalarValue[][] = [];
  let val = startV.value;
  for (let r = 0; r < rowsV.value; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colsV.value; c++) {
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
  const result: ScalarValue[][] = [];
  for (let r = 0; r < rowsV.value; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colsV.value; c++) {
      const v = minV.value + Math.random() * (maxV.value - minV.value);
      row.push(rvNumber(wholeV.value ? Math.floor(v) : v));
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
  const result: ScalarValue[][] = [];
  for (const a of args) {
    if (a.kind === RVKind.Array) {
      for (let r = 0; r < a.height; r++) {
        const row: ScalarValue[] = [];
        for (let c = 0; c < a.width; c++) {
          row.push(a.rows[r][c]);
        }
        result.push(row);
      }
    } else {
      result.push([topLeft(a)]);
    }
  }
  return result.length > 0 ? rvArray(result) : ERRORS.VALUE;
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
  const pad: ScalarValue = args.length > 3 ? topLeft(args[3]) : ERRORS.NA;
  const result: ScalarValue[][] = [];
  for (let r = 0; r < rowsV.value; r++) {
    const row: ScalarValue[] = [];
    for (let c = 0; c < colsV.value; c++) {
      row.push(r < d.height && c < d.width ? getCell(d, r, c) : pad);
    }
    result.push(row);
  }
  return rvArray(result);
}
