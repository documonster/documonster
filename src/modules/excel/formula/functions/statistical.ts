/**
 * Statistical Functions
 *
 * Native RuntimeValue implementations.
 */

import type { RuntimeValue, ScalarValue, NumberValue, ErrorValue } from "../runtime/values";
import {
  RVKind,
  ERRORS,
  rvNumber,
  rvArray,
  toNumberRV,
  toBooleanRV,
  topLeft,
  isError
} from "../runtime/values";

// ============================================================================
// Local Helpers
// ============================================================================

type NF = (args: RuntimeValue[]) => RuntimeValue;

/**
 * Flatten args to numbers (RuntimeValue equivalent of flattenNumbers).
 * For Array args: iterate rows/cells, keep only RVKind.Number (skip booleans/strings/blanks).
 * For direct scalars: coerce via toNumberRV.
 * Errors are always propagated.
 */
function flattenNumbers(args: RuntimeValue[]): (number | ErrorValue)[] {
  const result: (number | ErrorValue)[] = [];
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            result.push(cell);
          } else if (cell.kind === RVKind.Number) {
            result.push(cell.value);
          }
          // skip booleans, strings, blanks in ranges
        }
      }
    } else {
      const sv = topLeft(arg);
      if (sv.kind === RVKind.Error) {
        result.push(sv);
      } else if (sv.kind === RVKind.Blank) {
        // skip blank direct args (Excel behavior for aggregates)
      } else {
        const n = toNumberRV(sv);
        if (n.kind === RVKind.Error) {
          result.push(n);
        } else {
          result.push(n.value);
        }
      }
    }
  }
  return result;
}

/**
 * Flatten all scalar values from args (RuntimeValue equivalent of flattenAll).
 */
function flattenAll(args: RuntimeValue[]): ScalarValue[] {
  const result: ScalarValue[] = [];
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          result.push(cell);
        }
      }
    } else {
      result.push(topLeft(arg));
    }
  }
  return result;
}

/** Return the first error in a list, or null. */
function firstError(values: (number | ErrorValue)[]): ErrorValue | null {
  for (const v of values) {
    if (typeof v !== "number") {
      return v;
    }
  }
  return null;
}

/**
 * Extract a number from a single arg (with coercion). Returns NumberValue or ErrorValue.
 */
function numArg(args: RuntimeValue[], idx: number): NumberValue | ErrorValue {
  return toNumberRV(topLeft(args[idx]));
}

/**
 * Extract a boolean from a single arg. Returns the boolean or ErrorValue.
 */
function boolArg(
  args: RuntimeValue[],
  idx: number
): { ok: true; value: boolean } | { ok: false; error: ErrorValue } {
  const rv = toBooleanRV(topLeft(args[idx]));
  if (rv.kind === RVKind.Error) {
    return { ok: false, error: rv };
  }
  return { ok: true, value: rv.value };
}

/**
 * Check if an arg is an array.
 */
function isArrayArg(arg: RuntimeValue): boolean {
  return arg.kind === RVKind.Array;
}

// ============================================================================
// MEDIAN, LARGE, SMALL, RANK
// ============================================================================

export const fnMEDIAN: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.NUM;
  }
  const sorted = (nums as number[]).slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return rvNumber(sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid]);
};

export const fnLARGE: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const nums = flattenNumbers([args[0]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const k = numArg(args, 1);
  if (k.kind === RVKind.Error) {
    return k;
  }
  const sorted = (nums as number[]).slice().sort((a, b) => b - a);
  const idx = Math.floor(k.value) - 1;
  return idx >= 0 && idx < sorted.length ? rvNumber(sorted[idx]) : ERRORS.NUM;
};

export const fnSMALL: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const nums = flattenNumbers([args[0]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const k = numArg(args, 1);
  if (k.kind === RVKind.Error) {
    return k;
  }
  const sorted = (nums as number[]).slice().sort((a, b) => a - b);
  const idx = Math.floor(k.value) - 1;
  return idx >= 0 && idx < sorted.length ? rvNumber(sorted[idx]) : ERRORS.NUM;
};

export const fnRANK: NF = args => {
  const num = numArg(args, 0);
  if (num.kind === RVKind.Error) {
    return num;
  }
  if (!isArrayArg(args[1])) {
    return ERRORS.VALUE;
  }
  const nums = flattenNumbers([args[1]]);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const orderRV = args.length > 2 ? numArg(args, 2) : rvNumber(0);
  if (orderRV.kind === RVKind.Error) {
    return orderRV;
  }
  const order = orderRV.value;
  const sorted =
    order === 0
      ? (nums as number[]).slice().sort((a, b) => b - a)
      : (nums as number[]).slice().sort((a, b) => a - b);
  const idx = sorted.indexOf(num.value);
  return idx === -1 ? ERRORS.NA : rvNumber(idx + 1);
};

// ============================================================================
// STDEV, STDEVP, VAR, VARP
// ============================================================================

export const fnSTDEV: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length < 2) {
    return ERRORS.DIV0;
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return rvNumber(Math.sqrt(sumSq / (n - 1)));
};

export const fnSTDEVP: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.DIV0;
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return rvNumber(Math.sqrt(sumSq / n));
};

export const fnVAR: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length < 2) {
    return ERRORS.DIV0;
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return rvNumber(sumSq / (n - 1));
};

export const fnVARP: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.DIV0;
  }
  const n = nums.length;
  let sum = 0;
  for (const v of nums) {
    sum += v as number;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of nums) {
    sumSq += ((v as number) - mean) ** 2;
  }
  return rvNumber(sumSq / n);
};

// ============================================================================
// Advanced Statistical Functions — private math helpers
// ============================================================================

// Peter Acklam's rational approximation for the standard normal inverse CDF.
// Accuracy: |error| < 1.15e-9 across the full range (0, 1).
function normSInv(p: number): number {
  if (p <= 0 || p >= 1) {
    return NaN;
  }
  if (p < 0.5) {
    return -normSInv(1 - p);
  }
  // Coefficients for rational approximation
  const a1 = -3.969683028665376e1;
  const a2 = 2.209460984245205e2;
  const a3 = -2.759285104469687e2;
  const a4 = 1.38357751867269e2;
  const a5 = -3.066479806614716e1;
  const a6 = 2.506628277459239;

  const b1 = -5.447609879822406e1;
  const b2 = 1.615858368580409e2;
  const b3 = -1.556989798598866e2;
  const b4 = 6.680131188771972e1;
  const b5 = -1.328068155288572e1;

  const c1 = -7.784894002430293e-3;
  const c2 = -3.223964580411365e-1;
  const c3 = -2.400758277161838;
  const c4 = -2.549732539343734;
  const c5 = 4.374664141464968;
  const c6 = 2.938163982698783;

  const d1 = 7.784695709041462e-3;
  const d2 = 3.224671290700398e-1;
  const d3 = 2.445134137142996;
  const d4 = 3.754408661907416;

  const pLow = 0.02425;
  const pHigh = 1 - pLow;

  if (p < pLow) {
    // Rational approximation for lower region
    const q = Math.sqrt(-2 * Math.log(p));
    return (
      (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
      ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    );
  }
  if (p <= pHigh) {
    // Rational approximation for central region
    const q = p - 0.5;
    const r = q * q;
    return (
      ((((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q) /
      (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    );
  }
  // Upper region — use symmetry
  const q = Math.sqrt(-2 * Math.log(1 - p));
  return -(
    (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
  );
}

// Standard normal CDF approximation (Abramowitz & Stegun)
function normSDist(x: number): number {
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x) / Math.SQRT2;
  const t = 1.0 / (1.0 + p * x);
  const y = 1.0 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  return 0.5 * (1.0 + sign * y);
}

// Standard normal PDF
function normSPdf(x: number): number {
  return Math.exp(-0.5 * x * x) / Math.sqrt(2 * Math.PI);
}

// ============================================================================
// Normal Distribution Functions
// ============================================================================

export const fnNORMSDIST: NF = args => {
  const z = numArg(args, 0);
  if (z.kind === RVKind.Error) {
    return z;
  }
  // Legacy NORM.S.DIST compatibility: single arg = CDF
  if (args.length > 1) {
    const cum = boolArg(args, 1);
    if (!cum.ok) {
      return cum.error;
    }
    return rvNumber(cum.value ? normSDist(z.value) : normSPdf(z.value));
  }
  return rvNumber(normSDist(z.value));
};

export const fnNORMDIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const stddev = numArg(args, 2);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  if (stddev.value <= 0) {
    return ERRORS.NUM;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  const zVal = (x.value - mean.value) / stddev.value;
  if (cum.value) {
    return rvNumber(normSDist(zVal));
  }
  return rvNumber(normSPdf(zVal) / stddev.value);
};

export const fnNORMSINV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  if (p.value <= 0 || p.value >= 1) {
    return ERRORS.NUM;
  }
  return rvNumber(normSInv(p.value));
};

export const fnNORMINV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const stddev = numArg(args, 2);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  if (p.value <= 0 || p.value >= 1 || stddev.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(mean.value + stddev.value * normSInv(p.value));
};

// ============================================================================
// PERCENTILE, QUARTILE, MODE
// ============================================================================

export const fnPERCENTILE: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const nums = flattenNumbers([args[0]]).filter((v): v is number => typeof v === "number");
  const k = numArg(args, 1);
  if (k.kind === RVKind.Error) {
    return k;
  }
  if (k.value < 0 || k.value > 1 || nums.length === 0) {
    return ERRORS.NUM;
  }
  nums.sort((a, b) => a - b);
  const n = nums.length;
  if (n === 1) {
    return rvNumber(nums[0]);
  }
  const rank = k.value * (n - 1);
  const lower = Math.floor(rank);
  const upper = Math.ceil(rank);
  const frac = rank - lower;
  return rvNumber(nums[lower] + frac * (nums[upper] - nums[lower]));
};

export const fnPERCENTILEEXC: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const nums = flattenNumbers([args[0]]).filter((v): v is number => typeof v === "number");
  const k = numArg(args, 1);
  if (k.kind === RVKind.Error) {
    return k;
  }
  const n = nums.length;
  if (k.value <= 0 || k.value >= 1 || n === 0) {
    return ERRORS.NUM;
  }
  if (k.value < 1 / (n + 1) || k.value > n / (n + 1)) {
    return ERRORS.NUM;
  }
  nums.sort((a, b) => a - b);
  const rank = k.value * (n + 1) - 1;
  const lower = Math.floor(rank);
  const upper = Math.ceil(rank);
  const frac = rank - lower;
  return rvNumber(
    nums[Math.max(0, lower)] + frac * (nums[Math.min(n - 1, upper)] - nums[Math.max(0, lower)])
  );
};

export const fnQUARTILE: NF = args => {
  const quart = numArg(args, 1);
  if (quart.kind === RVKind.Error) {
    return quart;
  }
  if (quart.value < 0 || quart.value > 4) {
    return ERRORS.NUM;
  }
  return fnPERCENTILE([args[0], rvNumber(quart.value / 4)]);
};

export const fnQUARTILEEXC: NF = args => {
  const quart = numArg(args, 1);
  if (quart.kind === RVKind.Error) {
    return quart;
  }
  if (quart.value < 1 || quart.value > 3) {
    return ERRORS.NUM;
  }
  return fnPERCENTILEEXC([args[0], rvNumber(quart.value / 4)]);
};

export const fnMODE: NF = args => {
  const all = flattenNumbers(args);
  const err = firstError(all);
  if (err) {
    return err;
  }
  const nums = all as number[];
  if (nums.length === 0) {
    return ERRORS.NA;
  }
  const counts = new Map<number, number>();
  let maxCount = 0;
  let mode = nums[0];
  for (const n of nums) {
    const c = (counts.get(n) ?? 0) + 1;
    counts.set(n, c);
    if (c > maxCount) {
      maxCount = c;
      mode = n;
    }
  }
  return maxCount > 1 ? rvNumber(mode) : ERRORS.NA;
};

// ============================================================================
// Paired-array functions: CORREL, SLOPE, INTERCEPT, RSQ, FORECAST
// ============================================================================

export const fnCORREL: NF = args => {
  if (!isArrayArg(args[0]) || !isArrayArg(args[1])) {
    return ERRORS.VALUE;
  }
  const xs = flattenNumbers([args[0]]).filter((v): v is number => typeof v === "number");
  const ys = flattenNumbers([args[1]]).filter((v): v is number => typeof v === "number");
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return ERRORS.DIV0;
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  let num = 0;
  let denomX = 0;
  let denomY = 0;
  for (let i = 0; i < n; i++) {
    const dx = xs[i] - meanX;
    const dy = ys[i] - meanY;
    num += dx * dy;
    denomX += dx * dx;
    denomY += dy * dy;
  }
  const denom = Math.sqrt(denomX * denomY);
  return denom === 0 ? ERRORS.DIV0 : rvNumber(num / denom);
};

export const fnSLOPE: NF = args => {
  if (!isArrayArg(args[0]) || !isArrayArg(args[1])) {
    return ERRORS.VALUE;
  }
  const ys = flattenNumbers([args[0]]).filter((v): v is number => typeof v === "number");
  const xs = flattenNumbers([args[1]]).filter((v): v is number => typeof v === "number");
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return ERRORS.DIV0;
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  let num = 0;
  let denom = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - meanX) * (ys[i] - meanY);
    denom += (xs[i] - meanX) ** 2;
  }
  return denom === 0 ? ERRORS.DIV0 : rvNumber(num / denom);
};

export const fnINTERCEPT: NF = args => {
  if (!isArrayArg(args[0]) || !isArrayArg(args[1])) {
    return ERRORS.VALUE;
  }
  const ys = flattenNumbers([args[0]]).filter((v): v is number => typeof v === "number");
  const xs = flattenNumbers([args[1]]).filter((v): v is number => typeof v === "number");
  const n = Math.min(xs.length, ys.length);
  if (n < 2) {
    return ERRORS.DIV0;
  }
  let sumX = 0;
  let sumY = 0;
  for (let i = 0; i < n; i++) {
    sumX += xs[i];
    sumY += ys[i];
  }
  const meanX = sumX / n;
  const meanY = sumY / n;
  const slopeResult = fnSLOPE(args);
  if (isError(slopeResult)) {
    return slopeResult;
  }
  const slope = (slopeResult as NumberValue).value;
  return rvNumber(meanY - slope * meanX);
};

export const fnRSQ: NF = args => {
  const r = fnCORREL(args);
  if (isError(r)) {
    return r;
  }
  return rvNumber((r as NumberValue).value ** 2);
};

export const fnFORECAST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const slope = fnSLOPE([args[1], args[2]]);
  if (isError(slope)) {
    return slope;
  }
  const intercept = fnINTERCEPT([args[1], args[2]]);
  if (isError(intercept)) {
    return intercept;
  }
  return rvNumber((intercept as NumberValue).value + (slope as NumberValue).value * x.value);
};

// ============================================================================
// FACT, FACTDOUBLE, COMBIN, COMBINA, PERMUT
// Re-exported from math.ts — canonical definitions live there.
// ============================================================================

export { fnFACT, fnFACTDOUBLE, fnCOMBIN, fnCOMBINA, fnPERMUT } from "./math";

// ============================================================================
// GEOMEAN, HARMEAN, TRIMMEAN, DEVSQ, AVEDEV
// ============================================================================

export const fnGEOMEAN: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.NUM;
  }
  let logSum = 0;
  for (const n of nums) {
    if ((n as number) <= 0) {
      return ERRORS.NUM;
    }
    logSum += Math.log(n as number);
  }
  return rvNumber(Math.exp(logSum / nums.length));
};

export const fnHARMEAN: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.NUM;
  }
  let recipSum = 0;
  for (const n of nums) {
    if ((n as number) <= 0) {
      return ERRORS.NUM;
    }
    recipSum += 1 / (n as number);
  }
  return rvNumber(nums.length / recipSum);
};

export const fnTRIMMEAN: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const all = flattenNumbers([args[0]]);
  const err = firstError(all);
  if (err) {
    return err;
  }
  const nums = all as number[];
  const pct = numArg(args, 1);
  if (pct.kind === RVKind.Error) {
    return pct;
  }
  if (pct.value < 0 || pct.value >= 1) {
    return ERRORS.NUM;
  }
  nums.sort((a, b) => a - b);
  const trimCount = Math.floor((nums.length * pct.value) / 2);
  const trimmed = nums.slice(trimCount, nums.length - trimCount);
  if (trimmed.length === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(trimmed.reduce((a, b) => a + b, 0) / trimmed.length);
};

export const fnDEVSQ: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  const mean = sum / nums.length;
  let result = 0;
  for (const n of nums) {
    result += ((n as number) - mean) ** 2;
  }
  return rvNumber(result);
};

export const fnAVEDEV: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.NUM;
  }
  let sum = 0;
  for (const n of nums) {
    sum += n as number;
  }
  const mean = sum / nums.length;
  let result = 0;
  for (const n of nums) {
    result += Math.abs((n as number) - mean);
  }
  return rvNumber(result / nums.length);
};

// ============================================================================
// CONFIDENCE, FISHER, AVERAGEA, MAXA, MINA
// ============================================================================

export const fnCONFIDENCENORM: NF = args => {
  const alpha = numArg(args, 0);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const stddev = numArg(args, 1);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  const size = numArg(args, 2);
  if (size.kind === RVKind.Error) {
    return size;
  }
  if (alpha.value <= 0 || alpha.value >= 1 || stddev.value <= 0 || size.value < 1) {
    return ERRORS.NUM;
  }
  return rvNumber((normSInv(1 - alpha.value / 2) * stddev.value) / Math.sqrt(size.value));
};

export const fnFISHER: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  if (x.value <= -1 || x.value >= 1) {
    return ERRORS.NUM;
  }
  return rvNumber(0.5 * Math.log((1 + x.value) / (1 - x.value)));
};

export const fnFISHERINV: NF = args => {
  const y = numArg(args, 0);
  if (y.kind === RVKind.Error) {
    return y;
  }
  const e2y = Math.exp(2 * y.value);
  return rvNumber((e2y - 1) / (e2y + 1));
};

export const fnAVERAGEA: NF = args => {
  const all = flattenAll(args);
  if (all.length === 0) {
    return ERRORS.DIV0;
  }
  let sum = 0;
  let count = 0;
  for (const v of all) {
    if (v.kind === RVKind.Blank) {
      continue;
    }
    if (v.kind === RVKind.Error) {
      return v;
    }
    if (v.kind === RVKind.Number) {
      sum += v.value;
    } else if (v.kind === RVKind.Boolean) {
      sum += v.value ? 1 : 0;
    }
    // Text = 0 for AVERAGEA (no addition needed)
    count++;
  }
  return count === 0 ? ERRORS.DIV0 : rvNumber(sum / count);
};

export const fnMAXA: NF = args => {
  const all = flattenAll(args);
  let max = -Infinity;
  let found = false;
  for (const v of all) {
    if (v.kind === RVKind.Blank) {
      continue;
    }
    if (v.kind === RVKind.Error) {
      return v;
    }
    let n: number;
    if (v.kind === RVKind.Number) {
      n = v.value;
    } else if (v.kind === RVKind.Boolean) {
      n = v.value ? 1 : 0;
    } else {
      n = 0;
    }
    if (n > max) {
      max = n;
    }
    found = true;
  }
  return rvNumber(found ? max : 0);
};

export const fnMINA: NF = args => {
  const all = flattenAll(args);
  let min = Infinity;
  let found = false;
  for (const v of all) {
    if (v.kind === RVKind.Blank) {
      continue;
    }
    if (v.kind === RVKind.Error) {
      return v;
    }
    let n: number;
    if (v.kind === RVKind.Number) {
      n = v.value;
    } else if (v.kind === RVKind.Boolean) {
      n = v.value ? 1 : 0;
    } else {
      n = 0;
    }
    if (n < min) {
      min = n;
    }
    found = true;
  }
  return rvNumber(found ? min : 0);
};

// ============================================================================
// Private helpers for distributions (pure number → number, unchanged)
// ============================================================================

function gammaFn(z: number): number {
  if (z < 0.5) {
    return Math.PI / (Math.sin(Math.PI * z) * gammaFn(1 - z));
  }
  z -= 1;
  const g = 7;
  const c = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
    -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
    1.5056327351493116e-7
  ];
  let x = c[0];
  for (let i = 1; i < g + 2; i++) {
    x += c[i] / (z + i);
  }
  const t = z + g + 0.5;
  return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * x;
}

function lnGamma(x: number): number {
  return Math.log(gammaFn(x));
}

function betaIncomplete(x: number, a: number, b: number): number {
  if (x <= 0) {
    return 0;
  }
  if (x >= 1) {
    return 1;
  }
  if (x > (a + 1) / (a + b + 2)) {
    return 1 - betaIncomplete(1 - x, b, a);
  }
  const lbeta = lnGamma(a) + lnGamma(b) - lnGamma(a + b);
  const front = Math.exp(Math.log(x) * a + Math.log(1 - x) * b - lbeta) / a;
  let f = 1;
  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1);
  if (Math.abs(d) < 1e-30) {
    d = 1e-30;
  }
  d = 1 / d;
  f = d;
  for (let m = 1; m <= 200; m++) {
    let num = (m * (b - m) * x) / ((a + 2 * m - 1) * (a + 2 * m));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) {
      c = 1e-30;
    }
    d = 1 / d;
    f *= c * d;
    num = -((a + m) * (a + b + m) * x) / ((a + 2 * m) * (a + 2 * m + 1));
    d = 1 + num * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    c = 1 + num / c;
    if (Math.abs(c) < 1e-30) {
      c = 1e-30;
    }
    d = 1 / d;
    const delta = c * d;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) {
      break;
    }
  }
  return front * f;
}

function gammaIncomplete(a: number, x: number): number {
  if (x < 0) {
    return 0;
  }
  if (x === 0) {
    return 0;
  }
  if (x < a + 1) {
    let sum = 1 / a;
    let term = 1 / a;
    for (let n = 1; n <= 200; n++) {
      term *= x / (a + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-14) {
        break;
      }
    }
    return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
  }
  let f = 1;
  const b0 = x + 1 - a;
  let ci = 1e30;
  let d = 1 / b0;
  f = d;
  for (let i = 1; i <= 200; i++) {
    const an = -i * (i - a);
    const bn = x + 2 * i + 1 - a;
    d = bn + an * d;
    if (Math.abs(d) < 1e-30) {
      d = 1e-30;
    }
    ci = bn + an / ci;
    if (Math.abs(ci) < 1e-30) {
      ci = 1e-30;
    }
    d = 1 / d;
    const delta = d * ci;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-10) {
      break;
    }
  }
  return 1 - f * Math.exp(-x + a * Math.log(x) - lnGamma(a));
}

// ============================================================================
// More Statistical Distribution Functions
// ============================================================================

export const fnPOISSON_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const cum = boolArg(args, 2);
  if (!cum.ok) {
    return cum.error;
  }
  const k = Math.floor(x.value);
  if (k < 0 || mean.value < 0) {
    return ERRORS.NUM;
  }
  if (!cum.value) {
    return rvNumber(Math.exp(-mean.value + k * Math.log(mean.value) - lnGamma(k + 1)));
  }
  return rvNumber(1 - gammaIncomplete(k + 1, mean.value));
};

export const fnBINOM_DIST: NF = args => {
  const numS = numArg(args, 0);
  if (numS.kind === RVKind.Error) {
    return numS;
  }
  const trials = numArg(args, 1);
  if (trials.kind === RVKind.Error) {
    return trials;
  }
  const probS = numArg(args, 2);
  if (probS.kind === RVKind.Error) {
    return probS;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  const k = Math.floor(numS.value);
  const n = Math.floor(trials.value);
  if (k < 0 || n < 0 || k > n || probS.value < 0 || probS.value > 1) {
    return ERRORS.NUM;
  }
  const pmf = (ki: number): number => {
    const lnC = lnGamma(n + 1) - lnGamma(ki + 1) - lnGamma(n - ki + 1);
    return Math.exp(lnC + ki * Math.log(probS.value) + (n - ki) * Math.log(1 - probS.value));
  };
  if (!cum.value) {
    return rvNumber(pmf(k));
  }
  let sum = 0;
  for (let i = 0; i <= k; i++) {
    sum += pmf(i);
  }
  return rvNumber(sum);
};

export const fnBINOM_INV: NF = args => {
  const trials = numArg(args, 0);
  if (trials.kind === RVKind.Error) {
    return trials;
  }
  const probS = numArg(args, 1);
  if (probS.kind === RVKind.Error) {
    return probS;
  }
  const alpha = numArg(args, 2);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const n = Math.floor(trials.value);
  if (n < 0 || probS.value < 0 || probS.value > 1 || alpha.value < 0 || alpha.value > 1) {
    return ERRORS.NUM;
  }
  let cdf = 0;
  for (let k = 0; k <= n; k++) {
    const lnC = lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);
    cdf += Math.exp(lnC + k * Math.log(probS.value) + (n - k) * Math.log(1 - probS.value));
    if (cdf >= alpha.value) {
      return rvNumber(k);
    }
  }
  return rvNumber(n);
};

export const fnHYPGEOM_DIST: NF = args => {
  const sampleS = numArg(args, 0);
  if (sampleS.kind === RVKind.Error) {
    return sampleS;
  }
  const numberSample = numArg(args, 1);
  if (numberSample.kind === RVKind.Error) {
    return numberSample;
  }
  const popS = numArg(args, 2);
  if (popS.kind === RVKind.Error) {
    return popS;
  }
  const numberPop = numArg(args, 3);
  if (numberPop.kind === RVKind.Error) {
    return numberPop;
  }
  const cum = boolArg(args, 4);
  if (!cum.ok) {
    return cum.error;
  }
  const ss = Math.floor(sampleS.value);
  const ns = Math.floor(numberSample.value);
  const ps = Math.floor(popS.value);
  const np = Math.floor(numberPop.value);
  const pmf = (k: number): number =>
    Math.exp(
      lnGamma(ps + 1) -
        lnGamma(k + 1) -
        lnGamma(ps - k + 1) +
        lnGamma(np - ps + 1) -
        lnGamma(ns - k + 1) -
        lnGamma(np - ps - ns + k + 1) -
        lnGamma(np + 1) +
        lnGamma(ns + 1) +
        lnGamma(np - ns + 1)
    );
  if (!cum.value) {
    return rvNumber(pmf(ss));
  }
  let sum = 0;
  for (let k = 0; k <= ss; k++) {
    sum += pmf(k);
  }
  return rvNumber(sum);
};

export const fnNEGBINOM_DIST: NF = args => {
  const numF = numArg(args, 0);
  if (numF.kind === RVKind.Error) {
    return numF;
  }
  const numS = numArg(args, 1);
  if (numS.kind === RVKind.Error) {
    return numS;
  }
  const probS = numArg(args, 2);
  if (probS.kind === RVKind.Error) {
    return probS;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  const f = Math.floor(numF.value);
  const s = Math.floor(numS.value);
  if (f < 0 || s < 1 || probS.value < 0 || probS.value > 1) {
    return ERRORS.NUM;
  }
  const pmf = (k: number): number => {
    const lnC = lnGamma(k + s) - lnGamma(s) - lnGamma(k + 1);
    return Math.exp(lnC + s * Math.log(probS.value) + k * Math.log(1 - probS.value));
  };
  if (!cum.value) {
    return rvNumber(pmf(f));
  }
  let sum = 0;
  for (let k = 0; k <= f; k++) {
    sum += pmf(k);
  }
  return rvNumber(sum);
};

export const fnCHISQ_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  const cum = boolArg(args, 2);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value < 0 || df.value < 1) {
    return ERRORS.NUM;
  }
  const k = Math.floor(df.value);
  if (cum.value) {
    return rvNumber(gammaIncomplete(k / 2, x.value / 2));
  }
  const halfK = k / 2;
  return rvNumber(Math.exp((halfK - 1) * Math.log(x.value / 2) - x.value / 2 - lnGamma(halfK)) / 2);
};

export const fnCHISQ_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (p.value < 0 || p.value >= 1 || df.value < 1) {
    return ERRORS.NUM;
  }
  let x = df.value;
  for (let iter = 0; iter < 100; iter++) {
    const cdf = gammaIncomplete(df.value / 2, x / 2);
    const halfK = df.value / 2;
    const pdf = Math.exp((halfK - 1) * Math.log(x / 2) - x / 2 - lnGamma(halfK)) / 2;
    if (Math.abs(pdf) < 1e-15) {
      break;
    }
    const delta = (cdf - p.value) / pdf;
    x -= delta;
    if (x <= 0) {
      x = 0.001;
    }
    if (Math.abs(delta) < 1e-10) {
      break;
    }
  }
  return rvNumber(x);
};

export const fnCHISQ_DIST_RT: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (x.value < 0 || df.value < 1) {
    return ERRORS.NUM;
  }
  return rvNumber(1 - gammaIncomplete(Math.floor(df.value) / 2, x.value / 2));
};

export const fnF_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df1 = numArg(args, 1);
  if (df1.kind === RVKind.Error) {
    return df1;
  }
  const df2 = numArg(args, 2);
  if (df2.kind === RVKind.Error) {
    return df2;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value < 0 || df1.value < 1 || df2.value < 1) {
    return ERRORS.NUM;
  }
  const d1 = Math.floor(df1.value);
  const d2 = Math.floor(df2.value);
  if (cum.value) {
    return rvNumber(betaIncomplete((d1 * x.value) / (d1 * x.value + d2), d1 / 2, d2 / 2));
  }
  const num =
    (Math.pow(d1 * x.value, d1 / 2) * Math.pow(d2, d2 / 2)) /
    Math.pow(d1 * x.value + d2, (d1 + d2) / 2);
  const denom = x.value * Math.exp(lnGamma(d1 / 2) + lnGamma(d2 / 2) - lnGamma((d1 + d2) / 2));
  return rvNumber(denom === 0 ? 0 : num / denom);
};

export const fnF_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const df1 = numArg(args, 1);
  if (df1.kind === RVKind.Error) {
    return df1;
  }
  const df2 = numArg(args, 2);
  if (df2.kind === RVKind.Error) {
    return df2;
  }
  if (p.value < 0 || p.value >= 1 || df1.value < 1 || df2.value < 1) {
    return ERRORS.NUM;
  }
  const d1 = Math.floor(df1.value);
  const d2 = Math.floor(df2.value);
  let lo = 0;
  let hi = 1000;
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    const cdf = betaIncomplete((d1 * mid) / (d1 * mid + d2), d1 / 2, d2 / 2);
    if (cdf < p.value) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-10) {
      break;
    }
  }
  return rvNumber((lo + hi) / 2);
};

export const fnT_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  const cum = boolArg(args, 2);
  if (!cum.ok) {
    return cum.error;
  }
  if (df.value < 1) {
    return ERRORS.NUM;
  }
  const v = Math.floor(df.value);
  if (cum.value) {
    const t = v / (v + x.value * x.value);
    const halfBeta = 0.5 * betaIncomplete(t, v / 2, 0.5);
    return rvNumber(x.value >= 0 ? 1 - halfBeta : halfBeta);
  }
  return rvNumber(
    Math.exp(lnGamma((v + 1) / 2) - lnGamma(v / 2)) /
      (Math.sqrt(v * Math.PI) * Math.pow(1 + (x.value * x.value) / v, (v + 1) / 2))
  );
};

export const fnT_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (p.value <= 0 || p.value >= 1 || df.value < 1) {
    return ERRORS.NUM;
  }
  let x = normSInv(p.value);
  const v = Math.floor(df.value);
  for (let iter = 0; iter < 100; iter++) {
    const t = v / (v + x * x);
    const halfBeta = 0.5 * betaIncomplete(t, v / 2, 0.5);
    const cdf = x >= 0 ? 1 - halfBeta : halfBeta;
    const pdf =
      Math.exp(lnGamma((v + 1) / 2) - lnGamma(v / 2)) /
      (Math.sqrt(v * Math.PI) * Math.pow(1 + (x * x) / v, (v + 1) / 2));
    if (Math.abs(pdf) < 1e-15) {
      break;
    }
    const delta = (cdf - p.value) / pdf;
    x -= delta;
    if (Math.abs(delta) < 1e-10) {
      break;
    }
  }
  return rvNumber(x);
};

export const fnT_DIST_2T: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (x.value < 0 || df.value < 1) {
    return ERRORS.NUM;
  }
  const v = Math.floor(df.value);
  return rvNumber(betaIncomplete(v / (v + x.value * x.value), v / 2, 0.5));
};

export const fnT_DIST_RT: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (df.value < 1) {
    return ERRORS.NUM;
  }
  const v = Math.floor(df.value);
  const halfBeta = 0.5 * betaIncomplete(v / (v + x.value * x.value), v / 2, 0.5);
  // T.DIST.RT(x, df) = right-tail = 1 - CDF(x)
  // For x >= 0: right-tail = halfBeta
  // For x < 0: right-tail = 1 - halfBeta
  return rvNumber(x.value >= 0 ? halfBeta : 1 - halfBeta);
};

export const fnT_INV_2T: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const df = numArg(args, 1);
  if (df.kind === RVKind.Error) {
    return df;
  }
  if (p.value <= 0 || p.value > 1 || df.value < 1) {
    return ERRORS.NUM;
  }
  const result = fnT_INV([rvNumber(1 - p.value / 2), args[1]]);
  if (isError(result)) {
    return result;
  }
  return rvNumber(Math.abs((result as NumberValue).value));
};

// ============================================================================
// BETA, GAMMA, EXPON, WEIBULL, LOGNORM distributions
// ============================================================================

export const fnBETA_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const alpha = numArg(args, 1);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const beta = numArg(args, 2);
  if (beta.kind === RVKind.Error) {
    return beta;
  }
  let cumVal = true;
  if (args.length > 3) {
    const cum = boolArg(args, 3);
    if (!cum.ok) {
      return cum.error;
    }
    cumVal = cum.value;
  }
  let A = 0;
  if (args.length > 4) {
    const aRV = numArg(args, 4);
    if (aRV.kind === RVKind.Error) {
      return aRV;
    }
    A = aRV.value;
  }
  let B = 1;
  if (args.length > 5) {
    const bRV = numArg(args, 5);
    if (bRV.kind === RVKind.Error) {
      return bRV;
    }
    B = bRV.value;
  }
  if (alpha.value <= 0 || beta.value <= 0 || B <= A) {
    return ERRORS.NUM;
  }
  const xn = (x.value - A) / (B - A);
  if (xn < 0 || xn > 1) {
    return ERRORS.NUM;
  }
  if (cumVal) {
    return rvNumber(betaIncomplete(xn, alpha.value, beta.value));
  }
  return rvNumber(
    Math.exp(
      (alpha.value - 1) * Math.log(xn) +
        (beta.value - 1) * Math.log(1 - xn) -
        lnGamma(alpha.value) -
        lnGamma(beta.value) +
        lnGamma(alpha.value + beta.value)
    ) /
      (B - A)
  );
};

export const fnBETA_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const alpha = numArg(args, 1);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const beta = numArg(args, 2);
  if (beta.kind === RVKind.Error) {
    return beta;
  }
  let A = 0;
  if (args.length > 3) {
    const aRV = numArg(args, 3);
    if (aRV.kind === RVKind.Error) {
      return aRV;
    }
    A = aRV.value;
  }
  let B = 1;
  if (args.length > 4) {
    const bRV = numArg(args, 4);
    if (bRV.kind === RVKind.Error) {
      return bRV;
    }
    B = bRV.value;
  }
  if (p.value < 0 || p.value > 1 || alpha.value <= 0 || beta.value <= 0) {
    return ERRORS.NUM;
  }
  let lo = 0;
  let hi = 1;
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    if (betaIncomplete(mid, alpha.value, beta.value) < p.value) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-12) {
      break;
    }
  }
  return rvNumber(A + ((lo + hi) / 2) * (B - A));
};

export const fnGAMMA: NF = args => {
  const n = numArg(args, 0);
  if (n.kind === RVKind.Error) {
    return n;
  }
  if (n.value <= 0 && n.value === Math.floor(n.value)) {
    return ERRORS.NUM;
  }
  return rvNumber(gammaFn(n.value));
};

export const fnGAMMALN: NF = args => {
  const n = numArg(args, 0);
  if (n.kind === RVKind.Error) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(lnGamma(n.value));
};

export const fnGAMMA_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const alpha = numArg(args, 1);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const beta = numArg(args, 2);
  if (beta.kind === RVKind.Error) {
    return beta;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value < 0 || alpha.value <= 0 || beta.value <= 0) {
    return ERRORS.NUM;
  }
  if (cum.value) {
    return rvNumber(gammaIncomplete(alpha.value, x.value / beta.value));
  }
  return rvNumber(
    Math.exp(
      (alpha.value - 1) * Math.log(x.value) -
        x.value / beta.value -
        alpha.value * Math.log(beta.value) -
        lnGamma(alpha.value)
    )
  );
};

export const fnGAMMA_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const alpha = numArg(args, 1);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const beta = numArg(args, 2);
  if (beta.kind === RVKind.Error) {
    return beta;
  }
  if (p.value < 0 || p.value >= 1 || alpha.value <= 0 || beta.value <= 0) {
    return ERRORS.NUM;
  }
  let lo = 0;
  let hi = Math.max(alpha.value * beta.value * 10, 100);
  for (let i = 0; i < 100; i++) {
    const mid = (lo + hi) / 2;
    if (gammaIncomplete(alpha.value, mid / beta.value) < p.value) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-10) {
      break;
    }
  }
  return rvNumber((lo + hi) / 2);
};

export const fnEXPON_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const lambda = numArg(args, 1);
  if (lambda.kind === RVKind.Error) {
    return lambda;
  }
  const cum = boolArg(args, 2);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value < 0 || lambda.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(
    cum.value
      ? 1 - Math.exp(-lambda.value * x.value)
      : lambda.value * Math.exp(-lambda.value * x.value)
  );
};

export const fnWEIBULL_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const alpha = numArg(args, 1);
  if (alpha.kind === RVKind.Error) {
    return alpha;
  }
  const beta = numArg(args, 2);
  if (beta.kind === RVKind.Error) {
    return beta;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value < 0 || alpha.value <= 0 || beta.value <= 0) {
    return ERRORS.NUM;
  }
  if (cum.value) {
    return rvNumber(1 - Math.exp(-Math.pow(x.value / beta.value, alpha.value)));
  }
  return rvNumber(
    (alpha.value / beta.value) *
      Math.pow(x.value / beta.value, alpha.value - 1) *
      Math.exp(-Math.pow(x.value / beta.value, alpha.value))
  );
};

export const fnLOGNORM_DIST: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const stddev = numArg(args, 2);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  const cum = boolArg(args, 3);
  if (!cum.ok) {
    return cum.error;
  }
  if (x.value <= 0 || stddev.value <= 0) {
    return ERRORS.NUM;
  }
  const z = (Math.log(x.value) - mean.value) / stddev.value;
  if (cum.value) {
    return rvNumber(normSDist(z));
  }
  return rvNumber(normSPdf(z) / (x.value * stddev.value));
};

export const fnLOGNORM_INV: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const stddev = numArg(args, 2);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  if (p.value <= 0 || p.value >= 1 || stddev.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.exp(mean.value + stddev.value * normSInv(p.value)));
};

export const fnPHI: NF = args => {
  const x = numArg(args, 0);
  return x.kind === RVKind.Error ? x : rvNumber(normSPdf(x.value));
};

export const fnGAUSS: NF = args => {
  const z = numArg(args, 0);
  return z.kind === RVKind.Error ? z : rvNumber(normSDist(z.value) - 0.5);
};

// ============================================================================
// ERF, ERFC, STANDARDIZE
// ============================================================================

function erfFn(x: number): number {
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const t = 1.0 / (1.0 + p * ax);
  const y = 1.0 - ((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.exp(-ax * ax);
  return sign * y;
}

export const fnERF: NF = args => {
  const lower = numArg(args, 0);
  if (lower.kind === RVKind.Error) {
    return lower;
  }
  if (args.length > 1) {
    const upper = numArg(args, 1);
    if (upper.kind === RVKind.Error) {
      return upper;
    }
    return rvNumber(erfFn(upper.value) - erfFn(lower.value));
  }
  return rvNumber(erfFn(lower.value));
};

export const fnERFC: NF = args => {
  const x = numArg(args, 0);
  return x.kind === RVKind.Error ? x : rvNumber(1 - erfFn(x.value));
};

export const fnSTANDARDIZE: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const mean = numArg(args, 1);
  if (mean.kind === RVKind.Error) {
    return mean;
  }
  const stddev = numArg(args, 2);
  if (stddev.kind === RVKind.Error) {
    return stddev;
  }
  if (stddev.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber((x.value - mean.value) / stddev.value);
};

// ============================================================================
// Array-returning functions: FREQUENCY, GROWTH, TREND, LINEST, LOGEST
// ============================================================================

export const fnFREQUENCY: NF = args => {
  if (!isArrayArg(args[0]) || !isArrayArg(args[1])) {
    return ERRORS.VALUE;
  }
  const rawData = flattenNumbers([args[0]]);
  const dataErr = firstError(rawData);
  if (dataErr) {
    return dataErr;
  }
  const data = rawData as number[];
  const rawBins = flattenNumbers([args[1]]);
  const binsErr = firstError(rawBins);
  if (binsErr) {
    return binsErr;
  }
  const bins = rawBins as number[];
  bins.sort((a, b) => a - b);
  const result: ScalarValue[][] = [];
  for (let i = 0; i <= bins.length; i++) {
    let count = 0;
    for (const d of data) {
      if (i === 0 && d <= bins[0]) {
        count++;
      } else if (i === bins.length && d > bins[bins.length - 1]) {
        count++;
      } else if (i > 0 && i < bins.length && d > bins[i - 1] && d <= bins[i]) {
        count++;
      }
    }
    result.push([rvNumber(count)]);
  }
  return rvArray(result);
};

export const fnGROWTH: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const rawY = flattenNumbers([args[0]]);
  const yErr = firstError(rawY);
  if (yErr) {
    return yErr;
  }
  const knownY = rawY as number[];
  let knownX: number[];
  if (args.length > 1 && isArrayArg(args[1])) {
    const rawX = flattenNumbers([args[1]]);
    const xErr = firstError(rawX);
    if (xErr) {
      return xErr;
    }
    knownX = rawX as number[];
  } else {
    knownX = knownY.map((_, i) => i + 1);
  }
  let newX: number[];
  if (args.length > 2 && isArrayArg(args[2])) {
    const rawNewX = flattenNumbers([args[2]]);
    const newXErr = firstError(rawNewX);
    if (newXErr) {
      return newXErr;
    }
    newX = rawNewX as number[];
  } else {
    newX = knownX;
  }
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return ERRORS.VALUE;
  }
  let sumX = 0,
    sumLnY = 0,
    sumXLnY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    if (knownY[i] <= 0) {
      return ERRORS.NUM;
    }
    sumX += knownX[i];
    sumLnY += Math.log(knownY[i]);
    sumXLnY += knownX[i] * Math.log(knownY[i]);
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return ERRORS.DIV0;
  }
  const lnM = (n * sumXLnY - sumX * sumLnY) / denom;
  const lnB = (sumLnY - lnM * sumX) / n;
  const rows: ScalarValue[][] = newX.map(x => [rvNumber(Math.exp(lnB + lnM * x))]);
  return rvArray(rows);
};

export const fnTREND: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const rawY = flattenNumbers([args[0]]);
  const yErr = firstError(rawY);
  if (yErr) {
    return yErr;
  }
  const knownY = rawY as number[];
  let knownX: number[];
  if (args.length > 1 && isArrayArg(args[1])) {
    const rawX = flattenNumbers([args[1]]);
    const xErr = firstError(rawX);
    if (xErr) {
      return xErr;
    }
    knownX = rawX as number[];
  } else {
    knownX = knownY.map((_, i) => i + 1);
  }
  let newX: number[];
  if (args.length > 2 && isArrayArg(args[2])) {
    const rawNewX = flattenNumbers([args[2]]);
    const newXErr = firstError(rawNewX);
    if (newXErr) {
      return newXErr;
    }
    newX = rawNewX as number[];
  } else {
    newX = knownX;
  }
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return ERRORS.VALUE;
  }
  let sumX = 0,
    sumY = 0,
    sumXY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    sumX += knownX[i];
    sumY += knownY[i];
    sumXY += knownX[i] * knownY[i];
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return ERRORS.DIV0;
  }
  const m = (n * sumXY - sumX * sumY) / denom;
  const b = (sumY - m * sumX) / n;
  const rows: ScalarValue[][] = newX.map(x => [rvNumber(b + m * x)]);
  return rvArray(rows);
};

export const fnLINEST: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const rawY = flattenNumbers([args[0]]);
  const yErr = firstError(rawY);
  if (yErr) {
    return yErr;
  }
  const knownY = rawY as number[];
  let knownX: number[];
  if (args.length > 1 && isArrayArg(args[1])) {
    const rawX = flattenNumbers([args[1]]);
    const xErr = firstError(rawX);
    if (xErr) {
      return xErr;
    }
    knownX = rawX as number[];
  } else {
    knownX = knownY.map((_, i) => i + 1);
  }
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return ERRORS.VALUE;
  }
  let sumX = 0,
    sumY = 0,
    sumXY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    sumX += knownX[i];
    sumY += knownY[i];
    sumXY += knownX[i] * knownY[i];
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return ERRORS.DIV0;
  }
  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  return rvArray([[rvNumber(slope), rvNumber(intercept)]]);
};

export const fnLOGEST: NF = args => {
  if (!isArrayArg(args[0])) {
    return ERRORS.VALUE;
  }
  const rawY = flattenNumbers([args[0]]);
  const yErr = firstError(rawY);
  if (yErr) {
    return yErr;
  }
  const knownY = rawY as number[];
  let knownX: number[];
  if (args.length > 1 && isArrayArg(args[1])) {
    const rawX = flattenNumbers([args[1]]);
    const xErr = firstError(rawX);
    if (xErr) {
      return xErr;
    }
    knownX = rawX as number[];
  } else {
    knownX = knownY.map((_, i) => i + 1);
  }
  const n = Math.min(knownX.length, knownY.length);
  if (n < 1) {
    return ERRORS.VALUE;
  }
  let sumX = 0,
    sumLnY = 0,
    sumXLnY = 0,
    sumX2 = 0;
  for (let i = 0; i < n; i++) {
    if (knownY[i] <= 0) {
      return ERRORS.NUM;
    }
    sumX += knownX[i];
    sumLnY += Math.log(knownY[i]);
    sumXLnY += knownX[i] * Math.log(knownY[i]);
    sumX2 += knownX[i] * knownX[i];
  }
  const denom = n * sumX2 - sumX * sumX;
  if (denom === 0) {
    return ERRORS.DIV0;
  }
  const lnM = (n * sumXLnY - sumX * sumLnY) / denom;
  const lnB = (sumLnY - lnM * sumX) / n;
  return rvArray([[rvNumber(Math.exp(lnM)), rvNumber(Math.exp(lnB))]]);
};

// ============================================================================
// F.DIST.RT, F.INV.RT — F-distribution right tail
// ============================================================================

/**
 * F.DIST.RT(x, d1, d2) — right-tail probability of the F-distribution.
 *
 * Equivalent to `1 - F.DIST(x, d1, d2, TRUE)`. Using the symmetry of the
 * regularized incomplete beta function, this can be expressed as
 * `I(d2/(d2 + d1*x), d2/2, d1/2)`, which avoids subtracting from 1 and
 * is numerically stable in the upper tail.
 */
export const fnF_DIST_RT: NF = args => {
  const x = numArg(args, 0);
  if (x.kind === RVKind.Error) {
    return x;
  }
  const df1 = numArg(args, 1);
  if (df1.kind === RVKind.Error) {
    return df1;
  }
  const df2 = numArg(args, 2);
  if (df2.kind === RVKind.Error) {
    return df2;
  }
  if (x.value <= 0 || df1.value < 1 || df2.value < 1) {
    return ERRORS.NUM;
  }
  const d1 = Math.floor(df1.value);
  const d2 = Math.floor(df2.value);
  // Right tail via symmetry: I(d2/(d2 + d1*x), d2/2, d1/2).
  return rvNumber(betaIncomplete(d2 / (d2 + d1 * x.value), d2 / 2, d1 / 2));
};

/**
 * F.INV.RT(p, d1, d2) — inverse right-tail of the F-distribution.
 * Returns x such that P(F > x) = p. Implemented via binary search on the
 * right-tail CDF (monotonically decreasing from 1 at x=0 to 0 at x=∞).
 */
export const fnF_INV_RT: NF = args => {
  const p = numArg(args, 0);
  if (p.kind === RVKind.Error) {
    return p;
  }
  const df1 = numArg(args, 1);
  if (df1.kind === RVKind.Error) {
    return df1;
  }
  const df2 = numArg(args, 2);
  if (df2.kind === RVKind.Error) {
    return df2;
  }
  if (p.value <= 0 || p.value > 1 || df1.value < 1 || df2.value < 1) {
    return ERRORS.NUM;
  }
  const d1 = Math.floor(df1.value);
  const d2 = Math.floor(df2.value);
  // Right-tail CDF at x: I(d2/(d2 + d1*x), d2/2, d1/2), decreases with x.
  let lo = 0;
  let hi = 1e6;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const rt = betaIncomplete(d2 / (d2 + d1 * mid), d2 / 2, d1 / 2);
    if (rt > p.value) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo < 1e-10) {
      break;
    }
  }
  return rvNumber((lo + hi) / 2);
};

// ============================================================================
// SKEW, SKEW.P, KURT
// ============================================================================

/**
 * SKEW — sample skewness.
 * Formula: n / ((n-1)(n-2)) * Σ((xi-mean)/s)^3, where s is the sample stdev.
 */
export const fnSKEW: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const xs = nums as number[];
  const n = xs.length;
  if (n < 3) {
    return ERRORS.DIV0;
  }
  let sum = 0;
  for (const v of xs) {
    sum += v;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of xs) {
    sumSq += (v - mean) ** 2;
  }
  const sampleStd = Math.sqrt(sumSq / (n - 1));
  if (sampleStd === 0) {
    return ERRORS.DIV0;
  }
  let sumCubed = 0;
  for (const v of xs) {
    sumCubed += ((v - mean) / sampleStd) ** 3;
  }
  return rvNumber((n / ((n - 1) * (n - 2))) * sumCubed);
};

/**
 * SKEW.P — population skewness.
 * Formula: (1/n) * Σ((xi-mean)/σ)^3, where σ is the population stdev.
 */
export const fnSKEW_P: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const xs = nums as number[];
  const n = xs.length;
  if (n < 1) {
    return ERRORS.DIV0;
  }
  let sum = 0;
  for (const v of xs) {
    sum += v;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of xs) {
    sumSq += (v - mean) ** 2;
  }
  const popStd = Math.sqrt(sumSq / n);
  if (popStd === 0) {
    return ERRORS.DIV0;
  }
  let sumCubed = 0;
  for (const v of xs) {
    sumCubed += ((v - mean) / popStd) ** 3;
  }
  return rvNumber(sumCubed / n);
};

/**
 * KURT — sample excess kurtosis.
 * Formula: n(n+1) / ((n-1)(n-2)(n-3)) * Σ((xi-mean)/s)^4 - 3(n-1)^2 / ((n-2)(n-3)).
 */
export const fnKURT: NF = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  const xs = nums as number[];
  const n = xs.length;
  if (n < 4) {
    return ERRORS.DIV0;
  }
  let sum = 0;
  for (const v of xs) {
    sum += v;
  }
  const mean = sum / n;
  let sumSq = 0;
  for (const v of xs) {
    sumSq += (v - mean) ** 2;
  }
  const sampleStd = Math.sqrt(sumSq / (n - 1));
  if (sampleStd === 0) {
    return ERRORS.DIV0;
  }
  let sumQuad = 0;
  for (const v of xs) {
    sumQuad += ((v - mean) / sampleStd) ** 4;
  }
  const term1 = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3));
  const term2 = (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  return rvNumber(term1 * sumQuad - term2);
};
