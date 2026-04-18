/**
 * Math / Aggregate Functions — Native RuntimeValue implementation.
 */

import type {
  RuntimeValue,
  NumberValue,
  ArrayValue,
  ErrorValue,
  ScalarValue
} from "../runtime/values";
import {
  RVKind,
  ERRORS,
  rvNumber,
  rvString,
  rvArray,
  toNumberRV,
  toStringRV,
  topLeft,
  isError,
  isArray
} from "../runtime/values";
import { argToNumber, flattenAll, flattenNumbers, firstError } from "./_shared";

// ============================================================================
// Native Function Type
// ============================================================================

export type NativeFn = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Internal Helpers
// ============================================================================

/**
 * Apply a rounding operation using an integer factor, avoiding the
 * floating-point drift that `Math.pow(10, d) / Math.pow(10, d)` introduces
 * when `d` is negative (e.g. `1/0.1` is not exactly `10`).
 *
 * For `digits >= 0` we scale up by `10^digits`, round, and scale back down.
 * For `digits < 0` we divide by `10^|digits|`, round, and multiply back up.
 * In both branches the scaling factor is an integer, so only one rounding
 * error is introduced.
 */
function applyRounding(value: number, digits: number, round: (n: number) => number): number {
  const d = Math.trunc(digits);
  // Clamp digits to Excel's documented range [-127, 127]. Without this
  // guard, extreme inputs blow `Math.pow(10, d)` up to Infinity; the
  // resulting `round(value * Infinity) / Infinity` is NaN, which then
  // leaks through the whole rounding family (ROUND / ROUNDDOWN /
  // ROUNDUP / TRUNC / MROUND) and persists to the worksheet. Return
  // `value` itself for well-behaved "nothing to do" edge cases (d very
  // large is effectively "no rounding"; d very negative rounds to 0).
  if (!Number.isFinite(value)) {
    return value;
  }
  if (d >= 308) {
    return value;
  }
  if (d <= -308) {
    return round(0);
  }
  if (d >= 0) {
    const factor = Math.pow(10, d);
    if (!Number.isFinite(factor) || factor === 0) {
      return value;
    }
    return round(value * factor) / factor;
  }
  const factor = Math.pow(10, -d);
  if (!Number.isFinite(factor) || factor === 0) {
    return value;
  }
  return round(value / factor) * factor;
}

/**
 * Round half away from zero, matching Excel's ROUND semantics.
 *
 * JavaScript's `Math.round` rounds half toward +∞ (so `Math.round(-0.5)`
 * returns `-0` instead of `-1`), whereas Excel rounds the magnitude and
 * preserves the sign: `ROUND(-0.5, 0) = -1`. This helper implements the
 * Excel rule by operating on the absolute value and restoring the sign.
 */
function roundHalfAwayFromZero(n: number): number {
  return n < 0 ? -Math.round(-n) : Math.round(n);
}

// ============================================================================
// Trigonometric Functions
// ============================================================================

export const fnSIN: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.sin(n.value));
};

export const fnCOS: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.cos(n.value));
};

export const fnTAN: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.tan(n.value));
};

export const fnASIN: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < -1 || n.value > 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.asin(n.value));
};

export const fnACOS: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < -1 || n.value > 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.acos(n.value));
};

export const fnATAN: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.atan(n.value));
};

export const fnATAN2: NativeFn = args => {
  const x = argToNumber(args[0]);
  if (isError(x)) {
    return x;
  }
  const y = argToNumber(args[1]);
  if (isError(y)) {
    return y;
  }
  if (x.value === 0 && y.value === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.atan2(y.value, x.value));
};

export const fnSINH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // sinh(x) overflows double for |x| > ~710. Excel returns #NUM! rather
  // than a silent Infinity (which would poison downstream arithmetic).
  const r = Math.sinh(n.value);
  return isFinite(r) ? rvNumber(r) : ERRORS.NUM;
};

export const fnCOSH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const r = Math.cosh(n.value);
  return isFinite(r) ? rvNumber(r) : ERRORS.NUM;
};

export const fnTANH: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.tanh(n.value));
};

export const fnASINH: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.asinh(n.value));
};

export const fnACOSH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.acosh(n.value));
};

export const fnATANH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= -1 || n.value >= 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.atanh(n.value));
};

/**
 * Secondary trigonometric family (SEC / CSC / COT and hyperbolic /
 * inverse variants). None of these exist on the JavaScript Math object
 * so we derive them from the standard sin / cos / tan primitives, with
 * explicit guards at the discontinuities (π/2 for SEC, multiples of π
 * for CSC / COT, zero for the H variants).
 */
export const fnSEC: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const c = Math.cos(n.value);
  return c === 0 ? ERRORS.DIV0 : rvNumber(1 / c);
};

export const fnCSC: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const s = Math.sin(n.value);
  return s === 0 ? ERRORS.DIV0 : rvNumber(1 / s);
};

export const fnCOT: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const s = Math.sin(n.value);
  if (s === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.cos(n.value) / s);
};

export const fnSECH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // sech(x) = 1/cosh(x); cosh is never zero for real x, but it does
  // overflow to Infinity for |x| > ~710 — return 0 (the mathematical
  // limit) rather than dividing and producing NaN.
  const c = Math.cosh(n.value);
  return Number.isFinite(c) ? rvNumber(1 / c) : rvNumber(0);
};

export const fnCSCH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value === 0) {
    return ERRORS.DIV0;
  }
  const s = Math.sinh(n.value);
  if (!Number.isFinite(s)) {
    return rvNumber(0);
  }
  return rvNumber(1 / s);
};

export const fnCOTH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(1 / Math.tanh(n.value));
};

export const fnACOT: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // Excel's ACOT returns values in (0, π), not (-π/2, π/2) like
  // Math.atan. Map: ACOT(x) = π/2 − atan(x).
  return rvNumber(Math.PI / 2 - Math.atan(n.value));
};

export const fnACOTH: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // Defined only for |x| > 1.
  if (n.value >= -1 && n.value <= 1) {
    return ERRORS.NUM;
  }
  return rvNumber(0.5 * Math.log((n.value + 1) / (n.value - 1)));
};

// ============================================================================
// Math / Aggregate Functions
// ============================================================================

export const fnSUM: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as NumberValue).value;
  }
  // Fail fast on overflow to Infinity; otherwise the result leaks into
  // any formula that aggregates it (AVERAGE, STDEV, etc.) and those
  // downstream callers would then fan #NUM! out across the graph.
  return Number.isFinite(sum) ? rvNumber(sum) : ERRORS.NUM;
};

export const fnAVERAGE: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return ERRORS.DIV0;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as NumberValue).value;
  }
  const avg = sum / nums.length;
  return Number.isFinite(avg) ? rvNumber(avg) : ERRORS.NUM;
};

export const fnMIN: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let min = Infinity;
  for (const n of nums) {
    if ((n as NumberValue).value < min) {
      min = (n as NumberValue).value;
    }
  }
  return rvNumber(min);
};

export const fnMAX: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let max = -Infinity;
  for (const n of nums) {
    if ((n as NumberValue).value > max) {
      max = (n as NumberValue).value;
    }
  }
  return rvNumber(max);
};

export const fnCOUNT: NativeFn = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v.kind === RVKind.Number) {
      count++;
    }
  }
  return rvNumber(count);
};

export const fnCOUNTA: NativeFn = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    // Count everything that is not blank and not empty string
    if (v.kind !== RVKind.Blank && !(v.kind === RVKind.String && v.value === "")) {
      count++;
    }
  }
  return rvNumber(count);
};

export const fnCOUNTBLANK: NativeFn = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v.kind === RVKind.Blank || (v.kind === RVKind.String && v.value === "")) {
      count++;
    }
  }
  return rvNumber(count);
};

export const fnPRODUCT: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let product = 1;
  for (const n of nums) {
    product *= (n as NumberValue).value;
  }
  // Excel surfaces an overflow as #NUM! rather than letting Infinity
  // propagate into subsequent arithmetic.
  return isFinite(product) ? rvNumber(product) : ERRORS.NUM;
};

export const fnSUMPRODUCT: NativeFn = args => {
  if (args.length === 0) {
    return ERRORS.VALUE;
  }
  // Promote scalar args to 1x1 arrays. Excel allows a scalar (or 1x1 array)
  // to broadcast to the surrounding array's shape, e.g. SUMPRODUCT(1, A1:A10)
  // is equivalent to SUMPRODUCT(A1:A10). All non-broadcast arrays must share
  // the same height/width — only 1x1 may be broadcast.
  const arrays: ArrayValue[] = [];
  for (const a of args) {
    if (isArray(a)) {
      arrays.push(a);
    } else {
      arrays.push(rvArray([[topLeft(a)]]));
    }
  }
  // Find the target dimensions: the max height/width across all non-1x1 arrays.
  // 1x1 arrays are eligible to broadcast to whatever size is chosen.
  let rows = 1;
  let cols = 1;
  for (const arr of arrays) {
    if (arr.height !== 1 || arr.width !== 1) {
      if (rows === 1 && cols === 1) {
        rows = arr.height;
        cols = arr.width;
      } else if (arr.height !== rows || arr.width !== cols) {
        // Two different non-1x1 shapes — incompatible.
        return ERRORS.VALUE;
      }
    }
  }
  let sum = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let product = 1;
      for (const arr of arrays) {
        // Broadcast 1x1 arrays to the target cell position.
        const val = arr.height === 1 && arr.width === 1 ? arr.rows[0][0] : arr.rows[r][c];
        if (val.kind === RVKind.Error) {
          return val;
        }
        const n =
          val.kind === RVKind.Number
            ? val.value
            : val.kind === RVKind.Boolean
              ? val.value
                ? 1
                : 0
              : 0;
        product *= n;
      }
      sum += product;
    }
  }
  if (!Number.isFinite(sum)) {
    return ERRORS.NUM;
  }
  return rvNumber(sum);
};

// ============================================================================
// Math Functions
// ============================================================================

export const fnABS: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.abs(n.value));
};

export const fnCEILING: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const sigRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(1);
  if (isError(sigRV)) {
    return sigRV;
  }
  const sig = sigRV.value;
  if (sig === 0) {
    return rvNumber(0);
  }
  // CEILING (the non-MATH variant) requires number and significance to share
  // a sign. `CEILING(2, -1)` in Excel is #NUM!; only `CEILING.MATH` tolerates
  // a negative significance with a positive number.
  if (num.value !== 0 && Math.sign(num.value) !== Math.sign(sig)) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.ceil(num.value / sig) * sig);
};

export const fnFLOOR: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const sigRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(1);
  if (isError(sigRV)) {
    return sigRV;
  }
  const sig = sigRV.value;
  if (sig === 0) {
    return ERRORS.DIV0;
  }
  if (num.value !== 0 && Math.sign(num.value) !== Math.sign(sig)) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.floor(num.value / sig) * sig);
};

export const fnINT: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.floor(n.value));
};

export const fnMOD: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const div = argToNumber(args[1]);
  if (isError(div)) {
    return div;
  }
  if (div.value === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(num.value - div.value * Math.floor(num.value / div.value));
};

export const fnPOWER: NativeFn = args => {
  const base = argToNumber(args[0]);
  if (isError(base)) {
    return base;
  }
  const exp = argToNumber(args[1]);
  if (isError(exp)) {
    return exp;
  }
  // Distinguish the two degenerate-base cases Excel handles separately:
  //   POWER(0, 0)       → 1 (by convention, matches Excel)
  //   POWER(0, <0)      → #DIV/0!
  //   POWER(<0, non-int) → #NUM! (complex result; Math.pow returns NaN)
  if (base.value === 0) {
    if (exp.value < 0) {
      return ERRORS.DIV0;
    }
    if (exp.value === 0) {
      return rvNumber(1);
    }
  }
  const result = Math.pow(base.value, exp.value);
  if (Number.isNaN(result)) {
    return ERRORS.NUM;
  }
  return !isFinite(result) ? ERRORS.NUM : rvNumber(result);
};

export const fnROUND: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  return rvNumber(applyRounding(num.value, digitsRV.value, roundHalfAwayFromZero));
};

export const fnROUNDDOWN: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  return rvNumber(applyRounding(num.value, digitsRV.value, Math.trunc));
};

export const fnROUNDUP: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  // ROUNDUP rounds away from zero for the fractional part at the requested
  // precision. We emulate this by ceiling the scaled absolute value and
  // restoring the sign.
  const sign = num.value >= 0 ? 1 : -1;
  const rounded = applyRounding(Math.abs(num.value), digitsRV.value, Math.ceil);
  return rvNumber(sign * rounded);
};

export const fnSQRT: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.sqrt(n.value));
};

export const fnLN: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log(n.value));
};

export const fnLOG: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  const baseRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(10);
  if (isError(baseRV)) {
    return baseRV;
  }
  if (baseRV.value <= 0 || baseRV.value === 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log(n.value) / Math.log(baseRV.value));
};

export const fnLOG10: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log10(n.value));
};

export const fnEXP: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // EXP(~710) overflows double to Infinity. Excel returns #NUM! in that
  // regime rather than letting the non-finite result propagate.
  const r = Math.exp(n.value);
  return isFinite(r) ? rvNumber(r) : ERRORS.NUM;
};

export const fnPI: NativeFn = () => rvNumber(Math.PI);

export const fnRAND: NativeFn = () => rvNumber(Math.random());

export const fnRANDBETWEEN: NativeFn = args => {
  const bottom = argToNumber(args[0]);
  if (isError(bottom)) {
    return bottom;
  }
  const top = argToNumber(args[1]);
  if (isError(top)) {
    return top;
  }
  const lo = Math.ceil(bottom.value);
  const hi = Math.floor(top.value);
  // Excel returns #NUM! when bottom > top; otherwise the formula below would
  // produce a garbage integer from a negative range.
  if (lo > hi) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.floor(Math.random() * (hi - lo + 1)) + lo);
};

export const fnSIGN: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  // `Math.sign(-0)` returns `-0` (preserving IEEE-754 sign bit). Excel's
  // SIGN normalises zero to +0 (since the documented result for a zero
  // input is 0, not distinguishing signed zero). The `|| 0` collapses
  // both ±0 into `+0` while leaving ±1 untouched.
  return rvNumber(Math.sign(n.value) || 0);
};

// ============================================================================
// Additional Math Functions
// ============================================================================

export const fnTRUNC: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  return rvNumber(applyRounding(num.value, digitsRV.value, Math.trunc));
};

export const fnSUMSQ: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as NumberValue).value ** 2;
  }
  return isFinite(sum) ? rvNumber(sum) : ERRORS.NUM;
};

export const fnGCD: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  // Excel rejects any negative argument with #NUM!; truncate toward zero
  // for non-integer positives (previous `Math.floor(-5.5) = -6` then
  // `Math.abs` produced 6 instead of #NUM!).
  const coerce = (v: number): number | ErrorValue => {
    if (v < 0) {
      return ERRORS.NUM;
    }
    return Math.trunc(v);
  };
  const first = coerce((nums[0] as NumberValue).value);
  if (typeof first !== "number") {
    return first;
  }
  let result = first;
  for (let i = 1; i < nums.length; i++) {
    const bi = coerce((nums[i] as NumberValue).value);
    if (typeof bi !== "number") {
      return bi;
    }
    let b = bi;
    while (b) {
      const t = b;
      b = result % b;
      result = t;
    }
  }
  return rvNumber(result);
};

export const fnLCM: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  const coerce = (v: number): number | ErrorValue => {
    if (v < 0) {
      return ERRORS.NUM;
    }
    return Math.trunc(v);
  };
  const first = coerce((nums[0] as NumberValue).value);
  if (typeof first !== "number") {
    return first;
  }
  let result = first;
  for (let i = 1; i < nums.length; i++) {
    const bi = coerce((nums[i] as NumberValue).value);
    if (typeof bi !== "number") {
      return bi;
    }
    const b = bi;
    if (result === 0 && b === 0) {
      result = 0;
    } else {
      let g = result;
      let t = b;
      while (t) {
        const tmp = t;
        t = g % t;
        g = tmp;
      }
      result = (result * b) / g;
    }
  }
  return rvNumber(result);
};

// ============================================================================
// More Math Functions
// ============================================================================

export const fnEVEN: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const sign = n.value >= 0 ? 1 : -1;
  const abs = Math.abs(n.value);
  const ceil = Math.ceil(abs);
  return rvNumber(sign * (ceil % 2 === 0 ? ceil : ceil + 1));
};

export const fnODD: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value === 0) {
    return rvNumber(1);
  }
  const sign = n.value >= 0 ? 1 : -1;
  const abs = Math.abs(n.value);
  const ceil = Math.ceil(abs);
  return rvNumber(sign * (ceil % 2 === 1 ? ceil : ceil + 1));
};

export const fnMROUND: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const multiple = argToNumber(args[1]);
  if (isError(multiple)) {
    return multiple;
  }
  if (multiple.value === 0) {
    return rvNumber(0);
  }
  if ((num.value > 0 && multiple.value < 0) || (num.value < 0 && multiple.value > 0)) {
    return ERRORS.NUM;
  }
  return rvNumber(roundHalfAwayFromZero(num.value / multiple.value) * multiple.value);
};

export const fnQUOTIENT: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const den = argToNumber(args[1]);
  if (isError(den)) {
    return den;
  }
  if (den.value === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.trunc(num.value / den.value));
};

export const fnBASE: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const radix = argToNumber(args[1]);
  if (isError(radix)) {
    return radix;
  }
  if (radix.value < 2 || radix.value > 36) {
    return ERRORS.NUM;
  }
  const minLenRV = args.length > 2 ? argToNumber(args[2]) : rvNumber(0);
  if (isError(minLenRV)) {
    return minLenRV;
  }
  const result = Math.floor(num.value).toString(Math.floor(radix.value)).toUpperCase();
  return rvString(minLenRV.value > 0 ? result.padStart(minLenRV.value, "0") : result);
};

export const fnDECIMAL: NativeFn = args => {
  const e = topLeft(args[0]);
  if (e.kind === RVKind.Error) {
    return e;
  }
  const text = toStringRV(e);
  const radix = argToNumber(args[1]);
  if (isError(radix)) {
    return radix;
  }
  if (radix.value < 2 || radix.value > 36) {
    return ERRORS.NUM;
  }
  // `parseInt("1G", 16)` returns 1 because JavaScript silently stops at
  // the first invalid digit. Excel's DECIMAL requires every character in
  // the input to be a valid digit for the given radix, so we validate
  // strictly before delegating.
  const base = Math.floor(radix.value);
  const trimmed = text.trim();
  if (trimmed === "") {
    return ERRORS.NUM;
  }
  // Digit alphabet up to base 36: 0-9, A-Z (case-insensitive).
  const match = /^[+-]?([0-9A-Za-z]+)$/.exec(trimmed);
  if (!match) {
    return ERRORS.NUM;
  }
  const digits = match[1].toUpperCase();
  for (const ch of digits) {
    const d = ch >= "0" && ch <= "9" ? ch.charCodeAt(0) - 48 : ch.charCodeAt(0) - 55;
    if (d < 0 || d >= base) {
      return ERRORS.NUM;
    }
  }
  const result = parseInt(trimmed, base);
  if (isNaN(result)) {
    return ERRORS.NUM;
  }
  return rvNumber(result);
};

export const fnROMAN: NativeFn = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  let n = Math.floor(num.value);
  if (n < 0 || n > 3999) {
    return ERRORS.VALUE;
  }
  if (n === 0) {
    return rvString("");
  }
  // The optional `form` argument (0=classic through 4=simplified) controls
  // how far Excel will collapse repeated characters into subtractive pairs.
  // Higher forms introduce additional patterns beyond the classic IV/IX/
  // XL/XC/CD/CM pairs — e.g. form 1 allows LM for 950, form 4 uses the
  // maximally short forms. A value of TRUE (1) / FALSE (0) maps to form 0
  // / 4 the way Excel does.
  let form = 0;
  if (args.length > 1 && args[1].kind !== RVKind.Blank) {
    const f = argToNumber(args[1]);
    if (isError(f)) {
      return f;
    }
    if (f.value === 1 && (f as NumberValue).value === 1) {
      // Boolean inputs flow through argToNumber as 0/1 already.
    }
    form = Math.floor(f.value);
    if (form < 0 || form > 4) {
      return ERRORS.VALUE;
    }
  }
  // Classic table (form 0) — subtractive pairs IV, IX, XL, XC, CD, CM.
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  // Forms 1-4 progressively introduce aggressive subtractive pairs
  // (LM=950, VC=95, IC=99, etc.). Each extra must only trigger at its
  // `minForm` or higher — and must NOT duplicate a value the classic
  // table already covers, or we'd emit the same subtractive pair twice
  // per occurrence (e.g. 1999 would become "CMCMCXCIX" instead of
  // "MCMXCIX").
  const extraByForm: { value: number; sym: string; minForm: number }[] = [
    { value: 995, sym: "VM", minForm: 4 },
    { value: 990, sym: "XM", minForm: 3 },
    { value: 950, sym: "LM", minForm: 1 },
    { value: 495, sym: "VD", minForm: 4 },
    { value: 490, sym: "XD", minForm: 3 },
    { value: 450, sym: "LD", minForm: 1 },
    { value: 99, sym: "IC", minForm: 2 },
    { value: 95, sym: "VC", minForm: 1 },
    { value: 49, sym: "IL", minForm: 2 },
    { value: 45, sym: "VL", minForm: 1 }
  ];
  let result = "";
  // Merge: walk from largest remainder, trying any applicable extra then
  // the classic table entry.
  while (n > 0) {
    let matched = false;
    for (const ex of extraByForm) {
      if (form >= ex.minForm && n >= ex.value) {
        result += ex.sym;
        n -= ex.value;
        matched = true;
        break;
      }
    }
    if (matched) {
      continue;
    }
    for (let i = 0; i < vals.length; i++) {
      if (n >= vals[i]) {
        result += syms[i];
        n -= vals[i];
        matched = true;
        break;
      }
    }
    if (!matched) {
      // Safety net — should not happen since the classic table covers 1.
      break;
    }
  }
  return rvString(result);
};

export const fnARABIC: NativeFn = args => {
  const s = topLeft(args[0]);
  if (s.kind === RVKind.Error) {
    return s;
  }
  const text = toStringRV(s).toUpperCase().trim();
  if (text === "") {
    return rvNumber(0);
  }
  const romanMap: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let result = 0;
  for (let i = 0; i < text.length; i++) {
    const current = romanMap[text[i]];
    const next = romanMap[text[i + 1]];
    if (current === undefined) {
      return ERRORS.VALUE;
    }
    if (next && current < next) {
      result -= current;
    } else {
      result += current;
    }
  }
  return rvNumber(result);
};

export const fnDEGREES: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber((n.value * 180) / Math.PI);
};

export const fnRADIANS: NativeFn = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber((n.value * Math.PI) / 180);
};

/**
 * Shared driver for SUMX2MY2 / SUMX2PY2 / SUMXMY2. Walks two arrays in
 * lock-step and reduces `combine(x, y)` across matching numeric cells.
 * Non-numeric cells are skipped (Excel behaviour) and errors propagate.
 */
function sumPairedArrays(
  args: RuntimeValue[],
  combine: (x: number, y: number) => number
): RuntimeValue {
  const a0 = args[0];
  const a1 = args[1];
  if (a0.kind !== RVKind.Array || a1.kind !== RVKind.Array) {
    return ERRORS.VALUE;
  }
  const h = Math.min(a0.height, a1.height);
  const w = Math.min(a0.width, a1.width);
  let sum = 0;
  for (let r = 0; r < h; r++) {
    for (let c = 0; c < w; c++) {
      const x = a0.rows[r][c];
      const y = a1.rows[r][c];
      if (x.kind === RVKind.Error) {
        return x;
      }
      if (y.kind === RVKind.Error) {
        return y;
      }
      if (x.kind !== RVKind.Number || y.kind !== RVKind.Number) {
        continue;
      }
      sum += combine(x.value, y.value);
    }
  }
  return rvNumber(sum);
}

export const fnSUMX2MY2: NativeFn = args => sumPairedArrays(args, (x, y) => x * x - y * y);

export const fnSUMX2PY2: NativeFn = args => sumPairedArrays(args, (x, y) => x * x + y * y);

export const fnSUMXMY2: NativeFn = args => sumPairedArrays(args, (x, y) => (x - y) ** 2);

export const fnMULTINOMIAL: NativeFn = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  // Work in log space so the numerator (sum! ≈ up to ~1e307 around sum = 170)
  // doesn't overflow before we divide. Summing logs avoids the NaN case
  // `Infinity / Infinity` that the direct product formulation would hit
  // for large inputs.
  let sum = 0;
  let lnDenom = 0;
  for (const n of nums) {
    const ni = Math.floor((n as NumberValue).value);
    if (ni < 0) {
      return ERRORS.NUM;
    }
    sum += ni;
    for (let i = 2; i <= ni; i++) {
      lnDenom += Math.log(i);
    }
  }
  let lnNumer = 0;
  for (let i = 2; i <= sum; i++) {
    lnNumer += Math.log(i);
  }
  const result = Math.exp(lnNumer - lnDenom);
  if (!isFinite(result)) {
    return ERRORS.NUM;
  }
  // Round to the nearest integer — the log-exp round-trip introduces sub-ulp
  // noise that would otherwise leave us with things like `20.0000000001`.
  return rvNumber(Math.round(result));
};

export const fnFACT: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const num = Math.floor(n.value);
  if (num < 0) {
    return ERRORS.NUM;
  }
  if (num > 170) {
    return ERRORS.NUM;
  }
  let result = 1;
  for (let i = 2; i <= num; i++) {
    result *= i;
  }
  return rvNumber(result);
};

export const fnFACTDOUBLE: NativeFn = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const num = Math.floor(n.value);
  if (num < -1) {
    return ERRORS.NUM;
  }
  if (num <= 0) {
    return rvNumber(1);
  }
  let result = 1;
  for (let i = num; i > 0; i -= 2) {
    result *= i;
    if (!isFinite(result)) {
      return ERRORS.NUM;
    }
  }
  return rvNumber(result);
};

export const fnCOMBIN: NativeFn = args => {
  const nRV = argToNumber(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const kRV = argToNumber(args[1]);
  if (isError(kRV)) {
    return kRV;
  }
  const ni = Math.floor(nRV.value);
  const ki = Math.floor(kRV.value);
  if (ni < 0 || ki < 0 || ki > ni) {
    return ERRORS.NUM;
  }
  // Use the `C(n, k) = C(n, k-1) * (n-k+1)/k` recurrence and pick the
  // smaller of k / (n-k) so the loop runs at most n/2 iterations. The
  // interleaved divide keeps the running value bounded and preserves
  // near-full double precision until the final product overflows
  // magnitude ~1e308 (at which point we surface #NUM!).
  const kEff = Math.min(ki, ni - ki);
  let result = 1;
  for (let i = 0; i < kEff; i++) {
    result = (result * (ni - i)) / (i + 1);
    if (!isFinite(result)) {
      return ERRORS.NUM;
    }
  }
  // Excel returns an integer for COMBIN when the value fits in a double
  // exactly (≤ 2^53); beyond that (e.g. COMBIN(100, 50) ≈ 1e29) the
  // result is fundamentally approximate, so rounding only makes sense
  // below 2^53.
  return rvNumber(result < 9.007199254740992e15 ? Math.round(result) : result);
};

export const fnCOMBINA: NativeFn = args => {
  const nRV = argToNumber(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const kRV = argToNumber(args[1]);
  if (isError(kRV)) {
    return kRV;
  }
  // Special case: Excel's COMBINA(0, 0) is 1 even though the delegated
  // `COMBIN(-1, 0)` would flag #NUM! under our generic validation.
  if (nRV.value === 0 && kRV.value === 0) {
    return rvNumber(1);
  }
  return fnCOMBIN([rvNumber(nRV.value + kRV.value - 1), kRV]);
};

export const fnPERMUT: NativeFn = args => {
  const nRV = argToNumber(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const kRV = argToNumber(args[1]);
  if (isError(kRV)) {
    return kRV;
  }
  const ni = Math.floor(nRV.value);
  const ki = Math.floor(kRV.value);
  if (ni < 0 || ki < 0 || ki > ni) {
    return ERRORS.NUM;
  }
  let result = 1;
  for (let i = 0; i < ki; i++) {
    result *= ni - i;
  }
  return rvNumber(result);
};

// ============================================================================
// Matrix functions: MMULT, MDETERM, MINVERSE, MUNIT
// ============================================================================

/**
 * Extract a numeric matrix from an ArrayValue. Returns #VALUE! when any
 * cell is non-numeric or an error propagates.
 */
function asNumericMatrix(v: RuntimeValue): number[][] | ErrorValue {
  if (!isArray(v)) {
    return ERRORS.VALUE;
  }
  const arr = v as ArrayValue;
  const out: number[][] = [];
  for (const row of arr.rows) {
    const r: number[] = [];
    for (const cell of row) {
      if (cell.kind === RVKind.Error) {
        return cell;
      }
      if (cell.kind !== RVKind.Number) {
        return ERRORS.VALUE;
      }
      r.push(cell.value);
    }
    out.push(r);
  }
  return out;
}

/**
 * MMULT(array1, array2) — matrix product. Dimensions must be
 * (m×k) × (k×n) = (m×n); mismatched sizes return #VALUE!.
 */
export const fnMMULT: NativeFn = args => {
  const a = asNumericMatrix(args[0]);
  if ("kind" in a) {
    return a;
  }
  const b = asNumericMatrix(args[1]);
  if ("kind" in b) {
    return b;
  }
  const m = a.length;
  const k = a[0]?.length ?? 0;
  const k2 = b.length;
  const n = b[0]?.length ?? 0;
  if (m === 0 || k === 0 || n === 0 || k !== k2) {
    return ERRORS.VALUE;
  }
  const rows: ScalarValue[][] = [];
  for (let i = 0; i < m; i++) {
    const row: ScalarValue[] = [];
    for (let j = 0; j < n; j++) {
      let sum = 0;
      for (let p = 0; p < k; p++) {
        sum += a[i][p] * b[p][j];
      }
      row.push(rvNumber(sum));
    }
    rows.push(row);
  }
  return rvArray(rows);
};

/**
 * MDETERM(array) — determinant of a square matrix via Gaussian
 * elimination with partial pivoting. Non-square or non-numeric input
 * returns #VALUE!.
 */
export const fnMDETERM: NativeFn = args => {
  const mat = asNumericMatrix(args[0]);
  if ("kind" in mat) {
    return mat;
  }
  const n = mat.length;
  if (n === 0 || mat[0].length !== n) {
    return ERRORS.VALUE;
  }
  // Copy rows so we don't mutate the caller's array.
  const a = mat.map(r => r.slice());
  let det = 1;
  for (let i = 0; i < n; i++) {
    // Partial pivot: find row with largest |a[r][i]| for r in [i, n).
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(a[r][i]) > Math.abs(a[pivot][i])) {
        pivot = r;
      }
    }
    if (Math.abs(a[pivot][i]) < 1e-14) {
      return rvNumber(0);
    }
    if (pivot !== i) {
      const tmp = a[i];
      a[i] = a[pivot];
      a[pivot] = tmp;
      det = -det;
    }
    det *= a[i][i];
    for (let r = i + 1; r < n; r++) {
      const factor = a[r][i] / a[i][i];
      for (let c = i; c < n; c++) {
        a[r][c] -= factor * a[i][c];
      }
    }
  }
  return rvNumber(det);
};

/**
 * MINVERSE(array) — inverse of a square matrix via Gauss-Jordan
 * elimination. Singular matrices return #NUM!; non-square return
 * #VALUE!.
 */
export const fnMINVERSE: NativeFn = args => {
  const mat = asNumericMatrix(args[0]);
  if ("kind" in mat) {
    return mat;
  }
  const n = mat.length;
  if (n === 0 || mat[0].length !== n) {
    return ERRORS.VALUE;
  }
  // Build augmented matrix [A | I].
  const aug: number[][] = [];
  for (let i = 0; i < n; i++) {
    const row = new Array<number>(2 * n).fill(0);
    for (let j = 0; j < n; j++) {
      row[j] = mat[i][j];
    }
    row[n + i] = 1;
    aug.push(row);
  }
  // Gauss-Jordan elimination with partial pivoting.
  for (let i = 0; i < n; i++) {
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(aug[r][i]) > Math.abs(aug[pivot][i])) {
        pivot = r;
      }
    }
    if (Math.abs(aug[pivot][i]) < 1e-14) {
      return ERRORS.NUM; // singular
    }
    if (pivot !== i) {
      const tmp = aug[i];
      aug[i] = aug[pivot];
      aug[pivot] = tmp;
    }
    const diag = aug[i][i];
    for (let c = 0; c < 2 * n; c++) {
      aug[i][c] /= diag;
    }
    for (let r = 0; r < n; r++) {
      if (r === i) {
        continue;
      }
      const factor = aug[r][i];
      if (factor === 0) {
        continue;
      }
      for (let c = 0; c < 2 * n; c++) {
        aug[r][c] -= factor * aug[i][c];
      }
    }
  }
  // Extract inverse from right half.
  const rows: ScalarValue[][] = [];
  for (let i = 0; i < n; i++) {
    const row: ScalarValue[] = [];
    for (let j = 0; j < n; j++) {
      row.push(rvNumber(aug[i][n + j]));
    }
    rows.push(row);
  }
  return rvArray(rows);
};

/**
 * MUNIT(dimension) — n×n identity matrix.
 */
export const fnMUNIT: NativeFn = args => {
  const nV = toNumberRV(topLeft(args[0]));
  if (isError(nV)) {
    return nV;
  }
  const n = Math.trunc(nV.value);
  if (n < 1) {
    return ERRORS.VALUE;
  }
  const rows: ScalarValue[][] = [];
  for (let i = 0; i < n; i++) {
    const row: ScalarValue[] = new Array<ScalarValue>(n);
    for (let j = 0; j < n; j++) {
      row[j] = rvNumber(i === j ? 1 : 0);
    }
    rows.push(row);
  }
  return rvArray(rows);
};

// ============================================================================
// SERIESSUM
// ============================================================================

/**
 * SERIESSUM(x, n, m, coefficients) — returns the sum of a power series
 *   x^n * coef[0] + x^(n+m) * coef[1] + x^(n+2m) * coef[2] + …
 */
export const fnSERIESSUM: NativeFn = args => {
  const xV = toNumberRV(topLeft(args[0]));
  if (isError(xV)) {
    return xV;
  }
  const nV = toNumberRV(topLeft(args[1]));
  if (isError(nV)) {
    return nV;
  }
  const mV = toNumberRV(topLeft(args[2]));
  if (isError(mV)) {
    return mV;
  }
  const coeffs = flattenNumbers([args[3]]);
  const err = firstError(coeffs);
  if (err) {
    return err;
  }
  if (coeffs.length === 0) {
    return ERRORS.VALUE;
  }

  let total = 0;
  let exponent = nV.value;
  for (const c of coeffs) {
    total += (c as NumberValue).value * Math.pow(xV.value, exponent);
    exponent += mV.value;
  }
  if (!isFinite(total)) {
    return ERRORS.NUM;
  }
  return rvNumber(total);
};
