/**
 * Math / Aggregate Functions — Native RuntimeValue implementation.
 */

import type {
  RuntimeValue,
  NumberValue,
  ErrorValue,
  ScalarValue,
  ArrayValue
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

// ============================================================================
// Native Function Type
// ============================================================================

export type NativeFunction = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Internal Helpers
// ============================================================================

/**
 * Flatten all arguments into a list of numbers (or errors).
 *
 * For Array arguments: only numeric cells are included — booleans, strings,
 * and blanks are skipped (Excel behavior for range args).
 * For direct scalar arguments: booleans and numeric strings are coerced via toNumberRV.
 */
function flattenNumbers(args: RuntimeValue[]): (NumberValue | ErrorValue)[] {
  const result: (NumberValue | ErrorValue)[] = [];
  for (const arg of args) {
    if (arg.kind === RVKind.Array) {
      for (const row of arg.rows) {
        for (const cell of row) {
          if (cell.kind === RVKind.Error) {
            result.push(cell);
          } else if (cell.kind === RVKind.Number) {
            result.push(cell);
          }
          // Skip booleans, strings, blanks in array context (Excel behavior)
        }
      }
    } else {
      // Direct scalar argument — coerce
      if (arg.kind === RVKind.Error) {
        result.push(arg);
      } else if (arg.kind !== RVKind.Blank && arg.kind !== RVKind.MissingArg) {
        const n = toNumberRV(arg);
        result.push(n);
      }
      // Skip blanks for scalar args
    }
  }
  return result;
}

/**
 * Flatten all arguments into a list of scalar values (for COUNTA, COUNTBLANK, etc.).
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

/**
 * Find the first error in a list of NumberValue | ErrorValue.
 */
function firstError(values: (NumberValue | ErrorValue)[]): ErrorValue | null {
  for (const v of values) {
    if (v.kind === RVKind.Error) {
      return v;
    }
  }
  return null;
}

/**
 * Coerce a single argument to a number. For scalars, applies topLeft then toNumberRV.
 */
function argToNumber(arg: RuntimeValue): NumberValue | ErrorValue {
  const s = topLeft(arg);
  if (s.kind === RVKind.Error) {
    return s;
  }
  return toNumberRV(s);
}

// ============================================================================
// Trigonometric Functions
// ============================================================================

export const fnSIN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.sin(n.value));
};

export const fnCOS: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.cos(n.value));
};

export const fnTAN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.tan(n.value));
};

export const fnASIN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < -1 || n.value > 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.asin(n.value));
};

export const fnACOS: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < -1 || n.value > 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.acos(n.value));
};

export const fnATAN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.atan(n.value));
};

export const fnATAN2: NativeFunction = args => {
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

export const fnSINH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.sinh(n.value));
};

export const fnCOSH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.cosh(n.value));
};

export const fnTANH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.tanh(n.value));
};

export const fnASINH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.asinh(n.value));
};

export const fnACOSH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.acosh(n.value));
};

export const fnATANH: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= -1 || n.value >= 1) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.atanh(n.value));
};

// ============================================================================
// Math / Aggregate Functions
// ============================================================================

export const fnSUM: NativeFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as NumberValue).value;
  }
  return rvNumber(sum);
};

export const fnAVERAGE: NativeFunction = args => {
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
  return rvNumber(sum / nums.length);
};

export const fnMIN: NativeFunction = args => {
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

export const fnMAX: NativeFunction = args => {
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

export const fnCOUNT: NativeFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v.kind === RVKind.Number) {
      count++;
    }
  }
  return rvNumber(count);
};

export const fnCOUNTA: NativeFunction = args => {
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

export const fnCOUNTBLANK: NativeFunction = args => {
  let count = 0;
  const all = flattenAll(args);
  for (const v of all) {
    if (v.kind === RVKind.Blank || (v.kind === RVKind.String && v.value === "")) {
      count++;
    }
  }
  return rvNumber(count);
};

export const fnPRODUCT: NativeFunction = args => {
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
  return rvNumber(product);
};

export const fnSUMPRODUCT: NativeFunction = args => {
  if (args.length === 0) {
    return ERRORS.VALUE;
  }
  // Promote scalar args to 1x1 arrays (Excel behavior)
  const arrays: ArrayValue[] = [];
  for (const a of args) {
    if (isArray(a)) {
      arrays.push(a);
    } else {
      arrays.push(rvArray([[topLeft(a)]]));
    }
  }
  const rows = arrays[0].height;
  const cols = arrays[0].width;
  // Verify all same dimensions
  for (const arr of arrays) {
    if (arr.height !== rows || arr.width !== cols) {
      return ERRORS.VALUE;
    }
  }
  let sum = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let product = 1;
      for (const arr of arrays) {
        const val = arr.rows[r][c];
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
  return rvNumber(sum);
};

// ============================================================================
// Math Functions
// ============================================================================

export const fnABS: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.abs(n.value));
};

export const fnCEILING: NativeFunction = args => {
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
  return rvNumber(Math.ceil(num.value / sig) * sig);
};

export const fnFLOOR: NativeFunction = args => {
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
  return rvNumber(Math.floor(num.value / sig) * sig);
};

export const fnINT: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.floor(n.value));
};

export const fnMOD: NativeFunction = args => {
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

export const fnPOWER: NativeFunction = args => {
  const base = argToNumber(args[0]);
  if (isError(base)) {
    return base;
  }
  const exp = argToNumber(args[1]);
  if (isError(exp)) {
    return exp;
  }
  const result = Math.pow(base.value, exp.value);
  return !isFinite(result) ? ERRORS.NUM : rvNumber(result);
};

export const fnROUND: NativeFunction = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  const factor = Math.pow(10, digitsRV.value);
  return rvNumber(Math.round(num.value * factor) / factor);
};

export const fnROUNDDOWN: NativeFunction = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  const factor = Math.pow(10, digitsRV.value);
  return rvNumber(Math.trunc(num.value * factor) / factor);
};

export const fnROUNDUP: NativeFunction = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  const factor = Math.pow(10, digitsRV.value);
  const truncated = Math.trunc(num.value * factor);
  return rvNumber(
    (num.value * factor === truncated ? truncated : truncated + (num.value >= 0 ? 1 : -1)) / factor
  );
};

export const fnSQRT: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value < 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.sqrt(n.value));
};

export const fnLN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log(n.value));
};

export const fnLOG: NativeFunction = args => {
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

export const fnLOG10: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  if (n.value <= 0) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.log10(n.value));
};

export const fnEXP: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.exp(n.value));
};

export const fnPI: NativeFunction = () => rvNumber(Math.PI);

export const fnRAND: NativeFunction = () => rvNumber(Math.random());

export const fnRANDBETWEEN: NativeFunction = args => {
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
  return rvNumber(Math.floor(Math.random() * (hi - lo + 1)) + lo);
};

export const fnSIGN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber(Math.sign(n.value));
};

// ============================================================================
// Additional Math Functions
// ============================================================================

export const fnTRUNC: NativeFunction = args => {
  const num = argToNumber(args[0]);
  if (isError(num)) {
    return num;
  }
  const digitsRV = args.length > 1 ? argToNumber(args[1]) : rvNumber(0);
  if (isError(digitsRV)) {
    return digitsRV;
  }
  const factor = Math.pow(10, digitsRV.value);
  return rvNumber(Math.trunc(num.value * factor) / factor);
};

export const fnSUMSQ: NativeFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  for (const n of nums) {
    sum += (n as NumberValue).value ** 2;
  }
  return rvNumber(sum);
};

export const fnGCD: NativeFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let result = Math.abs(Math.floor((nums[0] as NumberValue).value));
  for (let i = 1; i < nums.length; i++) {
    let b = Math.abs(Math.floor((nums[i] as NumberValue).value));
    while (b) {
      const t = b;
      b = result % b;
      result = t;
    }
  }
  return rvNumber(result);
};

export const fnLCM: NativeFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  if (nums.length === 0) {
    return rvNumber(0);
  }
  let result = Math.abs(Math.floor((nums[0] as NumberValue).value));
  for (let i = 1; i < nums.length; i++) {
    const b = Math.abs(Math.floor((nums[i] as NumberValue).value));
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

export const fnEVEN: NativeFunction = args => {
  const n = argToNumber(args[0]);
  if (isError(n)) {
    return n;
  }
  const sign = n.value >= 0 ? 1 : -1;
  const abs = Math.abs(n.value);
  const ceil = Math.ceil(abs);
  return rvNumber(sign * (ceil % 2 === 0 ? ceil : ceil + 1));
};

export const fnODD: NativeFunction = args => {
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

export const fnMROUND: NativeFunction = args => {
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
  return rvNumber(Math.round(num.value / multiple.value) * multiple.value);
};

export const fnQUOTIENT: NativeFunction = args => {
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

export const fnBASE: NativeFunction = args => {
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

export const fnDECIMAL: NativeFunction = args => {
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
  const result = parseInt(text, Math.floor(radix.value));
  if (isNaN(result)) {
    return ERRORS.NUM;
  }
  return rvNumber(result);
};

export const fnROMAN: NativeFunction = args => {
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
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let result = "";
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) {
      result += syms[i];
      n -= vals[i];
    }
  }
  return rvString(result);
};

export const fnARABIC: NativeFunction = args => {
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

export const fnDEGREES: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber((n.value * 180) / Math.PI);
};

export const fnRADIANS: NativeFunction = args => {
  const n = argToNumber(args[0]);
  return isError(n) ? n : rvNumber((n.value * Math.PI) / 180);
};

export const fnSUMX2MY2: NativeFunction = args => {
  const a0 = args[0],
    a1 = args[1];
  if (a0.kind !== RVKind.Array || a1.kind !== RVKind.Array) {
    return ERRORS.VALUE;
  }
  let sum = 0;
  for (let r = 0; r < Math.min(a0.height, a1.height); r++) {
    for (let c = 0; c < Math.min(a0.width, a1.width); c++) {
      const x = a0.rows[r]?.[c];
      const y = a1.rows[r]?.[c];
      if (!x || !y) {
        continue;
      }
      if (x.kind === RVKind.Error) {
        return x;
      }
      if (y.kind === RVKind.Error) {
        return y;
      }
      if (x.kind !== RVKind.Number || y.kind !== RVKind.Number) {
        continue;
      }
      sum += x.value * x.value - y.value * y.value;
    }
  }
  return rvNumber(sum);
};

export const fnSUMX2PY2: NativeFunction = args => {
  const a0 = args[0],
    a1 = args[1];
  if (a0.kind !== RVKind.Array || a1.kind !== RVKind.Array) {
    return ERRORS.VALUE;
  }
  let sum = 0;
  for (let r = 0; r < Math.min(a0.height, a1.height); r++) {
    for (let c = 0; c < Math.min(a0.width, a1.width); c++) {
      const x = a0.rows[r]?.[c];
      const y = a1.rows[r]?.[c];
      if (!x || !y) {
        continue;
      }
      if (x.kind === RVKind.Error) {
        return x;
      }
      if (y.kind === RVKind.Error) {
        return y;
      }
      if (x.kind !== RVKind.Number || y.kind !== RVKind.Number) {
        continue;
      }
      sum += x.value * x.value + y.value * y.value;
    }
  }
  return rvNumber(sum);
};

export const fnSUMXMY2: NativeFunction = args => {
  const a0 = args[0],
    a1 = args[1];
  if (a0.kind !== RVKind.Array || a1.kind !== RVKind.Array) {
    return ERRORS.VALUE;
  }
  let sum = 0;
  for (let r = 0; r < Math.min(a0.height, a1.height); r++) {
    for (let c = 0; c < Math.min(a0.width, a1.width); c++) {
      const x = a0.rows[r]?.[c];
      const y = a1.rows[r]?.[c];
      if (!x || !y) {
        continue;
      }
      if (x.kind === RVKind.Error) {
        return x;
      }
      if (y.kind === RVKind.Error) {
        return y;
      }
      if (x.kind !== RVKind.Number || y.kind !== RVKind.Number) {
        continue;
      }
      sum += (x.value - y.value) ** 2;
    }
  }
  return rvNumber(sum);
};

export const fnMULTINOMIAL: NativeFunction = args => {
  const nums = flattenNumbers(args);
  const err = firstError(nums);
  if (err) {
    return err;
  }
  let sum = 0;
  let denom = 1;
  for (const n of nums) {
    const ni = Math.floor((n as NumberValue).value);
    if (ni < 0) {
      return ERRORS.NUM;
    }
    sum += ni;
    for (let i = 2; i <= ni; i++) {
      denom *= i;
    }
  }
  let numer = 1;
  for (let i = 2; i <= sum; i++) {
    numer *= i;
  }
  return rvNumber(numer / denom);
};

export const fnFACT: NativeFunction = args => {
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

export const fnFACTDOUBLE: NativeFunction = args => {
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
  }
  return rvNumber(result);
};

export const fnCOMBIN: NativeFunction = args => {
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
    result = (result * (ni - i)) / (i + 1);
  }
  return rvNumber(Math.round(result));
};

export const fnCOMBINA: NativeFunction = args => {
  const nRV = argToNumber(args[0]);
  if (isError(nRV)) {
    return nRV;
  }
  const kRV = argToNumber(args[1]);
  if (isError(kRV)) {
    return kRV;
  }
  return fnCOMBIN([rvNumber(nRV.value + kRV.value - 1), kRV]);
};

export const fnPERMUT: NativeFunction = args => {
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
