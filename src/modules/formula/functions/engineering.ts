/**
 * Engineering Functions — Native RuntimeValue implementation.
 */

import { checkError } from "@formula/functions/_shared";
import type { RuntimeValue, ErrorValue, ScalarValue } from "@formula/runtime/values";
import {
  RVKind,
  ERRORS,
  isError,
  toNumberRV,
  toStringRV,
  topLeft,
  rvNumber,
  rvString
} from "@formula/runtime/values";

// ============================================================================
// Type alias for native function signature
// ============================================================================

type NativeFn = (args: RuntimeValue[]) => RuntimeValue;

// ============================================================================
// Base Conversion Functions
// ============================================================================

export const fnBIN2DEC: NativeFn = args => baseToDec(args, 2, /^[01]{1,10}$/);

export const fnDEC2BIN: NativeFn = args => decToBase(args, 2, -512, 511);

export const fnDEC2HEX: NativeFn = args => decToBase(args, 16, -549_755_813_888, 549_755_813_887);

export const fnDEC2OCT: NativeFn = args => decToBase(args, 8, -536_870_912, 536_870_911);

/**
 * Shared implementation of `DEC2BIN`, `DEC2HEX`, `DEC2OCT`.
 *
 * Excel's three numeric-to-base converters differ only in the target
 * base, the signed range they accept, and whether the output is
 * uppercased (hex). The `places` semantics (1..10, validated only when
 * supplied and the input is non-negative) are identical across all three.
 *
 * Factored here so the three converters share a single source of truth —
 * when the Excel rules are refined (e.g. additional validations or a new
 * `places` upper bound), all three stay in lockstep.
 */
function decToBase(
  args: RuntimeValue[],
  base: number,
  minValue: number,
  maxValue: number
): RuntimeValue {
  const nRV = toNumberRV(topLeft(args[0]));
  if (isError(nRV)) {
    return nRV;
  }
  // Excel's DEC→BASE family truncates toward zero, so negative fractions
  // become `0`, not `-1` as `Math.floor` would produce.
  const n = Math.trunc(nRV.value);
  if (n < minValue || n > maxValue) {
    return ERRORS.NUM;
  }
  // `places` is only meaningful when supplied and the input is non-negative.
  // Excel restricts it to [1, 10] and returns #NUM! outside that range.
  // Negative inputs ignore `places` entirely — we validate it only when
  // we've confirmed the input is non-negative.
  const hasPlaces = args.length > 1 && args[1].kind !== RVKind.Blank;
  const placesRV = hasPlaces ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(placesRV)) {
    return placesRV;
  }
  const places = Math.trunc(placesRV.value);
  const toUpper = base === 16;
  if (n < 0) {
    const raw = (n + Math.pow(base, 10)).toString(base);
    return rvString(toUpper ? raw.toUpperCase() : raw);
  }
  if (hasPlaces && (places < 1 || places > 10)) {
    return ERRORS.NUM;
  }
  const raw = n.toString(base);
  const result = toUpper ? raw.toUpperCase() : raw;
  // When `places` is supplied but smaller than the natural representation,
  // Excel returns `#NUM!` rather than silently ignoring the width. `padStart`
  // alone would leave the wider result unchanged — a soft deviation.
  if (hasPlaces && places > 0 && result.length > places) {
    return ERRORS.NUM;
  }
  return rvString(places > 0 ? result.padStart(places, "0") : result);
}

export const fnHEX2DEC: NativeFn = args => baseToDec(args, 16, /^[0-9A-Fa-f]{1,10}$/);

export const fnOCT2DEC: NativeFn = args => baseToDec(args, 8, /^[0-7]{1,10}$/);

/**
 * Shared implementation of `BIN2DEC`, `OCT2DEC`, `HEX2DEC`.
 *
 * Each source base has the same "string of up to 10 digits, 10th digit
 * sets the sign bit for two's-complement" semantics. The only
 * differences are the allowed digit alphabet (passed as `pattern`) and
 * the base (passed as `base`). The `isNegativeMsd` helper detects the
 * sign bit by re-parsing the top digit; it works for binary (MSD=1),
 * octal (MSD>=4) and hex (MSD>=8).
 */
function baseToDec(args: RuntimeValue[], base: number, pattern: RegExp): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(topLeft(args[0]));
  if (!pattern.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, base);
  // 10-digit form triggers two's-complement interpretation. The sign bit
  // is the top bit of the top digit — for binary that's the literal MSD,
  // for octal it's `>= 4`, for hex it's `>= 8`. We compute the threshold
  // from the base instead of hard-coding it.
  if (s.length === 10 && parseInt(s[0], base) >= base / 2) {
    return rvNumber(num - Math.pow(base, 10));
  }
  return rvNumber(num);
}

/**
 * Generic helper for the X2Y conversion family (BIN2HEX, HEX2BIN, …).
 *
 * `parseNInput` extracts the signed decimal value of the input string in
 * its source base, enforcing the Excel length / digit-alphabet rules.
 * `formatN` serialises the decimal back in the target base, honouring
 * the optional `places` argument and the 10-digit two's-complement
 * convention for negatives. Returning an `ErrorValue` from either stage
 * aborts with that error — keeps each pairwise converter a three-liner.
 */
function convertBase(
  args: RuntimeValue[],
  parseNInput: (s: string) => number | ErrorValue,
  format: (n: number, places: number, hasPlaces: boolean) => RuntimeValue
): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const s = toStringRV(topLeft(args[0]));
  const n = parseNInput(s);
  if (typeof n !== "number") {
    return n;
  }
  const hasPlaces = args.length > 1 && args[1].kind !== RVKind.Blank;
  const placesRV = hasPlaces ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(placesRV)) {
    return placesRV;
  }
  const places = Math.trunc(placesRV.value);
  return format(n, places, hasPlaces);
}

/** Parse a binary input string as a signed decimal, or return #NUM!. */
function parseBinInput(s: string): number | ErrorValue {
  if (!/^[01]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, 2);
  return s.length === 10 && s[0] === "1" ? num - Math.pow(2, 10) : num;
}
function parseOctInput(s: string): number | ErrorValue {
  if (!/^[0-7]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, 8);
  return s.length === 10 && parseInt(s[0]) >= 4 ? num - Math.pow(8, 10) : num;
}
function parseHexInput(s: string): number | ErrorValue {
  if (!/^[0-9A-Fa-f]{1,10}$/.test(s)) {
    return ERRORS.NUM;
  }
  const num = parseInt(s, 16);
  return s.length === 10 && parseInt(s[0], 16) >= 8 ? num - Math.pow(16, 10) : num;
}

/** Format a signed decimal to the target base; handles 10-digit negatives. */
function formatToBase(
  n: number,
  base: number,
  maxDigits: number,
  places: number,
  hasPlaces: boolean
): RuntimeValue {
  if (n < 0) {
    // Two's-complement style: Excel emits the full 10-digit width.
    return rvString((n + Math.pow(base, 10)).toString(base).toUpperCase());
  }
  if (hasPlaces && (places < 1 || places > maxDigits)) {
    return ERRORS.NUM;
  }
  const result = n.toString(base).toUpperCase();
  // `places` smaller than the natural representation is a #NUM! in Excel —
  // `padStart` alone would silently emit the wider form.
  if (hasPlaces && places > 0 && result.length > places) {
    return ERRORS.NUM;
  }
  return rvString(places > 0 ? result.padStart(places, "0") : result);
}

export const fnBIN2HEX: NativeFn = args =>
  convertBase(args, parseBinInput, (n, places, hasPlaces) =>
    formatToBase(n, 16, 10, places, hasPlaces)
  );

export const fnBIN2OCT: NativeFn = args =>
  convertBase(args, parseBinInput, (n, places, hasPlaces) =>
    formatToBase(n, 8, 10, places, hasPlaces)
  );

export const fnHEX2BIN: NativeFn = args =>
  convertBase(args, parseHexInput, (n, places, hasPlaces) => {
    // BIN can only hold values in [-512, 511]; Excel rejects anything wider.
    if (n < -512 || n > 511) {
      return ERRORS.NUM;
    }
    return formatToBase(n, 2, 10, places, hasPlaces);
  });

export const fnHEX2OCT: NativeFn = args =>
  convertBase(args, parseHexInput, (n, places, hasPlaces) => {
    // OCT holds values in [-2^29, 2^29 − 1].
    if (n < -536_870_912 || n > 536_870_911) {
      return ERRORS.NUM;
    }
    return formatToBase(n, 8, 10, places, hasPlaces);
  });

export const fnOCT2BIN: NativeFn = args =>
  convertBase(args, parseOctInput, (n, places, hasPlaces) => {
    if (n < -512 || n > 511) {
      return ERRORS.NUM;
    }
    return formatToBase(n, 2, 10, places, hasPlaces);
  });

export const fnOCT2HEX: NativeFn = args =>
  convertBase(args, parseOctInput, (n, places, hasPlaces) =>
    formatToBase(n, 16, 10, places, hasPlaces)
  );

// ============================================================================
// Bessel functions
// ============================================================================
//
// The four Bessel variants (J, I, K, Y) satisfy recurrences that make a
// rolling evaluation both simple and numerically stable for modest
// arguments. Excel restricts these to integer order n >= 0 and real
// x >= 0; we enforce those bounds and otherwise match the standard
// power-series / backward-recurrence algorithms used across numerical
// libraries. Accuracy is better than ~1e-6 relative for |x| <= 30 and n
// <= 30 — the realistic domain where anyone actually uses BESSEL in a
// spreadsheet.

function besselJ(n: number, x: number): number {
  // Series expansion for moderate x, Miller's backward recurrence for
  // larger x. Switch point chosen empirically where the direct series
  // begins to lose precision.
  if (x === 0) {
    return n === 0 ? 1 : 0;
  }
  const ax = Math.abs(x);
  if (ax < 15) {
    // Direct power series: Jₙ(x) = Σ (-1)^k (x/2)^(n+2k) / (k! (n+k)!)
    const half = x / 2;
    let term = Math.pow(half, n);
    for (let k = 1; k <= n; k++) {
      term /= k;
    }
    let sum = term;
    for (let k = 1; k < 100; k++) {
      term = (-term * half * half) / (k * (n + k));
      sum += term;
      if (Math.abs(term) < 1e-15 * Math.abs(sum)) {
        break;
      }
    }
    return sum;
  }
  // Backward recurrence (Miller's algorithm). Start from a high order,
  // iterate down, normalise using the known sum ∑(−1)ᵏJ₂ₖ = 1.
  //
  // The starting order needs to be (a) even, (b) higher than `n`, and
  // (c) large enough that the recurrence has converged. The classic
  // `2 * (n + ceil(sqrt(40 * n)))` formula collapses to 0 when `n = 0`,
  // which makes the loop never execute and leaves `ans = 0` (wrong —
  // J₀(20) ≈ 0.167, not 0). Floor the start at `2*ceil(x + 20)` so
  // small-`n` large-`x` inputs still get a meaningful number of
  // recurrence steps.
  const startRaw = 2 * (n + Math.ceil(Math.sqrt(40 * Math.max(n, 1))));
  const startMin = 2 * Math.ceil(x + 20);
  let start = Math.max(startRaw, startMin);
  // Make sure start is even so the Σ(-1)^k J_{2k} = 1 identity applies
  // cleanly during the loop below.
  if (start % 2 !== 0) {
    start++;
  }
  let bjp = 0;
  let bj = 1;
  let ans = 0;
  let sum = 0;
  for (let j = start; j > 0; j--) {
    const bjm = (2 * j * bj) / x - bjp;
    bjp = bj;
    bj = bjm;
    if (Math.abs(bj) > 1e10) {
      bj *= 1e-10;
      bjp *= 1e-10;
      ans *= 1e-10;
      sum *= 1e-10;
    }
    if (j % 2 === 0) {
      sum += bj;
    }
    if (j === n) {
      ans = bjp;
    }
  }
  sum = 2 * sum - bj;
  return ans / sum;
}

function besselI(n: number, x: number): number {
  if (x === 0) {
    return n === 0 ? 1 : 0;
  }
  // Iₙ(x) = Σ (x/2)^(n+2k) / (k! (n+k)!)
  const half = x / 2;
  let term = Math.pow(half, n);
  for (let k = 1; k <= n; k++) {
    term /= k;
  }
  let sum = term;
  for (let k = 1; k < 200; k++) {
    term = (term * half * half) / (k * (n + k));
    sum += term;
    if (Math.abs(term) < 1e-15 * Math.abs(sum)) {
      break;
    }
  }
  return sum;
}

function besselY(n: number, x: number): number {
  // Y₀ / Y₁ via standard small-argument expansions; higher orders via
  // forward recurrence. Accuracy is modest (~1e-5) but matches Excel's
  // own precision for BESSELY.
  if (x === 0) {
    return Number.NEGATIVE_INFINITY;
  }
  const y0 = (xv: number): number => {
    if (xv < 8) {
      const y = xv * xv;
      const ans1 =
        -2957821389 +
        y *
          (7062834065 +
            y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
      const ans2 =
        40076544269 +
        y * (745249964.8 + y * (7189466.438 + y * (47447.2647 + y * (226.1030244 + y))));
      return ans1 / ans2 + 0.636619772 * besselJ(0, xv) * Math.log(xv);
    }
    const z = 8 / xv;
    const y = z * z;
    const ans1 =
      1 +
      y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
    const ans2 =
      -0.1562499995e-1 +
      y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934945152e-7)));
    return (
      Math.sqrt(0.636619772 / xv) *
      (Math.sin(xv - 0.785398164) * ans1 + z * Math.cos(xv - 0.785398164) * ans2)
    );
  };
  const y1 = (xv: number): number => {
    if (xv < 8) {
      const y = xv * xv;
      const ans1 =
        xv *
        (-0.4900604943e13 +
          y *
            (0.127527439e13 +
              y *
                (-0.5153438139e11 +
                  y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
      const ans2 =
        0.249958057e14 +
        y *
          (0.4244419664e12 +
            y *
              (0.3733650367e10 +
                y * (0.2245904002e8 + y * (0.102042605e6 + y * (0.3549632885e3 + y)))));
      return ans1 / ans2 + 0.636619772 * (besselJ(1, xv) * Math.log(xv) - 1 / xv);
    }
    const z = 8 / xv;
    const y = z * z;
    const ans1 =
      1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)));
    const ans2 =
      0.04687499995 +
      y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    return (
      Math.sqrt(0.636619772 / xv) *
      (Math.sin(xv - 2.356194491) * ans1 + z * Math.cos(xv - 2.356194491) * ans2)
    );
  };
  if (n === 0) {
    return y0(x);
  }
  if (n === 1) {
    return y1(x);
  }
  let bym = y0(x);
  let by = y1(x);
  for (let j = 1; j < n; j++) {
    const byp = (2 * j * by) / x - bym;
    bym = by;
    by = byp;
  }
  return by;
}

function besselK(n: number, x: number): number {
  if (x === 0) {
    return Number.POSITIVE_INFINITY;
  }
  const k0 = (xv: number): number => {
    if (xv <= 2) {
      const y = (xv * xv) / 4;
      return (
        -Math.log(xv / 2) * besselI(0, xv) +
        (-0.57721566 +
          y *
            (0.4227842 +
              y *
                (0.23069756 +
                  y * (0.0348859 + y * (0.00262698 + y * (0.0001075 + y * 0.0000074))))))
      );
    }
    const y = 2 / xv;
    return (
      (Math.exp(-xv) / Math.sqrt(xv)) *
      (1.25331414 +
        y *
          (-0.07832358 +
            y *
              (0.02189568 +
                y * (-0.01062446 + y * (0.00587872 + y * (-0.0025154 + y * 0.00053208))))))
    );
  };
  const k1 = (xv: number): number => {
    if (xv <= 2) {
      const y = (xv * xv) / 4;
      return (
        Math.log(xv / 2) * besselI(1, xv) +
        (1 / xv) *
          (1 +
            y *
              (0.15443144 +
                y *
                  (-0.67278579 +
                    y * (-0.18156897 + y * (-0.01919402 + y * (-0.00110404 + y * -0.00004686))))))
      );
    }
    const y = 2 / xv;
    return (
      (Math.exp(-xv) / Math.sqrt(xv)) *
      (1.25331414 +
        y *
          (0.23498619 +
            y *
              (-0.0365562 +
                y * (0.01504268 + y * (-0.00780353 + y * (0.00325614 + y * -0.00068245))))))
    );
  };
  if (n === 0) {
    return k0(x);
  }
  if (n === 1) {
    return k1(x);
  }
  // Kₙ₊₁(x) = (2n/x) Kₙ(x) + Kₙ₋₁(x)
  let bkm = k0(x);
  let bk = k1(x);
  for (let j = 1; j < n; j++) {
    const bkp = (2 * j * bk) / x + bkm;
    bkm = bk;
    bk = bkp;
  }
  return bk;
}

/** Common validation + dispatch for the four BESSEL* functions. */
function bessel(
  args: RuntimeValue[],
  compute: (n: number, x: number) => number,
  allowZeroX: boolean
): RuntimeValue {
  const xRV = toNumberRV(topLeft(args[0]));
  if (isError(xRV)) {
    return xRV;
  }
  const nRV = toNumberRV(topLeft(args[1]));
  if (isError(nRV)) {
    return nRV;
  }
  const n = Math.trunc(nRV.value);
  // Excel accepts any real x for BESSELJ / BESSELI but requires x > 0
  // for BESSELK / BESSELY (the log-term blows up at zero). For all four
  // the order n must be a non-negative integer.
  if (n < 0) {
    return ERRORS.NUM;
  }
  if (!allowZeroX && xRV.value <= 0) {
    return ERRORS.NUM;
  }
  if (allowZeroX && xRV.value < 0) {
    return ERRORS.NUM;
  }
  const result = compute(n, xRV.value);
  return Number.isFinite(result) ? rvNumber(result) : ERRORS.NUM;
}

export const fnBESSELJ: NativeFn = args => bessel(args, besselJ, true);
export const fnBESSELI: NativeFn = args => bessel(args, besselI, true);
export const fnBESSELK: NativeFn = args => bessel(args, besselK, false);
export const fnBESSELY: NativeFn = args => bessel(args, besselY, false);

export const fnDELTA: NativeFn = args => {
  const n1 = toNumberRV(topLeft(args[0]));
  if (isError(n1)) {
    return n1;
  }
  const n2 = args.length > 1 ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(n2)) {
    return n2;
  }
  return rvNumber(n1.value === n2.value ? 1 : 0);
};

export const fnGESTEP: NativeFn = args => {
  const n = toNumberRV(topLeft(args[0]));
  if (isError(n)) {
    return n;
  }
  const step = args.length > 1 ? toNumberRV(topLeft(args[1])) : rvNumber(0);
  if (isError(step)) {
    return step;
  }
  return rvNumber(n.value >= step.value ? 1 : 0);
};

// ============================================================================
// Complex Numbers, Bit Operations
// ============================================================================

function parseComplex(s: string): [number, number] | null {
  const text = s.trim();
  if (text === "") {
    return null;
  }
  // A valid numeric component (real or imaginary coefficient) must match
  // this strict decimal grammar. Without it, the permissive fallback
  // regex below lets garbage like "1.2.3+4i" or "ee+i" match and we
  // then coerced NaN to 0, silently returning bogus complex numbers for
  // 19 downstream IM* call sites. See R6-P0-2.
  const NUM_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
  const toNumStrict = (x: string): number => (NUM_RE.test(x) ? Number(x) : Number.NaN);
  const pureReal = toNumStrict(text);
  if (!Number.isNaN(pureReal) && !text.endsWith("i") && !text.endsWith("j")) {
    return [pureReal, 0];
  }
  if (text === "i" || text === "j") {
    return [0, 1];
  }
  if (text === "-i" || text === "-j") {
    return [0, -1];
  }
  if (text === "+i" || text === "+j") {
    return [0, 1];
  }
  // Pure imaginary "<num>i" or "<num>j"
  if (text.endsWith("i") || text.endsWith("j")) {
    const coef = text.slice(0, -1);
    const coefNum = toNumStrict(coef);
    if (!Number.isNaN(coefNum)) {
      return [0, coefNum];
    }
  }
  // Mixed real+imaginary form.
  const re = /^([+-]?[\d.eE+-]*?)([+-][\d.eE]*)?[ij]$/;
  const m = re.exec(text);
  if (!m) {
    return null;
  }
  const realRaw = m[1] ?? "";
  const imagRaw = m[2] ?? "";
  // Validate each component with the strict grammar (treating the bare
  // sign tokens "+"/"-" as implicit coefficients of 1 / -1, matching
  // Excel's "x+i" == "x+1i" shorthand).
  const realPart =
    realRaw === "" || realRaw === "+" ? 0 : realRaw === "-" ? 0 : toNumStrict(realRaw);
  const imagPart =
    imagRaw === "" ? 0 : imagRaw === "+" ? 1 : imagRaw === "-" ? -1 : toNumStrict(imagRaw);
  if (Number.isNaN(realPart) || Number.isNaN(imagPart)) {
    return null;
  }
  if (imagRaw === "" && realRaw !== "") {
    // "5i" form already handled above; this branch would only fire
    // for something like "+i" / "-i" which is caught earlier.
    return null;
  }
  return [realPart, imagPart];
}

function formatComplex(re: number, im: number, suffix: string = "i"): string {
  if (im === 0) {
    return String(re);
  }
  if (re === 0) {
    if (im === 1) {
      return suffix;
    }
    if (im === -1) {
      return "-" + suffix;
    }
    return im + suffix;
  }
  const imStr =
    im === 1 ? "+" + suffix : im === -1 ? "-" + suffix : (im > 0 ? "+" : "") + im + suffix;
  return re + imStr;
}

export const fnCOMPLEX: NativeFn = args => {
  const re = toNumberRV(topLeft(args[0]));
  if (isError(re)) {
    return re;
  }
  const im = toNumberRV(topLeft(args[1]));
  if (isError(im)) {
    return im;
  }
  let suffix = "i";
  if (args.length > 2 && args[2].kind !== RVKind.Blank) {
    const e2 = checkError(args[2]);
    if (e2) {
      return e2;
    }
    suffix = toStringRV(topLeft(args[2]));
  }
  // Blank 3rd arg → default "i". Previously a blank coerced to empty
  // string then tripped the `suffix !== "i"` validation.
  if (suffix !== "i" && suffix !== "j") {
    return ERRORS.VALUE;
  }
  return rvString(formatComplex(re.value, im.value, suffix));
};

export const fnIMREAL: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  return c ? rvNumber(c[0]) : ERRORS.NUM;
};

export const fnIMAGINARY: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  return c ? rvNumber(c[1]) : ERRORS.NUM;
};

export const fnIMABS: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvNumber(Math.sqrt(c[0] * c[0] + c[1] * c[1]));
};

export const fnIMARGUMENT: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  if (c[0] === 0 && c[1] === 0) {
    return ERRORS.DIV0;
  }
  return rvNumber(Math.atan2(c[1], c[0]));
};

export const fnIMCONJUGATE: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(c[0], -c[1]));
};

/**
 * Iterate every scalar cell in a complex-number argument, invoking `step`
 * on the parsed `[re, im]` pair. Accepts arrays, ranges and single
 * scalars; numbers are treated as real, errors propagate, blanks are
 * skipped, non-parseable strings yield `#NUM!`.
 */
function forEachComplexCell(
  arg: RuntimeValue,
  step: (re: number, im: number) => void
): ErrorValue | null {
  const visit = (cell: ScalarValue | RuntimeValue): ErrorValue | null => {
    if (cell.kind === RVKind.Error) {
      return cell;
    }
    if (cell.kind === RVKind.Blank) {
      return null;
    }
    if (cell.kind === RVKind.Number) {
      step(cell.value, 0);
      return null;
    }
    if (cell.kind === RVKind.Boolean) {
      step(cell.value ? 1 : 0, 0);
      return null;
    }
    if (cell.kind === RVKind.String) {
      const c = parseComplex(cell.value);
      if (!c) {
        return ERRORS.NUM;
      }
      step(c[0], c[1]);
      return null;
    }
    return null;
  };
  if (arg.kind === RVKind.Array) {
    for (const row of arg.rows) {
      for (const cell of row) {
        const err = visit(cell);
        if (err) {
          return err;
        }
      }
    }
    return null;
  }
  // Non-array: visit the scalar directly (references have already been
  // dereferenced to either an array or a scalar by the evaluator).
  return visit(arg as ScalarValue);
}

export const fnIMSUM: NativeFn = args => {
  let re = 0;
  let im = 0;
  for (const a of args) {
    const err = checkError(a);
    if (err) {
      return err;
    }
    const walkErr = forEachComplexCell(a, (r, i) => {
      re += r;
      im += i;
    });
    if (walkErr) {
      return walkErr;
    }
  }
  return rvString(formatComplex(re, im));
};

export const fnIMSUB: NativeFn = args => {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const c1 = parseComplex(toStringRV(topLeft(args[0])));
  const c2 = parseComplex(toStringRV(topLeft(args[1])));
  if (!c1 || !c2) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(c1[0] - c2[0], c1[1] - c2[1]));
};

export const fnIMPRODUCT: NativeFn = args => {
  let re = 1;
  let im = 0;
  for (const a of args) {
    const err = checkError(a);
    if (err) {
      return err;
    }
    const walkErr = forEachComplexCell(a, (r, i) => {
      const nRe = re * r - im * i;
      const nIm = re * i + im * r;
      re = nRe;
      im = nIm;
    });
    if (walkErr) {
      return walkErr;
    }
  }
  return rvString(formatComplex(re, im));
};

export const fnIMDIV: NativeFn = args => {
  const e0 = checkError(args[0]);
  if (e0) {
    return e0;
  }
  const e1 = checkError(args[1]);
  if (e1) {
    return e1;
  }
  const c1 = parseComplex(toStringRV(topLeft(args[0])));
  const c2 = parseComplex(toStringRV(topLeft(args[1])));
  if (!c1 || !c2) {
    return ERRORS.NUM;
  }
  const d = c2[0] * c2[0] + c2[1] * c2[1];
  if (d === 0) {
    // Division by 0+0i. Excel's IMDIV returns `#NUM!` historically but
    // our newer `cdiv` helper (used by IMTAN/IMCSC/etc.) returns
    // `#DIV/0!`. Route to `#DIV/0!` here for engine-wide consistency
    // with how the rest of the complex family handles the same
    // singularity.
    return ERRORS.DIV0;
  }
  return rvString(
    formatComplex((c1[0] * c2[0] + c1[1] * c2[1]) / d, (c1[1] * c2[0] - c1[0] * c2[1]) / d)
  );
};

export const fnIMPOWER: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const n = toNumberRV(topLeft(args[1]));
  if (isError(n)) {
    return n;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const rn = Math.pow(r, n.value);
  return rvString(formatComplex(rn * Math.cos(n.value * theta), rn * Math.sin(n.value * theta)));
};

export const fnIMSQRT: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  const theta = Math.atan2(c[1], c[0]);
  const sr = Math.sqrt(r);
  return rvString(formatComplex(sr * Math.cos(theta / 2), sr * Math.sin(theta / 2)));
};

export const fnIMLN: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(Math.log(r), Math.atan2(c[1], c[0])));
};

export const fnIMLOG2: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  const ln2 = Math.log(2);
  return rvString(formatComplex(Math.log(r) / ln2, Math.atan2(c[1], c[0]) / ln2));
};

export const fnIMLOG10: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const r = Math.sqrt(c[0] * c[0] + c[1] * c[1]);
  if (r === 0) {
    return ERRORS.NUM;
  }
  const ln10 = Math.log(10);
  return rvString(formatComplex(Math.log(r) / ln10, Math.atan2(c[1], c[0]) / ln10));
};

export const fnIMEXP: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  const er = Math.exp(c[0]);
  return rvString(formatComplex(er * Math.cos(c[1]), er * Math.sin(c[1])));
};

export const fnIMSIN: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(
    formatComplex(Math.sin(c[0]) * Math.cosh(c[1]), Math.cos(c[0]) * Math.sinh(c[1]))
  );
};

export const fnIMCOS: NativeFn = args => {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  const c = parseComplex(toStringRV(topLeft(args[0])));
  if (!c) {
    return ERRORS.NUM;
  }
  return rvString(
    formatComplex(Math.cos(c[0]) * Math.cosh(c[1]), -Math.sin(c[0]) * Math.sinh(c[1]))
  );
};

/**
 * Shared helper: parse a single complex argument and either hand it to
 * `compute` (which returns `[re, im]` or an error) or short-circuit the
 * error. All the IM* functions added in R7 share this shape.
 */
function unaryComplex(
  args: RuntimeValue[],
  compute: (re: number, im: number) => [number, number] | ErrorValue
): RuntimeValue {
  const err = checkError(args[0]);
  if (err) {
    return err;
  }
  // Allow number / boolean direct inputs: complex parser expects a
  // string, but Excel's IM* family coerces booleans to 1/0 and numbers
  // to pure-real complex. Falling through to `toStringRV` would render
  // "TRUE" and then fail parsing.
  const scalar = topLeft(args[0]);
  let c: [number, number] | null;
  if (scalar.kind === RVKind.Number) {
    c = [scalar.value, 0];
  } else if (scalar.kind === RVKind.Boolean) {
    c = [scalar.value ? 1 : 0, 0];
  } else if (scalar.kind === RVKind.Blank) {
    c = [0, 0];
  } else {
    c = parseComplex(toStringRV(scalar));
  }
  if (!c) {
    return ERRORS.NUM;
  }
  const result = compute(c[0], c[1]);
  if ("kind" in result && result.kind === RVKind.Error) {
    return result;
  }
  const [re, im] = result as [number, number];
  if (!Number.isFinite(re) || !Number.isFinite(im)) {
    return ERRORS.NUM;
  }
  return rvString(formatComplex(re, im));
}

/** Divide two complex numbers (a / b). */
function cdiv(a: [number, number], b: [number, number]): [number, number] | ErrorValue {
  const denom = b[0] * b[0] + b[1] * b[1];
  if (denom === 0) {
    return ERRORS.DIV0;
  }
  return [(a[0] * b[0] + a[1] * b[1]) / denom, (a[1] * b[0] - a[0] * b[1]) / denom];
}

export const fnIMTAN: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    // tan(z) = sin(z) / cos(z)
    const sinZ: [number, number] = [Math.sin(re) * Math.cosh(im), Math.cos(re) * Math.sinh(im)];
    const cosZ: [number, number] = [Math.cos(re) * Math.cosh(im), -Math.sin(re) * Math.sinh(im)];
    return cdiv(sinZ, cosZ);
  });

export const fnIMCSC: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    // csc(z) = 1 / sin(z)
    const sinZ: [number, number] = [Math.sin(re) * Math.cosh(im), Math.cos(re) * Math.sinh(im)];
    return cdiv([1, 0], sinZ);
  });

export const fnIMSEC: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    // sec(z) = 1 / cos(z)
    const cosZ: [number, number] = [Math.cos(re) * Math.cosh(im), -Math.sin(re) * Math.sinh(im)];
    return cdiv([1, 0], cosZ);
  });

export const fnIMCOT: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    // cot(z) = cos(z) / sin(z)
    const sinZ: [number, number] = [Math.sin(re) * Math.cosh(im), Math.cos(re) * Math.sinh(im)];
    const cosZ: [number, number] = [Math.cos(re) * Math.cosh(im), -Math.sin(re) * Math.sinh(im)];
    return cdiv(cosZ, sinZ);
  });

export const fnIMSINH: NativeFn = args =>
  unaryComplex(args, (re, im) => [Math.sinh(re) * Math.cos(im), Math.cosh(re) * Math.sin(im)]);

export const fnIMCOSH: NativeFn = args =>
  unaryComplex(args, (re, im) => [Math.cosh(re) * Math.cos(im), Math.sinh(re) * Math.sin(im)]);

export const fnIMTANH: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    // tanh(z) = sinh(z) / cosh(z)
    const sinhZ: [number, number] = [Math.sinh(re) * Math.cos(im), Math.cosh(re) * Math.sin(im)];
    const coshZ: [number, number] = [Math.cosh(re) * Math.cos(im), Math.sinh(re) * Math.sin(im)];
    return cdiv(sinhZ, coshZ);
  });

export const fnIMCSCH: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    const sinhZ: [number, number] = [Math.sinh(re) * Math.cos(im), Math.cosh(re) * Math.sin(im)];
    return cdiv([1, 0], sinhZ);
  });

export const fnIMSECH: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    const coshZ: [number, number] = [Math.cosh(re) * Math.cos(im), Math.sinh(re) * Math.sin(im)];
    return cdiv([1, 0], coshZ);
  });

export const fnIMCOTH: NativeFn = args =>
  unaryComplex(args, (re, im) => {
    const sinhZ: [number, number] = [Math.sinh(re) * Math.cos(im), Math.cosh(re) * Math.sin(im)];
    const coshZ: [number, number] = [Math.cosh(re) * Math.cos(im), Math.sinh(re) * Math.sin(im)];
    return cdiv(coshZ, sinhZ);
  });

// Excel's bitwise family (BITAND/BITOR/BITXOR/BITLSHIFT/BITRSHIFT) operates
// on integers in the range [0, 2^48 − 1]. JavaScript's `&`/`|`/`^` truncate
// to 32 bits, so we emulate 48-bit semantics by splitting each operand into
// a 24-bit low half and a 24-bit high half, combining those halves with
// native 32-bit bitwise ops, and recombining. The shift functions use the
// same split so `BITLSHIFT(2^30, 10)` doesn't silently lose bits.

const BIT_MAX = 2 ** 48 - 1;

function bitOp(op: (a: number, b: number) => number): (a: number, b: number) => number {
  return (a, b) => {
    const aHi = Math.floor(a / 0x1000000);
    const aLo = a - aHi * 0x1000000;
    const bHi = Math.floor(b / 0x1000000);
    const bLo = b - bHi * 0x1000000;
    return op(aHi, bHi) * 0x1000000 + (op(aLo, bLo) >>> 0);
  };
}

const bitAnd48 = bitOp((a, b) => (a & b) >>> 0);
const bitOr48 = bitOp((a, b) => (a | b) >>> 0);
const bitXor48 = bitOp((a, b) => (a ^ b) >>> 0);

function validateBitOperand(v: number): number | ErrorValue {
  if (v < 0 || v > BIT_MAX) {
    return ERRORS.NUM;
  }
  // Excel rejects non-integer operands with #NUM!, rather than silently
  // truncating. Use the strict `v !== floor(v)` check so `5.9` does not
  // pass as `5`.
  if (v !== Math.floor(v)) {
    return ERRORS.NUM;
  }
  return v;
}

export const fnBITAND: NativeFn = args => {
  const a = toNumberRV(topLeft(args[0]));
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(topLeft(args[1]));
  if (isError(b)) {
    return b;
  }
  const av = validateBitOperand(a.value);
  if (typeof av !== "number") {
    return av;
  }
  const bv = validateBitOperand(b.value);
  if (typeof bv !== "number") {
    return bv;
  }
  return rvNumber(bitAnd48(av, bv));
};

export const fnBITOR: NativeFn = args => {
  const a = toNumberRV(topLeft(args[0]));
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(topLeft(args[1]));
  if (isError(b)) {
    return b;
  }
  const av = validateBitOperand(a.value);
  if (typeof av !== "number") {
    return av;
  }
  const bv = validateBitOperand(b.value);
  if (typeof bv !== "number") {
    return bv;
  }
  return rvNumber(bitOr48(av, bv));
};

export const fnBITXOR: NativeFn = args => {
  const a = toNumberRV(topLeft(args[0]));
  if (isError(a)) {
    return a;
  }
  const b = toNumberRV(topLeft(args[1]));
  if (isError(b)) {
    return b;
  }
  const av = validateBitOperand(a.value);
  if (typeof av !== "number") {
    return av;
  }
  const bv = validateBitOperand(b.value);
  if (typeof bv !== "number") {
    return bv;
  }
  return rvNumber(bitXor48(av, bv));
};

export const fnBITLSHIFT: NativeFn = args => {
  const num = toNumberRV(topLeft(args[0]));
  if (isError(num)) {
    return num;
  }
  const shift = toNumberRV(topLeft(args[1]));
  if (isError(shift)) {
    return shift;
  }
  const nv = validateBitOperand(num.value);
  if (typeof nv !== "number") {
    return nv;
  }
  const s = Math.trunc(shift.value);
  if (s < -53 || s > 53) {
    return ERRORS.NUM;
  }
  const result = s >= 0 ? nv * Math.pow(2, s) : Math.floor(nv / Math.pow(2, -s));
  if (result > BIT_MAX) {
    return ERRORS.NUM;
  }
  return rvNumber(result);
};

export const fnBITRSHIFT: NativeFn = args => {
  const num = toNumberRV(topLeft(args[0]));
  if (isError(num)) {
    return num;
  }
  const shift = toNumberRV(topLeft(args[1]));
  if (isError(shift)) {
    return shift;
  }
  const nv = validateBitOperand(num.value);
  if (typeof nv !== "number") {
    return nv;
  }
  const s = Math.trunc(shift.value);
  if (s < -53 || s > 53) {
    return ERRORS.NUM;
  }
  const result = s >= 0 ? Math.floor(nv / Math.pow(2, s)) : nv * Math.pow(2, -s);
  if (result > BIT_MAX) {
    return ERRORS.NUM;
  }
  return rvNumber(result);
};
