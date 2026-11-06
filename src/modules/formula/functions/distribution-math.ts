/**
 * The special functions the statistical distributions and their inverses are built on, to double precision.
 *
 * The approximations these replaced were good to about seven digits (`normSDist`), or broke outright past a point:
 * `lnGamma` was `Math.log(Γ(x))`, and Γ overflows past 171, so every distribution with large degrees of freedom became
 * NaN — `CONFIDENCE.T(0.05,8.27,350)` was `#NUM!`. The inverses are checked against values Excel computed, to 1e-9
 * relative (`__tests__/statistical-inverses.test.ts`), which needs every layer below them to be accurate.
 *
 * Each distribution function is available as a lower and an upper tail, computed directly rather than as `1 - other`:
 * an inverse at `p = 1e-10` in the upper tail cannot be found from a lower tail that has rounded to 1.
 */

const EPSILON = 1e-16;
const TINY = 1e-300;
const LN_SQRT_2PI = 0.5 * Math.log(2 * Math.PI);

const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313,
  -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6,
  1.5056327351493116e-7
];

/** ln Γ(x) for x > 0, without overflow. */
export function logGamma(x: number): number {
  if (x < 0.5) {
    // Reflection: Γ(x)Γ(1-x) = π / sin(πx).
    return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - logGamma(1 - x);
  }
  const z = x - 1;
  let sum = LANCZOS[0]!;
  for (let i = 1; i < LANCZOS.length; i++) {
    sum += LANCZOS[i]! / (z + i);
  }
  const t = z + 7.5;
  return LN_SQRT_2PI + (z + 0.5) * Math.log(t) - t + Math.log(sum);
}

/** ln B(a, b). */
function logBeta(a: number, b: number): number {
  return logGamma(a) + logGamma(b) - logGamma(a + b);
}

/** The continued fraction of the regularized incomplete beta (modified Lentz). */
function betaContinuedFraction(x: number, a: number, b: number): number {
  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1);
  if (Math.abs(d) < TINY) {
    d = TINY;
  }
  d = 1 / d;
  let f = d;
  for (let m = 1; m <= 100000; m++) {
    const m2 = 2 * m;
    let num = (m * (b - m) * x) / ((a + m2 - 1) * (a + m2));
    d = 1 + num * d;
    c = 1 + num / c;
    d = 1 / (Math.abs(d) < TINY ? TINY : d);
    c = Math.abs(c) < TINY ? TINY : c;
    f *= d * c;
    num = (-(a + m) * (a + b + m) * x) / ((a + m2) * (a + m2 + 1));
    d = 1 + num * d;
    c = 1 + num / c;
    d = 1 / (Math.abs(d) < TINY ? TINY : d);
    c = Math.abs(c) < TINY ? TINY : c;
    const delta = d * c;
    f *= delta;
    if (Math.abs(delta - 1) < EPSILON) {
      break;
    }
  }
  return f;
}

/**
 * The regularized incomplete beta I_x(a, b) (`upper` false) or 1 − I_x(a, b) (`upper` true), with `y = 1 − x` passed
 * separately so a caller that knows it exactly does not lose it to rounding.
 */
export function regularizedBeta(x: number, y: number, a: number, b: number, upper = false): number {
  if (x <= 0) {
    return upper ? 1 : 0;
  }
  if (y <= 0) {
    return upper ? 0 : 1;
  }
  const front = Math.exp(a * Math.log(x) + b * Math.log(y) - logBeta(a, b));
  // The fraction converges fast below the mean; above it, the symmetric form I_x(a,b) = 1 − I_y(b,a) is evaluated.
  if (x < (a + 1) / (a + b + 2)) {
    const lower = (front * betaContinuedFraction(x, a, b)) / a;
    return upper ? 1 - lower : lower;
  }
  const complement = (front * betaContinuedFraction(y, b, a)) / b;
  return upper ? complement : 1 - complement;
}

/** The regularized lower incomplete gamma P(a, x) (`upper` false) or Q(a, x) = 1 − P(a, x) (`upper` true). */
export function regularizedGamma(a: number, x: number, upper = false): number {
  if (x <= 0) {
    return upper ? 1 : 0;
  }
  if (!Number.isFinite(x)) {
    return upper ? 0 : 1;
  }
  const logFront = a * Math.log(x) - x - logGamma(a);
  if (x < a + 1) {
    let term = 1 / a;
    let sum = term;
    for (let n = 1; n <= 100000; n++) {
      term *= x / (a + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * EPSILON) {
        break;
      }
    }
    const lower = sum * Math.exp(logFront);
    return upper ? 1 - lower : lower;
  }
  let b = x + 1 - a;
  let c = 1 / TINY;
  let d = 1 / b;
  let f = d;
  for (let i = 1; i <= 100000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    d = 1 / (Math.abs(d) < TINY ? TINY : d);
    c = b + an / c;
    c = Math.abs(c) < TINY ? TINY : c;
    const delta = d * c;
    f *= delta;
    if (Math.abs(delta - 1) < EPSILON) {
      break;
    }
  }
  const complement = Math.exp(logFront) * f;
  return upper ? complement : 1 - complement;
}

/** The standard normal distribution below `x` (`upper` false) or above it. */
export function normalTail(x: number, upper = false): number {
  // Φ(x) = Q(1/2, x²/2) / 2 for x ≤ 0, which keeps the far tails exact.
  const half = 0.5 * regularizedGamma(0.5, (x * x) / 2, true);
  const below = x <= 0;
  return below !== upper ? half : 1 - half;
}

/**
 * The beta arguments of Student's t at `t` ≥ 0: x = v / (v + t²) and y = t² / (v + t²), written through r = √v / t so a
 * t past 1e154 does not square to infinity.
 */
function studentBetaArgs(t: number, v: number): [x: number, y: number] {
  if (t === 0) {
    return [1, 0];
  }
  const r = Math.sqrt(v) / t;
  const r2 = r * r;
  return [r2 / (1 + r2), 1 / (1 + r2)];
}

/** Student's t above `t` ≥ 0 with `v` degrees of freedom. */
export function studentUpperTail(t: number, v: number): number {
  if (v === 1) {
    // Cauchy, in closed form: past t = 1e154 the beta argument below underflows.
    return Math.atan2(1, t) / Math.PI;
  }
  const [x, y] = studentBetaArgs(t, v);
  return 0.5 * regularizedBeta(x, y, v / 2, 0.5);
}

/** The probability that Student's t lies within ±`t`, `t` ≥ 0 — exact where the upper tail is close to 1/2. */
export function studentCentral(t: number, v: number): number {
  if (v === 1) {
    return (2 * Math.atan(t)) / Math.PI;
  }
  const [x, y] = studentBetaArgs(t, v);
  return regularizedBeta(y, x, 0.5, v / 2);
}

/**
 * The x at which a monotonically increasing `f` reaches `target`, between `lo` and `hi` — or, for a decreasing `f`, the
 * same with `decreasing`. Bisection, geometric while the bracket spans orders of magnitude, until the bracket is as
 * narrow as a double allows: slower than a Newton step and never wrong about which side of the answer it is on.
 * An answer beyond the largest double is `Infinity`, which the caller reports as `#NUM!`.
 */
export function solveMonotone(
  f: (x: number) => number,
  target: number,
  lo: number,
  hi: number,
  decreasing = false
): number {
  const below = (x: number): boolean => (decreasing ? f(x) > target : f(x) < target);
  // Widen an open upper end until it holds the answer.
  while (below(hi)) {
    if (hi > Number.MAX_VALUE / 16) {
      return Infinity;
    }
    lo = hi;
    hi *= 16;
  }
  for (let i = 0; i < 2000; i++) {
    const mid =
      lo > 0 && hi / lo > 4
        ? Math.sqrt(lo) * Math.sqrt(hi)
        : lo === 0 && hi > 1
          ? hi / 16
          : lo + (hi - lo) / 2;
    if (mid <= lo || mid >= hi) {
      break;
    }
    if (below(mid)) {
      lo = mid;
    } else {
      hi = mid;
    }
    if (hi - lo <= 2e-16 * Math.max(Math.abs(lo), Math.abs(hi))) {
      break;
    }
  }
  return lo + (hi - lo) / 2;
}
