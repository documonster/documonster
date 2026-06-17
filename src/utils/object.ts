/**
 * Generic object utilities — deep structural equality and deep merge.
 *
 * No native equivalent exists: `JSON.stringify` comparison is key-order
 * sensitive and mishandles `undefined`/`NaN`/`Date`; `{...a, ...b}` is a
 * shallow merge. These are intentionally small, dependency-free helpers.
 */

/**
 * Deep structural equality for plain JSON-like values (primitives, arrays,
 * plain objects). Reference-equal short-circuits first.
 *
 * Note: compares `Date`/`Map`/`Set`/`RegExp` by enumerable own keys (i.e. not
 * specially), and `NaN !== NaN` — sufficient for the plain style/config
 * objects it is used on, not a general structural-clone-grade comparator.
 */
export function deepEqual(a: unknown, b: unknown): boolean {
  // Fast path: identical references or primitives
  if (a === b) {
    return true;
  }
  if (a == null || b == null) {
    return a === b;
  }

  const aType = typeof a;
  if (aType !== typeof b) {
    return false;
  }
  // Primitives already handled by ===
  if (aType !== "object") {
    return false;
  }

  const aIsArray = Array.isArray(a);
  const bIsArray = Array.isArray(b);
  if (aIsArray !== bIsArray) {
    return false;
  }

  if (aIsArray) {
    const arrA = a as unknown[];
    const arrB = b as unknown[];
    const len = arrA.length;
    if (len !== arrB.length) {
      return false;
    }
    for (let i = 0; i < len; i++) {
      if (!deepEqual(arrA[i], arrB[i])) {
        return false;
      }
    }
    return true;
  }

  const objA = a as Record<string, unknown>;
  const objB = b as Record<string, unknown>;
  const aKeys = Object.keys(objA);
  const bKeys = Object.keys(objB);
  if (aKeys.length !== bKeys.length) {
    return false;
  }
  for (let i = 0, len = aKeys.length; i < len; i++) {
    const key = aKeys[i];
    if (!Object.prototype.hasOwnProperty.call(objB, key)) {
      return false;
    }
    if (!deepEqual(objA[key], objB[key])) {
      return false;
    }
  }
  return true;
}

function isPlainObject(val: unknown): val is Record<string, unknown> {
  return val !== null && typeof val === "object" && !Array.isArray(val);
}

/**
 * Deep-merge `sources` into the first argument and return it. Arrays and plain
 * objects merge recursively; `undefined` values are skipped; prototype-polluting
 * keys (`__proto__`/`constructor`/`prototype`) are ignored.
 */
export function deepMerge<T = Record<string, unknown>>(...args: unknown[]): T {
  const target = (args[0] as Record<string, unknown>) || {};
  const len = args.length;

  for (let i = 1; i < len; i++) {
    const arg = args[i];
    if (!arg) {
      continue;
    }

    if (Array.isArray(arg)) {
      const tgt = target as unknown as unknown[];
      for (let j = 0, jLen = arg.length; j < jLen; j++) {
        const val = arg[j];
        if (val === undefined) {
          continue;
        }
        const src = tgt[j];
        if (Array.isArray(val)) {
          tgt[j] = deepMerge(Array.isArray(src) ? src : [], val);
        } else if (isPlainObject(val)) {
          tgt[j] = deepMerge(isPlainObject(src) ? src : {}, val);
        } else {
          tgt[j] = val;
        }
      }
    } else {
      const obj = arg as Record<string, unknown>;
      const keys = Object.keys(obj);
      for (let j = 0, jLen = keys.length; j < jLen; j++) {
        const key = keys[j];
        if (key === "__proto__" || key === "constructor" || key === "prototype") {
          continue;
        }
        const val = obj[key];
        if (val === undefined) {
          continue;
        }
        const src = target[key];
        if (Array.isArray(val)) {
          target[key] = deepMerge(Array.isArray(src) ? src : [], val);
        } else if (isPlainObject(val)) {
          target[key] = deepMerge(isPlainObject(src) ? src : {}, val);
        } else {
          target[key] = val;
        }
      }
    }
  }
  return target as T;
}
