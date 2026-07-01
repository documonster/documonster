/**
 * Generic object utilities — deep structural equality and deep merge.
 *
 * No native equivalent exists: `JSON.stringify` comparison is key-order
 * sensitive and mishandles `undefined`/`NaN`/`Date`; `{...a, ...b}` is a
 * shallow merge. These are intentionally small, dependency-free helpers.
 */

/**
 * Property names that must never be assigned from untrusted input, to avoid
 * prototype pollution. Covers the direct pollution vectors
 * (`__proto__`/`constructor`/`prototype`) plus the legacy `Object.prototype`
 * accessor-defining methods (`__defineGetter__` etc.), which are dangerous
 * keys on prototyped objects and never legitimate data keys.
 */
const FORBIDDEN_KEYS: ReadonlySet<string> = new Set([
  "__proto__",
  "constructor",
  "prototype",
  "__defineGetter__",
  "__defineSetter__",
  "__lookupGetter__",
  "__lookupSetter__"
]);

/**
 * True for object keys that must never be assigned from untrusted input, to
 * avoid prototype pollution. Use this to guard every dynamic
 * `obj[key] = value` where `key` derives from external data (parsed files,
 * worker messages, user config) and the key is copied from an existing object.
 *
 * For keys built from raw external strings (CSV headers, worker configs),
 * prefer {@link isSafeDynamicKey}, which additionally bounds length and
 * rejects control characters.
 */
export function isForbiddenKey(key: string): boolean {
  return FORBIDDEN_KEYS.has(key);
}

/**
 * Maximum accepted length for a dynamic object key derived from untrusted
 * input. Real column names / aliases are short; anything longer is not a
 * legitimate key and is rejected to bound memory and avoid abuse.
 */
export const MAX_DYNAMIC_KEY_LENGTH = 512;

/**
 * True when `key` is safe to use as a dynamic object property assigned from
 * untrusted input (parsed files, worker messages, user config). Rejects:
 *
 * - forbidden prototype-pollution keys (see {@link isForbiddenKey}),
 * - keys longer than {@link MAX_DYNAMIC_KEY_LENGTH},
 * - keys containing control characters (U+0000–U+001F, U+007F), which are
 *   never valid header/alias names and can corrupt downstream serialization.
 *
 * The empty string is allowed (empty CSV headers are legitimate and harmless
 * as keys). Unicode letters, digits, spaces and ordinary punctuation are
 * allowed, so column names like `"First Name"` or `"价格"` pass unchanged.
 */
export function isSafeDynamicKey(key: string): boolean {
  if (isForbiddenKey(key)) {
    return false;
  }
  if (key.length > MAX_DYNAMIC_KEY_LENGTH) {
    return false;
  }
  for (let i = 0; i < key.length; i++) {
    const code = key.charCodeAt(i);
    if (code <= 0x1f || code === 0x7f) {
      return false;
    }
  }
  return true;
}

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
      const tgt = target as Record<number, unknown>;
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
        if (isForbiddenKey(key)) {
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
