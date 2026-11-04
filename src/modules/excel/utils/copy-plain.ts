/**
 * A deep copy of plain model data — arrays, plain objects and `Date`s — for the readers that hand out a snapshot.
 * Anything else (a typed array, a Temporal value, a class instance) is immutable here or owned by the caller, and is
 * shared rather than copied, which also keeps a value `structuredClone` would refuse (a function inside a JSON cell's
 * raw value) from making a read throw.
 */
export function copyPlain<T>(value: T): T {
  if (value instanceof Date) {
    return new Date(value.getTime()) as T;
  }
  if (Array.isArray(value)) {
    return value.map(copyPlain) as T;
  }
  if (typeof value === "object" && value !== null) {
    const proto = Object.getPrototypeOf(value) as unknown;
    if (proto === Object.prototype || proto === null) {
      const out: Record<string, unknown> = proto === null ? Object.create(null) : {};
      for (const [key, item] of Object.entries(value)) {
        out[key] = copyPlain(item);
      }
      return out as T;
    }
  }
  return value;
}
