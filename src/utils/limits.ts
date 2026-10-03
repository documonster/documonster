/**
 * Validation for size-limit options (`maxOutputLength`, `maxEntrySize`,
 * `maxDecodedBytes`, …), applied once where a public option enters.
 *
 * A limit is a byte or item count: `undefined` selects the default, any
 * non-negative number — `Infinity` included, meaning "unbounded" — is used as
 * given. NaN and negatives used to fall through to whatever comparison each
 * code path happened to make (a NaN bound disables a `>` check; a negative one
 * was clamped on one path and rejected on another), so they are rejected here.
 */

/** Throw a `RangeError` unless `value` is `undefined` or a non-negative number. */
export function assertLimitOption(name: string, value: unknown): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "number" || Number.isNaN(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative number or Infinity, got ${String(value)}`);
  }
}
