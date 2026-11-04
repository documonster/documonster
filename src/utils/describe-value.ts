/**
 * How an argument error shows the value it refused: `must be <expected>, got <description>`.
 *
 * The description has to tell the refused value apart from what was wanted, or the message contradicts itself. So a
 * built-in object is named by its kind (`a Map`, `a Uint8Array`, `an ArrayBuffer`) rather than "an object", and a
 * `Date` that holds no time is `an invalid Date` — "must be a valid Date, got a Date" reads as a contradiction. A
 * primitive is shown as its value; a string is quoted and cut short, since a rejected document (base64, XML) can be
 * megabytes long.
 */

/** The longest string shown whole; a longer one is cut here and marked with `…`. */
const MAX_STRING = 60;

/** A short, readable rendering of a rejected value, for `<field> must be <expected>, got <this>`. */
export function describeValue(value: unknown): string {
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint" ||
    value === undefined ||
    value === null
  ) {
    return typeof value === "bigint" ? `${value}n` : String(value);
  }
  if (typeof value === "string") {
    return JSON.stringify(value.length > MAX_STRING ? `${value.slice(0, MAX_STRING)}…` : value);
  }
  if (typeof value === "symbol") {
    return value.toString();
  }
  if (typeof value === "function") {
    return "a function";
  }
  if (Array.isArray(value)) {
    return "an array";
  }
  // A built-in object is named by its tag — which a value from another realm carries too.
  const tag = Object.prototype.toString.call(value).slice(8, -1);
  if (tag === "Date" && Number.isNaN((value as Date).getTime())) {
    return "an invalid Date";
  }
  if (tag === "Object") {
    return "an object";
  }
  // "an ArrayBuffer", "an Int8Array", but "a Uint8Array".
  return /^[AEIO]/.test(tag) ? `an ${tag}` : `a ${tag}`;
}
