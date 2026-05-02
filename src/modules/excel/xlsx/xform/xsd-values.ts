/**
 * Shared parsers for OOXML / XSD-typed attribute values.
 *
 * These helpers are the single source of truth for attribute parsing
 * across every xform. Multiple xforms used to carry private copies
 * that drifted on edge cases:
 *   - chartsheet-xform read `tabSelected="true"` as `false` because it
 *     only accepted `"1"` (spec permits `"true"` per `xsd:boolean`);
 *   - some older code accepted `NaN` from `parseFloat` and serialised
 *     it back as the literal `"NaN"` attribute — an XML Schema
 *     violation that Excel's strict reader rejects.
 *
 * All parsers return `undefined` for missing / unparseable input so
 * callers distinguish "absent" from "present but invalid" and can
 * pipe through `?? default` to apply the spec's default value at the
 * call site (element-specific, so the parser shouldn't bake one in).
 */

/**
 * Parse an OOXML boolean attribute (`xsd:boolean`).
 *
 * Accepts the four canonical forms — `"1"`, `"0"`, `"true"`, `"false"` —
 * per the XSD 1.1 spec. Leading / trailing whitespace is stripped
 * first because XML 1.0 §3.3.3 attribute-value normalisation isn't
 * guaranteed by every parser, and third-party writers legitimately
 * emit `" 1 "` / `"0\n"` / etc. Without the trim, those values would
 * silently downgrade to the schema default (typically `false`),
 * losing state on round-trip.
 *
 * Returns `undefined` for any other value so call sites can
 * distinguish "attribute absent / unrecognised" from "explicit false"
 * and apply the element's schema default via `?? true` / `?? false`.
 */
export function parseXsdBoolean(value: string | undefined): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "1" || trimmed === "true") {
    return true;
  }
  if (trimmed === "0" || trimmed === "false") {
    return false;
  }
  return undefined;
}

/**
 * Parse an OOXML integer attribute (`xsd:int`, bounded at ±2^31-1 per
 * ECMA-376 / W3C XSD).
 *
 * Rejects trailing garbage — `parseInt("100%", 10)` silently returns
 * `100`, which round-trips back to the wire as a valid-looking
 * attribute even though the source was corrupt. Gate on a strict
 * regex so only the lexical space defined by XSD is accepted; the
 * surrounding trim handles the same third-party-writer whitespace
 * cases {@link parseXsdBoolean} does.
 *
 * Returns `undefined` for missing / empty / non-numeric input rather
 * than producing `NaN` (which would propagate through the model and
 * surface as `"NaN"` in the round-trip XML — rejected by strict
 * readers). Values outside `xsd:int` range are also rejected; callers
 * that genuinely need 64-bit integers should use a `BigInt` path.
 */
export function parseXsdInt(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "" || !/^[+-]?\d+$/.test(trimmed)) {
    return undefined;
  }
  const n = parseInt(trimmed, 10);
  if (!Number.isFinite(n)) {
    return undefined;
  }
  if (n > 2147483647 || n < -2147483648) {
    return undefined;
  }
  return n;
}

/**
 * Parse an OOXML floating-point attribute (`xsd:double` / `xsd:float`).
 * See {@link parseXsdInt} for the rationale — returns `undefined`
 * instead of `NaN` / `Infinity` so the value never escapes back into
 * the wire as a schema-invalid attribute.
 *
 * Strict lexical-space regex matches `xsd:double` production:
 * optional sign, integer / decimal, optional exponent. Rejects
 * trailing garbage for the same reason as {@link parseXsdInt}
 * (corrupt-source preservation is worse than loud rejection).
 */
export function parseXsdFloat(value: string | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "" || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    return undefined;
  }
  const n = parseFloat(trimmed);
  return Number.isFinite(n) ? n : undefined;
}
