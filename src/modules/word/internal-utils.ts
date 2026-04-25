/**
 * DOCX Module - Internal utilities (shared helpers)
 */

/**
 * Encode a byte array to base64. Uses native Node `Buffer` if available;
 * otherwise falls back to chunked `btoa()` for browser environments.
 *
 * The chunked approach avoids the O(n²) string concatenation cost that naïve
 * `String.fromCharCode(...data)` implementations exhibit on large buffers.
 */
export function bytesToBase64(data: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }
  const CHUNK = 0x8000;
  const parts: string[] = [];
  for (let i = 0; i < data.length; i += CHUNK) {
    const slice = data.subarray(i, i + CHUNK);
    // `apply` on Uint8Array works in all modern JS engines
    parts.push(String.fromCharCode.apply(null, slice as unknown as number[]));
  }
  return btoa(parts.join(""));
}

/**
 * Decode a base64 string to a byte array. Uses native Node `Buffer` if available.
 */
export function base64ToBytes(s: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(s, "base64"));
  }
  const binary = atob(s);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Build an XML attributes object from an array of [name, value] entries,
 * omitting entries whose value is `undefined` or `false`. Numeric and boolean
 * values are stringified: `number → String(n)`, `boolean → "1"`.
 *
 * Reduces repetitive `if (x !== undefined) attrs["w:x"] = String(x)` patterns.
 */
export function buildAttrs(
  entries: ReadonlyArray<[string, string | number | boolean | undefined | null]>
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of entries) {
    if (v === undefined || v === null || v === false) {
      continue;
    }
    out[k] = typeof v === "boolean" ? "1" : String(v);
  }
  return out;
}
