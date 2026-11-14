/**
 * Hexadecimal (base16) with one reading on every platform, as `Uint8Array.fromHex` and
 * `Uint8Array.prototype.toHex` (ES2026) define it — and through them where the runtime has them.
 * They are absent from Node 22 and 24 and from browsers older than Chrome 140, Firefox 133 and
 * Safari 18.2.
 *
 * - Decoding takes pairs of `0-9`, `a-f` and `A-F`, nothing else: no white space, no `0x`, no
 *   odd final digit.
 * - Encoding writes lower case.
 */

/** `Uint8Array.fromHex` and `Uint8Array.prototype.toHex`; not yet in TypeScript's `lib`. */
interface NativeDecoder {
  readonly fromHex?: (this: typeof Uint8Array, text: string) => Uint8Array<ArrayBuffer>;
}
interface NativeEncoder {
  readonly toHex?: (this: Uint8Array) => string;
}

/** Hex digits only; the even length is checked apart, so no engine has a group to repeat. */
const HEX_DIGITS = /^[0-9a-fA-F]*$/;

/** The ASCII code of each lower-case hex digit, by value. */
const DIGIT_CODE = /* @__PURE__ */ new TextEncoder().encode("0123456789abcdef");

const asciiDecoder = /* @__PURE__ */ new TextDecoder();

/** The value of a hex digit by char code, or -1. */
const DIGIT_VALUE = /* @__PURE__ */ (() => {
  const table = new Int8Array(128).fill(-1);
  for (let i = 0; i < 16; i++) {
    table["0123456789abcdef".charCodeAt(i)] = i;
    table["0123456789ABCDEF".charCodeAt(i)] = i;
  }
  return table;
})();

/** Whether `text` is hex by the rules above, without decoding it. */
export function isHex(text: string): boolean {
  return text.length % 2 === 0 && HEX_DIGITS.test(text);
}

/** The bytes `text` encodes, or `undefined` when it is not hex by the rules above. */
export function decodeHex(text: string): Uint8Array<ArrayBuffer> | undefined {
  if (!isHex(text)) {
    return undefined;
  }
  const native = (Uint8Array as NativeDecoder).fromHex;
  if (typeof native === "function") {
    return native.call(Uint8Array, text);
  }
  const bytes = new Uint8Array(text.length / 2);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] =
      (DIGIT_VALUE[text.charCodeAt(2 * i)]! << 4) | DIGIT_VALUE[text.charCodeAt(2 * i + 1)]!;
  }
  return bytes;
}

/** `bytes` as lower-case hex. */
export function encodeHex(bytes: Uint8Array): string {
  const native = (bytes as NativeEncoder).toHex;
  if (typeof native === "function") {
    return native.call(bytes);
  }
  // The digits as ASCII bytes, decoded once: measured at 35 ms for 16 MB against 288 ms for a
  // string per byte joined, which also held one string object per input byte
  const ascii = new Uint8Array(bytes.length * 2);
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i]!;
    ascii[2 * i] = DIGIT_CODE[b >> 4]!;
    ascii[2 * i + 1] = DIGIT_CODE[b & 15]!;
  }
  return asciiDecoder.decode(ascii);
}
