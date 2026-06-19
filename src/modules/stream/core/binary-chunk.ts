import { stringToUint8Array, textEncoder } from "@utils/binary";

// =============================================================================
// String-to-bytes with encoding (for Readable.push / unshift parity with Node.js)
// =============================================================================

/**
 * Encode a string into Uint8Array respecting the specified encoding.
 *
 * `TextEncoder` only supports UTF-8, so we need manual implementations for
 * other Node.js encodings (latin1, ascii, hex, base64, base64url).
 * This is used by the browser `Readable.push()` and `unshift()` to match
 * Node.js `Buffer.from(string, encoding)` behavior.
 */
export function stringToEncodedBytes(str: string, encoding?: string): Uint8Array {
  const enc = (encoding ?? "utf-8").trim().toLowerCase();
  switch (enc) {
    case "":
    case "utf8":
    case "utf-8":
      return textEncoder.encode(str);

    case "latin1":
    case "binary": {
      const bytes = new Uint8Array(str.length);
      for (let i = 0; i < str.length; i++) {
        bytes[i] = str.charCodeAt(i) & 0xff;
      }
      return bytes;
    }

    case "ascii": {
      const bytes = new Uint8Array(str.length);
      for (let i = 0; i < str.length; i++) {
        bytes[i] = str.charCodeAt(i) & 0x7f;
      }
      return bytes;
    }

    case "hex": {
      const len = str.length >>> 1;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = parseInt(str.substring(i * 2, i * 2 + 2), 16);
      }
      return bytes;
    }

    case "base64": {
      const binaryStr = atob(str);
      const bytes = new Uint8Array(binaryStr.length);
      for (let i = 0; i < binaryStr.length; i++) {
        bytes[i] = binaryStr.charCodeAt(i);
      }
      return bytes;
    }

    case "base64url": {
      // Convert base64url to standard base64
      let b64 = str.replace(/-/g, "+").replace(/_/g, "/");
      while (b64.length % 4 !== 0) {
        b64 += "=";
      }
      const binaryStr = atob(b64);
      const bytes = new Uint8Array(binaryStr.length);
      for (let i = 0; i < binaryStr.length; i++) {
        bytes[i] = binaryStr.charCodeAt(i);
      }
      return bytes;
    }

    case "utf16le":
    case "utf-16le":
    case "ucs2":
    case "ucs-2": {
      const bytes = new Uint8Array(str.length * 2);
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        bytes[i * 2] = code & 0xff;
        bytes[i * 2 + 1] = (code >>> 8) & 0xff;
      }
      return bytes;
    }

    default:
      // Match Node.js: unknown encoding throws ERR_UNKNOWN_ENCODING
      throw createUnknownEncodingError(enc);
  }
}

function createUnknownEncodingError(enc: string): Error & { code: string } {
  const err = new TypeError(`Unknown encoding: ${enc}`) as Error & { code: string };
  err.code = "ERR_UNKNOWN_ENCODING";
  return err;
}

// =============================================================================
// Binary chunk utilities
// =============================================================================

/**
 * Normalize a binary-like value into Uint8Array.
 */
export const toBinaryChunk = (value: unknown): Uint8Array | null => {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value);
  }
  if (ArrayBuffer.isView(value)) {
    const view = value as ArrayBufferView;
    return new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
  }
  return null;
};

/**
 * Convert any stream chunk to bytes for text decoding.
 * Handles: string, Uint8Array, ArrayBuffer, TypedArray, Array, array-like objects.
 * Returns null if the chunk type is not recognized.
 *
 * Shared by both Node.js and browser streamToString / streamToBuffer.
 */
export const toStreamBytes = (chunk: unknown): Uint8Array | null => {
  if (typeof chunk === "string") {
    return stringToUint8Array(chunk);
  }
  if (Array.isArray(chunk)) {
    return new Uint8Array(chunk);
  }
  const binary = toBinaryChunk(chunk);
  if (binary) {
    return binary;
  }
  return toArrayLikeBytes(chunk);
};

/**
 * Convert an array-like object (e.g. {0: 65, 1: 66, length: 2}) to Uint8Array.
 * Returns null if the value is not a valid array-like of numbers.
 */
const toArrayLikeBytes = (chunk: unknown): Uint8Array | null => {
  if (chunk == null || typeof chunk !== "object") {
    return null;
  }

  const lengthValue = (chunk as { length?: unknown }).length;
  if (
    typeof lengthValue !== "number" ||
    !Number.isFinite(lengthValue) ||
    lengthValue < 0 ||
    !Number.isInteger(lengthValue)
  ) {
    return null;
  }

  const result = new Uint8Array(lengthValue);
  const source = chunk as Record<number, unknown>;
  for (let index = 0; index < lengthValue; index++) {
    const value = source[index];
    if (typeof value !== "number" || !Number.isFinite(value)) {
      return null;
    }
    result[index] = value;
  }

  return result;
};
