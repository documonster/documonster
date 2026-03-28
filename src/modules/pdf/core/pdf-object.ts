/**
 * Low-level PDF object primitives.
 *
 * PDF is built from a small set of object types: booleans, numbers, strings,
 * names, arrays, dictionaries, streams, and indirect object references.
 * This module provides serialization for all of them.
 *
 * @see PDF Reference 1.7, Chapter 3 - Objects
 */

// =============================================================================
// PDF Object Serialization
// =============================================================================

/**
 * Escape a string for PDF string literal (parentheses-delimited).
 * Escapes backslash, parentheses, and non-printable characters.
 */
export function pdfString(value: string): string {
  if (!isAsciiPdfString(value)) {
    return pdfHexString(encodePdfUtf16String(value));
  }
  const escaped = value
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/\r\n/g, "\\n")
    .replace(/\r/g, "\\n")
    .replace(/\n/g, "\\n");
  return `(${escaped})`;
}

/**
 * Format a PDF hex string from raw bytes.
 */
export function pdfHexString(bytes: Uint8Array): string {
  let hex = "<";
  for (let i = 0; i < bytes.length; i++) {
    hex += bytes[i].toString(16).padStart(2, "0");
  }
  return hex + ">";
}

/**
 * Format a PDF name object. Names are prefixed with /.
 * Characters outside the printable ASCII range are encoded as #XX.
 */
export function pdfName(name: string): string {
  let result = "/";
  for (let i = 0; i < name.length; i++) {
    const code = name.charCodeAt(i);
    if (
      code < 0x21 ||
      code > 0x7e ||
      code === 0x23 || // #
      code === 0x28 || // (
      code === 0x29 || // )
      code === 0x3c || // <
      code === 0x3e || // >
      code === 0x5b || // [
      code === 0x5d || // ]
      code === 0x7b || // {
      code === 0x7d || // }
      code === 0x2f // /
    ) {
      result += "#" + code.toString(16).padStart(2, "0");
    } else {
      result += name[i];
    }
  }
  return result;
}

/**
 * Format a PDF number. Integers are output without decimal point.
 * Floats are rounded to 4 decimal places to avoid floating-point noise.
 */
export function pdfNumber(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }
  if (Number.isInteger(value)) {
    return value.toString();
  }
  // Round to 4 decimal places to avoid floating point artifacts
  const rounded = Math.round(value * 10000) / 10000;
  return rounded.toString();
}

/**
 * Format a PDF boolean.
 */
export function pdfBoolean(value: boolean): string {
  return value ? "true" : "false";
}

/**
 * Format a PDF array from pre-serialized elements.
 */
export function pdfArray(elements: string[]): string {
  return "[" + elements.join(" ") + "]";
}

/**
 * Format a PDF indirect object reference.
 */
export function pdfRef(objectNumber: number, generation = 0): string {
  return `${objectNumber} ${generation} R`;
}

/**
 * Format a PDF date string conforming to the PDF date format.
 * Format: D:YYYYMMDDHHmmSSOHH'mm
 */
export function pdfDate(date: Date): string {
  const year = date.getUTCFullYear().toString().padStart(4, "0");
  const month = (date.getUTCMonth() + 1).toString().padStart(2, "0");
  const day = date.getUTCDate().toString().padStart(2, "0");
  const hours = date.getUTCHours().toString().padStart(2, "0");
  const minutes = date.getUTCMinutes().toString().padStart(2, "0");
  const seconds = date.getUTCSeconds().toString().padStart(2, "0");
  return `(D:${year}${month}${day}${hours}${minutes}${seconds}Z)`;
}

function isAsciiPdfString(value: string): boolean {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code > 0x7f) {
      return false;
    }
  }
  return true;
}

function encodePdfUtf16String(value: string): Uint8Array {
  const bytes = new Uint8Array(2 + value.length * 2);
  bytes[0] = 0xfe;
  bytes[1] = 0xff;
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    bytes[2 + i * 2] = (code >> 8) & 0xff;
    bytes[3 + i * 2] = code & 0xff;
  }
  return bytes;
}

// =============================================================================
// PDF Dictionary Builder
// =============================================================================

/**
 * Builds a PDF dictionary object from key-value pairs.
 * Values are already-serialized PDF strings.
 */
export class PdfDict {
  private entries: Array<[string, string]> = [];

  /**
   * Set a dictionary entry. The key should NOT include the leading /.
   * The value should be a pre-serialized PDF value string.
   */
  set(key: string, value: string): this {
    const idx = this.entries.findIndex(([k]) => k === key);
    if (idx >= 0) {
      this.entries[idx] = [key, value];
    } else {
      this.entries.push([key, value]);
    }
    return this;
  }

  /**
   * Conditionally set a dictionary entry.
   */
  setIf(condition: boolean, key: string, value: string): this {
    if (condition) {
      this.set(key, value);
    }
    return this;
  }

  /**
   * Serialize to a PDF dictionary string.
   */
  toString(): string {
    const parts = ["<<"];
    for (const [key, value] of this.entries) {
      parts.push(`${pdfName(key)} ${value}`);
    }
    parts.push(">>");
    return parts.join("\n");
  }
}
