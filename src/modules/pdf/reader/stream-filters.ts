/**
 * PDF stream filter decoder chain.
 *
 * Decodes PDF stream data by applying the appropriate filter(s)
 * specified in the stream dictionary's /Filter entry.
 *
 * Supported filters:
 * - /FlateDecode (zlib/deflate compression)
 * - /ASCII85Decode (ASCII base-85 encoding)
 * - /ASCIIHexDecode (ASCII hexadecimal encoding)
 * - /LZWDecode (LZW compression)
 * - /RunLengthDecode (run-length encoding)
 *
 * @see PDF Reference 1.7, §3.3 - Filters
 */

import { unzlibSync } from "@archive/compression/compress";
import { inflateRaw } from "@archive/compression/deflate-fallback";
import { ArchiveLimitError } from "@archive/core/errors";
import { undoPngFilters } from "@pdf/core/png-filters";
import { PdfLimitExceededError } from "@pdf/errors";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";
import { dictGetNumber, isPdfDict, isPdfArray } from "@pdf/reader/pdf-parser";

// =============================================================================
// Public API
// =============================================================================

/**
 * Default ceiling on the decoded size of a single stream: 256 MiB.
 *
 * Large enough for any realistic page content or image (a 8000×8000 RGBA
 * raster is 244 MiB), small enough that a few-KB "zip bomb" stream cannot
 * exhaust the heap of a browser tab or a default Node process.
 */
const DEFAULT_MAX_DECODED_BYTES = 256 * 1024 * 1024;

/**
 * Ceiling on the number of filters in one /Filter array.
 *
 * Real producers use one or two (e.g. `[/ASCII85Decode /FlateDecode]`);
 * a long chain only serves to multiply decode work.
 *
 * This bounds repeated work (CPU), not memory. The byte limit applies to each
 * filter's output separately, so it alone would let a stream run hundreds of
 * passes that each stay under it — e.g. an ASCIIHex/ASCII85 chain re-expanding
 * the data at every step. The two limits are independent and neither covers
 * the other.
 */
const MAX_FILTER_CHAIN = 16;

/**
 * Decode stream data by applying the filter chain from the stream dictionary.
 *
 * @param maxBytes - Most bytes any single filter in the chain may produce.
 * @throws PdfLimitExceededError when a filter's output exceeds `maxBytes`
 * or the chain is longer than {@link MAX_FILTER_CHAIN}.
 */
export function decodeStreamFilters(
  data: Uint8Array,
  dict: PdfDictValue,
  maxBytes = DEFAULT_MAX_DECODED_BYTES
): Uint8Array {
  const filter = dict.get("Filter");
  if (filter === undefined || filter === null) {
    return data;
  }

  const decodeParms = dict.get("DecodeParms") ?? dict.get("DP");

  if (typeof filter === "string") {
    // Single filter
    const parms = isPdfDict(decodeParms) ? decodeParms : undefined;
    return applyFilter(data, filter, parms, maxBytes);
  }

  if (isPdfArray(filter)) {
    if (filter.length > MAX_FILTER_CHAIN) {
      throw new PdfLimitExceededError(
        `Stream filter chain has ${filter.length} filters; the limit is ${MAX_FILTER_CHAIN}`
      );
    }
    // Filter chain — apply in order
    let result = data;
    const parmsArray = isPdfArray(decodeParms) ? decodeParms : [];
    for (let i = 0; i < filter.length; i++) {
      const filterName = filter[i] as string;
      const parm = parmsArray[i];
      const parmDict = isPdfDict(parm) ? parm : undefined;
      result = applyFilter(result, filterName, parmDict, maxBytes);
    }
    return result;
  }

  return data;
}

function decodedTooLarge(filter: string, maxBytes: number): PdfLimitExceededError {
  return new PdfLimitExceededError(
    `Decoded ${filter} stream exceeds the limit of ${maxBytes} bytes (maxDecodedBytes)`
  );
}

/**
 * Output buffer for a decoder whose output size is not known up front. It
 * grows geometrically, never past `maxBytes`, and throws the limit error
 * before a write would cross it — so a hostile stream costs at most
 * `maxBytes`, not the 8× of a boxed `number[]`.
 */
interface BoundedOutput {
  /** Reserve `count` bytes and return the offset to write them at. */
  reserve(count: number): number;
  /** The live buffer; re-read after every `reserve`, which may replace it. */
  readonly buffer: Uint8Array;
  /** The bytes written so far. */
  result(): Uint8Array;
}

function createBoundedOutput(filter: string, maxBytes: number, initial: number): BoundedOutput {
  let buffer = new Uint8Array(Math.min(Math.max(initial, 64), maxBytes));
  let length = 0;
  return {
    reserve(count: number): number {
      const required = length + count;
      if (required > maxBytes) {
        throw decodedTooLarge(filter, maxBytes);
      }
      if (required > buffer.length) {
        const grown = new Uint8Array(Math.min(Math.max(required, buffer.length * 2), maxBytes));
        grown.set(buffer.subarray(0, length));
        buffer = grown;
      }
      const at = length;
      length = required;
      return at;
    },
    get buffer() {
      return buffer;
    },
    result: () => buffer.subarray(0, length)
  };
}

// =============================================================================
// Filter Application
// =============================================================================

function applyFilter(
  data: Uint8Array,
  filterName: string,
  parms: PdfDictValue | undefined,
  maxBytes: number
): Uint8Array {
  switch (filterName) {
    case "FlateDecode":
    case "Fl":
      return decodeFlateDecode(data, parms, maxBytes);
    case "ASCII85Decode":
    case "A85":
      return decodeAscii85(data, maxBytes);
    case "ASCIIHexDecode":
    case "AHx":
      return decodeAsciiHex(data, maxBytes);
    case "LZWDecode":
    case "LZW":
      return decodeLzw(data, parms, maxBytes);
    case "RunLengthDecode":
    case "RL":
      return decodeRunLength(data, maxBytes);
    case "DCTDecode":
    case "DCT":
      // JPEG data — return as-is (used for image XObjects)
      return data;
    case "JPXDecode":
      // JPEG 2000 — return as-is
      return data;
    case "CCITTFaxDecode":
    case "CCF":
      // CCITT fax — return as-is (would need full CCITT decoder)
      return data;
    case "JBIG2Decode":
      // JBIG2 — return as-is
      return data;
    case "Crypt":
      // Handled by decryption layer — pass through
      return data;
    default:
      // Unknown filter — return as-is
      return data;
  }
}

// =============================================================================
// FlateDecode
// =============================================================================

function decodeFlateDecode(
  data: Uint8Array,
  parms: PdfDictValue | undefined,
  maxBytes: number
): Uint8Array {
  if (data.length === 0) {
    return data;
  }

  // The bound is handed to the inflater so a hostile stream stops as soon as
  // it crosses it, instead of being fully materialised and measured after.
  const bounded = { maxOutputLength: maxBytes };
  let decompressed: Uint8Array;
  try {
    // Try zlib (RFC 1950) first — has 2-byte header
    decompressed = unzlibSync(data, bounded);
  } catch (zlibErr) {
    if (zlibErr instanceof ArchiveLimitError) {
      throw decodedTooLarge("FlateDecode", maxBytes);
    }
    try {
      // Fall back to raw deflate
      decompressed = inflateRaw(data, maxBytes);
    } catch (rawErr) {
      if (rawErr instanceof ArchiveLimitError) {
        throw decodedTooLarge("FlateDecode", maxBytes);
      }
      // Last resort: return as-is
      return data;
    }
  }

  // Apply predictor if specified
  if (parms) {
    const predictor = dictGetNumber(parms, "Predictor") ?? 1;
    if (predictor > 1) {
      decompressed = undoPredictor(decompressed, parms);
    }
  }

  return decompressed;
}

/**
 * Undo PNG/TIFF predictors used in FlateDecode and LZWDecode streams.
 *
 * @see PDF Reference 1.7, Table 3.8
 */
function undoPredictor(data: Uint8Array, parms: PdfDictValue): Uint8Array {
  const predictor = dictGetNumber(parms, "Predictor") ?? 1;
  const columns = dictGetNumber(parms, "Columns") ?? 1;
  const colors = dictGetNumber(parms, "Colors") ?? 1;
  const bitsPerComponent = dictGetNumber(parms, "BitsPerComponent") ?? 8;

  if (predictor === 1) {
    return data; // No prediction
  }

  if (predictor === 2) {
    // TIFF predictor 2
    return undoTiffPredictor(data, columns, colors, bitsPerComponent);
  }

  if (predictor >= 10 && predictor <= 15) {
    // PNG predictors (10-15)
    return undoPngPredictor(data, columns, colors, bitsPerComponent);
  }

  return data;
}

/**
 * Undo TIFF Predictor 2 (horizontal differencing).
 */
function undoTiffPredictor(
  data: Uint8Array,
  columns: number,
  colors: number,
  bitsPerComponent: number
): Uint8Array {
  const bytesPerPixel = Math.ceil((colors * bitsPerComponent) / 8);
  const rowBytes = Math.ceil((columns * colors * bitsPerComponent) / 8);
  const rows = Math.floor(data.length / rowBytes);
  const result = new Uint8Array(data.length);

  for (let row = 0; row < rows; row++) {
    const rowStart = row * rowBytes;
    // First pixel is unmodified
    for (let i = 0; i < bytesPerPixel; i++) {
      result[rowStart + i] = data[rowStart + i];
    }
    // Subsequent pixels: add previous pixel
    for (let i = bytesPerPixel; i < rowBytes; i++) {
      result[rowStart + i] = (data[rowStart + i] + result[rowStart + i - bytesPerPixel]) & 0xff;
    }
  }

  return result;
}

/**
 * Undo PNG row filters.
 * Each row is preceded by a filter type byte.
 */
function undoPngPredictor(
  data: Uint8Array,
  columns: number,
  colors: number,
  bitsPerComponent: number
): Uint8Array {
  const bytesPerPixel = Math.max(1, Math.ceil((colors * bitsPerComponent) / 8));
  const rowBytes = Math.ceil((columns * colors * bitsPerComponent) / 8);
  const rows = Math.floor(data.length / (rowBytes + 1)); // +1 for the filter byte
  return undoPngFilters(data, rows, rowBytes, bytesPerPixel);
}

// =============================================================================
// ASCII85Decode
// =============================================================================

function isPdfWhitespace(b: number): boolean {
  return b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d || b === 0x0c || b === 0x00;
}

function decodeAscii85(data: Uint8Array, maxBytes: number): Uint8Array {
  // Every input byte yields at most 4 output bytes ('z' → 4 zeros; a 5-char
  // group → 4 bytes), so 4n bounds the output; the limit bounds it further.
  const output = new Uint8Array(Math.min(4 * data.length, maxBytes + 4));
  let outLen = 0;
  const group = [0, 0, 0, 0, 0];
  let i = 0;

  while (i < data.length) {
    const b = data[i];

    if (isPdfWhitespace(b)) {
      i++;
      continue;
    }

    // End of data marker ~>
    if (b === 0x7e) {
      break;
    }

    // Special 'z' character = four zero bytes (array is zero-initialised)
    if (b === 0x7a) {
      if (outLen + 4 > maxBytes) {
        throw decodedTooLarge("ASCII85Decode", maxBytes);
      }
      outLen += 4;
      i++;
      continue;
    }

    // Decode 5-character group into 4 bytes
    let count = 0;
    while (count < 5 && i < data.length) {
      const c = data[i];
      if (c === 0x7e) {
        break; // EOD
      }
      i++;
      if (c < 0x21 || c > 0x75) {
        continue; // Whitespace or invalid — skip
      }
      group[count++] = c - 0x21;
    }

    if (count === 0) {
      break;
    }

    // Pad short final group with 'u' (84) values
    const numBytes = count - 1;
    for (let k = count; k < 5; k++) {
      group[k] = 84;
    }

    if (outLen + numBytes > maxBytes) {
      throw decodedTooLarge("ASCII85Decode", maxBytes);
    }

    const value = (((group[0] * 85 + group[1]) * 85 + group[2]) * 85 + group[3]) * 85 + group[4];
    for (let j = 0; j < numBytes; j++) {
      output[outLen++] = (value >>> (24 - 8 * j)) & 0xff;
    }
  }

  return output.subarray(0, outLen);
}

// =============================================================================
// ASCIIHexDecode
// =============================================================================

function decodeAsciiHex(data: Uint8Array, maxBytes: number): Uint8Array {
  // Two hex digits per byte, plus one for an odd trailing digit: ceil(n/2).
  const output = new Uint8Array(Math.min(Math.ceil(data.length / 2), maxBytes + 1));
  let outLen = 0;
  let highNibble = -1;

  for (let i = 0; i < data.length; i++) {
    const b = data[i];

    // End of data marker >
    if (b === 0x3e) {
      break;
    }

    let val: number;
    if (b >= 0x30 && b <= 0x39) {
      val = b - 0x30;
    } else if (b >= 0x41 && b <= 0x46) {
      val = b - 0x41 + 10;
    } else if (b >= 0x61 && b <= 0x66) {
      val = b - 0x61 + 10;
    } else {
      continue; // Whitespace or invalid — skip
    }

    if (highNibble < 0) {
      highNibble = val;
    } else {
      if (outLen >= maxBytes) {
        throw decodedTooLarge("ASCIIHexDecode", maxBytes);
      }
      output[outLen++] = (highNibble << 4) | val;
      highNibble = -1;
    }
  }

  // Odd digit — pad with 0
  if (highNibble >= 0) {
    if (outLen >= maxBytes) {
      throw decodedTooLarge("ASCIIHexDecode", maxBytes);
    }
    output[outLen++] = highNibble << 4;
  }

  return output.subarray(0, outLen);
}

// =============================================================================
// LZWDecode
// =============================================================================

function decodeLzw(
  data: Uint8Array,
  parms: PdfDictValue | undefined,
  maxBytes: number
): Uint8Array {
  const earlyChange = parms ? (dictGetNumber(parms, "EarlyChange") ?? 1) : 1;
  const output = createBoundedOutput("LZWDecode", maxBytes, data.length * 4);

  // LZW bit reader
  let bitPos = 0;

  function readBits(n: number): number {
    let result = 0;
    for (let i = 0; i < n; i++) {
      const byteIdx = (bitPos + i) >> 3;
      const bitIdx = 7 - ((bitPos + i) & 7); // MSB first
      if (byteIdx < data.length) {
        result = (result << 1) | ((data[byteIdx] >> bitIdx) & 1);
      }
    }
    bitPos += n;
    return result;
  }

  const CLEAR_TABLE = 256;
  const EOD = 257;
  let codeSize = 9;
  let nextCode = 258;
  let table: Uint8Array[] = [];

  // Initialize table
  function resetTable(): void {
    table = [];
    for (let i = 0; i < 256; i++) {
      table[i] = new Uint8Array([i]);
    }
    table[CLEAR_TABLE] = new Uint8Array(0);
    table[EOD] = new Uint8Array(0);
    nextCode = 258;
    codeSize = 9;
  }

  resetTable();

  let prevEntry: Uint8Array | null = null;

  while (bitPos < data.length * 8) {
    const code = readBits(codeSize);

    if (code === EOD) {
      break;
    }

    if (code === CLEAR_TABLE) {
      resetTable();
      prevEntry = null;
      continue;
    }

    let entry: Uint8Array;
    if (code < nextCode && table[code]) {
      entry = table[code];
    } else if (code === nextCode && prevEntry) {
      // Special case: code not in table yet
      entry = new Uint8Array(prevEntry.length + 1);
      entry.set(prevEntry);
      entry[prevEntry.length] = prevEntry[0];
    } else {
      // Invalid code — bail
      break;
    }

    output.buffer.set(entry, output.reserve(entry.length));

    // Add new entry to table
    if (prevEntry !== null) {
      const newEntry = new Uint8Array(prevEntry.length + 1);
      newEntry.set(prevEntry);
      newEntry[prevEntry.length] = entry[0];
      table[nextCode] = newEntry;
      nextCode++;

      // Increase code size
      const threshold = earlyChange ? nextCode : nextCode + 1;
      if (threshold >= 1 << codeSize && codeSize < 12) {
        codeSize++;
      }
    }

    prevEntry = entry;
  }

  let result = output.result();

  // Apply predictor if specified
  if (parms) {
    const predictor = dictGetNumber(parms, "Predictor") ?? 1;
    if (predictor > 1) {
      result = undoPredictor(result, parms);
    }
  }

  return result;
}

// =============================================================================
// RunLengthDecode
// =============================================================================

function decodeRunLength(data: Uint8Array, maxBytes: number): Uint8Array {
  const output = createBoundedOutput("RunLengthDecode", maxBytes, data.length * 2);
  let i = 0;

  while (i < data.length) {
    const length = data[i];
    i++;

    if (length === 128) {
      // EOD
      break;
    }

    if (length < 128) {
      // Copy (length + 1) literal bytes
      // A truncated final run copies what is there
      const literal = data.subarray(i, i + length + 1);
      output.buffer.set(literal, output.reserve(literal.length));
      i += literal.length;
    } else {
      // Repeat next byte (257 - length) times
      if (i < data.length) {
        const count = 257 - length;
        const at = output.reserve(count);
        output.buffer.fill(data[i], at, at + count);
        i++;
      }
    }
  }

  return output.result();
}
