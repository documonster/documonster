/**
 * Tiny binary reader for Uint8Array-backed DataView.
 * Shared by ZIP parsers.
 */

import { decodeCp437 } from "@archive/core/text";
import { uint8ArrayToString as decodeUtf8 } from "@utils/binary";

export function writeUint32LE(value: number): Uint8Array {
  const out = new Uint8Array(4);
  const view = new DataView(out.buffer);
  view.setUint32(0, value >>> 0, true);
  return out;
}

export function readUint32LE(data: Uint8Array, offset: number): number {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return view.getUint32(offset, true);
}

export class BinaryReader {
  private view: DataView;
  private offset: number;
  private data: Uint8Array;

  constructor(data: Uint8Array, offset = 0) {
    this.data = data;
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.offset = offset;
  }

  get position(): number {
    return this.offset;
  }

  set position(value: number) {
    this.offset = value;
  }

  get remaining(): number {
    return this.data.length - this.offset;
  }

  readUint8(): number {
    const value = this.view.getUint8(this.offset);
    this.offset += 1;
    return value;
  }

  readUint16(): number {
    const value = this.view.getUint16(this.offset, true);
    this.offset += 2;
    return value;
  }

  readUint32(): number {
    const value = this.view.getUint32(this.offset, true);
    this.offset += 4;
    return value;
  }

  readBigUint64(): bigint {
    const value = this.view.getBigUint64(this.offset, true);
    this.offset += 8;
    return value;
  }

  readBytes(length: number): Uint8Array {
    const bytes = this.data.subarray(this.offset, this.offset + length);
    this.offset += length;
    return bytes;
  }

  readString(length: number, utf8 = true): string {
    const bytes = this.readBytes(length);
    return utf8 ? decodeUtf8(bytes) : decodeCp437(bytes);
  }

  skip(length: number): void {
    this.offset += length;
  }

  slice(start: number, end: number): Uint8Array {
    return this.data.subarray(start, end);
  }

  peekUint32(offset: number): number {
    return this.view.getUint32(offset, true);
  }
}

// =============================================================================
// Format-based parsing (legacy-style declarative parser)
// =============================================================================

/**
 * Parses sequential unsigned little endian numbers from the head of the passed buffer according to
 * the specified format passed. If the buffer is not large enough to satisfy the full format,
 * null values will be assigned to the remaining keys.
 * @param buffer The buffer to sequentially extract numbers from.
 * @param format Expected format to follow when extracting values from the buffer. A list of list entries
 * with the following structure:
 * [
 *   [
 *     <key>,  // Name of the key to assign the extracted number to.
 *     <size>  // The size in bytes of the number to extract. possible values are 1, 2, 4, 8.
 *   ],
 *   ...
 * ]
 * @returns An object with keys set to their associated extracted values.
 */
export function parseFormatted(
  buffer: Uint8Array,
  format: [string, number][]
): Record<string, number | null> {
  const result: Record<string, number | null> = {};
  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  let offset = 0;
  for (const [key, size] of format) {
    if (buffer.length >= offset + size) {
      switch (size) {
        case 1:
          result[key] = view.getUint8(offset);
          break;
        case 2:
          result[key] = view.getUint16(offset, true);
          break;
        case 4:
          result[key] = view.getUint32(offset, true);
          break;
        case 8: {
          // Keep behavior (Number) while avoiding BigInt costs.
          const low = view.getUint32(offset, true);
          const high = view.getUint32(offset + 4, true);
          result[key] = high * 0x100000000 + low;
          break;
        }
        default:
          throw new Error("Unsupported UInt LE size!");
      }
    } else {
      result[key] = null;
    }
    offset += size;
  }
  return result;
}

export function parseFormattedTyped<T>(buffer: Uint8Array, format: [string, number][]): T {
  return parseFormatted(buffer, format) as T;
}
