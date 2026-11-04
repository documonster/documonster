/**
 * CRC32 calculation core (shared between Node.js and browser).
 *
 * Implements CRC-32 IEEE 802.3 using the reversed polynomial 0xEDB88320.
 *
 * Notes:
 * - The byte-wise table (256 entries) serves the single-byte update ZipCrypto needs.
 * - The bulk update is "slicing-by-16": sixteen derived tables let one step fold sixteen
 *   input bytes, which measured about 6× the byte-wise loop in V8 (≈1.4 GB/s against
 *   250 MB/s on 10 MiB). It is the browser's path; Node uses native `zlib.crc32`.
 * - Exposes an incremental update API for streaming use cases.
 */

let _crc32Table: Uint32Array | null = null;
let _sliceTables: Uint32Array | null = null;

function getCrc32Table(): Uint32Array {
  if (_crc32Table) {
    return _crc32Table;
  }

  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let crc = i;
    for (let j = 0; j < 8; j++) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[i] = crc;
  }

  _crc32Table = table;
  return table;
}

/**
 * Sixteen 256-entry tables, back to back: table `k` advances a byte by `k` further bytes of
 * zeros, so the byte `k` positions before the end of a 16-byte block is looked up in table `k`.
 */
function getSliceTables(): Uint32Array {
  if (_sliceTables) {
    return _sliceTables;
  }
  const base = getCrc32Table();
  const tables = new Uint32Array(256 * 16);
  tables.set(base);
  for (let k = 1; k < 16; k++) {
    for (let i = 0; i < 256; i++) {
      const previous = tables[(k - 1) * 256 + i]!;
      tables[k * 256 + i] = (previous >>> 8) ^ base[previous & 0xff]!;
    }
  }
  _sliceTables = tables;
  return tables;
}

let _littleEndian: boolean | undefined;

/**
 * Whether a `Uint32Array` view reads bytes least-significant first, which the word-at-a-time
 * loop relies on. Every engine that runs this in practice is little-endian; the byte-wise
 * slicing loop covers any that is not.
 */
function isLittleEndian(): boolean {
  return (_littleEndian ??= new Uint8Array(new Uint32Array([1]).buffer)[0] === 1);
}

/** Below this many aligned bytes a `Uint32Array` view costs more than it saves. */
const WORD_PATH_MIN_BYTES = 256;

/**
 * Update a CRC32 value with a new data chunk.
 *
 * The CRC state here is the non-finalized (inverted) form:
 * - initial state: 0xffffffff
 * - finalize: xor with 0xffffffff
 */
export function crc32UpdateJS(crc: number, data: Uint8Array): number {
  return crc32UpdateSliced(crc, data, isLittleEndian());
}

/**
 * {@link crc32UpdateJS} with the loop chosen by the caller: `wordLoads` reads the input as
 * aligned 32-bit words, which is only correct on a little-endian engine. Exported so that
 * both loops are tested wherever the tests run.
 *
 * @internal
 */
export function crc32UpdateSliced(crc: number, data: Uint8Array, wordLoads: boolean): number {
  const length = data.length;
  if (length < 16) {
    const table = getCrc32Table();
    for (let i = 0; i < length; i++) {
      crc = table[(crc ^ data[i]!) & 0xff]! ^ (crc >>> 8);
    }
    return crc >>> 0;
  }

  const t = getSliceTables();
  let i = 0;

  if (wordLoads && length >= WORD_PATH_MIN_BYTES) {
    // Byte-wise up to a 4-byte boundary, then whole 16-byte blocks as four aligned words
    while ((data.byteOffset + i) & 3) {
      crc = t[(crc ^ data[i]!) & 0xff]! ^ (crc >>> 8);
      i++;
    }
    const blocks = (length - i) >>> 4;
    const words = new Uint32Array(data.buffer, data.byteOffset + i, blocks * 4);
    for (let w = 0; w < words.length; w += 4) {
      const a = crc ^ words[w]!;
      const b = words[w + 1]!;
      const c = words[w + 2]!;
      const d = words[w + 3]!;
      crc =
        t[3840 + (a & 0xff)]! ^
        t[3584 + ((a >>> 8) & 0xff)]! ^
        t[3328 + ((a >>> 16) & 0xff)]! ^
        t[3072 + (a >>> 24)]! ^
        t[2816 + (b & 0xff)]! ^
        t[2560 + ((b >>> 8) & 0xff)]! ^
        t[2304 + ((b >>> 16) & 0xff)]! ^
        t[2048 + (b >>> 24)]! ^
        t[1792 + (c & 0xff)]! ^
        t[1536 + ((c >>> 8) & 0xff)]! ^
        t[1280 + ((c >>> 16) & 0xff)]! ^
        t[1024 + (c >>> 24)]! ^
        t[768 + (d & 0xff)]! ^
        t[512 + ((d >>> 8) & 0xff)]! ^
        t[256 + ((d >>> 16) & 0xff)]! ^
        t[d >>> 24]!;
    }
    i += blocks * 16;
  } else {
    for (; i + 16 <= length; i += 16) {
      const a =
        crc ^ (data[i]! | (data[i + 1]! << 8) | (data[i + 2]! << 16) | (data[i + 3]! << 24));
      const b = data[i + 4]! | (data[i + 5]! << 8) | (data[i + 6]! << 16) | (data[i + 7]! << 24);
      const c = data[i + 8]! | (data[i + 9]! << 8) | (data[i + 10]! << 16) | (data[i + 11]! << 24);
      const d =
        data[i + 12]! | (data[i + 13]! << 8) | (data[i + 14]! << 16) | (data[i + 15]! << 24);
      crc =
        t[3840 + (a & 0xff)]! ^
        t[3584 + ((a >>> 8) & 0xff)]! ^
        t[3328 + ((a >>> 16) & 0xff)]! ^
        t[3072 + (a >>> 24)]! ^
        t[2816 + (b & 0xff)]! ^
        t[2560 + ((b >>> 8) & 0xff)]! ^
        t[2304 + ((b >>> 16) & 0xff)]! ^
        t[2048 + (b >>> 24)]! ^
        t[1792 + (c & 0xff)]! ^
        t[1536 + ((c >>> 8) & 0xff)]! ^
        t[1280 + ((c >>> 16) & 0xff)]! ^
        t[1024 + (c >>> 24)]! ^
        t[768 + (d & 0xff)]! ^
        t[512 + ((d >>> 8) & 0xff)]! ^
        t[256 + ((d >>> 16) & 0xff)]! ^
        t[d >>> 24]!;
    }
  }

  for (; i < length; i++) {
    crc = t[(crc ^ data[i]!) & 0xff]! ^ (crc >>> 8);
  }
  // Unsigned, as the native update returns it, so the state is the same number on every platform
  return crc >>> 0;
}

/**
 * Finalize CRC32 calculation.
 */
export function crc32Finalize(crc: number): number {
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Compute CRC32 for the full input using the JS implementation.
 */
export function crc32JS(data: Uint8Array): number {
  return crc32Finalize(crc32UpdateJS(0xffffffff, data));
}

/**
 * Single-byte CRC32 update for ZipCrypto key derivation.
 * This uses the non-inverted form (raw table lookup).
 *
 * @param crc - Current CRC value (non-inverted)
 * @param byte - Single byte to process
 * @returns Updated CRC value
 */
export function crc32UpdateByte(crc: number, byte: number): number {
  const table = getCrc32Table();
  return (table[(crc ^ byte) & 0xff]! ^ (crc >>> 8)) >>> 0;
}
