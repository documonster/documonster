/**
 * Pure JavaScript DEFLATE implementation for browsers without CompressionStream support
 *
 * This fallback supports:
 * - Decompression: Full DEFLATE decompression (RFC 1951)
 * - Compression: STORE mode only (no compression, but valid DEFLATE format)
 *
 * Used automatically when CompressionStream with "deflate-raw" is unavailable:
 * - Firefox < 113
 * - Safari < 16.4
 * - Chrome < 103
 */

import { ArchiveError } from "@archive/core/errors";
import { concatUint8Arrays } from "@utils/binary";

// ============================================================================
// DEFLATE Decompression (Full implementation)
// ============================================================================

// Fixed Huffman code lengths for literals/lengths (RFC 1951)
const FIXED_LITERAL_LENGTHS = /* @__PURE__ */ (() => {
  const t = new Uint8Array(288);
  for (let i = 0; i <= 143; i++) {
    t[i] = 8;
  }
  for (let i = 144; i <= 255; i++) {
    t[i] = 9;
  }
  for (let i = 256; i <= 279; i++) {
    t[i] = 7;
  }
  for (let i = 280; i <= 287; i++) {
    t[i] = 8;
  }
  return t;
})();

// Fixed Huffman code lengths for distances
const FIXED_DISTANCE_LENGTHS = /* @__PURE__ */ (() => new Uint8Array(32).fill(5))();

// Length base values and extra bits (codes 257-285)
const LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258
];
const LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
];

// Distance base values and extra bits (codes 0-29)
const DISTANCE_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12289, 16385, 24577
];
const DISTANCE_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
];

// Distance code table from RFC 1951.
// Each entry: [maxDistance, code, extraBits]
const DIST_TABLE: ReadonlyArray<readonly [number, number, number]> = [
  [1, 0, 0],
  [2, 1, 0],
  [3, 2, 0],
  [4, 3, 0],
  [6, 4, 1],
  [8, 5, 1],
  [12, 6, 2],
  [16, 7, 2],
  [24, 8, 3],
  [32, 9, 3],
  [48, 10, 4],
  [64, 11, 4],
  [96, 12, 5],
  [128, 13, 5],
  [192, 14, 6],
  [256, 15, 6],
  [384, 16, 7],
  [512, 17, 7],
  [768, 18, 8],
  [1024, 19, 8],
  [1536, 20, 9],
  [2048, 21, 9],
  [3072, 22, 10],
  [4096, 23, 10],
  [6144, 24, 11],
  [8192, 25, 11],
  [12288, 26, 12],
  [16384, 27, 12],
  [24576, 28, 13],
  [32768, 29, 13]
];

// Code length order for dynamic Huffman tables
const CODE_LENGTH_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

/**
 * Huffman tree node
 */
interface HuffmanNode {
  symbol?: number;
  left?: HuffmanNode;
  right?: HuffmanNode;
}

/**
 * Build Huffman tree from code lengths
 */
function buildHuffmanTree(lengths: Uint8Array, maxSymbol: number): HuffmanNode {
  // Count codes of each length
  const blCount = new Uint16Array(16);
  for (let i = 0; i < maxSymbol; i++) {
    if (lengths[i] > 0) {
      blCount[lengths[i]]++;
    }
  }

  // Find first code value for each length
  const nextCode = new Uint16Array(16);
  let code = 0;
  for (let bits = 1; bits <= 15; bits++) {
    code = (code + blCount[bits - 1]) << 1;
    nextCode[bits] = code;
  }

  // Build tree
  const root: HuffmanNode = {};

  for (let symbol = 0; symbol < maxSymbol; symbol++) {
    const len = lengths[symbol];
    if (len === 0) {
      continue;
    }

    code = nextCode[len]++;
    let node = root;

    for (let bit = len - 1; bit >= 0; bit--) {
      const b = (code >> bit) & 1;
      if (b === 0) {
        if (!node.left) {
          node.left = {};
        }
        node = node.left;
      } else {
        if (!node.right) {
          node.right = {};
        }
        node = node.right;
      }
    }
    node.symbol = symbol;
  }

  return root;
}

/**
 * Bit reader for DEFLATE streams
 */
class BitReader {
  private data: Uint8Array;
  private pos: number;
  private bitBuf: number;
  private bitCount: number;

  constructor(data: Uint8Array) {
    this.data = data;
    this.pos = 0;
    this.bitBuf = 0;
    this.bitCount = 0;
  }

  /**
   * Read n bits (LSB first)
   */
  readBits(n: number): number {
    while (this.bitCount < n) {
      if (this.pos >= this.data.length) {
        throw new ArchiveError("Unexpected end of DEFLATE data");
      }
      this.bitBuf |= this.data[this.pos++] << this.bitCount;
      this.bitCount += 8;
    }
    const result = this.bitBuf & ((1 << n) - 1);
    this.bitBuf >>= n;
    this.bitCount -= n;
    return result;
  }

  /**
   * Decode a symbol using Huffman tree
   */
  decodeSymbol(tree: HuffmanNode): number {
    let node = tree;
    while (node.symbol === undefined) {
      const bit = this.readBits(1);
      node = bit === 0 ? node.left! : node.right!;
      if (!node) {
        throw new ArchiveError("Invalid Huffman code");
      }
    }
    return node.symbol;
  }

  /**
   * Align to byte boundary
   */
  alignToByte(): void {
    this.bitBuf = 0;
    this.bitCount = 0;
  }

  /**
   * Read a byte directly (must be aligned)
   */
  readByte(): number {
    if (this.pos >= this.data.length) {
      throw new ArchiveError("Unexpected end of data");
    }
    return this.data[this.pos++];
  }

  /**
   * Read 16-bit little-endian value (must be aligned)
   */
  readUint16(): number {
    return this.readByte() | (this.readByte() << 8);
  }
}

/**
 * Decompress DEFLATE data (raw format, no zlib header)
 *
 * @param data - Compressed data in deflate-raw format
 * @returns Decompressed data
 */
export function inflateRaw(data: Uint8Array): Uint8Array {
  const reader = new BitReader(data);
  // Output accumulator. DEFLATE has no length prefix, so we grow a typed array
  // geometrically instead of pushing into a `number[]` (which boxes every byte
  // at ~8x the memory and forces a final full copy). Seed with a heuristic
  // based on the compressed size; deflate ratios on OOXML/text are commonly
  // 3-5x, so 4x is a reasonable starting capacity that avoids most regrowths.
  let output = new Uint8Array(Math.max(64, data.length * 4));
  let outLen = 0;

  const ensureCapacity = (additional: number): void => {
    const required = outLen + additional;
    if (required <= output.length) {
      return;
    }
    let capacity = output.length;
    while (capacity < required) {
      capacity *= 2;
    }
    const grown = new Uint8Array(capacity);
    grown.set(output.subarray(0, outLen));
    output = grown;
  };

  let isFinal = false;

  while (!isFinal) {
    isFinal = reader.readBits(1) === 1;
    const blockType = reader.readBits(2);

    if (blockType === 0) {
      // Stored block (no compression)
      reader.alignToByte();
      const len = reader.readUint16();
      const nlen = reader.readUint16();

      if ((len ^ nlen) !== 0xffff) {
        throw new ArchiveError("Invalid stored block length");
      }

      ensureCapacity(len);
      for (let i = 0; i < len; i++) {
        output[outLen++] = reader.readByte();
      }
    } else if (blockType === 1 || blockType === 2) {
      // Compressed block
      let literalTree: HuffmanNode;
      let distanceTree: HuffmanNode;

      if (blockType === 1) {
        // Fixed Huffman codes
        literalTree = buildHuffmanTree(FIXED_LITERAL_LENGTHS, 288);
        distanceTree = buildHuffmanTree(FIXED_DISTANCE_LENGTHS, 32);
      } else {
        // Dynamic Huffman codes
        const hlit = reader.readBits(5) + 257;
        const hdist = reader.readBits(5) + 1;
        const hclen = reader.readBits(4) + 4;

        // Read code length code lengths
        const codeLengthLengths = new Uint8Array(19);
        for (let i = 0; i < hclen; i++) {
          codeLengthLengths[CODE_LENGTH_ORDER[i]] = reader.readBits(3);
        }

        const codeLengthTree = buildHuffmanTree(codeLengthLengths, 19);

        // Decode literal/length and distance code lengths
        const allLengths = new Uint8Array(hlit + hdist);
        let i = 0;

        while (i < hlit + hdist) {
          const symbol = reader.decodeSymbol(codeLengthTree);

          if (symbol < 16) {
            allLengths[i++] = symbol;
          } else if (symbol === 16) {
            // Copy previous length 3-6 times
            const repeat = reader.readBits(2) + 3;
            const prev = allLengths[i - 1];
            for (let j = 0; j < repeat; j++) {
              allLengths[i++] = prev;
            }
          } else if (symbol === 17) {
            // Repeat 0 for 3-10 times
            const repeat = reader.readBits(3) + 3;
            for (let j = 0; j < repeat; j++) {
              allLengths[i++] = 0;
            }
          } else if (symbol === 18) {
            // Repeat 0 for 11-138 times
            const repeat = reader.readBits(7) + 11;
            for (let j = 0; j < repeat; j++) {
              allLengths[i++] = 0;
            }
          }
        }

        literalTree = buildHuffmanTree(allLengths.subarray(0, hlit), hlit);
        distanceTree = buildHuffmanTree(allLengths.subarray(hlit), hdist);
      }

      // Decode compressed data
      while (true) {
        const symbol = reader.decodeSymbol(literalTree);

        if (symbol < 256) {
          // Literal byte
          ensureCapacity(1);
          output[outLen++] = symbol;
        } else if (symbol === 256) {
          // End of block
          break;
        } else {
          // Length/distance pair
          const lengthCode = symbol - 257;
          const length = LENGTH_BASE[lengthCode] + reader.readBits(LENGTH_EXTRA[lengthCode]);

          const distCode = reader.decodeSymbol(distanceTree);
          const distance = DISTANCE_BASE[distCode] + reader.readBits(DISTANCE_EXTRA[distCode]);

          // Copy from output buffer. Overlapping copies (distance < length) are
          // intentional and must be done byte-by-byte so repeated runs expand
          // correctly — do not use set()/copyWithin() here.
          ensureCapacity(length);
          const start = outLen - distance;
          for (let i = 0; i < length; i++) {
            output[outLen++] = output[start + i];
          }
        }
      }
    } else {
      throw new ArchiveError("Invalid DEFLATE block type: " + blockType);
    }
  }

  // Return an exactly-sized copy so callers never observe the over-allocated
  // tail and the growth buffer can be garbage-collected.
  return output.slice(0, outLen);
}

// ============================================================================
// DEFLATE Compression (STORE mode only - no actual compression)
// ============================================================================

/**
 * Compress data using DEFLATE STORE mode (no compression)
 *
 * This creates valid DEFLATE data but without actual compression.
 * Files will be larger but this works on all browsers.
 *
 * @param data - Data to "compress"
 * @returns DEFLATE-formatted data (stored, not compressed)
 */
export function deflateRawStore(data: Uint8Array): Uint8Array {
  // Maximum stored block size is 65535 bytes
  const MAX_BLOCK_SIZE = 65535;
  const numBlocks = Math.ceil(data.length / MAX_BLOCK_SIZE) || 1;

  // Calculate output size: 5 bytes header per block + data
  const outputSize = numBlocks * 5 + data.length;
  const output = new Uint8Array(outputSize);
  let outPos = 0;
  let inPos = 0;

  for (let block = 0; block < numBlocks; block++) {
    const isLast = block === numBlocks - 1;
    const blockSize = Math.min(MAX_BLOCK_SIZE, data.length - inPos);

    // Block header: BFINAL (1 bit) + BTYPE=00 (2 bits) = stored block
    // Then align to byte boundary (5 bits padding)
    output[outPos++] = isLast ? 0x01 : 0x00;

    // LEN (16-bit little-endian)
    output[outPos++] = blockSize & 0xff;
    output[outPos++] = (blockSize >> 8) & 0xff;

    // NLEN (one's complement of LEN)
    output[outPos++] = ~blockSize & 0xff;
    output[outPos++] = (~blockSize >> 8) & 0xff;

    // Data
    output.set(data.subarray(inPos, inPos + blockSize), outPos);
    outPos += blockSize;
    inPos += blockSize;
  }

  return output.subarray(0, outPos);
}

// ============================================================================
// LZ77 + Huffman Compression
// ============================================================================

// Hash table size must be a power of 2. 32768 entries keeps memory reasonable
// while providing a good distribution for the 3-byte hash.
const HASH_SIZE = 32768;
const HASH_MASK = HASH_SIZE - 1;

// Minimum match length for LZ77 (RFC 1951 minimum).
const MIN_MATCH = 3;

// Maximum match length (RFC 1951 maximum).
const MAX_MATCH = 258;

// Maximum back-reference distance (RFC 1951 / 32 KB sliding window).
const MAX_DIST = 32768;

// =============================================================================
// Level-based LZ77 configuration (modelled after zlib's deflate parameters)
// =============================================================================

interface LZ77Config {
  maxChainLen: number;
  /** Lazy match: emit current match immediately if length >= this (skip lazy). */
  goodLen: number;
  /** Lazy match: only try lazy if current match length < niceLen. */
  niceLen: number;
  /** Enable lazy matching (false = greedy). */
  lazy: boolean;
}

/**
 * Get LZ77 configuration for the given compression level (1-9).
 * Modelled after zlib's configuration_table.
 */
function getLZ77Config(level: number): LZ77Config {
  // Level 0 should be handled by the caller (store mode).
  if (level <= 1) {
    return { maxChainLen: 4, goodLen: 4, niceLen: 8, lazy: false };
  }
  if (level <= 3) {
    return { maxChainLen: 8, goodLen: 8, niceLen: 32, lazy: true };
  }
  if (level <= 5) {
    return { maxChainLen: 32, goodLen: 16, niceLen: 128, lazy: true };
  }
  if (level <= 7) {
    return { maxChainLen: 64, goodLen: 32, niceLen: 258, lazy: true };
  }
  // level 8-9
  return { maxChainLen: 128, goodLen: 64, niceLen: 258, lazy: true };
}

/**
 * Hash function for 3-byte sequences.
 * Uses a multiplicative hash for better distribution than the naive
 * shift-or approach. The constant 0x1e35a7bd is chosen for good avalanche
 * properties in the lower bits.
 */
function hash3(a: number, b: number, c: number): number {
  return ((((a << 16) | (b << 8) | c) * 0x1e35a7bd) >>> 17) & HASH_MASK;
}

/**
 * Compress data using DEFLATE with Dynamic Huffman codes (BTYPE=2).
 *
 * Uses LZ77 with hash chains and lazy matching for match finding, then builds
 * optimal Huffman trees from the symbol frequencies for entropy coding.
 *
 * @param data - Data to compress
 * @param level - Compression level (1-9, default 6)
 * @returns Compressed data in deflate-raw format
 */
export function deflateRawCompressed(data: Uint8Array, level = 6): Uint8Array {
  if (data.length === 0) {
    // Empty input: single final block with just end-of-block symbol
    return new Uint8Array([0x03, 0x00]);
  }

  // For small data, use STORE mode
  if (data.length < 100) {
    return deflateRawStore(data);
  }

  const config = getLZ77Config(level);

  // --- Phase 1: LZ77 match finding → collect symbols ---
  const lz77Symbols = lz77Compress(data, 0, data.length, config, null);

  // --- Phase 2: Encode as a single final DEFLATE block ---
  const output = new BitWriter();
  emitDynamicBlock(output, lz77Symbols, true);

  return output.finish();
}

/**
 * Bit writer for DEFLATE output
 */
class BitWriter {
  private static readonly CHUNK_SIZE = 65536;
  private chunks: Uint8Array[] = [];
  // Current chunk written by index instead of pushing into a `number[]`
  // (which boxes every byte). Flushed into `chunks` when full.
  private buffer = new Uint8Array(BitWriter.CHUNK_SIZE);
  private bufLen = 0;
  private bitBuf = 0;
  private bitCount = 0;

  /**
   * Align to the next byte boundary by padding with zero bits.
   */
  alignToByte(): void {
    if (this.bitCount > 0) {
      this.writeBits(0, 8 - this.bitCount);
    }
  }

  private pushByte(byte: number): void {
    this.buffer[this.bufLen++] = byte;
    if (this.bufLen >= BitWriter.CHUNK_SIZE) {
      this.chunks.push(this.buffer);
      this.buffer = new Uint8Array(BitWriter.CHUNK_SIZE);
      this.bufLen = 0;
    }
  }

  writeBits(value: number, count: number): void {
    this.bitBuf |= value << this.bitCount;
    this.bitCount += count;

    while (this.bitCount >= 8) {
      this.pushByte(this.bitBuf & 0xff);
      this.bitBuf >>= 8;
      this.bitCount -= 8;
    }
  }

  writeBitsReverse(value: number, count: number): void {
    // Write bits in reverse order (MSB first, used for Huffman codes)
    let reversed = 0;
    for (let i = 0; i < count; i++) {
      reversed = (reversed << 1) | ((value >> i) & 1);
    }
    this.writeBits(reversed, count);
  }

  finish(): Uint8Array {
    // Flush remaining bits
    if (this.bitCount > 0) {
      this.pushByte(this.bitBuf & 0xff);
      this.bitBuf = 0;
      this.bitCount = 0;
    }

    if (this.chunks.length === 0) {
      return this.buffer.slice(0, this.bufLen);
    }

    this.chunks.push(this.buffer.subarray(0, this.bufLen));
    return concatUint8Arrays(this.chunks);
  }

  /**
   * Return all fully completed bytes, leaving the partial byte intact.
   * Used by SyncDeflater.write() to emit output between blocks while
   * preserving the bit-stream state for the next block.
   */
  flushBytes(): Uint8Array {
    if (this.chunks.length === 0 && this.bufLen === 0) {
      return new Uint8Array(0);
    }

    let result: Uint8Array;
    if (this.chunks.length === 0) {
      result = this.buffer.slice(0, this.bufLen);
    } else {
      this.chunks.push(this.buffer.subarray(0, this.bufLen));
      result = concatUint8Arrays(this.chunks);
      this.chunks = [];
    }
    this.buffer = new Uint8Array(BitWriter.CHUNK_SIZE);
    this.bufLen = 0;
    return result;
  }
}

// Fixed Huffman code tables
const LITERAL_CODES = /* @__PURE__ */ (() => {
  const codes: Array<[number, number]> = [];

  // Build fixed literal/length Huffman codes
  for (let i = 0; i <= 287; i++) {
    let code: number;
    let len: number;

    if (i <= 143) {
      // 00110000 - 10111111 (8 bits)
      code = 0x30 + i;
      len = 8;
    } else if (i <= 255) {
      // 110010000 - 111111111 (9 bits)
      code = 0x190 + (i - 144);
      len = 9;
    } else if (i <= 279) {
      // 0000000 - 0010111 (7 bits)
      code = i - 256;
      len = 7;
    } else {
      // 11000000 - 11000111 (8 bits)
      code = 0xc0 + (i - 280);
      len = 8;
    }

    codes[i] = [code, len];
  }

  return codes;
})();

/**
 * Write a literal or end-of-block symbol using fixed Huffman codes
 */
function writeLiteralCode(output: BitWriter, symbol: number): void {
  const [code, len] = LITERAL_CODES[symbol];
  output.writeBitsReverse(code, len);
}

// ============================================================================
// Dynamic Huffman Encoding (RFC 1951 §3.2.7)
// ============================================================================

/**
 * LZ77 symbol: either a literal byte or a (length, distance) pair.
 * Stored compactly: if dist === 0, it's a literal (value = litOrLen).
 * Otherwise it's a match with length = litOrLen, distance = dist.
 */
interface LZ77Symbol {
  litOrLen: number;
  dist: number;
}

/**
 * Compute the DEFLATE length code (257..285) and extra bits for a given
 * match length (3..258).
 */
function getLengthSymbol(length: number): { code: number; extra: number; extraBits: number } {
  for (let i = 0; i < LENGTH_BASE.length; i++) {
    if (i === LENGTH_BASE.length - 1 || length < LENGTH_BASE[i + 1]) {
      return {
        code: 257 + i,
        extra: length - LENGTH_BASE[i],
        extraBits: LENGTH_EXTRA[i]
      };
    }
  }
  return { code: 285, extra: 0, extraBits: 0 };
}

/**
 * Compute the DEFLATE distance code (0..29) and extra bits for a given
 * distance (1..32768).
 */
function getDistSymbol(distance: number): { code: number; extra: number; extraBits: number } {
  for (let i = 0; i < DIST_TABLE.length; i++) {
    const [maxDist, c, extraBitsCount] = DIST_TABLE[i];
    if (distance <= maxDist) {
      const baseVal = i === 0 ? 1 : DIST_TABLE[i - 1][0] + 1;
      return { code: c, extra: distance - baseVal, extraBits: extraBitsCount };
    }
  }
  // Fallback (should not reach for valid distances)
  return { code: 29, extra: 0, extraBits: 13 };
}

/**
 * Build canonical Huffman code lengths from symbol frequencies.
 * Uses a bottom-up approach: build a Huffman tree from a priority queue,
 * then extract depths. Limits maximum code length to maxBits using
 * the algorithm from zlib's build_tree() / gen_bitlen().
 *
 * Returns an array of code lengths indexed by symbol.
 */
function buildCodeLengths(freqs: Uint32Array, maxBits: number): Uint8Array {
  const n = freqs.length;
  const codeLens = new Uint8Array(n);

  // Count symbols with non-zero frequency
  const activeSymbols: Array<{ sym: number; freq: number }> = [];
  for (let i = 0; i < n; i++) {
    if (freqs[i] > 0) {
      activeSymbols.push({ sym: i, freq: freqs[i] });
    }
  }

  if (activeSymbols.length === 0) {
    return codeLens;
  }
  // RFC 1951 requires a complete prefix code. For a single symbol, we need
  // at least 2 entries to form a valid tree. We assign code length 1 to the
  // symbol — the decoder uses only 1 bit but the tree is valid because
  // DEFLATE decoders handle this as per the spec (the other 1-bit code is
  // simply unused). This matches zlib's behavior.
  if (activeSymbols.length === 1) {
    codeLens[activeSymbols[0].sym] = 1;
    return codeLens;
  }

  // Sort by frequency (ascending), then by symbol (ascending) for stability
  activeSymbols.sort((a, b) => a.freq - b.freq || a.sym - b.sym);

  // Build Huffman tree using a two-queue approach (Moffat & Turpin).
  // O(n) for pre-sorted input. Queue 1: leaf nodes. Queue 2: internal nodes.
  interface TreeNode {
    freq: number;
    sym: number; // -1 for internal
    left: TreeNode | null;
    right: TreeNode | null;
  }

  const nodes: TreeNode[] = activeSymbols.map(s => ({
    freq: s.freq,
    sym: s.sym,
    left: null,
    right: null
  }));

  let leafIdx = 0;
  let intIdx = 0;
  const intNodes: TreeNode[] = [];

  function getMin(): TreeNode {
    const hasLeaf = leafIdx < nodes.length;
    const hasInt = intIdx < intNodes.length;

    if (hasLeaf && hasInt) {
      if (nodes[leafIdx].freq <= intNodes[intIdx].freq) {
        return nodes[leafIdx++];
      }
      return intNodes[intIdx++];
    }
    if (hasLeaf) {
      return nodes[leafIdx++];
    }
    return intNodes[intIdx++];
  }

  const totalNodes = activeSymbols.length;
  for (let i = 0; i < totalNodes - 1; i++) {
    const a = getMin();
    const b = getMin();
    const merged: TreeNode = {
      freq: a.freq + b.freq,
      sym: -1,
      left: a,
      right: b
    };
    intNodes.push(merged);
  }

  // Extract depths from the root (last internal node)
  const root = intNodes[intNodes.length - 1];

  function extractDepths(node: TreeNode, depth: number): void {
    if (node.sym >= 0) {
      codeLens[node.sym] = depth;
      return;
    }
    if (node.left) {
      extractDepths(node.left, depth + 1);
    }
    if (node.right) {
      extractDepths(node.right, depth + 1);
    }
  }
  extractDepths(root, 0);

  // --- Length limiting using the zlib bl_count redistribution algorithm ---
  // Count code lengths at each bit depth
  const blCount = new Uint16Array(maxBits + 1);

  for (let i = 0; i < n; i++) {
    if (codeLens[i] > 0) {
      if (codeLens[i] > maxBits) {
        blCount[maxBits]++;
        codeLens[i] = maxBits;
      } else {
        blCount[codeLens[i]]++;
      }
    }
  }

  // Check Kraft inequality: sum of 2^(maxBits - len) must equal 2^maxBits
  let kraft = 0;
  for (let bits = 1; bits <= maxBits; bits++) {
    kraft += blCount[bits] << (maxBits - bits);
  }
  const target = 1 << maxBits;

  if (kraft === target) {
    return codeLens; // Already valid
  }

  // Redistribute to satisfy Kraft's inequality.
  // Strategy: move symbols from shorter lengths to maxBits until balanced.
  // Each symbol moved from length `bits` to `maxBits` reduces kraft by
  // (2^(maxBits-bits) - 1) — we remove a large weight and add a weight of 1.
  while (kraft > target) {
    // Find a code length < maxBits that has symbols we can push down.
    // Start from maxBits-1 to minimize the damage per move.
    let bits = maxBits - 1;
    while (bits > 0 && blCount[bits] === 0) {
      bits--;
    }
    if (bits === 0) {
      break; // Can't redistribute further
    }
    // Move one symbol from length `bits` to length `maxBits`
    blCount[bits]--;
    blCount[maxBits]++;
    // Kraft change: removed 2^(maxBits-bits), added 2^0 = 1
    kraft -= (1 << (maxBits - bits)) - 1;
  }

  // If kraft < target (under-allocated), add dummy codes at maxBits.
  // This can happen when we overshoot during redistribution.
  while (kraft < target) {
    blCount[maxBits]++;
    kraft++;
  }

  // Reassign code lengths to symbols (preserve relative order: longer
  // codes go to less frequent symbols, matching the Huffman property).
  // Sort symbols by their original code length (longest first), then by
  // frequency (rarest first) for same length.
  const symbolsByLen: Array<{ sym: number; origLen: number; freq: number }> = [];
  for (let i = 0; i < n; i++) {
    if (codeLens[i] > 0) {
      symbolsByLen.push({ sym: i, origLen: codeLens[i], freq: freqs[i] });
    }
  }
  symbolsByLen.sort((a, b) => b.origLen - a.origLen || a.freq - b.freq);

  // Assign new lengths from the bl_count distribution
  codeLens.fill(0);
  let symIdx = 0;
  for (let bits = maxBits; bits >= 1; bits--) {
    for (let count = blCount[bits]; count > 0; count--) {
      if (symIdx < symbolsByLen.length) {
        codeLens[symbolsByLen[symIdx].sym] = bits;
        symIdx++;
      }
    }
  }

  return codeLens;
}

/**
 * Build canonical Huffman codes from code lengths (RFC 1951 §3.2.2).
 * Returns [code, length] pairs indexed by symbol.
 */
function buildCanonicalCodes(codeLens: Uint8Array): Array<[number, number]> {
  const n = codeLens.length;
  const codes: Array<[number, number]> = new Array(n);

  const blCount = new Uint16Array(16);
  for (let i = 0; i < n; i++) {
    if (codeLens[i] > 0) {
      blCount[codeLens[i]]++;
    }
  }

  const nextCode = new Uint16Array(16);
  let code = 0;
  for (let bits = 1; bits <= 15; bits++) {
    code = (code + blCount[bits - 1]) << 1;
    nextCode[bits] = code;
  }

  for (let i = 0; i < n; i++) {
    const len = codeLens[i];
    if (len > 0) {
      codes[i] = [nextCode[len]++, len];
    } else {
      codes[i] = [0, 0];
    }
  }

  return codes;
}

/**
 * Emit a Dynamic Huffman DEFLATE block (BTYPE=2).
 *
 * Takes the LZ77 symbol sequence, builds optimal Huffman trees,
 * encodes the tree descriptions, then encodes the symbols.
 */
function emitDynamicBlock(out: BitWriter, symbols: LZ77Symbol[], isFinal: boolean): void {
  // --- Step 1: Collect frequencies ---
  const litLenFreqs = new Uint32Array(286);
  const distFreqs = new Uint32Array(30);

  // Always include EOB
  litLenFreqs[256] = 1;

  for (const sym of symbols) {
    if (sym.dist === 0) {
      litLenFreqs[sym.litOrLen]++;
    } else {
      const ls = getLengthSymbol(sym.litOrLen);
      litLenFreqs[ls.code]++;
      const ds = getDistSymbol(sym.dist);
      distFreqs[ds.code]++;
    }
  }

  // --- Step 2: Build Huffman trees ---
  const litLenLens = buildCodeLengths(litLenFreqs, 15);
  let distLens = buildCodeLengths(distFreqs, 15);

  // DEFLATE requires at least 1 distance code even if unused.
  // Assign two codes at length 1 to form a complete prefix code.
  let hasDistCodes = false;
  for (let i = 0; i < distLens.length; i++) {
    if (distLens[i] > 0) {
      hasDistCodes = true;
      break;
    }
  }
  if (!hasDistCodes) {
    distLens = new Uint8Array(30);
    distLens[0] = 1;
    distLens[1] = 1;
  }

  const litLenCodes = buildCanonicalCodes(litLenLens);
  const distCodes = buildCanonicalCodes(distLens);

  // --- Step 3: Determine HLIT and HDIST ---
  let hlit = 286;
  while (hlit > 257 && litLenLens[hlit - 1] === 0) {
    hlit--;
  }
  let hdist = 30;
  while (hdist > 1 && distLens[hdist - 1] === 0) {
    hdist--;
  }

  // --- Step 4: Run-length encode the code lengths ---
  const combined = new Uint8Array(hlit + hdist);
  combined.set(litLenLens.subarray(0, hlit));
  combined.set(distLens.subarray(0, hdist), hlit);

  const clSymbols: Array<{ sym: number; extra: number; extraBits: number }> = [];
  const clFreqs = new Uint32Array(19);

  for (let i = 0; i < combined.length; ) {
    const val = combined[i];

    if (val === 0) {
      let run = 1;
      while (i + run < combined.length && combined[i + run] === 0) {
        run++;
      }

      while (run > 0) {
        if (run >= 11) {
          const repeat = Math.min(run, 138);
          clSymbols.push({ sym: 18, extra: repeat - 11, extraBits: 7 });
          clFreqs[18]++;
          run -= repeat;
          i += repeat;
        } else if (run >= 3) {
          const repeat = Math.min(run, 10);
          clSymbols.push({ sym: 17, extra: repeat - 3, extraBits: 3 });
          clFreqs[17]++;
          run -= repeat;
          i += repeat;
        } else {
          clSymbols.push({ sym: 0, extra: 0, extraBits: 0 });
          clFreqs[0]++;
          run--;
          i++;
        }
      }
    } else {
      clSymbols.push({ sym: val, extra: 0, extraBits: 0 });
      clFreqs[val]++;
      i++;

      let run = 0;
      while (i + run < combined.length && combined[i + run] === val) {
        run++;
      }
      while (run >= 3) {
        const repeat = Math.min(run, 6);
        clSymbols.push({ sym: 16, extra: repeat - 3, extraBits: 2 });
        clFreqs[16]++;
        run -= repeat;
        i += repeat;
      }
      while (run > 0) {
        clSymbols.push({ sym: val, extra: 0, extraBits: 0 });
        clFreqs[val]++;
        run--;
        i++;
      }
    }
  }

  // --- Step 5: Build code-length Huffman tree ---
  const clLens = buildCodeLengths(clFreqs, 7);
  const clCodes = buildCanonicalCodes(clLens);

  let hclen = 19;
  while (hclen > 4 && clLens[CODE_LENGTH_ORDER[hclen - 1]] === 0) {
    hclen--;
  }

  // --- Step 6: Write block header ---
  out.writeBits(isFinal ? 1 : 0, 1); // BFINAL
  out.writeBits(2, 2); // BTYPE = 10 (dynamic Huffman)

  out.writeBits(hlit - 257, 5);
  out.writeBits(hdist - 1, 5);
  out.writeBits(hclen - 4, 4);

  for (let i = 0; i < hclen; i++) {
    out.writeBits(clLens[CODE_LENGTH_ORDER[i]], 3);
  }

  for (const cls of clSymbols) {
    const [clCode, clLen] = clCodes[cls.sym];
    out.writeBitsReverse(clCode, clLen);
    if (cls.extraBits > 0) {
      out.writeBits(cls.extra, cls.extraBits);
    }
  }

  // --- Step 7: Write compressed data ---
  for (const sym of symbols) {
    if (sym.dist === 0) {
      const [lCode, lLen] = litLenCodes[sym.litOrLen];
      out.writeBitsReverse(lCode, lLen);
    } else {
      const ls = getLengthSymbol(sym.litOrLen);
      const [lCode, lLen] = litLenCodes[ls.code];
      out.writeBitsReverse(lCode, lLen);
      if (ls.extraBits > 0) {
        out.writeBits(ls.extra, ls.extraBits);
      }

      const ds = getDistSymbol(sym.dist);
      const [dCode, dLen] = distCodes[ds.code];
      out.writeBitsReverse(dCode, dLen);
      if (ds.extraBits > 0) {
        out.writeBits(ds.extra, ds.extraBits);
      }
    }
  }

  // End of block
  const [eobCode, eobLen] = litLenCodes[256];
  out.writeBitsReverse(eobCode, eobLen);
}

// ============================================================================
// Shared LZ77 Engine
// ============================================================================

/**
 * Hash chain state that persists across chunks for the streaming deflater.
 */
interface LZ77State {
  head: Int32Array;
  prev: Int32Array;
  window: Uint8Array;
  windowLen: number;
  totalIn: number;
  hasPrevMatch: boolean;
  prevMatchLen: number;
  prevMatchDist: number;
  prevLiteral: number;
}

/**
 * Run LZ77 match-finding on `data[start..end)`.
 *
 * When `state` is null, performs one-shot compression with fresh hash tables.
 * When `state` is provided, maintains sliding window and hash chains across calls.
 *
 * Returns an array of LZ77 symbols (literals + length/distance pairs).
 */
function lz77Compress(
  data: Uint8Array,
  start: number,
  end: number,
  config: LZ77Config,
  state: LZ77State | null
): LZ77Symbol[] {
  const symbols: LZ77Symbol[] = [];
  const maxChainLen = config.maxChainLen;
  const goodLen = config.goodLen;
  const niceLen = config.niceLen;
  const useLazy = config.lazy;

  let head: Int32Array;
  let prevArr: Int32Array;
  let window: Uint8Array | null;
  let wLen: number;
  let totalIn: number;
  let hasPrevMatch: boolean;
  let prevMatchLen: number;
  let prevMatchDist: number;
  let prevLiteral: number;

  if (state) {
    head = state.head;
    prevArr = state.prev;
    window = state.window;
    wLen = state.windowLen;
    totalIn = state.totalIn;
    hasPrevMatch = state.hasPrevMatch;
    prevMatchLen = state.prevMatchLen;
    prevMatchDist = state.prevMatchDist;
    prevLiteral = state.prevLiteral;
  } else {
    head = new Int32Array(HASH_SIZE);
    prevArr = new Int32Array(MAX_DIST);
    window = null;
    wLen = 0;
    totalIn = 0;
    hasPrevMatch = false;
    prevMatchLen = 0;
    prevMatchDist = 0;
    prevLiteral = 0;
  }

  const getByte = state
    ? (globalPos: number): number => {
        const localPos = globalPos - totalIn;
        if (localPos >= start && localPos < end) {
          return data[localPos];
        }
        return window![globalPos & (MAX_DIST - 1)];
      }
    : (globalPos: number): number => data[globalPos];

  const insertHash = state
    ? (localPos: number): void => {
        if (localPos + 2 >= end) {
          return;
        }
        const h = hash3(data[localPos], data[localPos + 1], data[localPos + 2]);
        const gp = totalIn + localPos;
        prevArr[gp & (MAX_DIST - 1)] = head[h];
        head[h] = gp + 1;
      }
    : (localPos: number): void => {
        if (localPos + 2 >= end) {
          return;
        }
        const h = hash3(data[localPos], data[localPos + 1], data[localPos + 2]);
        prevArr[localPos & (MAX_DIST - 1)] = head[h];
        head[h] = localPos + 1;
      };

  const insertWindow = state
    ? (localPos: number, count: number): void => {
        for (let i = 0; i < count; i++) {
          window![(wLen + i) & (MAX_DIST - 1)] = data[localPos + i];
        }
        wLen += count;
      }
    : (_localPos: number, _count: number): void => {};

  let pos = start;

  for (; pos < end; ) {
    let bestLen = 0;
    let bestDist = 0;

    if (pos + 2 < end) {
      const h = hash3(data[pos], data[pos + 1], data[pos + 2]);
      const globalPos = state ? totalIn + pos : pos;

      // When we already have a good match from a previous lazy evaluation,
      // reduce the chain search length (matching zlib's good_length behavior).
      let chainRemaining =
        useLazy && hasPrevMatch && prevMatchLen >= goodLen ? maxChainLen >> 2 : maxChainLen;
      let matchHead = head[h];

      while (matchHead > 0 && chainRemaining-- > 0) {
        const mGlobalPos = matchHead - 1;
        const dist = globalPos - mGlobalPos;
        if (dist > MAX_DIST || dist <= 0) {
          break;
        }

        if (bestLen >= MIN_MATCH) {
          const checkGlobal = mGlobalPos + bestLen;
          if (getByte(checkGlobal) !== data[pos + bestLen]) {
            matchHead = prevArr[mGlobalPos & (MAX_DIST - 1)];
            continue;
          }
        }

        const maxLen = Math.min(MAX_MATCH, end - pos);
        let len = 0;
        while (len < maxLen) {
          if (getByte(mGlobalPos + len) !== data[pos + len]) {
            break;
          }
          len++;
        }

        if (len > bestLen) {
          bestLen = len;
          bestDist = dist;
          if (len >= niceLen) {
            break;
          }
        }

        matchHead = prevArr[mGlobalPos & (MAX_DIST - 1)];
      }

      if (state) {
        prevArr[globalPos & (MAX_DIST - 1)] = head[h];
        head[h] = globalPos + 1;
      } else {
        prevArr[pos & (MAX_DIST - 1)] = head[h];
        head[h] = pos + 1;
      }
    }

    if (useLazy && hasPrevMatch) {
      if (bestLen > prevMatchLen) {
        symbols.push({ litOrLen: prevLiteral, dist: 0 });
        prevMatchLen = bestLen;
        prevMatchDist = bestDist;
        prevLiteral = data[pos];
        insertWindow(pos, 1);
        pos++;
      } else {
        symbols.push({ litOrLen: prevMatchLen, dist: prevMatchDist });
        const matchEnd = Math.min(pos - 1 + prevMatchLen, end);
        for (let i = pos; i < matchEnd; i++) {
          insertHash(i);
        }
        insertWindow(pos, matchEnd - pos);
        pos = matchEnd;
        hasPrevMatch = false;
        prevMatchLen = 0;
      }
    } else if (bestLen >= MIN_MATCH) {
      if (useLazy) {
        hasPrevMatch = true;
        prevMatchLen = bestLen;
        prevMatchDist = bestDist;
        prevLiteral = data[pos];
        insertWindow(pos, 1);
        pos++;
      } else {
        symbols.push({ litOrLen: bestLen, dist: bestDist });
        const matchEnd = Math.min(pos + bestLen, end);
        for (let i = pos + 1; i < matchEnd; i++) {
          insertHash(i);
        }
        insertWindow(pos, matchEnd - pos);
        pos = matchEnd;
      }
    } else {
      if (hasPrevMatch) {
        // Non-lazy mode shouldn't reach here, but handle gracefully
        symbols.push({ litOrLen: prevMatchLen, dist: prevMatchDist });
        hasPrevMatch = false;
        prevMatchLen = 0;
      }
      symbols.push({ litOrLen: data[pos], dist: 0 });
      insertWindow(pos, 1);
      pos++;
    }
  }

  // Flush pending lazy match
  if (hasPrevMatch) {
    symbols.push({ litOrLen: prevMatchLen, dist: prevMatchDist });
    const matchEnd = Math.min(pos - 1 + prevMatchLen, end);
    for (let i = pos; i < matchEnd; i++) {
      insertHash(i);
    }
    insertWindow(pos, matchEnd - pos);
    hasPrevMatch = false;
    prevMatchLen = 0;
  }

  if (state) {
    state.windowLen = wLen;
    state.totalIn = totalIn + (end - start);
    state.hasPrevMatch = hasPrevMatch;
    state.prevMatchLen = prevMatchLen;
    state.prevMatchDist = prevMatchDist;
    state.prevLiteral = prevLiteral;
  }

  return symbols;
}

// ============================================================================
// Stateful Streaming Deflater
// ============================================================================

/**
 * Stateful synchronous DEFLATE compressor with Dynamic Huffman encoding.
 *
 * Unlike `deflateRawCompressed` (which is a one-shot function), this class
 * maintains state across multiple `write()` calls:
 *
 *  - **LZ77 sliding window**: back-references can span across chunks.
 *  - **Hash chains**: match positions persist across chunks with typed-array
 *    hash tables for fast lookup.
 *  - **Lazy matching**: configurable per compression level.
 *  - **Dynamic Huffman**: each block builds optimal Huffman trees from
 *    actual symbol frequencies (BTYPE=2), producing significantly smaller
 *    output than fixed Huffman (BTYPE=1).
 *  - **Bit writer**: bit position is preserved, so consecutive blocks form
 *    a single valid DEFLATE bit-stream without alignment issues.
 *
 * Each `write()` emits one non-final Dynamic Huffman block (BFINAL=0).
 * `finish()` emits a final empty fixed-Huffman block (BFINAL=1).
 *
 * This is the pure-JS equivalent of Node.js `zlib.deflateRawSync` with
 * `Z_SYNC_FLUSH`, used by the streaming ZIP writer (`pushSync`) to achieve
 * constant-memory streaming in both Node.js and browsers.
 *
 * @param level - Compression level (0-9). Level 0 emits STORE blocks.
 *                Default: 6 (matching zlib default).
 */
export class SyncDeflater {
  private _output = new BitWriter();
  private _config: LZ77Config;
  private _level: number;
  private _state: LZ77State = {
    head: new Int32Array(HASH_SIZE),
    prev: new Int32Array(MAX_DIST),
    window: new Uint8Array(MAX_DIST),
    windowLen: 0,
    totalIn: 0,
    hasPrevMatch: false,
    prevMatchLen: 0,
    prevMatchDist: 0,
    prevLiteral: 0
  };

  constructor(level = 6) {
    this._level = Math.max(0, Math.min(9, level));
    this._config = getLZ77Config(this._level);
  }

  /**
   * Compress a chunk and return the compressed bytes produced so far.
   * The output is a valid prefix of a DEFLATE stream (one or more non-final blocks).
   */
  write(data: Uint8Array): Uint8Array {
    if (data.length === 0) {
      return new Uint8Array(0);
    }

    const out = this._output;

    if (this._level === 0) {
      // Store mode: emit uncompressed block(s)
      this._writeStore(data);
      return out.flushBytes();
    }

    // LZ77 + Dynamic Huffman
    const symbols = lz77Compress(data, 0, data.length, this._config, this._state);
    emitDynamicBlock(out, symbols, false);

    return out.flushBytes();
  }

  /**
   * Finalize the DEFLATE stream. Emits a final empty fixed-Huffman block
   * and returns any remaining bytes (including partial-byte padding).
   */
  finish(): Uint8Array {
    const out = this._output;
    // Final block: BFINAL=1, BTYPE=01, immediately followed by EOB (symbol 256)
    out.writeBits(1, 1); // BFINAL = 1
    out.writeBits(1, 2); // BTYPE  = 01 (fixed Huffman)
    writeLiteralCode(out, 256);
    return out.finish();
  }

  /**
   * Write STORE (uncompressed) blocks for level=0.
   * Each block is non-final (BFINAL=0); the final block is emitted by finish().
   */
  private _writeStore(data: Uint8Array): void {
    const out = this._output;
    const MAX_BLOCK_SIZE = 65535;
    let offset = 0;

    while (offset < data.length) {
      const remaining = data.length - offset;
      const blockSize = Math.min(MAX_BLOCK_SIZE, remaining);

      // Align to byte boundary before stored block header
      out.alignToByte();

      out.writeBits(0, 1); // BFINAL = 0 (never final; finish() handles that)
      out.writeBits(0, 2); // BTYPE = 00 (stored)

      // Align to byte boundary after block header (3 bits → pad to 8)
      out.alignToByte();

      // LEN
      out.writeBits(blockSize & 0xff, 8);
      out.writeBits((blockSize >> 8) & 0xff, 8);
      // NLEN
      out.writeBits(~blockSize & 0xff, 8);
      out.writeBits((~blockSize >> 8) & 0xff, 8);

      // Data
      for (let i = 0; i < blockSize; i++) {
        out.writeBits(data[offset + i], 8);
      }

      offset += blockSize;
    }
  }
}
