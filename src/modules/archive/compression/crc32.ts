/**
 * CRC32 calculation utility for ZIP files
 *
 * - Node.js: Uses native zlib.crc32 (C++ implementation, ~100x faster)
 * - Browser: Uses lookup table optimization
 *
 * The polynomial used is the standard CRC-32 IEEE 802.3:
 * x^32 + x^26 + x^23 + x^22 + x^16 + x^12 + x^11 + x^10 + x^8 + x^7 + x^5 + x^4 + x^2 + x + 1
 * Represented as 0xEDB88320 in reversed (LSB-first) form
 */

import type * as zlibType from "zlib";
import { isNode } from "@utils/env";
import { crc32JS, crc32UpdateJS, crc32Finalize } from "@archive/compression/crc32.base";

// Lazy-loaded zlib module for Node.js
let _zlib: typeof zlibType | null = null;
let _zlibLoading: Promise<typeof zlibType | null> | null = null;
let _zlibInitStarted = false;

/**
 * Lazily initialize zlib loading in Node.js.
 * Called on first use rather than at module load time.
 */
function ensureZlibLoading(): void {
  if (_zlibInitStarted) {
    return;
  }
  _zlibInitStarted = true;
  if (isNode()) {
    _zlibLoading = import("zlib")
      .then(module => {
        _zlib = (module as { default?: typeof zlibType }).default ?? (module as typeof zlibType);
        return _zlib;
      })
      .catch(() => {
        _zlib = null;
        return null;
      });
  }
}

/**
 * Synchronously ensure zlib is loaded for Node.js.
 * Used by the sync deflate path where the async dynamic import may not have
 * resolved yet. Falls back to `require()` which is synchronous in Node.js.
 */
export function ensureZlibSync(): void {
  if (_zlib || !isNode()) {
    return;
  }
  try {
    // oxlint-disable-next-line typescript/no-require-imports
    _zlib = require("zlib") as typeof zlibType;
    _zlibInitStarted = true;
  } catch {
    // Bundler or non-Node environment — JS fallback will be used
  }
}

/**
 * Calculate CRC32 checksum for the given data
 * Uses native zlib.crc32 in Node.js for ~100x better performance
 *
 * @param data - Input data as Uint8Array or Buffer
 * @returns CRC32 checksum as unsigned 32-bit integer
 *
 * @example
 * ```ts
 * const data = new TextEncoder().encode("Hello, World!");
 * const checksum = crc32(data);
 * console.log(checksum.toString(16)); // "ec4ac3d0"
 * ```
 */
export function crc32(data: Uint8Array): number {
  ensureZlibLoading();
  // Use native zlib.crc32 if available (Node.js)
  if (_zlib && typeof _zlib.crc32 === "function") {
    return _zlib.crc32(data) >>> 0;
  }
  // Fallback to JS implementation
  return crc32JS(data);
}

/**
 * Ensure zlib is loaded (for use before calling crc32)
 */
export async function ensureCrc32(): Promise<void> {
  ensureZlibLoading();
  if (_zlibLoading) {
    await _zlibLoading;
  }
}

/**
 * Calculate CRC32 incrementally (useful for streaming)
 * Call with initial crc of 0xffffffff, then finalize with crc32Finalize
 * In Node.js, this uses native zlib.crc32 when available for performance.
 * The internal CRC state remains the same as the JS table implementation:
 * - initial state: 0xffffffff
 * - finalize: xor with 0xffffffff
 *
 * @param crc - Current CRC value (start with 0xffffffff)
 * @param data - Input data chunk
 * @returns Updated CRC value (not finalized)
 *
 * @example
 * ```ts
 * let crc = 0xffffffff;
 * crc = crc32Update(crc, chunk1);
 * crc = crc32Update(crc, chunk2);
 * const checksum = crc32Finalize(crc);
 * ```
 */
export function crc32Update(crc: number, data: Uint8Array): number {
  ensureZlibLoading();
  // If available, use native zlib.crc32 but preserve our internal state shape.
  // zlib.crc32 returns a finalized CRC value and can accept the previous finalized
  // CRC as the second parameter (chainable). Our internal state is the inverted
  // (non-finalized) CRC, so convert with xor before/after.
  if (_zlib && typeof _zlib.crc32 === "function") {
    const prevFinal = (crc ^ 0xffffffff) >>> 0;
    const nextFinal = _zlib.crc32(data, prevFinal) >>> 0;
    return (nextFinal ^ 0xffffffff) >>> 0;
  }

  return crc32UpdateJS(crc, data);
}

/**
 * Finalize CRC32 calculation
 * XOR with 0xffffffff and convert to unsigned 32-bit
 *
 * @param crc - CRC value from crc32Update
 * @returns Final CRC32 checksum
 */
export { crc32Finalize };
