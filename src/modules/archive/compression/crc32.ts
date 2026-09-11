/**
 * CRC32 for ZIP files — Node variant.
 *
 * `zlib` is imported **statically**, and that is the invariant this file exists to hold. Loading it lazily
 * leaves a window in which CRC32 runs on the portable lookup table instead, and the window cannot be closed
 * synchronously from an ES module. `crc32.browser.ts` is the sibling that has no `zlib` and uses the table.
 *
 * `zlib.crc32` has existed since Node 22.2 and `engines.node` is `>=22.13`, so it needs no capability probe.
 *
 * The polynomial is the standard CRC-32 IEEE 802.3:
 * x^32 + x^26 + x^23 + x^22 + x^16 + x^12 + x^11 + x^10 + x^8 + x^7 + x^5 + x^4 + x^2 + x + 1
 * Represented as 0xEDB88320 in reversed (LSB-first) form.
 */

import { crc32 as nativeCrc32 } from "zlib";

import { crc32Finalize } from "@archive/compression/crc32.base";

/**
 * Calculate CRC32 checksum for the given data
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
  return nativeCrc32(data) >>> 0;
}

/**
 * Calculate CRC32 incrementally (useful for streaming)
 * Call with initial crc of 0xffffffff, then finalize with crc32Finalize
 *
 * The internal state matches the lookup-table implementation, so the two are interchangeable:
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
  // `zlib.crc32` takes and returns a *finalized* CRC and chains through its second argument. The state here
  // is the inverted (non-finalized) form, so convert on the way in and out.
  const prevFinal = (crc ^ 0xffffffff) >>> 0;
  const nextFinal = nativeCrc32(data, prevFinal) >>> 0;
  return (nextFinal ^ 0xffffffff) >>> 0;
}

/**
 * Finalize CRC32 calculation
 * XOR with 0xffffffff and convert to unsigned 32-bit
 *
 * @param crc - CRC value from crc32Update
 * @returns Final CRC32 checksum
 */
export { crc32Finalize };
