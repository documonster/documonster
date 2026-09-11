/**
 * CRC32 for ZIP files — browser variant.
 *
 * The portable lookup-table implementation, with no `zlib` and no Node built-in of any kind. `crc32.ts` is
 * the sibling that uses native `zlib.crc32`.
 *
 * The polynomial is the standard CRC-32 IEEE 802.3:
 * x^32 + x^26 + x^23 + x^22 + x^16 + x^12 + x^11 + x^10 + x^8 + x^7 + x^5 + x^4 + x^2 + x + 1
 * Represented as 0xEDB88320 in reversed (LSB-first) form.
 *
 * @example
 * ```ts
 * let crc = 0xffffffff;
 * crc = crc32Update(crc, chunk1);
 * crc = crc32Update(crc, chunk2);
 * const checksum = crc32Finalize(crc);
 * ```
 */

export {
  crc32JS as crc32,
  crc32UpdateJS as crc32Update,
  crc32Finalize
} from "@archive/compression/crc32.base";
