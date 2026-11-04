/**
 * The portable CRC-32 — the browser's implementation — against a plain bit-by-bit reference,
 * over both of its loops (aligned words, which assume a little-endian engine, and bytes), at
 * every alignment and across the lengths where it changes strategy. Run in Node and in a
 * browser.
 */

import { crc32Finalize, crc32JS, crc32UpdateSliced } from "@archive/compression/crc32.base";
import { describe, expect, it } from "vitest";

function reference(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let k = 0; k < 8; k++) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function sample(length: number): Uint8Array {
  const out = new Uint8Array(length);
  let s = 7;
  for (let i = 0; i < length; i++) {
    s = (s * 1103515245 + 12345) >>> 0;
    out[i] = s >>> 24;
  }
  return out;
}

export function defineCrc32PortableSuite(): void {
  describe("portable CRC-32 (slicing-by-16)", () => {
    const data = sample(5000);
    const lengths = [0, 1, 3, 15, 16, 17, 31, 255, 256, 257, 271, 1024, 4093];

    for (const wordLoads of [true, false]) {
      it(`matches the reference at every alignment (${wordLoads ? "word" : "byte"} loads)`, () => {
        for (let offset = 0; offset < 8; offset++) {
          for (const length of lengths) {
            const view = data.subarray(offset, offset + length);
            expect(crc32Finalize(crc32UpdateSliced(0xffffffff, view, wordLoads))).toBe(
              reference(view)
            );
          }
        }
      });

      it(`chains across chunks of any size (${wordLoads ? "word" : "byte"} loads)`, () => {
        for (const step of [1, 7, 16, 100, 300, 1000]) {
          let state = 0xffffffff;
          for (let i = 0; i < data.length; i += step) {
            state = crc32UpdateSliced(state, data.subarray(i, i + step), wordLoads);
            expect(state >>> 0).toBe(state);
          }
          expect(crc32Finalize(state)).toBe(reference(data));
        }
      });
    }

    it("is the published one-shot form", () => {
      expect(crc32JS(new TextEncoder().encode("Hello, World!"))).toBe(0xec4ac3d0);
      expect(crc32JS(data)).toBe(reference(data));
    });
  });
}
