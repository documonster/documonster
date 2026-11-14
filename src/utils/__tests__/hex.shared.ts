/**
 * `@utils/hex` reads and writes as `Uint8Array.fromHex` / `toHex` do, whether the platform has
 * them or not: each case runs with the native methods and with them removed.
 */
import { decodeHex, encodeHex, isHex } from "@utils/hex";
import { describe, expect, it } from "vitest";

const DECODES: ReadonlyArray<readonly [string, number[]]> = [
  ["", []],
  ["00", [0]],
  ["ff", [255]],
  ["FF", [255]],
  ["0aBc", [10, 188]],
  ["deadbeef", [0xde, 0xad, 0xbe, 0xef]]
];

const REFUSED = ["0", "abc", "0x00", "00 ff", "gg", "0g", " 00", "00\n", "é0"];

function variants(): Array<[string, () => () => void]> {
  const ctor = Uint8Array as unknown as { fromHex?: unknown };
  const out: Array<[string, () => () => void]> = [];
  if (typeof ctor.fromHex === "function") {
    out.push(["Uint8Array.fromHex", () => () => {}]);
  }
  out.push([
    "the fallback",
    () => {
      const from = Object.getOwnPropertyDescriptor(Uint8Array, "fromHex");
      const to = Object.getOwnPropertyDescriptor(Uint8Array.prototype, "toHex");
      delete ctor.fromHex;
      delete (Uint8Array.prototype as { toHex?: unknown }).toHex;
      return () => {
        if (from) {
          Object.defineProperty(Uint8Array, "fromHex", from);
        }
        if (to) {
          Object.defineProperty(Uint8Array.prototype, "toHex", to);
        }
      };
    }
  ]);
  return out;
}

export function defineHexSuite(): void {
  describe("@utils/hex", () => {
    for (const [name, setUp] of variants()) {
      describe(`with ${name}`, () => {
        const run = (body: () => void) => () => {
          const restore = setUp();
          try {
            body();
          } finally {
            restore();
          }
        };

        it.each(DECODES)("reads %j", (text, bytes) =>
          run(() => {
            expect([...decodeHex(text)!]).toEqual(bytes);
            expect(isHex(text)).toBe(true);
          })()
        );

        it.each(REFUSED)("refuses %j", text =>
          run(() => {
            expect(decodeHex(text)).toBeUndefined();
            expect(isHex(text)).toBe(false);
          })()
        );

        it(
          "writes lower case, a view's own bytes only",
          run(() => {
            expect(encodeHex(Uint8Array.of(0, 10, 171, 255))).toBe("000aabff");
            expect(encodeHex(Uint8Array.of(1, 2, 3).subarray(1, 2))).toBe("02");
            expect(encodeHex(new Uint8Array(0))).toBe("");
          })
        );

        it(
          "handles a large input",
          run(() => {
            const bytes = new Uint8Array(8 * 1024 * 1024).map((_, i) => (i * 7) & 255);
            const text = encodeHex(bytes);
            expect(text.length).toBe(bytes.length * 2);
            expect(text.slice(0, 8)).toBe("00070e15");
            expect(decodeHex(text)).toEqual(bytes);
            expect(decodeHex(`${text}0`)).toBeUndefined();
          })
        );

        it(
          "round-trips every byte value",
          run(() => {
            const all = Uint8Array.from({ length: 256 }, (_, i) => i);
            expect(decodeHex(encodeHex(all))).toEqual(all);
            expect(decodeHex(encodeHex(all).toUpperCase())).toEqual(all);
          })
        );
      });
    }
  });
}
