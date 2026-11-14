/**
 * `@utils/base64` reads the same text the same way whichever decoder the platform offers: the
 * native `Uint8Array.fromBase64`, Node's `Buffer`, or `atob`. Each case runs once per decoder
 * available here, with the faster ones removed to reach the next.
 */
import { decodeBase64, encodeBase64, isBase64 } from "@utils/base64";
import { describe, expect, it } from "vitest";

const DECODES: ReadonlyArray<readonly [string, number[]]> = [
  ["", []],
  ["YQ", [97]],
  ["YQ==", [97]],
  ["YWJj", [97, 98, 99]],
  ["YWJjZA", [97, 98, 99, 100]],
  ["YWJjZA==", [97, 98, 99, 100]],
  ["YW Jj\n", [97, 98, 99]],
  ["Y Q=\t=", [97]],
  // Bits left over in the last character are ignored, as every platform decoder ignores them
  ["YR==", [97]],
  ["+/+/", [251, 255, 191]]
];

const REFUSED = ["Y", "YQ=", "YWJj=", "YWJj===", "YQ==YQ==", "YW!Jj", "-_8", "=", "YQ=a"];

/** The ways a platform can decode, each reached by removing the ones before it. */
function decoderVariants(): Array<[string, () => () => void]> {
  const proto = Uint8Array as unknown as { fromBase64?: unknown };
  const removeNative = (): (() => void) => {
    const from = Object.getOwnPropertyDescriptor(Uint8Array, "fromBase64");
    const to = Object.getOwnPropertyDescriptor(Uint8Array.prototype, "toBase64");
    delete proto.fromBase64;
    delete (Uint8Array.prototype as { toBase64?: unknown }).toBase64;
    return () => {
      if (from) {
        Object.defineProperty(Uint8Array, "fromBase64", from);
      }
      if (to) {
        Object.defineProperty(Uint8Array.prototype, "toBase64", to);
      }
    };
  };
  const variants: Array<[string, () => () => void]> = [];
  if (typeof proto.fromBase64 === "function") {
    variants.push(["Uint8Array.fromBase64", () => () => {}]);
  }
  variants.push([
    typeof (globalThis as { Buffer?: unknown }).Buffer === "function" ? "Buffer" : "atob",
    removeNative
  ]);
  return variants;
}

export function defineBase64Suite(): void {
  describe("@utils/base64", () => {
    for (const [decoder, setUp] of decoderVariants()) {
      describe(`with ${decoder}`, () => {
        it.each(DECODES)("reads %j", (text, bytes) => {
          const restore = setUp();
          try {
            expect([...decodeBase64(text)!]).toEqual(bytes);
            expect(isBase64(text)).toBe(true);
          } finally {
            restore();
          }
        });

        it.each(REFUSED)("refuses %j", text => {
          const restore = setUp();
          try {
            expect(decodeBase64(text)).toBeUndefined();
            expect(isBase64(text)).toBe(false);
          } finally {
            restore();
          }
        });

        it("returns memory of its own, never a view of a shared pool", () => {
          const restore = setUp();
          try {
            const bytes = decodeBase64("YWJj")!;
            expect(bytes.constructor).toBe(Uint8Array);
            expect(bytes.byteOffset).toBe(0);
            expect(bytes.buffer.byteLength).toBe(3);
          } finally {
            restore();
          }
        });

        it("encodes in either alphabet, with or without padding", () => {
          const restore = setUp();
          try {
            const bytes = Uint8Array.of(251, 255, 191, 97);
            expect(encodeBase64(bytes)).toBe("+/+/YQ==");
            expect(encodeBase64(bytes, "base64url")).toBe("-_-_YQ==");
            expect(encodeBase64(bytes, "base64url", true)).toBe("-_-_YQ");
            // A view encodes its own bytes, not its buffer's
            expect(encodeBase64(new Uint8Array([0, 97, 0]).subarray(1, 2))).toBe("YQ==");
          } finally {
            restore();
          }
        });

        it("round-trips a large input", () => {
          const restore = setUp();
          try {
            const bytes = new Uint8Array(200_000).map((_, i) => (i * 31) & 255);
            expect(decodeBase64(encodeBase64(bytes))).toEqual(bytes);
          } finally {
            restore();
          }
        });
      });
    }

    it("reads the URL-safe alphabet only when asked, and either when asked for either", () => {
      expect([...decodeBase64("-_8", "base64url")!]).toEqual([251, 255]);
      expect(decodeBase64("+/8", "base64url")).toBeUndefined();
      expect(decodeBase64("-_8")).toBeUndefined();
      expect([...decodeBase64("-/+_", "either")!]).toEqual([251, 255, 191]);
    });
  });
}
