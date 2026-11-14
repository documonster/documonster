/**
 * The browser's synchronous `hash` is pure JS, and agrees with Web Crypto byte for byte on every algorithm an Office
 * encryption header may name — the lengths around a block boundary included, where padding goes wrong.
 */
import { hash } from "@utils/crypto";
import { describe, expect, it } from "vitest";

async function webCryptoDigest(algorithm: string, data: Uint8Array): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest(algorithm, data.slice()));
}

describe("hash (browser)", () => {
  const lengths = [0, 1, 3, 55, 56, 57, 63, 64, 65, 111, 112, 113, 127, 128, 129, 1000, 4096];

  it.each(["SHA-1", "SHA-256", "SHA-384", "SHA-512"])("%s matches Web Crypto", async algorithm => {
    for (const length of lengths) {
      const data = new Uint8Array(length).map((_, i) => (i * 31 + length) % 256);
      expect(hash(algorithm, data), `${algorithm}, ${length} bytes`).toEqual(
        await webCryptoDigest(algorithm, data)
      );
    }
  });

  // A large input runs many blocks through one state; a view at an offset must be read from its offset.
  it.each(["SHA-1", "SHA-256", "SHA-384", "SHA-512"])(
    "%s matches Web Crypto on a large view",
    async algorithm => {
      const backing = new Uint8Array((1 << 20) + 11).map((_, i) => (i * 7) % 256);
      const data = backing.subarray(5, backing.length - 3);
      expect(hash(algorithm, data)).toEqual(await webCryptoDigest(algorithm, data));
    }
  );

  it("refuses an algorithm it does not have", () => {
    expect(() => hash("SHA-224", new Uint8Array(0))).toThrow(/unsupported algorithm/);
  });
});
