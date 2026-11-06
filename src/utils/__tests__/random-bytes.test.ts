import { randomBytes } from "@utils/random-bytes";
import { describe, expect, it } from "vitest";

describe("randomBytes", () => {
  it("returns the length asked for", () => {
    expect(randomBytes(0)).toHaveLength(0);
    expect(randomBytes(16)).toHaveLength(16);
  });

  // Web Crypto's `getRandomValues` refuses more than 65,536 bytes in one call; the browser form
  // called it once, so a larger length threw `QuotaExceededError` there.
  it("fills a length above Web Crypto's per-call quota", () => {
    const bytes = randomBytes(65_536 * 2 + 3);
    expect(bytes).toHaveLength(131_075);
    // Every quota-sized piece was filled, the 3-byte tail included: an unfilled piece stays all zero
    // (a filled one does so with probability 2^-24 at worst).
    for (const start of [0, 65_536, 131_072]) {
      expect(bytes.subarray(start, start + 65_536).some(b => b !== 0)).toBe(true);
    }
  });
});
