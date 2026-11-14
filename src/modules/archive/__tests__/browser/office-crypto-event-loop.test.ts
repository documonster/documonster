/**
 * Encrypting and decrypting an Office package let the event loop run in a browser too, where the password spin is
 * hashed in pure JS — SHA-512 by default, SHA-1 for an older file — rather than one awaited Web Crypto digest per step.
 */
import { decryptPackage, encryptPackage } from "@archive/office-crypto/agile";
import { scheduleMacrotask } from "@utils/yield";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// The budget is measured with `performance.now`, so on a fast enough machine a whole read could fit inside it and
// never yield — a test that passes or fails by the speed of the machine. A clock that advances a millisecond per
// reading makes the budget run out by the amount of work done instead: what is checked is that the reader asks
// between its pieces.
beforeEach(() => {
  let clock = 0;
  vi.spyOn(performance, "now").mockImplementation(() => (clock += 1));
});
afterEach(() => {
  vi.restoreAllMocks();
});

async function turnsDuring(work: () => Promise<unknown>): Promise<number> {
  let turns = 0;
  let running = true;
  const tick = (): void => {
    if (running) {
      turns++;
      scheduleMacrotask(tick);
    }
  };
  scheduleMacrotask(tick);
  const before = turns;
  try {
    await work();
  } finally {
    running = false;
  }
  return turns - before;
}

describe("Agile encryption and the event loop (browser)", () => {
  const zip = new Uint8Array(64 * 1024).map((_, i) => i % 251);

  // SHA-1's 20-byte digest is too short for a 256-bit key, so it is paired with AES-128 as Office 2010 paired it.
  it.each([
    ["SHA-512", 256],
    ["SHA-1", 128]
  ] as const)(
    "lets the event loop run while it encrypts and decrypts with %s",
    async (hashAlgorithm, keyBits) => {
      let encrypted: Uint8Array | undefined;
      expect(
        await turnsDuring(
          async () => (encrypted = await encryptPackage(zip, "pw", { hashAlgorithm, keyBits }))
        )
      ).toBeGreaterThan(1);
      let decrypted: Uint8Array | undefined;
      expect(
        await turnsDuring(async () => (decrypted = await decryptPackage(encrypted!, "pw")))
      ).toBeGreaterThan(1);
      expect(decrypted).toEqual(zip);
    }
  );
});
