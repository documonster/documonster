/**
 * Encrypting and decrypting an Office package let the event loop run.
 *
 * The password derivation is 100,000 digests, and it ran as one synchronous loop — three times per file, once for each
 * key it derives, though all three share the spin. The package segments and the data-integrity HMAC over the whole
 * package were synchronous too. A file opened on a server stopped every other request for a quarter of a second.
 */

import { decryptPackage, deriveAgileKey, encryptPackage } from "@archive/office-crypto/agile";
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

/** How many turns the event loop takes while `work` runs. */
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
  // Taken after the first turn has been queued, so only turns `work` lets through are counted.
  const before = turns;
  try {
    await work();
  } finally {
    running = false;
  }
  return turns - before;
}

describe("Agile encryption and the event loop", () => {
  const zip = new Uint8Array(64 * 1024).map((_, i) => i % 251);

  it("lets the event loop run while it encrypts and while it decrypts", async () => {
    let encrypted: Uint8Array | undefined;
    expect(
      await turnsDuring(async () => (encrypted = await encryptPackage(zip, "pw")))
    ).toBeGreaterThan(1);
    let decrypted: Uint8Array | undefined;
    expect(
      await turnsDuring(async () => (decrypted = await decryptPackage(encrypted!, "pw")))
    ).toBeGreaterThan(1);
    expect(decrypted).toEqual(zip);
  });

  // The encrypt and decrypt above yield over the package's segments too; the spin has nothing else in it.
  it("lets the event loop run while it spins a password", async () => {
    const info = {
      keySalt: new Uint8Array(16),
      spinCount: 100_000,
      hashAlgorithm: "SHA512",
      keyBits: 256
    };
    expect(await turnsDuring(() => deriveAgileKey("pw", info, new Uint8Array(8)))).toBeGreaterThan(
      5
    );
  });
});
