import { loadIife } from "@test/browser/load-iife";
import { describe, it, expect, beforeAll } from "vitest";

/**
 * Smoke test for the shipped `documonster.archive.iife.min.js` bundle: asserts
 * the `Documonster.Archive` surface and a gzip → gunzip round-trip in a
 * browser. (Source-import browser tests already cover archive internals; this
 * verifies the delivered IIFE artifact itself runs.)
 */
describe("Documonster.Archive IIFE bundle", () => {
  let Archive: any;

  beforeAll(async () => {
    Archive = await loadIife("archive", "Archive");
  }, 60000);

  it("exposes core compression helpers", () => {
    for (const member of ["gzip", "gunzip", "crc32", "Archive"]) {
      expect(Archive[member], `Documonster.Archive.${member}`).toBeTruthy();
    }
  });

  it("round-trips data through gzip → gunzip", async () => {
    const input = new TextEncoder().encode("Documonster archive IIFE round-trip");
    const compressed: Uint8Array = await Archive.gzip(input);
    expect(compressed).toBeInstanceOf(Uint8Array);
    expect(compressed.byteLength).toBeGreaterThan(0);

    const restored: Uint8Array = await Archive.gunzip(compressed);
    expect(new TextDecoder().decode(restored)).toBe("Documonster archive IIFE round-trip");
  });
});
