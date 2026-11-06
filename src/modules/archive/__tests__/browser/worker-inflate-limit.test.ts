import { compressSync, decompress } from "@archive/compression/compress";
import { hasWorkerSupport, inflateWithPool } from "@archive/compression/worker-pool/index.browser";
import { ArchiveLimitExceededError } from "@archive/core/errors";
import { describe, it, expect } from "vitest";

// The worker enforces maxOutputLength on its own side of postMessage, and the
// pool has to rebuild the hit as an ArchiveLimitExceededError — an Error's class does
// not survive structured clone. Only a real browser worker exercises that.
describe("worker inflate - maxOutputLength", () => {
  const original = new Uint8Array(200_000).fill(7);
  const deflated = compressSync(original);

  it("is a real worker environment", () => {
    expect(hasWorkerSupport()).toBe(true);
  });

  it("rejects with ArchiveLimitExceededError when the output exceeds the bound", async () => {
    const err = await inflateWithPool(deflated, { maxOutputLength: 1000 }).catch(e => e);
    expect(err).toBeInstanceOf(ArchiveLimitExceededError);
    expect((err as ArchiveLimitExceededError).limit).toBe("maxOutputLength");
    expect((err as ArchiveLimitExceededError).allowed).toBe(1000);
  }, 20_000);

  it("inflates normally when the output fits", async () => {
    const out = await inflateWithPool(deflated, { maxOutputLength: original.length });
    expect(out).toEqual(original);
  }, 20_000);

  it("surfaces the limit through decompress({ useWorker: true }) without falling back", async () => {
    await expect(
      decompress(deflated, { useWorker: true, maxOutputLength: 1000 })
    ).rejects.toBeInstanceOf(ArchiveLimitExceededError);
    expect(
      await decompress(deflated, { useWorker: true, maxOutputLength: original.length })
    ).toEqual(original);
  }, 20_000);
});
