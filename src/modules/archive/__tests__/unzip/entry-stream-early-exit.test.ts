import { Archive } from "@archive";
import { describe, expect, it } from "vitest";

/** Compressible but not trivially so: several inflate output chunks per entry. */
const payload = new Uint8Array(400_000).map((_, i) => ((i * 2654435761) >>> 26) + (i % 7));

/** Give a late decoder `error` event the chance to surface as an unhandled error. */
const settle = (): Promise<void> => new Promise(resolve => setTimeout(resolve, 20));

describe("leaving an entry's stream() early", () => {
  it("does not hang, and the next entry still reads", { timeout: 10_000 }, async () => {
    const zip = await Archive.packZip([
      ["a.bin", payload],
      ["b.txt", "beta"]
    ]);
    const seen: Array<[string, number]> = [];
    for await (const entry of Archive.unzip(zip).entries()) {
      if (entry.path === "a.bin") {
        for await (const chunk of entry.stream()) {
          seen.push([entry.path, chunk.length > 0 ? 1 : 0]);
          break;
        }
      } else {
        seen.push([entry.path, (await entry.bytes()).length]);
      }
    }
    expect(seen).toEqual([
      ["a.bin", 1],
      ["b.txt", 4]
    ]);
  });

  it("does not raise an unhandled error over damaged data", { timeout: 10_000 }, async () => {
    const zip = await Archive.packZip([["a.bin", payload]]);
    const damaged = zip.slice();
    for (let i = 3000; i < 3100; i++) {
      damaged[i]! ^= 0x55;
    }
    for await (const entry of Archive.unzip(damaged).entries()) {
      try {
        for await (const _ of entry.stream()) {
          break;
        }
      } catch {
        // A decode error is a fine answer; an unhandled one is not
      }
    }
    await settle();
  });
});
