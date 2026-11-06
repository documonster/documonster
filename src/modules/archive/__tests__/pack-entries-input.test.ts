/**
 * The one-shot builders take their entries as a `ReadonlyMap` of name → source (any map the
 * caller may not mutate, not only a `Map`), as the other keyed inputs of the library do, or as
 * an iterable of `[name, source]` pairs / `{ name, source }` objects.
 */
import { Archive } from "@archive/index";
import type { ArchiveSource } from "@archive/index";
import { describe, expect, it } from "vitest";

const decoder = new TextDecoder();

/** A read-only map that is not a `Map`. */
class FrozenEntries implements ReadonlyMap<string, string> {
  readonly #map: Map<string, string>;
  constructor(entries: Array<[string, string]>) {
    this.#map = new Map(entries);
  }
  get size() {
    return this.#map.size;
  }
  get(key: string) {
    return this.#map.get(key);
  }
  has(key: string) {
    return this.#map.has(key);
  }
  forEach(cb: (value: string, key: string, map: ReadonlyMap<string, string>) => void) {
    this.#map.forEach((v, k) => cb(v, k, this));
  }
  entries() {
    return this.#map.entries();
  }
  keys() {
    return this.#map.keys();
  }
  values() {
    return this.#map.values();
  }
  [Symbol.iterator]() {
    return this.#map[Symbol.iterator]();
  }
}

describe("pack* take a ReadonlyMap", () => {
  const readonlyMap: ReadonlyMap<string, ArchiveSource> = new Map([["a.txt", "A"]]);
  const sync: ReadonlyMap<string, string> = new FrozenEntries([["a.txt", "A"]]);

  it("packZip / packZipSync", async () => {
    const zips = [
      await Archive.packZip(readonlyMap),
      await Archive.packZip(new FrozenEntries([["a.txt", "A"]])),
      Archive.packZipSync(sync)
    ];
    for (const zip of zips) {
      expect(decoder.decode(Archive.unpackZipSync(zip).get("a.txt")!.data)).toBe("A");
    }
  });

  it("packTar / packTarSync / packTarGz / packTarGzSync", async () => {
    const tars = [await Archive.packTar(readonlyMap), Archive.packTarSync(sync)];
    for (const tar of tars) {
      expect(decoder.decode(Archive.unpackTarSync(tar).get("a.txt")!.data)).toBe("A");
    }
    const tgzs = [await Archive.packTarGz(readonlyMap), Archive.packTarGzSync(sync)];
    for (const tgz of tgzs) {
      expect(decoder.decode(Archive.unpackTarGzSync(tgz).get("a.txt")!.data)).toBe("A");
    }
  });
});
