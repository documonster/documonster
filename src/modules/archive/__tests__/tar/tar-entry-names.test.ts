import { Archive } from "@archive";
import type { TarEntry, UnpackedTarEntry } from "@archive";
import { describe, expect, expectTypeOf, it } from "vitest";

async function tarBytes(): Promise<Uint8Array> {
  return Archive.packTar([
    ["a.txt", "alpha"],
    ["b.txt", "bravo"]
  ]);
}

async function collect(iterable: AsyncIterable<TarEntry>): Promise<TarEntry[]> {
  const out: TarEntry[] = [];
  for await (const entry of iterable) {
    out.push(entry);
  }
  return out;
}

describe("parsed TAR entries read with the reader-entry names", () => {
  it("bytes() / discard() on every parse form, and the deprecated data() / skip() agree", async () => {
    const tar = await tarBytes();
    const tgz = await Archive.packTarGz([
      ["a.txt", "alpha"],
      ["b.txt", "bravo"]
    ]);
    const forms: TarEntry[][] = [
      Archive.parseTar(tar),
      await collect(Archive.parseTarStream(tar)),
      await Archive.parseTarGz(tgz),
      await collect(Archive.parseTarGzStream(tgz))
    ];
    for (const entries of forms) {
      expect(entries.map(e => e.info.path)).toEqual(["a.txt", "b.txt"]);
      const [first] = entries;
      const bytes = await first.bytes();
      expect(new TextDecoder().decode(bytes)).toBe("alpha");
      expect(await first.text()).toBe("alpha");
      // Deprecated names keep their behaviour.
      expect(await first.data()).toEqual(bytes);
      await expect(first.skip()).resolves.toBeUndefined();
      expect(first.discard()).toBeUndefined();
    }
  });

  it("parseTarStream and parseTarGzStream yield the same named entry type", () => {
    expectTypeOf(Archive.parseTarStream).returns.toEqualTypeOf<AsyncGenerator<TarEntry>>();
    expectTypeOf(Archive.parseTarGzStream).returns.toEqualTypeOf<AsyncGenerator<TarEntry>>();
  });

  it("unpackTar's map values are UnpackedTarEntry", async () => {
    const map = await Archive.unpackTar(await tarBytes());
    expectTypeOf(map).toEqualTypeOf<Map<string, UnpackedTarEntry>>();
    expectTypeOf(Archive.unpackTarSync).returns.toEqualTypeOf<Map<string, UnpackedTarEntry>>();
    expectTypeOf(Archive.unpackTarGzSync).returns.toEqualTypeOf<Map<string, UnpackedTarEntry>>();
    expect(new TextDecoder().decode(map.get("b.txt")!.data)).toBe("bravo");
  });
});
