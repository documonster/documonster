import { Archive, Compression } from "@archive";
import { describe, expect, it } from "vitest";

describe("TarGzArchive.bytesSync", () => {
  it("returns gzip-compressed bytes, like bytes()", async () => {
    const archive = Archive.createTarGz();
    archive.add("a.txt", "hello");
    const sync = archive.toBufferSync();
    expect(Compression.isGzipData(sync)).toBe(true);
  });
});

/** Node's zlib produces the same deflate stream however its input is split; browsers do not. */
const isNode =
  typeof (globalThis as { process?: { versions?: { node?: string } } }).process?.versions?.node ===
  "string";

describe("tar.gz sync and async output agree", () => {
  const modTime = new Date(0);
  // Compressible text, plus pseudo-random bytes long enough to span several stored blocks
  const noise = new Uint8Array(200_000).map((_, i) => (i * 2654435761) >>> 24);
  const entries: Array<[string, string | Uint8Array]> = [
    ["a.txt", "hello ".repeat(2000)],
    ["dir/noise.bin", noise],
    ["empty.txt", ""]
  ];

  /** Byte equality as a boolean: a deep-equality diff of two large arrays is very slow. */
  function same(a: Uint8Array, b: Uint8Array): boolean {
    return a.length === b.length && a.every((byte, i) => byte === b[i]);
  }

  async function streamed(level: number): Promise<Uint8Array> {
    const archive = Archive.createTarGz({ modTime, level });
    for (const [name, data] of entries) {
      archive.add(name, data);
    }
    const parts: Uint8Array[] = [];
    for await (const chunk of archive.stream()) {
      parts.push(chunk);
    }
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let offset = 0;
    for (const part of parts) {
      out.set(part, offset);
      offset += part.length;
    }
    return out;
  }

  for (const level of [0, 1, 6, 9]) {
    it(`level ${level}`, async () => {
      const sync = Archive.packTarGzSync(entries, { modTime, level });
      const outputs = [await Archive.packTarGz(entries, { modTime, level }), await streamed(level)];
      const tar = Archive.packTarSync(entries, { modTime });
      expect(same(Compression.gunzipSync(sync), tar)).toBe(true);
      for (const output of outputs) {
        expect(same(Compression.gunzipSync(output), tar)).toBe(true);
        // Stored (level 0) gzip is framed by this library, so it is identical everywhere
        if (isNode || level === 0) {
          expect([output.length, same(output, sync)]).toEqual([sync.length, true]);
        }
      }
    });
  }
});
