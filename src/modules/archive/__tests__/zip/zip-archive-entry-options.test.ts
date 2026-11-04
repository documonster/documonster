/**
 * `ZipArchive` writes the same entries whichever output it is asked for. The in-memory fast
 * path behind `toBuffer()` used to forward only name, data, level, modTime, comment and
 * encoding — so `mode`, `atime`/`ctime`/`birthTime`, the MS-DOS and external attributes and
 * `versionMadeBy` reached `stream()` and vanished from `toBuffer()` — and `toBufferSync()`
 * dropped `level` too. A per-entry `zip64` was honoured by the stream alone.
 */
import { Archive } from "@archive";
import type { ZipEntryOptions } from "@archive/zip/zip-archive";
import { describe, expect, it } from "vitest";

const MOD = new Date(Date.UTC(2021, 4, 6, 7, 8, 10));
const ACCESS = new Date(Date.UTC(2022, 1, 2, 3, 4, 6));
const BIRTH = new Date(Date.UTC(2020, 0, 1, 0, 0, 0));
const TEXT = "compress me ".repeat(200);

const ENTRIES: Array<[string, string, ZipEntryOptions]> = [
  ["stored.txt", TEXT, { level: 0, mode: 0o100600, modTime: MOD, comment: "first" }],
  [
    "deflated.txt",
    TEXT,
    { level: 9, mode: 0o100755, modTime: MOD, atime: ACCESS, ctime: ACCESS, birthTime: BIRTH }
  ],
  ["dos.txt", "x", { msDosAttributes: 0x01, versionMadeBy: 20, modTime: MOD }],
  ["ext.txt", "y", { externalAttributes: 0x81a40000, modTime: MOD }],
  ["z64.txt", "z", { zip64: true, modTime: MOD }]
];

function build(): ReturnType<typeof Archive.createZip> {
  const zip = Archive.createZip({ timestamps: "dos+utc+ntfs", modTime: MOD, smartStore: false });
  for (const [name, text, options] of ENTRIES) {
    zip.add(name, text, options);
  }
  return zip;
}

async function headers(bytes: Uint8Array) {
  const reader = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(bytes));
  // Sorted: the buffered builder orders entries by name unless `noSort`, the stream keeps
  // input order (documented on `ZipOptions.noSort`).
  const entries = [...reader.getEntries()].sort((a, b) => a.path.localeCompare(b.path));
  return entries.map(entry => ({
    path: entry.path,
    compressionMethod: entry.compressionMethod,
    // Not `compressedSize`: the streaming encoder flushes per chunk, which is a different
    // but equally valid DEFLATE stream of the same data.
    uncompressedSize: entry.uncompressedSize,
    crc32: entry.crc32,
    comment: entry.comment,
    externalAttributes: entry.externalAttributes,
    mode: entry.mode,
    versionMadeBy: entry.versionMadeBy,
    dosTime: entry.dosTime,
    extraFields: extraFields(entry.extraField ?? new Uint8Array(0))
  }));
}

/**
 * The extra fields by header ID. A ZIP64 field (0x0001) keeps only its uncompressed size: the
 * compressed size and the local-header offset it also carries depend on the DEFLATE stream
 * and on the data descriptors a stream writes, which differ by design.
 */
function extraFields(extra: Uint8Array): Array<[number, number[]]> {
  const fields: Array<[number, number[]]> = [];
  for (let i = 0; i + 4 <= extra.length;) {
    const id = extra[i]! | (extra[i + 1]! << 8);
    const size = extra[i + 2]! | (extra[i + 3]! << 8);
    const data = Array.from(extra.subarray(i + 4, i + 4 + size));
    fields.push([id, id === 0x0001 ? data.slice(0, 8) : data]);
    i += 4 + size;
  }
  return fields;
}

async function streamed(zip: ReturnType<typeof Archive.createZip>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of zip.stream()) {
    chunks.push(chunk);
  }
  return Archive.concatUint8Arrays(chunks);
}

describe("ZipArchive entry options on every output path", () => {
  it("writes identical central-directory headers from toBuffer, toBufferSync and stream", async () => {
    const fromStream = await headers(await streamed(build()));
    expect(await headers(await build().toBuffer())).toEqual(fromStream);
    expect(await headers(build().toBufferSync())).toEqual(fromStream);
  });

  it("reads every entry back from every output", async () => {
    for (const bytes of [
      await streamed(build()),
      await build().toBuffer(),
      build().toBufferSync()
    ]) {
      const files = await Archive.unpackZip(bytes);
      for (const [name, text] of ENTRIES) {
        expect(new TextDecoder().decode(files.get(name)?.data)).toBe(text);
      }
    }
  });

  it("honours the options it is given", async () => {
    const [deflated, dos, ext, stored, z64] = await headers(build().toBufferSync());
    expect(stored!.compressionMethod).toBe(0);
    expect(deflated!.compressionMethod).toBe(8);
    expect(stored!.mode & 0o777).toBe(0o600);
    expect(deflated!.mode & 0o777).toBe(0o755);
    expect(dos!.externalAttributes & 0xff).toBe(0x01);
    expect(ext!.externalAttributes >>> 0).toBe(0x81a40000);
    expect(stored!.comment).toBe("first");
    // The ZIP64 extended-information extra field, and a size read through it.
    expect(z64!.extraFields.map(([id]) => id)).toContain(0x0001);
  });
});
