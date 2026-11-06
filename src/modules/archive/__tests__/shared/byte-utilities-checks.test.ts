/**
 * `Archive`'s byte utilities, `unpackTarGzSync` and the `Compression` stream factories check
 * their arguments like the rest of the namespace: a wrong one is an `ArchiveError`, never a
 * silently wrong result or a raw `TypeError` / `RangeError`.
 */
import { Archive, ArchiveError, Compression } from "@archive";
import { describe, expect, it } from "vitest";

const bytes = new TextEncoder().encode("abc");

describe("Archive byte utilities", () => {
  it.each<[string, () => unknown]>([
    ["crc32(string)", () => Archive.crc32("abc" as never)],
    ["crc32(null)", () => Archive.crc32(null as never)],
    ["crc32Update(NaN, bytes)", () => Archive.crc32Update(Number.NaN, bytes)],
    ["crc32Update(-1, bytes)", () => Archive.crc32Update(-1, bytes)],
    ["crc32Update(0, string)", () => Archive.crc32Update(0, "abc" as never)],
    ["crc32Finalize(1.5)", () => Archive.crc32Finalize(1.5)],
    ["base64ToUint8Array(42)", () => Archive.base64ToUint8Array(42 as never)],
    ["base64ToUint8Array('!!!')", () => Archive.base64ToUint8Array("!!!")],
    ["uint8ArrayToBase64('abc')", () => Archive.uint8ArrayToBase64("abc" as never)],
    ["concatUint8Arrays([a, b], 1)", () => Archive.concatUint8Arrays([bytes, bytes], 1)],
    ["concatUint8Arrays([a], 100)", () => Archive.concatUint8Arrays([bytes], 100)],
    ["concatUint8Arrays(['x'])", () => Archive.concatUint8Arrays(["x"] as never)],
    ["stringToUint8Array(42)", () => Archive.stringToUint8Array(42 as never)],
    ["uint8ArrayToString(bytes, 'bogus')", () => Archive.uint8ArrayToString(bytes, "bogus")],
    ["uint8ArrayToString('abc')", () => Archive.uint8ArrayToString("abc" as never)]
  ])("%s is an ArchiveError", (_name, call) => {
    expect(call).toThrow(ArchiveError);
  });

  it("still computes what it computed", () => {
    const crc = Archive.crc32(bytes);
    expect(Archive.crc32Finalize(Archive.crc32Update(0xffffffff, bytes))).toBe(crc);
    expect(Archive.uint8ArrayToBase64(bytes)).toBe("YWJj");
    expect([...Archive.base64ToUint8Array("YWJj")]).toEqual([...bytes]);
    expect([...Archive.base64ToUint8Array("YW Jj\n")]).toEqual([...bytes]);
    expect([...Archive.concatUint8Arrays([bytes, bytes], 6)]).toEqual([...bytes, ...bytes]);
    expect([...Archive.concatUint8Arrays([bytes, bytes])]).toEqual([...bytes, ...bytes]);
    expect(Archive.uint8ArrayToString(Archive.stringToUint8Array("é"))).toBe("é");
    expect(Archive.uint8ArrayToString(new Uint8Array([0xe9]), "latin1")).toBe("é");
  });
});

describe("unpackTarGzSync sources", () => {
  const tgz = Archive.packTarGzSync([["a.txt", "A"]]);

  it("reads an ArrayBuffer or any ArrayBufferView, as unpackTarSync does", () => {
    expect([...Archive.unpackTarGzSync(tgz.slice().buffer as never).keys()]).toEqual(["a.txt"]);
    expect([...Archive.unpackTarGzSync(new DataView(tgz.slice().buffer)).keys()]).toEqual([
      "a.txt"
    ]);
  });

  it.each([null, "abc", 42, {}])("refuses %s with the sibling's message", source => {
    expect(() => Archive.unpackTarGzSync(source as never)).toThrow(
      /^Archive\.unpackTarGzSync: source must be bytes \(an ArrayBufferView or ArrayBuffer\)/
    );
    expect(() => Archive.unpackTarSync(source as never)).toThrow(
      /^Archive\.unpackTarSync: source must be bytes \(an ArrayBufferView or ArrayBuffer\)/
    );
  });
});

describe("Compression stream factories", () => {
  it.each([
    ["createDeflateStream", Compression.createDeflateStream],
    ["createGzipStream", Compression.createGzipStream],
    ["createZlibStream", Compression.createZlibStream],
    ["createInflateStream", Compression.createInflateStream],
    ["createGunzipStream", Compression.createGunzipStream],
    ["createUnzlibStream", Compression.createUnzlibStream]
  ] as const)("%s refuses options that are not an object", (_name, create) => {
    expect(() => create(42 as never)).toThrow(ArchiveError);
    expect(() => create("fast" as never)).toThrow(ArchiveError);
    expect(() => create()).not.toThrow();
  });
});
