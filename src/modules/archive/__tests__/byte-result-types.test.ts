import { Archive, Compression } from "@archive";
import { describe, expect, expectTypeOf, it } from "vitest";

describe("byte results are typed over ArrayBuffer", () => {
  it("can be handed to Blob without a cast", async () => {
    const zipped = await Archive.packZip([["a.txt", new Uint8Array([1, 2, 3])]]);
    expectTypeOf(zipped).toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expect(zipped.constructor).toBe(Uint8Array);
    expect(new Blob([zipped]).size).toBe(zipped.length);

    const unpacked = await Archive.unpackZip(zipped);
    const data = unpacked.get("a.txt")!.data;
    expectTypeOf(data).toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expect(data.constructor).toBe(Uint8Array);

    const gz = await Compression.gzip(new Uint8Array([1]));
    expectTypeOf(gz).toEqualTypeOf<Uint8Array<ArrayBuffer>>();
    expect(gz.constructor).toBe(Uint8Array);
  });
});
