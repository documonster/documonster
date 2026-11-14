/**
 * The browser form of `foreign-realm-bytes.node.test.ts`: a `Uint8Array` made in an iframe fails
 * `instanceof Uint8Array` here, and every call below must treat it as the bytes it is — on the
 * browser variants (CompressionStream, Web Crypto), not only on Node's.
 */
import { Archive, Compression } from "@archive";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const text = "documonster bytes from another realm ".repeat(8);
const local = new TextEncoder().encode(text);

let frame: HTMLIFrameElement;
let ForeignUint8Array: Uint8ArrayConstructor;

function foreign(bytes: Uint8Array = local): Uint8Array {
  const copy = new ForeignUint8Array(bytes.length);
  copy.set(bytes);
  return copy;
}

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

/** Write `chunk` through a codec stream in a few pieces and collect what comes out. */
function pump(
  codec: ReturnType<typeof Compression.createDeflateStream>,
  chunk: Uint8Array
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const out: Uint8Array[] = [];
    codec.on("data", (part: Uint8Array) => out.push(part));
    codec.on("error", reject);
    codec.on("end", () => {
      const all = new Uint8Array(out.reduce((n, part) => n + part.length, 0));
      let offset = 0;
      for (const part of out) {
        all.set(part, offset);
        offset += part.length;
      }
      resolve(all);
    });
    const half = Math.floor(chunk.length / 2);
    codec.write(chunk.subarray(0, half));
    codec.write(chunk.subarray(half));
    codec.end();
  });
}

/** A `ReadableStream` that yields `bytes` — still the other realm's — in two chunks. */
function streamOf(bytes: Uint8Array): ReadableStream<Uint8Array> {
  const half = Math.floor(bytes.length / 2);
  return new ReadableStream({
    start(controller) {
      controller.enqueue(bytes.subarray(0, half));
      controller.enqueue(bytes.subarray(half));
      controller.close();
    }
  });
}

beforeAll(() => {
  frame = document.createElement("iframe");
  document.body.append(frame);
  ForeignUint8Array = (frame.contentWindow as unknown as { Uint8Array: Uint8ArrayConstructor })
    .Uint8Array;
});

afterAll(() => {
  frame.remove();
});

describe("bytes from an iframe", () => {
  it("is not an instance of this realm's Uint8Array", () => {
    expect(foreign() instanceof Uint8Array).toBe(false);
  });

  it("compress and decompress as local bytes do", async () => {
    expect(await Compression.deflate(foreign())).toEqual(await Compression.deflate(local));
    expect(decode(await Compression.inflate(foreign(await Compression.deflate(local))))).toBe(text);
    expect(decode(await Compression.gunzip(foreign(await Compression.gzip(foreign()))))).toBe(text);
    expect(decode(Compression.inflateSync(foreign(Compression.deflateSync(foreign()))))).toBe(text);
    expect(Compression.isGzipData(foreign(Compression.gzipSync(local)))).toBe(true);
    expect(Archive.crc32(foreign())).toBe(Archive.crc32(local));
  });

  it("pack into a ZIP and read back", async () => {
    const zip = await Archive.packZip([["a.txt", foreign()]]);
    expect(decode((await Archive.unpackZip(foreign(zip))).get("a.txt")!.data)).toBe(text);
    expect(await Archive.openZip(foreign(zip)).list()).toEqual(["a.txt"]);
  });

  it("encrypt and decrypt a ZIP with a password from another realm", async () => {
    const password = foreign(new TextEncoder().encode("pw"));
    for (const encryptionMethod of ["zipcrypto", "aes-256"] as const) {
      const zip = await Archive.packZip([["a.txt", foreign()]], { encryptionMethod, password });
      const files = await Archive.unpackZip(foreign(zip), { password: foreign(password) });
      expect(decode(files.get("a.txt")!.data)).toBe(text);
    }
  });

  it("edit a ZIP opened from another realm's bytes", async () => {
    const zip = await Archive.packZip([["a.txt", text]]);
    const editor = await Archive.ZipEditor.open(foreign(zip));
    editor.set("b.txt", foreign());
    for (const out of [await editor.bytes(), await editor.toBuffer()]) {
      const files = await Archive.unpackZip(out);
      expect(decode(files.get("a.txt")!.data)).toBe(text);
      expect(decode(files.get("b.txt")!.data)).toBe(text);
    }
  });

  it("compress on a worker as on this thread", async () => {
    expect(await Compression.deflate(foreign(), { useWorker: true })).toEqual(
      await Compression.deflate(local)
    );
    const compressed = await Compression.deflate(local);
    expect(decode(await Compression.inflate(foreign(compressed), { useWorker: true }))).toBe(text);
  });

  it("stream through the codec streams", async () => {
    const compressed = await pump(Compression.createDeflateStream(), foreign());
    expect(compressed).toEqual(await Compression.deflate(local));
    expect(decode(await pump(Compression.createInflateStream(), foreign(compressed)))).toBe(text);
    const gzipped = await pump(Compression.createGzipStream(), foreign());
    expect(decode(await pump(Compression.createGunzipStream(), foreign(gzipped)))).toBe(text);
  });

  it("pack and read a ZIP streamed in chunks from another realm", async () => {
    const zip = await Archive.packZip([["a.txt", streamOf(foreign())]]);
    expect(decode((await Archive.unpackZip(zip)).get("a.txt")!.data)).toBe(text);
    const entries: string[] = [];
    for await (const entry of Archive.openZip(streamOf(foreign(zip))).entries()) {
      entries.push(decode(await entry.bytes()));
    }
    expect(entries).toEqual([text]);
  });

  it("pack into a TAR and read back", async () => {
    const tar = await Archive.packTar([["a.txt", foreign()]]);
    expect(decode((await Archive.unpackTar(foreign(tar))).get("a.txt")!.data)).toBe(text);
  });

  it("ZipCrypto takes a password and random bytes from another realm", () => {
    const crc = Archive.crc32(local);
    const random = new Uint8Array(11).fill(7);
    expect(Archive.Crypto.createZipCryptoKeys(foreign())).toEqual(
      Archive.Crypto.createZipCryptoKeys(local)
    );
    expect(Archive.Crypto.encryptZipCrypto(foreign(), "pw", crc, () => foreign(random))).toEqual(
      Archive.Crypto.encryptZipCrypto(local, "pw", crc, () => random)
    );
  });
});
