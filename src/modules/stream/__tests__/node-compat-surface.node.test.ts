import * as nodeStream from "node:stream";
/**
 * On Node the flat `isReadable` / `isWritable` answer as `node:stream`'s do, Web streams
 * included — as type guards, so a value that is not a stream is `false` where Node says
 * `null`. `addAbortSignal` refuses what is not a stream, as Node's does, and `consumers`
 * has every reader Node's has.
 */
import * as nodeConsumers from "node:stream/consumers";

import * as entry from "@stream/index";
import { describe, expect, it } from "vitest";

const cases: Array<[string, () => unknown]> = [
  ["a Readable", () => nodeStream.Readable.from(["a"])],
  ["a destroyed Readable", () => nodeStream.Readable.from(["a"]).destroy()],
  ["a Writable", () => new nodeStream.Writable({ write: (_c, _e, cb) => cb() })],
  [
    "an ended Writable",
    () => {
      const w = new nodeStream.Writable({ write: (_c, _e, cb) => cb() });
      w.end();
      return w;
    }
  ],
  ["a Web ReadableStream", () => new ReadableStream()],
  ["a Web WritableStream", () => new WritableStream()],
  ["an object", () => ({})],
  ["null", () => null]
];

describe("isReadable / isWritable on Node", () => {
  it.each(cases)("agree with node:stream for %s", (_name, make) => {
    expect(entry.isReadable(make())).toBe(nodeStream.isReadable(make() as never) === true);
    expect(entry.isWritable(make())).toBe(nodeStream.isWritable(make() as never) === true);
  });
});

describe("addAbortSignal", () => {
  it("refuses something that is not a stream", () => {
    expect(() => entry.addAbortSignal(new AbortController().signal, {} as never)).toThrow(
      entry.StreamArgumentError
    );
  });
});

describe("consumers", () => {
  it("has every reader node:stream/consumers has", () => {
    const names = Object.keys(nodeConsumers).filter(name => name !== "default");
    expect(Object.keys(entry.consumers).sort()).toEqual(expect.arrayContaining(names.sort()));
  });

  it("bytes() resolves to a Uint8Array", async () => {
    const bytes = await entry.consumers.bytes(nodeStream.Readable.from([Buffer.from("ab")]));
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect([...bytes]).toEqual([97, 98]);
  });

  it("buffer() resolves to a Buffer on Node, as Node's does", async () => {
    expect(Buffer.isBuffer(await entry.consumers.buffer(nodeStream.Readable.from(["ab"])))).toBe(
      true
    );
  });
});

describe("Stream", () => {
  // Not published: on Node it would be the whole `node:stream` module beside this entry's own
  // members, and the browser kernel's class is not that. `Streams.isStream` answers "is it a
  // stream" on both.
  it("is not exported", () => {
    expect("Stream" in entry).toBe(false);
  });
});

// `consumers.bytes` is newer than Node 22; there, what it would answer is the `arrayBuffer` as bytes.
const nodeBytes: (source: never) => Promise<Uint8Array> =
  "bytes" in nodeConsumers
    ? nodeConsumers.bytes
    : async source => new Uint8Array(await nodeConsumers.arrayBuffer(source));

describe("consumers.text / json / bytes agree with node:stream/consumers", () => {
  const bom = new Uint8Array([0xef, 0xbb, 0xbf, 0x61]);
  const sources: Array<[string, () => unknown]> = [
    ["a BOM", () => nodeStream.Readable.from([bom])],
    ["a BOM in the middle", () => [new Uint8Array([0x61, 0xef, 0xbb, 0xbf])]],
    ["a string", () => "abc"],
    ["an array of Uint8Array", () => [new Uint8Array([97]), new Uint8Array([98])]],
    [
      "string chunks",
      () =>
        (async function* () {
          yield "a";
          yield "b";
        })()
    ],
    ["mixed chunks", () => ["a", new Uint8Array([98])]],
    ["UTF-8 split across chunks", () => [new Uint8Array([0xc3]), new Uint8Array([0xa9])]],
    ["invalid UTF-8", () => [new Uint8Array([0xff])]],
    ["a Web ReadableStream", () => new Blob([bom]).stream()]
  ];

  it.each(sources)("text of %s", async (_name, make) => {
    expect(await entry.consumers.text(make() as never)).toBe(
      await nodeConsumers.text(make() as never)
    );
    expect(await entry.Streams.text(make() as never)).toBe(
      await nodeConsumers.text(make() as never)
    );
  });

  it.each(sources.slice(0, 3))("bytes of %s", async (_name, make) => {
    expect([...(await entry.consumers.bytes(make() as never))]).toEqual([
      ...(await nodeBytes(make() as never))
    ]);
  });

  it("json strips a BOM", async () => {
    const json = () => [new Uint8Array([0xef, 0xbb, 0xbf, 0x31])];
    expect(await entry.consumers.json(json() as never)).toBe(
      await nodeConsumers.json(json() as never)
    );
  });

  it("is one implementation with the Streams readers", () => {
    expect(entry.consumers.bytes).toBe(entry.Streams.bytes);
    expect(entry.consumers.json).toBe(entry.Streams.json);
  });

  it("still refuses what is not a stream, with a stream error", async () => {
    await expect(entry.consumers.text(null as never)).rejects.toBeInstanceOf(
      entry.StreamUnsupportedTypeError
    );
    await expect(entry.consumers.text(42 as never)).rejects.toBeInstanceOf(
      entry.StreamUnsupportedTypeError
    );
  });
});
