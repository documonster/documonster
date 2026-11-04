/**
 * The archive readers answer the same questions the same way: a one-shot source read twice is
 * refused rather than read as empty, a duplicated name resolves to the entry every extractor
 * keeps, a lookup name is checked, and the size limits are the same options on every reader.
 */
import {
  Archive,
  ArchiveArgumentError,
  ArchiveError,
  FileTooLargeError,
  type ArchiveBytes
} from "@archive/index";
import { describe, expect, it } from "vitest";

const decoder = new TextDecoder();
const text = (bytes: Uint8Array | null): string | null => (bytes ? decoder.decode(bytes) : null);

const streamOf = (bytes: Uint8Array): ReadableStream<Uint8Array> =>
  new ReadableStream({
    start(controller) {
      controller.enqueue(bytes);
      controller.close();
    }
  });
const iterableOf = (bytes: Uint8Array): AsyncIterable<Uint8Array> =>
  (async function* () {
    yield bytes;
  })();

async function paths(entries: AsyncIterable<{ path: string; bytes(): Promise<Uint8Array> }>) {
  const out: string[] = [];
  for await (const entry of entries) {
    out.push(entry.path);
    await entry.bytes();
  }
  return out;
}

const ENTRIES: Array<[string, string]> = [
  ["a.txt", "A"],
  ["b.txt", "B"]
];

describe("a one-shot source is read once", () => {
  const zip = Archive.packZipSync(ENTRIES);
  const tar = Archive.packTarSync(ENTRIES);

  it.each([
    ["ReadableStream", streamOf],
    ["async iterable", iterableOf]
  ] as const)("openTar over a %s refuses a second pass instead of answering []", async (_n, of) => {
    const reader = Archive.openTar(of(tar));
    expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
    await expect(paths(reader.entries())).rejects.toThrow(/already consumed/);
    await expect(reader.list()).rejects.toBeInstanceOf(ArchiveError);
  });

  it.each([
    ["ReadableStream", streamOf],
    ["async iterable", iterableOf]
  ] as const)("openZip over a %s refuses a second pass with the same error", async (_n, of) => {
    const reader = Archive.openZip(of(zip));
    expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
    await expect(paths(reader.entries())).rejects.toThrow(
      /^Archive\.ZipReader#entries: .*already consumed/
    );
    await expect(reader.list()).rejects.toThrow(/^Archive\.ZipReader#list: .*already consumed/);
  });

  it("a list / get over a one-shot source keeps what it read for every later call", async () => {
    const zipReader = Archive.openZip(streamOf(zip));
    expect(await zipReader.list()).toEqual(["a.txt", "b.txt"]);
    expect(text(await zipReader.bytes("b.txt"))).toBe("B");
    expect(await paths(zipReader.entries())).toEqual(["a.txt", "b.txt"]);

    const tarReader = Archive.openTar(streamOf(tar));
    expect(await tarReader.list()).toEqual(["a.txt", "b.txt"]);
    expect(text(await tarReader.bytes("b.txt"))).toBe("B");
    expect(await paths(tarReader.entries())).toEqual(["a.txt", "b.txt"]);
  });

  it("bytes in memory are read as often as asked", async () => {
    for (const reader of [Archive.openZip(zip), Archive.openTar(tar)]) {
      expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
      expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
    }
  });

  it("entriesStream() and operation() consume it as entries() does, once iteration starts", async () => {
    const drain = async (stream: ReadableStream<{ path: string }>) => {
      const reader = stream.getReader();
      while (!(await reader.read()).done) {
        // read every entry
      }
    };
    for (const open of [
      () => Archive.openZip(streamOf(zip)),
      () => Archive.openTar(streamOf(tar))
    ]) {
      const viaStream = open();
      await drain(viaStream.entriesStream());
      await expect(viaStream.list()).rejects.toThrow(/already consumed/);

      const viaOperation = open();
      expect(await paths(viaOperation.operation().iterable)).toEqual(["a.txt", "b.txt"]);
      await expect(viaOperation.get("a.txt")).rejects.toThrow(/already consumed/);
      await expect(viaOperation.bytes("a.txt")).rejects.toThrow(/already consumed/);

      // Asking for the iterable reads nothing; a list first then buffers as usual
      const notStarted = open();
      notStarted.entries();
      notStarted.operation();
      expect(await notStarted.list()).toEqual(["a.txt", "b.txt"]);
    }
  });

  it("after a list / get over a one-shot source, entries() can be iterated again and again", async () => {
    for (const reader of [Archive.openZip(streamOf(zip)), Archive.openTar(streamOf(tar))]) {
      expect(await reader.get("a.txt")).not.toBeNull();
      expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
      expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
    }
  });

  it("TarReader.close() drops what a one-shot source buffered; the source cannot be read again", async () => {
    const reader = Archive.openTar(streamOf(tar));
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
    await reader.close();
    await expect(reader.list()).rejects.toThrow(/already consumed/);
  });

  it("openTarGz keeps its documented replay", async () => {
    const reader = Archive.openTarGz(streamOf(Archive.packTarGzSync(ENTRIES)));
    expect(await paths(reader.entries())).toEqual(["a.txt", "b.txt"]);
    expect(await reader.list()).toEqual(["a.txt", "b.txt"]);
  });
});

describe("duplicate entry names", () => {
  it("a ZIP builder refuses a name it already has", async () => {
    expect(() => Archive.createZip().add("a.txt", "1").add("a.txt", "2")).toThrow(
      ArchiveArgumentError
    );
    expect(() =>
      Archive.packZipSync([
        ["a.txt", "1"],
        ["a.txt", "2"]
      ])
    ).toThrow(/^Archive\.packZipSync: the archive already has an entry named "a\.txt"/);
    await expect(
      Archive.packZip([
        ["a.txt", "1"],
        ["a.txt", "2"]
      ])
    ).rejects.toThrow(ArchiveArgumentError);
    expect(() => Archive.createZip().addDirectory("d").addDirectory("d/")).toThrow(
      ArchiveArgumentError
    );
    // Under a path mode the stored (normalised) names are what must differ.
    expect(() =>
      Archive.createZip({ path: { mode: "legacy" } })
        .add("dir\\a.txt", "1")
        .add("dir/a.txt", "2")
    ).toThrow(ArchiveArgumentError);
  });

  // A ZIP with one name twice, as other tools write: two local headers and records.
  const dupZip = (() => {
    const one = Archive.packZipSync([
      ["a.txt", "first"],
      ["b.txt", "second"]
    ]);
    // Rename the second record (and its local header) to "a.txt": same length, so offsets hold.
    const bytes = one.slice();
    const name = new TextEncoder().encode("b.txt");
    const swap = new TextEncoder().encode("a.txt");
    for (let i = 0; i + name.length <= bytes.length; i++) {
      if (name.every((byte, k) => bytes[i + k] === byte)) {
        bytes.set(swap, i);
      }
    }
    return bytes;
  })();

  it("every ZIP reader returns the later of two same-named entries", async () => {
    expect(text((await Archive.unpackZip(dupZip)).get("a.txt")!.data)).toBe("second");
    expect(text(await Archive.openZip(dupZip).bytes("a.txt"))).toBe("second");
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(dupZip));
    expect(text(await remote.extract("a.txt"))).toBe("second");
  });

  it("RemoteZipReader counts and lists the same entries", async () => {
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(dupZip));
    expect(remote.listFiles()).toEqual(["a.txt", "a.txt"]);
    expect(remote.getFileCount()).toBe(remote.listFiles().length);
    const withDir = await Archive.RemoteZipReader.fromReader(
      Archive.createBufferReader(
        Archive.packZipSync([
          ["d/", ""],
          ["d/x.txt", "x"]
        ])
      )
    );
    expect(withDir.listFiles()).toEqual(["d/x.txt"]);
    expect(withDir.getFileCount()).toBe(1);
  });

  it("an unedited editZip keeps every entry, and an edit of the name leaves one", async () => {
    const untouched = await (await Archive.editZip(dupZip)).toBuffer();
    expect(await Archive.openZip(untouched).list()).toEqual(["a.txt", "a.txt"]);
    expect((await Archive.editZip(dupZip)).getEntries().length).toBe(2);

    const edited = await Archive.editZip(dupZip);
    edited.set("a.txt", "third");
    const out = await edited.toBuffer();
    expect(await Archive.openZip(out).list()).toEqual(["a.txt"]);
    expect(text(await Archive.openZip(out).bytes("a.txt"))).toBe("third");
  });

  it("TarReader.get / bytes return the later member, as unpackTar does", async () => {
    const tar = Archive.packTarSync([
      ["a.txt", "first"],
      ["a.txt", "second"]
    ]);
    expect(text(Archive.unpackTarSync(tar).get("a.txt")!.data)).toBe("second");
    expect(text(await Archive.openTar(tar).bytes("a.txt"))).toBe("second");
    expect(text(await (await Archive.openTar(tar).get("a.txt"))!.bytes())).toBe("second");
    const tgz = Archive.packTarGzSync([
      ["a.txt", "first"],
      ["a.txt", "second"]
    ]);
    expect(text(await Archive.openTarGz(tgz).bytes("a.txt"))).toBe("second");
  });
});

describe("lookup names are checked", () => {
  const zip = Archive.packZipSync(ENTRIES);
  const tar = Archive.packTarSync(ENTRIES);

  it.each([42, undefined, ""])("every reader refuses %j as a name", async bad => {
    const name = bad as never;
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip));
    const editor = await Archive.editZip(zip);
    for (const call of [
      () => Archive.openZip(zip).get(name),
      () => Archive.openZip(zip).bytes(name),
      () => Archive.openTar(tar).get(name),
      () => Archive.openTar(tar).bytes(name)
    ]) {
      await expect(call()).rejects.toBeInstanceOf(ArchiveArgumentError);
    }
    for (const call of [
      () => remote.getEntry(name),
      () => remote.hasEntry(name),
      () => editor.has(name),
      () => editor.delete(name)
    ]) {
      expect(call).toThrow(ArchiveArgumentError);
    }
    await expect(remote.extract(name)).rejects.toBeInstanceOf(ArchiveArgumentError);
  });

  it("names the call and the parameter", async () => {
    await expect(Archive.openZip(zip).get(42 as never)).rejects.toMatchObject({
      field: "path",
      message: expect.stringMatching(/^Archive\.ZipReader#get: path must be a non-empty string/)
    });
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip));
    await expect(remote.extractMultiple(["a.txt", 7 as never])).rejects.toMatchObject({
      field: "paths[1]"
    });
  });
});

describe("entry-size limits", () => {
  const big = new Uint8Array(4096);
  const zip = Archive.packZipSync([["big.bin", big]]);
  const tar = Archive.packTarSync([["big.bin", big]]);

  it("openZip honours maxEntrySize over bytes and over a stream", async () => {
    await expect(Archive.openZip(zip, { maxEntrySize: 100 }).bytes("big.bin")).rejects.toThrow(
      FileTooLargeError
    );
    const streamed = Archive.openZip(streamOf(zip), { maxEntrySize: 100 });
    await expect(
      (async () => {
        for await (const entry of streamed.entries()) {
          await entry.bytes();
        }
      })()
    ).rejects.toThrow(FileTooLargeError);
    expect((await Archive.openZip(zip, { maxEntrySize: 4096 }).bytes("big.bin"))!.length).toBe(
      4096
    );
  });

  it("RemoteZipReader honours maxEntrySize and maxEntries", async () => {
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip), {
      maxEntrySize: 100
    });
    await expect(remote.extract("big.bin")).rejects.toThrow(FileTooLargeError);
    await expect(
      Archive.RemoteZipReader.fromReader(Archive.createBufferReader(Archive.packZipSync(ENTRIES)), {
        maxEntries: 1
      })
    ).rejects.toThrow(/maxEntries/);
  });

  it("TAR takes the same name, maxEntrySize, and keeps maxFileSize", async () => {
    await expect(Archive.unpackTar(tar, { maxEntrySize: 100 })).rejects.toThrow(FileTooLargeError);
    await expect(Archive.openTar(tar, { maxEntrySize: 100 }).list()).rejects.toThrow(
      FileTooLargeError
    );
    expect(() => Archive.parseTar(tar, { maxEntrySize: 100 })).toThrow(FileTooLargeError);
    await expect(Archive.unpackTar(tar, { maxFileSize: 100 })).rejects.toThrow(FileTooLargeError);
  });

  it("refuses a limit that is not a count", () => {
    expect(() => Archive.openZip(zip, { maxEntrySize: -1 })).toThrow(ArchiveArgumentError);
    expect(() => Archive.openTar(tar, { maxEntrySize: Number.NaN })).toThrow(
      /^Archive\.openTar: options\.maxEntrySize/
    );
  });
});

describe("RemoteZipExtractOptions.onProgress", () => {
  it("is called, as the deprecated onprogress still is", async () => {
    const remote = await Archive.RemoteZipReader.fromReader(
      Archive.createBufferReader(Archive.packZipSync(ENTRIES))
    );
    const seen: number[] = [];
    await remote.extract("a.txt", { onProgress: (current: number) => seen.push(current) });
    const legacy: number[] = [];
    await remote.extract("a.txt", { onprogress: (current: number) => legacy.push(current) });
    expect(seen.length).toBeGreaterThan(0);
    expect(legacy).toEqual(seen);
    await expect(remote.extract("a.txt", { onProgress: 42 as never })).rejects.toBeInstanceOf(
      ArchiveArgumentError
    );
  });
});

describe("ZIP and TAR reader entries have the same shape", () => {
  it("UnzipEntry.isDirectory, TarReaderEntry.readableStream / pipeTo(WritableStream)", async () => {
    const zip = Archive.packZipSync([
      ["d/", ""],
      ["d/a.txt", "A"]
    ]);
    const kinds: Array<[string, boolean]> = [];
    for await (const entry of Archive.openZip(zip).entries()) {
      kinds.push([entry.path, entry.isDirectory]);
      entry.discard();
    }
    expect(kinds).toEqual([
      ["d/", true],
      ["d/a.txt", false]
    ]);

    const entry = (await Archive.openTar(Archive.packTarSync([["a.txt", "A"]])).get("a.txt"))!;
    expect(text(new Uint8Array(await new Response(entry.readableStream()).arrayBuffer()))).toBe(
      "A"
    );
    const chunks: Uint8Array[] = [];
    await entry.pipeTo(new WritableStream({ write: chunk => void chunks.push(chunk) }));
    expect(text(chunks[0]!)).toBe("A");
  });

  it("TarReader.entriesStream, as ZipReader.entriesStream", async () => {
    const reader = Archive.openTar(Archive.packTarSync(ENTRIES)).entriesStream().getReader();
    const seen: string[] = [];
    for (let step = await reader.read(); !step.done; step = await reader.read()) {
      seen.push(step.value.path);
    }
    expect(seen).toEqual(["a.txt", "b.txt"]);
  });
});

describe("byte readers take any ArrayBufferView or ArrayBuffer", () => {
  const zip = Archive.packZipSync(ENTRIES);
  const tar = Archive.packTarSync(ENTRIES);
  const views = (bytes: Uint8Array): Array<[string, ArchiveBytes]> => {
    const copy = bytes.slice();
    return [
      ["DataView", new DataView(copy.buffer)],
      ["ArrayBuffer", copy.buffer],
      ["offset Uint8Array", new Uint8Array([9, 9, ...bytes]).subarray(2)]
    ];
  };

  it.each(views(zip))("ZIP readers read a %s", async (_name, bytes) => {
    expect([...Archive.unpackZipSync(bytes).keys()]).toEqual(["a.txt", "b.txt"]);
    expect([...(await Archive.unpackZip(bytes)).keys()]).toEqual(["a.txt", "b.txt"]);
    expect(await Archive.openZip(bytes).list()).toEqual(["a.txt", "b.txt"]);
    expect((await Archive.editZip(bytes)).getEntries().length).toBe(2);
    const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(bytes));
    expect(text(await remote.extract("b.txt"))).toBe("B");
  });

  it.each(views(tar))("TAR readers read a %s", async (_name, bytes) => {
    expect([...Archive.unpackTarSync(bytes).keys()]).toEqual(["a.txt", "b.txt"]);
    expect(Archive.parseTar(bytes).length).toBe(2);
    expect(await Archive.openTar(bytes).list()).toEqual(["a.txt", "b.txt"]);
  });
});
