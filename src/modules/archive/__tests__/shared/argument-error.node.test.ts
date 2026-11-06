/**
 * A wrong argument is an `ArchiveArgumentError` — still an `ArchiveError`, so existing
 * `catch` and `isArchiveError` checks keep working — naming the argument or option at fault
 * in `field`. Damaged archive data stays a plain `ArchiveError`. A handle passed to a public
 * function (entries to pack, a source to read, a stream to adapt) is checked as cheaply as
 * any other argument, never left to fail with a raw `TypeError` inside.
 */
import { Archive, ArchiveArgumentError, ArchiveError, Compression, isArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

function caught(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  return undefined;
}

async function caughtAsync(call: () => unknown): Promise<unknown> {
  try {
    await call();
  } catch (error) {
    return error;
  }
  return undefined;
}

describe("ArchiveArgumentError", () => {
  it("is an ArchiveError with a field", () => {
    const error = new ArchiveArgumentError("level", "createZip: level must be …");
    expect(error).toBeInstanceOf(ArchiveError);
    expect(isArchiveError(error)).toBe(true);
    expect(error.name).toBe("ArchiveArgumentError");
    expect(error.field).toBe("level");
    expect(error.message).toBe("createZip: level must be …");
  });

  it.each<[string, string, () => unknown]>([
    // `field` is the path from the parameter name: an option is `options.<name>`.
    ["options.level", "createZip level", () => Archive.createZip({ level: 42 })],
    ["name", "ZipArchive.add name", () => Archive.createZip().add("../x", "x")],
    ["name", "an empty name", () => Archive.createZip().add("", "x")],
    [
      "options.modTime",
      "ZIP modTime",
      () => Archive.createZip({ modTime: new Date(Date.UTC(1960, 0)) })
    ],
    ["options.encoding", "ZIP encoding", () => Archive.createZip({ encoding: "bogus" as never })],
    ["source", "TarArchive.add source", () => Archive.createTar().add("a", 42 as never)],
    [
      "options.uname",
      "TAR uname",
      () => Archive.createTar().add("a", "x", { uname: "u".repeat(40) })
    ],
    ["options.type", "TAR type", () => Archive.createTar().add("a", "x", { type: "z" as never })],
    ["data", "parseTar data", () => Archive.parseTar(42 as never)],
    ["source", "unpackTarSync source", () => Archive.unpackTarSync(42 as never)],
    ["source", "unpackTarGzSync source", () => Archive.unpackTarGzSync(42 as never)],
    ["source", "unpackZipSync source", () => Archive.unpackZipSync(42 as never)],
    ["entries", "packZipSync entries", () => Archive.packZipSync(42 as never)],
    ["entries", "packTarSync entries", () => Archive.packTarSync(null as never)],
    ["entries", "packTarGzSync entries", () => Archive.packTarGzSync({} as never)],
    ["entries[0]", "packZipSync entry", () => Archive.packZipSync([42] as never)],
    ["data", "createBufferReader data", () => Archive.createBufferReader(null as never)],
    ["base64", "base64ToUint8Array", () => Archive.base64ToUint8Array("!!")],
    ["totalLength", "concatUint8Arrays", () => Archive.concatUint8Arrays([new Uint8Array(1)], 9)],
    [
      "options.level",
      "Compression.deflateSync",
      () => Compression.deflateSync(new Uint8Array(1), { level: 42 })
    ],
    ["data", "Compression.inflateSync", () => Compression.inflateSync(null as never)],
    [
      "options",
      "Compression.createInflateStream",
      () => Compression.createInflateStream(42 as never)
    ],
    ["rgba", "encodePng", () => Archive.encodePng(null as never, 1, 1)],
    ["width", "encodePng dimensions", () => Archive.encodePng(new Uint8Array(4), 0, 1)],
    [
      "keyStrength",
      "Crypto.getAesEncryptedSize",
      () => Archive.Crypto.getAesEncryptedSize(1, 7 as never)
    ],
    [
      "options.format",
      "ArchiveFile format",
      () => new Archive.ArchiveFile({ format: "rar" } as never)
    ],
    ["source", "toAsyncIterable source", () => Archive.toReadableStream(42 as never)],
    ["source", "toNodeReadable source", () => Archive.toNodeReadable(42 as never)]
  ])("names %s for %s", (field, _name, call) => {
    const error = caught(call);
    expect(error).toBeInstanceOf(ArchiveArgumentError);
    expect((error as ArchiveArgumentError).field).toBe(field);
  });

  it.each<[string, string, () => unknown]>([
    ["entries", "packZip", () => Archive.packZip(42 as never)],
    ["entries", "packTar", () => Archive.packTar(null as never)],
    ["entries", "packTarGz", () => Archive.packTarGz(42 as never)],
    ["source", "unpackTar", () => Archive.unpackTar(42 as never)],
    ["source", "unpackZip", () => Archive.unpackZip(42 as never)],
    ["source", "openZip", () => Archive.openZip(42 as never)],
    ["url", "HttpRangeReader.open", () => Archive.HttpRangeReader.open(42 as never)],
    ["source", "ZipEditor.open", () => Archive.ZipEditor.open(42 as never)],
    [
      "source",
      "Archive.toAsyncIterable",
      async () => {
        for await (const _ of Archive.toAsyncIterable(42 as never)) {
          // drain
        }
      }
    ]
  ])("names %s for %s (async)", async (field, _name, call) => {
    const error = await caughtAsync(call);
    expect(error).toBeInstanceOf(ArchiveArgumentError);
    expect((error as ArchiveArgumentError).field).toBe(field);
  });

  it("leaves damaged data a plain ArchiveError", async () => {
    const bad = caught(() => Archive.unpackTarSync(new Uint8Array(512).fill(65)));
    expect(bad).toBeInstanceOf(ArchiveError);
    expect(bad).not.toBeInstanceOf(ArchiveArgumentError);
    const gz = caught(() => Compression.gunzipSync(new TextEncoder().encode("not gzip at all")));
    expect(gz).not.toBeInstanceOf(ArchiveArgumentError);
    const zip = await caughtAsync(() => Archive.unpackZip(new Uint8Array(64)));
    expect(zip).toBeInstanceOf(ArchiveError);
    expect(zip).not.toBeInstanceOf(ArchiveArgumentError);
  });
});

describe("handles are checked where they are given", () => {
  const zip = Archive.packZipSync([["a", "A"]]);

  it.each<[string, string, () => unknown]>([
    ["plan", "ZipEditor.apply", async () => (await Archive.ZipEditor.open(zip)).apply(42 as never)],
    ["other", "ZipEditPlan.concat", () => new Archive.ZipEditPlan().concat(42 as never)],
    [
      "target",
      "ZipEditPlan.applyTo",
      () => new Archive.ZipEditPlan().set("a", "b").applyTo(42 as never)
    ],
    ["data", "ZipEditPlan.deserialize", () => Archive.ZipEditPlan.deserialize(42 as never)],
    ["json", "ZipEditPlan.fromJson", () => Archive.ZipEditPlan.fromJson(42 as never)],
    ["reader", "RemoteZipReader.fromReader", () => Archive.RemoteZipReader.fromReader(42 as never)],
    [
      "sink",
      "ZipArchive.pipeTo",
      () =>
        Archive.createZip()
          .add("a", "x")
          .pipeTo(42 as never)
    ],
    [
      "sink",
      "TarArchive.pipeTo",
      () =>
        Archive.createTar()
          .add("a", "x")
          .pipeTo(null as never)
    ],
    [
      "sink",
      "ArchiveFile.pipeTo",
      () => new Archive.ArchiveFile().addText("x", "a").pipeTo(42 as never)
    ],
    [
      "sink",
      "UnzipEntry.pipeTo",
      async () => {
        for await (const entry of Archive.openZip(zip).entries()) {
          await entry.pipeTo(42 as never);
        }
      }
    ],
    [
      "info",
      "createTarReaderEntry",
      () => Archive.createTarReaderEntry(null as never, new Uint8Array(1))
    ],
    [
      "data",
      "createTarReaderEntry data",
      () => Archive.createTarReaderEntry({ path: "a" } as never, 42 as never)
    ],
    ["entry", "isTarFile", () => Archive.isTarFile(null as never)],
    ["entry", "isTarDirectory", () => Archive.isTarDirectory(42 as never)]
  ])("%s is checked for %s", async (field, _name, call) => {
    const error = await caughtAsync(call);
    expect(error).toBeInstanceOf(ArchiveArgumentError);
    expect((error as ArchiveArgumentError).field).toBe(field);
  });
});
