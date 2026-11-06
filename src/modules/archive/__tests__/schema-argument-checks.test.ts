/**
 * Refusals that changed when the archive argument checks moved to `@utils/schema`: every
 * refusal now ends with the value it got, an array is not an options object, and the stream
 * options every archive operation takes (`signal`, `onProgress`, `progressIntervalMs`) are
 * checked where they are given.
 */
import { Archive, Compression } from "@archive/index";
import { describe, expect, it } from "vitest";

async function refusal(
  call: () => unknown
): Promise<{ name: string; field: string; message: string }> {
  try {
    await call();
  } catch (err) {
    const e = err as { name: string; field: string; message: string };
    return { name: e.name, field: e.field, message: e.message };
  }
  throw new Error("expected a refusal");
}

describe("archive argument checks declared as schemas", () => {
  it("names the value a stream compressor's options were", async () => {
    expect(await refusal(() => Compression.createDeflateStream(5 as never))).toEqual({
      name: "ArchiveArgumentError",
      field: "options",
      message: "Compression.createDeflateStream: options must be an object, got 5"
    });
  });

  it("describes a refused source by its value", async () => {
    expect((await refusal(() => Archive.toAsyncIterable(5 as never))).message).toBe(
      "Archive.toAsyncIterable: source must be a Uint8Array, ArrayBuffer, string, Blob, ReadableStream or async iterable, got 5"
    );
  });

  it("refuses an array as options", async () => {
    expect(await refusal(() => Archive.openTar(new Uint8Array(1024), [] as never))).toEqual({
      name: "ArchiveArgumentError",
      field: "options",
      message: "Archive.openTar: options must be an options object, got an array"
    });
  });

  it("checks signal, onProgress and progressIntervalMs where they are given", async () => {
    const tar = new Uint8Array(1024);
    expect((await refusal(() => Archive.openTar(tar, { signal: {} as never }))).field).toBe(
      "options.signal"
    );
    expect((await refusal(() => Archive.createTar({ onProgress: 1 as never }))).field).toBe(
      "options.onProgress"
    );
    expect((await refusal(() => Archive.createTarGz({ progressIntervalMs: -1 }))).message).toBe(
      "Archive.createTarGz: options.progressIntervalMs must be a number >= 0, got -1"
    );
  });

  it("checks the options of a remote reader before reading", async () => {
    const reader = Archive.createBufferReader(new Uint8Array(22));
    expect(
      await refusal(() => Archive.RemoteZipReader.fromReader(reader, { password: 1 as never }))
    ).toEqual({
      name: "ArchiveArgumentError",
      field: "options.password",
      message:
        "Archive.RemoteZipReader.fromReader: options.password must be a string or a Uint8Array, got 1"
    });
  });

  it("reports a wrong edit operation field where it is", async () => {
    expect(
      (
        await refusal(
          () => new Archive.ZipEditPlan([{ type: "rename", from: "a", to: 1 }] as never)
        )
      ).field
    ).toBe("ops[0].to");
  });

  it("names the value a key strength or a url was", async () => {
    expect(
      (await refusal(() => Archive.Crypto.getEncryptionMethodForAesKeyStrength(5 as never))).message
    ).toContain("must be 128, 192 or 256, got 5");
    expect((await refusal(() => Archive.HttpRangeReader.open(5 as never))).message).toBe(
      "Archive.HttpRangeReader.open: url must be a string, got 5"
    );
  });
});
