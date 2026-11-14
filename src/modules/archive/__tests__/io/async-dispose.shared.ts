/**
 * Every archive object that holds something to release is disposed of by `await using`, which calls the `close` it
 * already had — a reader opened inside a function is closed on every way out of it, without a `try`/`finally`.
 */
import { Archive } from "@archive/index";
import type { RandomAccessReader } from "@archive/index";
import { asyncDisposeKey } from "@utils/dispose";
import { describe, expect, it, vi } from "vitest";

const zip = Archive.packZipSync(new Map([["a.txt", new TextEncoder().encode("a")]]));

/** A random-access reader over `bytes` whose `close` is counted. */
function countingReader(bytes: Uint8Array): RandomAccessReader & { closed: number } {
  const reader = Archive.createBufferReader(bytes);
  const counting = {
    ...reader,
    size: reader.size,
    read: reader.read.bind(reader),
    closed: 0,
    close: async () => {
      counting.closed++;
    }
  };
  return counting;
}

/** The tests, for the Node and the browser runs to register. */
export function describeAsyncDispose(): void {
  describe("await using an archive object", () => {
    it.each([
      [
        "Archive.RemoteZipReader",
        (source: RandomAccessReader) => Archive.RemoteZipReader.fromReader(source)
      ],
      ["Archive.ZipEditor", (source: RandomAccessReader) => Archive.ZipEditor.openReader(source)]
    ])("closes %s, and the source it reads, on leaving the block", async (_name, open) => {
      const source = countingReader(zip);
      {
        await using opened = await open(source);
        expect(opened).toBeDefined();
        expect(source.closed).toBe(0);
      }
      expect(source.closed).toBe(1);
    });

    // Each `await using` below also type-checks the class as `AsyncDisposable`.
    it("calls the close of a ZIP reader, a TAR reader and an HTTP range reader", async () => {
      const closes = [
        vi.spyOn(Archive.ZipReader.prototype, "close"),
        vi.spyOn(Archive.TarReader.prototype, "close"),
        vi.spyOn(Archive.HttpRangeReader.prototype, "close")
      ];
      {
        await using zipReader = Archive.openZip(zip);
        await using tarReader = Archive.openTar(Archive.packTarSync(new Map([["a.txt", "a"]])));
        await using rangeReader = await Archive.HttpRangeReader.open("https://example.test/a.zip", {
          size: zip.length,
          validateRangeSupport: false
        });
        void [zipReader, tarReader, rangeReader];
      }
      for (const close of closes) {
        expect(close).toHaveBeenCalledTimes(1);
      }
      vi.restoreAllMocks();
    });

    it("disposes of a reader already closed without failing", async () => {
      const reader = await Archive.RemoteZipReader.fromReader(countingReader(zip));
      await reader.close();
      await expect(reader[Symbol.asyncDispose]()).resolves.toBeUndefined();
    });
  });

  describe("the key await using disposes by", () => {
    it("is the runtime's Symbol.asyncDispose where it has one", () => {
      expect(asyncDisposeKey(Symbol)).toBe(Symbol.asyncDispose);
    });

    // esbuild and Babel lower `await using` to a lookup under this key where the runtime has no symbol of its own.
    it("is the registered Symbol.for('Symbol.asyncDispose') where it has none", () => {
      expect(asyncDisposeKey({})).toBe(Symbol.for("Symbol.asyncDispose"));
    });
  });
}
