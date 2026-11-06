import { Archive, Compression } from "@archive";
import { describe, expect, it } from "vitest";

describe("deprecated names keep their call graph", () => {
  it("TarReader.bytes goes through an overridden get", async () => {
    class Reader extends Archive.TarReader {
      override async get(path: string) {
        return super.get(path === "alias" ? "a.txt" : path);
      }
    }
    const reader = new Reader(await Archive.packTar([["a.txt", "alpha"]]));
    expect(new TextDecoder().decode((await reader.bytes("alias"))!)).toBe("alpha");
  });

  it("sync build errors name the method that was called", () => {
    const stream = new ReadableStream<Uint8Array>();
    for (const archive of [Archive.createZip(), Archive.createTar()]) {
      archive.add("a", stream);
      expect(() => archive.toBufferSync()).toThrow(/^Archive\.\w+#toBufferSync: /);
      expect(() => archive.bytesSync()).toThrow(/^Archive\.\w+#bytesSync: /);
    }
  });

  it("TarGzArchive.bytesSync and toBufferSync both return gzip", () => {
    const archive = Archive.createTarGz();
    archive.add("a.txt", "alpha");
    const out = archive.toBufferSync();
    expect(Compression.gunzipSync(out).length).toBeGreaterThan(0);
    expect(archive.bytesSync()).toEqual(out);
  });
});
