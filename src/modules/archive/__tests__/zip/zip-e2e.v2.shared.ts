/**
 * ZIP End-to-End (E2E) Shared Tests
 *
 * These tests validate the new public `zip()` / `unzip()` API surface and run
 * in both Node.js and Browser environments.
 */

import type { ArchiveSource, UnzipOptions, ZipArchive, ZipOptions, ZipReader } from "@archive";
import { hasSignature } from "@archive/__tests__/zip/zip-test-utils";
import { LOCAL_FILE_HEADER_SIG, END_OF_CENTRAL_DIR_SIG } from "@archive/zip-spec/zip-records";
import { describe, it, expect } from "vitest";

export interface ZipE2EModuleImports {
  zip: (options?: ZipOptions) => ZipArchive;
  unzip: (source: ArchiveSource, options?: UnzipOptions) => ZipReader;
}

function makeTestEntries(): Array<{ name: string; data: Uint8Array }> {
  const text = new TextEncoder();

  const binary256 = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    binary256[i] = i;
  }

  return [
    { name: "empty.txt", data: new Uint8Array(0) },
    { name: "hello.txt", data: text.encode("Hello, ZIP!\n") },
    { name: "dir/subdir/newline.txt", data: text.encode("a\n\nb\r\nc\r\n") },
    { name: "unicode/文件-🌍.txt", data: text.encode("你好世界 🌍 مرحبا العالم") },
    { name: "binary/bytes-0-255.bin", data: binary256 }
  ];
}

async function unzipToMap(
  imports: ZipE2EModuleImports,
  source: ArchiveSource,
  options?: { parse?: { forceStream?: boolean; thresholdBytes?: number } }
): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  const reader = imports.unzip(source, options);
  for await (const entry of reader.entries()) {
    if (entry.type === "directory") {
      entry.discard();
      continue;
    }
    out.set(entry.path, await entry.bytes());
  }
  return out;
}

export function runZipE2ETests(imports: ZipE2EModuleImports): void {
  describe("zip()/unzip() e2e", () => {
    it("zip().bytesSync() roundtrips via unzip().get()", async () => {
      const entries = makeTestEntries();
      const z = imports.zip({ level: 6, comment: "zip-comment" });
      for (const e of entries) {
        z.add(e.name, e.data, { comment: "entry-comment" });
      }
      const zipBytes = z.bytesSync();

      expect(zipBytes).toBeInstanceOf(Uint8Array);
      expect(zipBytes.length).toBeGreaterThan(0);
      expect(hasSignature(zipBytes, LOCAL_FILE_HEADER_SIG, 0, 64)).toBe(true);
      expect(
        hasSignature(zipBytes, END_OF_CENTRAL_DIR_SIG, zipBytes.length - 256, zipBytes.length)
      ).toBe(true);

      const reader = imports.unzip(zipBytes);
      for (const e of entries) {
        const got = await reader.get(e.name);
        expect(got).not.toBeNull();
        expect(await got!.bytes()).toEqual(e.data);
      }
    });

    it("zip().bytes() (async) roundtrips via unzip().entries()", async () => {
      const entries = makeTestEntries();
      const z = imports.zip({ level: 6 });
      for (const e of entries) {
        z.add(e.name, e.data);
      }
      const zipBytes = await z.bytes();

      const extracted = await unzipToMap(imports, zipBytes);
      expect(extracted.size).toBe(entries.length);
      for (const e of entries) {
        expect(extracted.get(e.name)).toEqual(e.data);
      }
    });

    it("unzip(stream) works in streaming parse mode", async () => {
      // Keep this case ASCII-only to validate the streaming parser path.
      // UTF-8 filenames are already covered by the buffer-mode unzip tests.
      const entries = makeTestEntries().filter(e => !e.name.includes("unicode/"));
      const z = imports.zip({ level: 6 });
      for (const e of entries) {
        z.add(e.name, e.data);
      }
      // Use bytesSync here to validate the streaming unzip path. The async ZIP
      // writer uses data descriptors and is covered by the buffer-mode tests above.
      const zipBytes = z.bytesSync();

      const source: AsyncIterable<Uint8Array> = (async function* () {
        yield zipBytes;
      })();

      const extracted = await unzipToMap(imports, source, { parse: { forceStream: true } });
      expect(extracted.size).toBe(entries.length);
      for (const e of entries) {
        expect(extracted.get(e.name)).toEqual(e.data);
      }
    });

    it("bytesSync and bytes outputs unzip identically", async () => {
      const entries = makeTestEntries();

      const zSync = imports.zip({ level: 6 });
      for (const e of entries) {
        zSync.add(e.name, e.data);
      }
      const bytesSync = zSync.bytesSync();

      const zAsync = imports.zip({ level: 6 });
      for (const e of entries) {
        zAsync.add(e.name, e.data);
      }
      const bytesAsync = await zAsync.bytes();

      const a = await unzipToMap(imports, bytesSync);
      const b = await unzipToMap(imports, bytesAsync);
      expect(a.size).toBe(b.size);
      for (const [k, v] of a.entries()) {
        expect(b.get(k)).toEqual(v);
      }
    });
  });
}
