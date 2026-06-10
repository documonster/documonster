/**
 * DOCX Module - CFB (Compound File Binary) Reader Tests
 */

import { describe, it, expect } from "vitest";

import { readCfb, writeCfb } from "../security/cfb-reader";
import type { CfbEntry } from "../security/cfb-reader";

describe("readCfb", () => {
  it("rejects file smaller than 512 bytes", () => {
    expect(() => readCfb(new Uint8Array(100))).toThrow(/too small/);
  });

  it("rejects file with invalid signature", () => {
    const bad = new Uint8Array(512);
    // Wrong magic
    bad[0] = 0x11;
    expect(() => readCfb(bad)).toThrow(/invalid signature/);
  });

  it("rejects empty buffer", () => {
    expect(() => readCfb(new Uint8Array(0))).toThrow();
  });
});

describe("writeCfb / readCfb roundtrip", () => {
  it("writes and reads a single entry", () => {
    const entries: CfbEntry[] = [{ name: "TestStream", data: new Uint8Array([1, 2, 3, 4, 5]) }];

    const buffer = writeCfb(entries);
    expect(buffer).toBeInstanceOf(Uint8Array);

    const readBack = readCfb(buffer);
    const found = readBack.find(e => e.name === "TestStream");
    expect(found).toBeDefined();
    expect(Array.from(found!.data)).toEqual([1, 2, 3, 4, 5]);
  });

  it("writes and reads multiple entries", () => {
    const entries: CfbEntry[] = [
      { name: "EncryptionInfo", data: new Uint8Array([0xaa, 0xbb]) },
      { name: "EncryptedPackage", data: new Uint8Array([0xcc, 0xdd, 0xee]) }
    ];

    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);

    const ei = readBack.find(e => e.name === "EncryptionInfo");
    const ep = readBack.find(e => e.name === "EncryptedPackage");
    expect(ei).toBeDefined();
    expect(ep).toBeDefined();
    expect(Array.from(ei!.data)).toEqual([0xaa, 0xbb]);
    expect(Array.from(ep!.data)).toEqual([0xcc, 0xdd, 0xee]);
  });

  it("handles large stream data (mini stream threshold)", () => {
    // Create data larger than 4096 (mini stream cutoff)
    const largeData = new Uint8Array(8192);
    for (let i = 0; i < largeData.length; i++) {
      largeData[i] = i & 0xff;
    }

    const entries: CfbEntry[] = [{ name: "LargeStream", data: largeData }];
    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);
    const found = readBack.find(e => e.name === "LargeStream");
    expect(found).toBeDefined();
    expect(found!.data.length).toBe(8192);
    expect(found!.data[0]).toBe(0);
    expect(found!.data[1]).toBe(1);
    expect(found!.data[8191]).toBe(8191 & 0xff);
  });

  it("handles small stream data (mini sector)", () => {
    const smallData = new Uint8Array([0xde, 0xad, 0xbe, 0xef]);
    const entries: CfbEntry[] = [{ name: "MiniStream", data: smallData }];

    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);
    const found = readBack.find(e => e.name === "MiniStream");
    expect(found).toBeDefined();
    expect(Array.from(found!.data)).toEqual([0xde, 0xad, 0xbe, 0xef]);
  });

  it("preserves stream names correctly", () => {
    const entries: CfbEntry[] = [
      { name: "EncryptionInfo", data: new Uint8Array([1]) },
      { name: "EncryptedPackage", data: new Uint8Array([2]) }
    ];

    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);

    const names = readBack.map(e => e.name).sort();
    expect(names).toContain("EncryptionInfo");
    expect(names).toContain("EncryptedPackage");
  });

  it("handles empty stream data", () => {
    const entries: CfbEntry[] = [{ name: "Empty", data: new Uint8Array(0) }];
    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);
    const found = readBack.find(e => e.name === "Empty");
    expect(found).toBeDefined();
    expect(found!.data.length).toBe(0);
  });

  it("roundtrips streams nested inside storages (DataSpaces-style)", () => {
    const small = new Uint8Array(40).map((_, i) => i & 0xff); // < 4096 → mini-stream
    const big = new Uint8Array(5000).map((_, i) => (i * 7) & 0xff); // regular sector
    const entries: CfbEntry[] = [
      { name: "Version", path: ["\u0006DataSpaces"], data: small },
      {
        name: "StrongEncryptionDataSpace",
        path: ["\u0006DataSpaces", "DataSpaceInfo"],
        data: small
      },
      {
        name: "\u0006Primary",
        path: ["\u0006DataSpaces", "TransformInfo", "StrongEncryptionTransform"],
        data: small
      },
      { name: "EncryptedPackage", data: big }
    ];

    const buffer = writeCfb(entries);
    const readBack = readCfb(buffer);

    const names = readBack.map(e => e.name);
    expect(names).toContain("Version");
    expect(names).toContain("StrongEncryptionDataSpace");
    expect(names).toContain("\u0006Primary");
    expect(names).toContain("EncryptedPackage");

    // Mini-stream entry round-trips byte-for-byte.
    const version = readBack.find(e => e.name === "Version")!;
    expect(Array.from(version.data)).toEqual(Array.from(small));
    // Regular-sector entry round-trips byte-for-byte.
    const pkg = readBack.find(e => e.name === "EncryptedPackage")!;
    expect(Array.from(pkg.data)).toEqual(Array.from(big));
  });

  it("throws rather than emit a corrupt container when the input needs a DIFAT chain", () => {
    // The minimal writer only fills the 109 inline DIFAT slots. A stream large
    // enough to require >109 FAT sectors (~7 MB) must fail loudly instead of
    // silently producing a file no reader can parse.
    const huge = new Uint8Array(8 * 1024 * 1024);
    expect(() => writeCfb([{ name: "EncryptedPackage", data: huge }])).toThrow(/DIFAT/);
  });
});

describe("readCfb — defends against oversized declared sizes", () => {
  it("does not allocate beyond the buffer when a directory entry lies", () => {
    // Build a normal CFB then corrupt the declared stream size in the
    // directory entry to a multi-GiB value. A naive reader would attempt
    // `new Uint8Array(huge)` and crash; the implementation must clamp.
    const entries: CfbEntry[] = [{ name: "Stream1", data: new Uint8Array([1, 2, 3, 4]) }];
    const buffer = writeCfb(entries);

    // Directory sectors live after the FAT; rather than reverse-engineer
    // the layout, scan for the UTF-16LE name "Stream1" and overwrite the
    // 4-byte size field at +120 from the directory entry start (entry
    // size = 128; name field is at offset 0).
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    const utf16 = (s: string): number[] => {
      const out: number[] = [];
      for (const ch of s) {
        out.push(ch.charCodeAt(0) & 0xff, (ch.charCodeAt(0) >> 8) & 0xff);
      }
      return out;
    };
    const needle = utf16("Stream1");
    let dirEntryStart = -1;
    for (let i = 0; i < buffer.length - needle.length; i += 1) {
      let match = true;
      for (let j = 0; j < needle.length; j++) {
        if (buffer[i + j] !== needle[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        dirEntryStart = i;
        break;
      }
    }
    expect(dirEntryStart).toBeGreaterThan(0);

    // size32 lives at +120 (V3 CFB) — set it to 2^31 - 1 (~2 GiB).
    view.setUint32(dirEntryStart + 120, 0x7fffffff, true);

    // The reader must terminate quickly without allocating ~2 GiB.
    const start = Date.now();
    const result = readCfb(buffer);
    expect(Date.now() - start).toBeLessThan(2000);
    const stream = result.find(e => e.name === "Stream1");
    expect(stream).toBeDefined();
    // Whatever data we get back must fit inside the original buffer.
    expect(stream!.data.length).toBeLessThanOrEqual(buffer.length);
  });
});
