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
});
