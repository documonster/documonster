import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterAll, describe, expect, it } from "vitest";

import { readFileRegionsSync } from "../fs";

const dir = mkdtempSync(join(tmpdir(), "fs-regions-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("readFileRegionsSync", () => {
  it("reads only the ranges asked for, each at its own offset, and leaves the rest zero", () => {
    const path = join(dir, "data.bin");
    writeFileSync(
      path,
      Uint8Array.from({ length: 64 }, (_, i) => i + 1)
    );
    const seen: number[] = [];
    const out = readFileRegionsSync(path, (read, size) => {
      expect(size).toBe(64);
      // A header that says where to read next.
      const header = read(0, 2);
      seen.push(...header);
      read(header[1]! * 10, 3);
    });
    expect(seen).toEqual([1, 2]);
    expect(out.length).toBe(64);
    expect([...out.subarray(0, 2)]).toEqual([1, 2]);
    expect([...out.subarray(20, 23)]).toEqual([21, 22, 23]);
    expect(out.subarray(2, 20).every(byte => byte === 0)).toBe(true);
    expect(out.subarray(23).every(byte => byte === 0)).toBe(true);
  });

  it("clamps a range to the file instead of reading past it", () => {
    const path = join(dir, "short.bin");
    writeFileSync(path, Uint8Array.of(7, 8, 9));
    readFileRegionsSync(path, read => {
      expect([...read(1, 100)]).toEqual([8, 9]);
      expect(read(10, 4).length).toBe(0);
      expect(read(-5, 1).length).toBe(1);
    });
  });
});
