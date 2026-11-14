/**
 * `readCfb` reads a stream whose sectors are scattered through the file — as a file Office has edited in place holds
 * them — the same as one written in consecutive sectors. It copies a run of consecutive sectors in one piece, so the
 * order the chain visits them in, not their order in the file, is what has to come out.
 */
import { readCfb, writeCfb } from "@archive/cfb/cfb";
import { describe, expect, it } from "vitest";

const SECTOR = 512;
const ENDOFCHAIN = 0xfffffffe;

/**
 * `file` with the sectors of the stream named `name` moved so the chain visits them in `order` (a permutation of the
 * chain's positions): the data is moved and the FAT and the directory entry are rewritten to match.
 */
function scatter(file: Uint8Array, name: string, order: (count: number) => number[]): Uint8Array {
  const out = file.slice();
  const view = new DataView(out.buffer);
  const fatSector = view.getUint32(76, true);
  const fatAt = (sector: number): number => (fatSector + 1) * SECTOR + sector * 4;
  const dirAt = (view.getUint32(48, true) + 1) * SECTOR;
  let entry = -1;
  for (let i = 0; i < SECTOR / 128; i++) {
    const nameBytes = view.getUint16(dirAt + i * 128 + 64, true) - 2;
    let entryName = "";
    for (let j = 0; j < nameBytes; j += 2) {
      entryName += String.fromCharCode(view.getUint16(dirAt + i * 128 + j, true));
    }
    if (entryName === name) {
      entry = dirAt + i * 128;
    }
  }
  expect(entry).toBeGreaterThan(-1);
  const chain: number[] = [];
  for (
    let s = view.getUint32(entry + 116, true);
    s !== ENDOFCHAIN;
    s = view.getUint32(fatAt(s), true)
  ) {
    chain.push(s);
  }
  // Position k of the stream goes to the physical sector `chain[order[k]]`.
  const target = order(chain.length).map(k => chain[k]!);
  for (let k = 0; k < chain.length; k++) {
    out.set(
      file.subarray((chain[k]! + 1) * SECTOR, (chain[k]! + 2) * SECTOR),
      (target[k]! + 1) * SECTOR
    );
  }
  for (let k = 0; k < chain.length; k++) {
    view.setUint32(fatAt(target[k]!), k + 1 < chain.length ? target[k + 1]! : ENDOFCHAIN, true);
  }
  view.setUint32(entry + 116, target[0]!, true);
  return out;
}

describe("readCfb sector chains", () => {
  const data = new Uint8Array(40 * SECTOR + 100).map((_, i) => (i * 31 + (i >> 9)) % 256);
  const file = writeCfb([{ name: "Big", data }]);

  it.each([
    ["in reverse", (n: number) => Array.from({ length: n }, (_, k) => n - 1 - k)],
    [
      "in two runs, swapped",
      (n: number) => Array.from({ length: n }, (_, k) => (k + (n >> 1)) % n)
    ],
    [
      "interleaved",
      (n: number) =>
        Array.from({ length: n }, (_, k) => (k % 2 === 0 ? k >> 1 : Math.ceil(n / 2) + (k >> 1)))
    ]
  ])("reads a stream whose sectors lie %s", (_name, order) => {
    const read = readCfb(scatter(file, "Big", order));
    expect(read.find(entry => entry.name === "Big")?.data).toEqual(data);
  });

  it("reads a stream written in consecutive sectors", () => {
    expect(readCfb(file).find(entry => entry.name === "Big")?.data).toEqual(data);
  });

  it("reads every stream of a file holding many", () => {
    const many = Array.from({ length: 300 }, (_, i) => ({
      name: `S${i}`,
      data: new Uint8Array(4096 + i).fill(i & 255)
    }));
    const read = readCfb(writeCfb(many));
    expect(read.map(entry => [entry.name, entry.data])).toEqual(many.map(e => [e.name, e.data]));
  });
});

describe("readCfb on a header whose counts are wrong", () => {
  const file = writeCfb([{ name: "S", data: new Uint8Array(10_000).fill(1) }]);
  const edited = (edit: (view: DataView) => void): Uint8Array => {
    const copy = file.slice();
    edit(new DataView(copy.buffer));
    return copy;
  };

  // A count only bounds the walk: tables that are whole read whatever the header says of them.
  it.each([
    ["more FAT sectors than the file has", (view: DataView) => view.setUint32(44, 4_000_000, true)],
    [
      "more DIFAT sectors than the file has",
      (view: DataView) => view.setUint32(72, 200_000_000, true)
    ]
  ])("reads a file claiming %s", (_name, edit) => {
    expect(readCfb(edited(edit)).map(entry => [entry.name, entry.data.length])).toEqual([
      ["S", 10_000]
    ]);
  });

  // The counts were trusted: an 11 KB file claiming two hundred million DIFAT sectors in a chain that named itself ran
  // for seconds and gigabytes before failing with `Invalid array length`.
  it("refuses a DIFAT chain that comes back on itself, at once", () => {
    const looped = edited(view => {
      view.setUint32(44, 200_000_000, true);
      view.setUint32(68, 0, true); // the DIFAT starts at sector 0…
      view.setUint32(72, 200_000_000, true);
      view.setUint32(512 + 127 * 4, 0, true); // …whose next is sector 0 again
    });
    expect(() => readCfb(looped)).toThrow(/comes back to sector 0/);
  });

  it("refuses a DIFAT listing more FAT sectors than the file holds", () => {
    const repeated = edited(view => {
      view.setUint32(44, 4_000_000, true);
      for (let i = 0; i < 109; i++) {
        view.setUint32(76 + i * 4, 0, true);
      }
    });
    expect(() => readCfb(repeated)).toThrow(/lists 109 FAT sectors; the file holds/);
  });
});

/**
 * A version 4 file — 4096-byte sectors — of one stream `S` holding `data`, its sectors in file order or reversed, and
 * `trailing` bytes after the last sector. The writer writes version 3 only.
 */
function version4File(data: Uint8Array, { reversed = false, trailing = 0 } = {}): Uint8Array {
  const size = 4096;
  const [free, end] = [0xffffffff, 0xfffffffe];
  const count = Math.ceil(data.length / size);
  const file = new Uint8Array(size * (3 + count) + trailing);
  const view = new DataView(file.buffer);
  file.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  for (const [offset, value] of [
    [24, 0x3e],
    [26, 4],
    [28, 0xfffe],
    [30, 12],
    [32, 6]
  ]) {
    view.setUint16(offset!, value!, true);
  }
  for (const [offset, value] of [
    [40, 1],
    [44, 1],
    [48, 1],
    [56, 4096],
    [60, end],
    [68, end]
  ]) {
    view.setUint32(offset!, value!, true);
  }
  for (let i = 0; i < 109; i++) {
    view.setUint32(76 + i * 4, i === 0 ? 0 : free, true);
  }
  // Sector 0 is the FAT, sector 1 the directory, then the stream's sectors.
  const fat = size;
  for (let i = 0; i < size / 4; i++) {
    view.setUint32(fat + i * 4, free, true);
  }
  view.setUint32(fat, 0xfffffffd, true);
  view.setUint32(fat + 4, end, true);
  const sectors = Array.from({ length: count }, (_, i) => (reversed ? 1 + count - i : 2 + i));
  sectors.forEach((sector, i) => {
    view.setUint32(fat + sector * 4, i === count - 1 ? end : sectors[i + 1]!, true);
    file.set(data.subarray(i * size, (i + 1) * size), (sector + 1) * size);
  });
  const entry = (
    index: number,
    name: string,
    type: number,
    start: number,
    length: number,
    child: number
  ) => {
    const at = 2 * size + index * 128;
    for (let i = 0; i < name.length; i++) {
      view.setUint16(at + i * 2, name.charCodeAt(i), true);
    }
    view.setUint16(at + 64, (name.length + 1) * 2, true);
    view.setUint8(at + 66, type);
    view.setUint32(at + 68, free, true);
    view.setUint32(at + 72, free, true);
    view.setUint32(at + 76, child, true);
    view.setUint32(at + 116, start, true);
    view.setBigUint64(at + 120, BigInt(length), true);
  };
  for (let i = 2; i < size / 128; i++) {
    entry(i, "", 0, free, 0, free);
  }
  entry(0, "Root Entry", 5, end, 0, 1);
  entry(1, "S", 2, sectors[0]!, data.length, free);
  return file;
}

describe("readCfb on a version 4 file", () => {
  const data = new Uint8Array(10_000).map((_, i) => (i * 7) & 255);

  it.each<[string, { reversed?: boolean; trailing?: number }]>([
    ["in file order", {}],
    ["in reverse order", { reversed: true }],
    ["followed by bytes past its last sector", { trailing: 100 }]
  ])("reads a stream whose sectors lie %s, copied or as a view", (_name, layout) => {
    const file = version4File(data, layout);
    for (const views of [false, true]) {
      const read = readCfb(file, { views }).find(entry => entry.name === "S")!.data;
      expect(read).toEqual(data);
      // A view only where the stream's sectors are one run; otherwise a copy.
      expect(read.buffer === file.buffer).toBe(views && layout.reversed !== true);
    }
  });
});

describe("readCfb on a directory that overstates its entries", () => {
  const file = writeCfb([
    { name: "A", data: new Uint8Array(8000).fill(1) },
    { name: "B", data: new Uint8Array(8000).fill(2) }
  ]);
  const view = (bytes: Uint8Array) =>
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const directory = (bytes: Uint8Array) => (view(bytes).getUint32(48, true) + 1) * 512;

  // A name size past the 64-byte name field read on through the entry, and past the last entry out of the directory.
  it("reads a name no longer than its field holds", () => {
    const copy = file.slice();
    const at = directory(copy) + 128; // the first stream's entry
    view(copy).setUint16(at + 64, 0xfffe, true);
    const names = readCfb(copy).map(entry => entry.name);
    expect(names[0]).toHaveLength(31);
  });

  // Each sector belongs to one stream; entries naming one chain copied it once each, a file's worth every time.
  it("refuses entries that share sectors beyond what the file holds", () => {
    const copy = file.slice();
    const at = directory(copy);
    const first = {
      start: view(copy).getUint32(at + 128 + 116, true),
      size: view(copy).getUint32(at + 128 + 120, true)
    };
    for (let i = 2; i < 4; i++) {
      view(copy).setUint8(at + i * 128 + 66, 2);
      view(copy).setUint16(at + i * 128, 0x43, true);
      view(copy).setUint16(at + i * 128 + 64, 4, true);
      view(copy).setUint32(at + i * 128 + 116, first.start, true);
      view(copy).setUint32(at + i * 128 + 120, first.size * 4, true);
    }
    expect(() => readCfb(copy)).toThrow(/streams hold more than the file/);
  });
});
