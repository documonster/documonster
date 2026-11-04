/**
 * The order of a package's entries is not part of what it means: the same workbook repackaged with its
 * entries in another order — sorted by name, say, which puts `image10` before `image2` — must read into
 * the same images, with the same ids.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Cell, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

/** A 1×1 PNG whose single pixel is `shade`, so each image's bytes differ. */
function png(shade: number): Uint8Array {
  const bytes = Uint8Array.from(
    atob(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
    ),
    c => c.charCodeAt(0)
  );
  // A trailing byte after IEND is ignored by every decoder and tells the images apart.
  return Uint8Array.from([...bytes, shade]);
}

async function book(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  for (let i = 0; i < 12; i++) {
    const id = Image.add(wb, { buffer: png(i), extension: i % 3 === 0 ? "jpeg" : "png" });
    Image.place(ws, id, `${String.fromCharCode(66 + i)}2:${String.fromCharCode(66 + i)}3`);
  }
  return Workbook.toBuffer(wb);
}

/** `bytes` repackaged with its entries in `order`. */
async function repackaged(
  bytes: Uint8Array,
  order: (names: string[]) => string[]
): Promise<Uint8Array> {
  const entries = await extractAll(bytes);
  return createZip(
    order([...entries.keys()]).map(name => ({ name, data: entries.get(name)!.data })),
    { noSort: true }
  );
}

/** Each image of sheet `S`: its id, where it sits, and the bytes of the medium it shows. */
async function images(bytes: Uint8Array) {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes);
  const media = Workbook.getModel(wb).media as { buffer: Uint8Array; extension: string }[];
  return Image.list(Workbook.getWorksheet(wb, "S")!).map(image => ({
    id: image.imageId,
    range: (Image.getModel(image) as { range?: unknown }).range,
    medium: [
      media[Number(image.imageId)].extension,
      ...media[Number(image.imageId)].buffer.slice(-1)
    ]
  }));
}

describe("media are read independently of the package's entry order", () => {
  it("gives the same ids however the entries are ordered", async () => {
    const original = await book();
    const expected = await images(original);
    expect(expected.map(image => Number(image.id))).toEqual(expected.map((_, i) => i));
    for (const order of [
      (names: string[]) => [...names].sort(),
      (names: string[]) => [...names].reverse()
    ]) {
      expect(await images(await repackaged(original, order))).toEqual(expected);
    }
  });

  it("keeps the ids of a file this library wrote through a write and a second read", async () => {
    const first = await images(await book());
    const wb = Workbook.create();
    await Workbook.read(wb, await book());
    expect(await images(await Workbook.toBuffer(wb))).toEqual(first);
  });
});
