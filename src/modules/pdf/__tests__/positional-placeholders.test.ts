/**
 * The deprecated positional forms beside an options object behave exactly as the positional
 * signatures did, for every value those accepted — a `null` placeholder, any iterable page list, a
 * positional `scale` that only `undefined` defaulted. Only a genuine options object selects the new
 * form.
 *
 * Expected values were observed against the last release (c3b2407c).
 */
import { Draw } from "@draw/index";
import type { DrawList } from "@draw/types";
import * as PdfMod from "@pdf/index";
import { describe, expect, it } from "vitest";

const { Pdf } = PdfMod;

function loose<T>(value: unknown): T {
  return value as T;
}

async function twoPages(): Promise<Uint8Array> {
  const one = await Pdf.fromTable([["a"]]);
  const editor = Pdf.Editor.read(one);
  editor.copyPagesFrom(one);
  return editor.toBuffer();
}

describe("PdfEditor.copyPagesFrom positional page list", () => {
  it("null copies every page; any iterable of indices is a page list", async () => {
    const source = await twoPages();
    const copied = async (pages: unknown, options?: { password?: string }) => {
      const editor = Pdf.Editor.read(await Pdf.fromTable([["b"]]));
      editor.copyPagesFrom(source, loose<number[]>(pages), options);
      return Pdf.Editor.read(await editor.toBuffer()).pageCount - 1;
    };
    expect(await copied(null)).toBe(2);
    expect(await copied(undefined)).toBe(2);
    expect(await copied([0])).toBe(1);
    expect(await copied(new Int32Array([0]))).toBe(1);
    expect(await copied(new Set([0]))).toBe(1);
    expect(await copied(null, { password: "x" })).toBe(2);
    expect(await copied({ pageIndices: [1] })).toBe(1);
  });
});

describe("Pdf.sign positional certificate", () => {
  it("is handed to the signer unchanged, without a key check of its own", async () => {
    const noPlaceholder = await Pdf.fromTable([["a"]]);
    await expect(
      Pdf.sign(noPlaceholder, new Uint8Array([1, 2]), loose<Uint8Array>(undefined))
    ).rejects.toThrow(/no signature placeholder: no \/ByteRange/);
  });
});

describe("createPdfDrawSurface positional scale", () => {
  const list: DrawList = {
    width: 10,
    height: 10,
    children: [
      { kind: "rect", x: 2, y: 3, width: 4, height: 5, paint: { fill: { r: 1, g: 0, b: 0, a: 1 } } }
    ]
  };

  function recorder() {
    const calls: unknown[][] = [];
    const record =
      (name: string) =>
      (...args: unknown[]) => {
        calls.push([name, ...args]);
      };
    const target = new Proxy({} as Record<string, unknown>, {
      get: (_, name: string) =>
        name === "getContentStream"
          ? () => new Proxy({}, { get: (__, op: string) => record(op) })
          : record(name)
    });
    return { calls, target };
  }

  it("only undefined defaults it: null collapses, a numeric string scales", () => {
    const drawn = (scale: unknown) => {
      const r = recorder();
      Draw.render(
        list,
        PdfMod.createPdfDrawSurface(
          loose<PdfMod.PdfDrawPage>(r.target),
          { x: 1, y: 1, width: 10, height: 10 },
          loose<number>(scale)
        )
      );
      return r.calls;
    };
    const fill = { r: 1, g: 0, b: 0 };
    expect(drawn(null)).toEqual([["drawRect", { x: 1, y: 11, width: 0, height: 0, fill }]]);
    expect(drawn(undefined)).toEqual([["drawRect", { x: 3, y: 3, width: 4, height: 5, fill }]]);
    expect(drawn("2")).toEqual([["drawRect", { x: 5, y: -5, width: 8, height: 10, fill }]]);
    expect(drawn(2)).toEqual(drawn("2"));
  });
});
