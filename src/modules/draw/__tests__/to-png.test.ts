/**
 * `Draw.toPng` is `Draw.toRgba` plus PNG encoding: the same pixels, the same coverage report, and
 * argument errors that name `Draw.toPng`. `Mermaid.toPng` is the same through the diagram renderer.
 */
import { Archive } from "@archive/index";
import { Draw, DrawArgumentError } from "@draw/index";
import type { DrawList } from "@draw/index";
import { Mermaid, MermaidArgumentError } from "@mermaid/index";
import { describe, expect, it } from "vitest";

const LIST: DrawList = {
  width: 20,
  height: 10,
  children: [
    { kind: "rect", x: 2, y: 2, width: 16, height: 6, paint: { fill: { r: 1, g: 0, b: 0, a: 1 } } }
  ]
};

const SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

function hasChunk(png: Uint8Array, type: string): boolean {
  const bytes = [...type].map(c => c.charCodeAt(0));
  for (let i = 8; i + 4 <= png.length; i++) {
    if (bytes.every((b, k) => png[i + k] === b)) {
      return true;
    }
  }
  return false;
}

describe("Draw.toPng", () => {
  it("encodes exactly what Draw.toRgba paints", () => {
    const image = Draw.toPng(LIST, { scale: 2 });
    const pixels = Draw.toRgba(LIST, { scale: 2 });
    expect([...image.png.subarray(0, 8)]).toEqual(SIGNATURE);
    expect(image.width).toBe(pixels.width);
    expect(image.height).toBe(pixels.height);
    expect(image.png).toEqual(Archive.encodePng(pixels.data, pixels.width, pixels.height));
    expect(image.uncoveredCodePoints).toEqual(pixels.uncoveredCodePoints);
    expect(image.textWarnings).toEqual(pixels.textWarnings);
    expect("data" in image).toBe(false);
  });

  it("writes the resolution when dpi is given", () => {
    expect(hasChunk(Draw.toPng(LIST).png, "pHYs")).toBe(false);
    expect(hasChunk(Draw.toPng(LIST, { dpi: 300 }).png, "pHYs")).toBe(true);
  });

  it("names itself and the field in an argument error", () => {
    expect(() => Draw.toPng(LIST, { dpi: 0 })).toThrow(DrawArgumentError);
    expect(() => Draw.toPng(LIST, { dpi: Number.NaN })).toThrow(/^Draw\.toPng: options\.dpi/);
    expect(() => Draw.toPng(LIST, { scale: -1 })).toThrow(/^Draw\.toPng: options\.scale/);
    expect(() => Draw.toPng({ ...LIST, width: 0 })).toThrow(/^Draw\.toPng: list\.width/);
  });
});

describe("Mermaid.toPng", () => {
  const SOURCE = "flowchart LR\n  A --> B";

  it("encodes the diagram's display list", () => {
    const options = { scale: 2, background: "white" } as const;
    const image = Mermaid.toPng(SOURCE, options);
    const expected = Draw.toPng(Mermaid.toDrawList(SOURCE, options), { scale: 2 });
    expect(image.png).toEqual(expected.png);
    expect(image.width).toBe(expected.width);
  });

  it("names itself in an argument error", () => {
    expect(() => Mermaid.toPng(SOURCE, { scale: 0 })).toThrow(/^Mermaid\.toPng: options\.scale/);
    expect(() => Mermaid.toPng(SOURCE, { fontSize: -1 })).toThrow(MermaidArgumentError);
  });
});
