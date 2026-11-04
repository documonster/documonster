import * as DrawEntry from "@draw/index";
import { Draw, Matrix } from "@draw/index";
import { describe, expect, it } from "vitest";

describe("documonster/draw namespaces", () => {
  it("publishes the renamed members", () => {
    expect(typeof Draw.toSvg).toBe("function");
    expect(typeof Draw.toRgba).toBe("function");
    expect(typeof Draw.render).toBe("function");
    expect(Draw.parseColor("#ff0000")).toEqual({ r: 1, g: 0, b: 0, a: 1 });
    expect(Matrix.getRotation(Matrix.rotate(30))).toBeCloseTo(30, 9);
    expect(Matrix.getUniformScale(Matrix.scale(2))).toBeCloseTo(2, 9);
  });

  it("Backend.SvgSurface accumulates markup", () => {
    const surface = new Draw.Backend.SvgSurface();
    surface.rect(0, 0, 2, 2, 0, { fill: Draw.parseColor("#f00") });
    expect(surface.markup()).toContain("<rect");
  });

  it("createRasterSurface takes { scale }", () => {
    const canvas = new Draw.Backend.BasicRasterCanvas(4, 4);
    const surface = Draw.Backend.createRasterSurface(canvas, { scale: 2 });
    expect(typeof surface.rect).toBe("function");
  });

  it("keeps the deprecated flat names as forwards", () => {
    expect(DrawEntry.cssColour).toBe(Draw.parseColor);
    expect(DrawEntry.renderDrawList).toBe(Draw.render);
    expect(DrawEntry.textFeaturesOf).toBe(Draw.getTextFeatures);
    expect(DrawEntry.SvgSurface).toBe(Draw.Backend.SvgSurface);
    expect(DrawEntry.BasicRasterCanvas).toBe(Draw.Backend.BasicRasterCanvas);
    expect(DrawEntry.measureText).toBe(Draw.measureText);
    expect(DrawEntry.IDENTITY).toBe(Matrix.IDENTITY);
    expect(DrawEntry.translate).toBe(Matrix.translate);
    expect(DrawEntry.rotationOf).toBe(Matrix.getRotation);
    expect(DrawEntry.uniformScale).toBe(Matrix.getUniformScale);
  });

  it("the deprecated rasterizeToRgba and toSvg keep their old, unchecked reading (R4P-D2)", () => {
    const empty = { width: 0, height: 0, children: [] };
    // `Draw.toRgba` refuses a list with no area; the old name rasterised it to one pixel.
    expect(() => Draw.toRgba(empty)).toThrow(/list\.width/);
    const image = DrawEntry.rasterizeToRgba(empty);
    expect([image.width, image.height]).toEqual([1, 1]);
    expect(Array.from(image.data)).toEqual([0, 0, 0, 0]);
    // `Draw.toSvg` refuses an unreadable background; the old name wrote what it was given.
    const list = { width: 10, height: 10, children: [] };
    expect(() => Draw.toSvg(list, { background: "nonsense" })).toThrow(/options\.background/);
    expect(DrawEntry.toSvg(list, { background: "nonsense" })).toContain('fill="nonsense"');
  });

  it("the deprecated createRasterSurface accepts a positional scale", () => {
    const paint = { fill: { r: 0, g: 0, b: 0, a: 1 } };
    const legacy = new Draw.Backend.BasicRasterCanvas(8, 8);
    const current = new Draw.Backend.BasicRasterCanvas(8, 8);
    DrawEntry.createRasterSurface(legacy, 2).rect(1, 1, 2, 2, 0, paint);
    Draw.Backend.createRasterSurface(current, { scale: 2 }).rect(1, 1, 2, 2, 0, paint);
    expect(legacy.data.some(v => v !== 0)).toBe(true);
    expect(legacy.data).toEqual(current.data);
  });

  // Observed against the last release (c3b2407c): its positional `scale = 1` defaulted only
  // `undefined`, so a `null` placeholder scaled everything to the origin.
  it("the deprecated createRasterSurface defaults only an undefined scale", () => {
    const calls: unknown[][] = [];
    const canvas = new Proxy({} as Draw.Backend.BasicRasterCanvas, {
      get:
        (_, name: string) =>
        (...args: unknown[]) => {
          calls.push([name, ...args]);
        }
    });
    const paint = { fill: { r: 1, g: 0, b: 0, a: 1 } };
    const drawn = (scale: unknown) => {
      calls.length = 0;
      DrawEntry.createRasterSurface(canvas, scale as number).rect(2, 3, 4, 5, 0, paint);
      return [...calls];
    };
    expect(drawn(null)).toEqual([["fillRect", 0, 0, 0, 0, "#ff0000ff"]]);
    expect(drawn(undefined)).toEqual([["fillRect", 2, 3, 4, 5, "#ff0000ff"]]);
    expect(drawn("2")).toEqual([["fillRect", 4, 6, 8, 10, "#ff0000ff"]]);
  });
});
