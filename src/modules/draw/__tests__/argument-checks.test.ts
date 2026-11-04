import { Draw, DrawArgumentError, DrawError, Matrix, isDrawError } from "@draw/index";
import type { DrawList, DrawNode } from "@draw/index";
import { describe, expect, it } from "vitest";

const RED = { r: 1, g: 0, b: 0, a: 1 };
const list = (children: unknown[], extra: Record<string, unknown> = {}): DrawList =>
  ({ width: 100, height: 50, children, ...extra }) as unknown as DrawList;
const rect = (over: Record<string, unknown> = {}): DrawNode =>
  ({ kind: "rect", x: 0, y: 0, width: 10, height: 10, paint: { fill: RED }, ...over }) as DrawNode;

function fieldOf(fn: () => unknown): string {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(DrawArgumentError);
    expect(isDrawError(err)).toBe(true);
    return (err as DrawArgumentError).field;
  }
  throw new Error("expected a DrawArgumentError");
}

describe("Draw.toSvg / toRgba / render check the display list", () => {
  it("rejects non-finite geometry instead of writing 0", () => {
    // Was `<rect x="0" y="0" width="0" height="0"/>`.
    expect(fieldOf(() => Draw.toSvg(list([rect({ x: Number.NaN })])))).toBe("list.children[0].x");
    expect(fieldOf(() => Draw.toSvg(list([], { width: Number.NaN })))).toBe("list.width");
    expect(fieldOf(() => Draw.toRgba(list([], { width: Number.NaN })))).toBe("list.width");
  });

  it("rejects a colour that is not an Rgba01", () => {
    // Was `fill="#NaNNaNNaN"`.
    expect(fieldOf(() => Draw.toSvg(list([rect({ paint: { fill: "#ff0000" } })])))).toBe(
      "list.children[0].paint.fill"
    );
    expect(
      fieldOf(() => Draw.toSvg(list([rect({ paint: { fill: { r: 2, g: 0, b: 0, a: 1 } } })])))
    ).toBe("list.children[0].paint.fill.r");
  });

  it("rejects a NaN dash, an unknown kind, a null child and a bad anchor", () => {
    const line = { kind: "line", x1: 0, y1: 0, x2: 10, y2: 0, paint: { stroke: RED } };
    expect(
      fieldOf(() => Draw.toSvg(list([{ ...line, paint: { stroke: RED, dash: [Number.NaN, 2] } }])))
    ).toBe("list.children[0].paint.dash[0]");
    expect(fieldOf(() => Draw.toSvg(list([{ kind: "triangle" }])))).toBe("list.children[0].kind");
    expect(fieldOf(() => Draw.toSvg(list([null])))).toBe("list.children[0]");
    expect(fieldOf(() => Draw.toSvg(undefined as unknown as DrawList))).toBe("list");
    const text = {
      kind: "text",
      x: 0,
      y: 10,
      lines: [{ text: "x", dy: 0 }],
      style: { size: 10, anchor: "center", fill: RED }
    };
    expect(fieldOf(() => Draw.toSvg(list([text])))).toBe("list.children[0].style.anchor");
  });

  it("checks nested group children", () => {
    const group = { kind: "group", children: [rect(), rect({ height: Infinity })] };
    expect(fieldOf(() => Draw.toSvg(list([group])))).toBe("list.children[0].children[1].height");
  });

  it("still renders a valid list", () => {
    expect(Draw.toSvg(list([rect()]))).toContain('<rect x="0" y="0" width="10" height="10"');
    expect(Draw.toRgba(list([rect()])).width).toBe(100);
  });

  it("checks raster options and raises typed errors", () => {
    expect(fieldOf(() => Draw.toRgba(list([]), null as never))).toBe("options");
    expect(fieldOf(() => Draw.toRgba(list([]), { scale: 0 }))).toBe("options.scale");
    expect(fieldOf(() => Draw.toRgba(list([]), { fonts: ["abc"] as never }))).toBe(
      "options.fonts[0]"
    );
    expect(fieldOf(() => Draw.toRgba(list([]), { background: "bogus" }))).toBe(
      "options.background"
    );
    expect(() => Draw.toRgba({ width: 1e6, height: 1e6, children: [] })).toThrow(DrawError);
  });

  it("takes a background as a CSS colour or an IR colour, in both backends (R4P-S6)", () => {
    const red = { r: 1, g: 0, b: 0, a: 1 };
    // The two used to take mutually exclusive forms of the same-named option.
    expect(Draw.toSvg(list([]), { background: red })).toContain('fill="#ff0000"');
    expect(Draw.toSvg(list([]), { background: "red" })).toContain('fill="red"');
    expect(Draw.toSvg(list([]), { background: { ...red, a: 0.5 } })).toContain(
      'fill-opacity="0.5"'
    );
    const pixel = (background: string | typeof red) =>
      Array.from(Draw.toRgba(list([]), { background }).data.slice(0, 4));
    expect(pixel("#ff0000")).toEqual([255, 0, 0, 255]);
    expect(pixel(red)).toEqual([255, 0, 0, 255]);
    expect(pixel("transparent")).toEqual([0, 0, 0, 0]);
    // An unreadable string is refused, as Mermaid refuses it, instead of becoming a fill
    // attribute no viewer reads.
    expect(fieldOf(() => Draw.toSvg(list([]), { background: "nonsense" }))).toBe(
      "options.background"
    );
    expect(fieldOf(() => Draw.toSvg(list([]), { background: { r: 2, g: 0, b: 0, a: 1 } }))).toBe(
      "options.background.r"
    );
  });

  it("starts each message with the public call (X6)", () => {
    for (const [run, prefix] of [
      [() => Draw.toSvg(42 as never), "Draw.toSvg: "],
      [() => Draw.toRgba(list([]), { scale: 0 }), "Draw.toRgba: "],
      [() => Draw.Backend.downsample(new Uint8Array(1), 1, 1, 0), "Draw.Backend.downsample: "],
      [() => Draw.Geometry.flattenPath(42 as never), "Draw.Geometry.flattenPath: "],
      [() => new Draw.Backend.BasicRasterCanvas(0, 1), "new Draw.Backend.BasicRasterCanvas: "],
      [() => Draw.parseRasterFont(42 as never), "Draw.parseRasterFont: "]
    ] as const) {
      try {
        run();
        throw new Error("expected a throw");
      } catch (err) {
        expect((err as Error).message.startsWith(prefix)).toBe(true);
      }
    }
  });

  it("checks the surface given to render", () => {
    expect(fieldOf(() => Draw.render(list([]), null as never))).toBe("surface");
  });
});

describe("Draw colour and text helpers check their arguments", () => {
  it("reads named colours and treats none/transparent as transparent, not black", () => {
    expect(Draw.parseColor("red")).toEqual({ r: 1, g: 0, b: 0, a: 1 });
    expect(Draw.parseColor("transparent")).toEqual({ r: 0, g: 0, b: 0, a: 0 });
    expect(Draw.parseColor("none")).toEqual({ r: 0, g: 0, b: 0, a: 0 });
    expect(Draw.parseColor("bogus")).toEqual({ r: 0, g: 0, b: 0, a: 1 });
    expect(Draw.withAlpha("transparent", 0.5).a).toBe(0);
    expect(fieldOf(() => Draw.parseColor(42 as never))).toBe("token");
  });

  it("rejects an alpha outside 0..1", () => {
    expect(fieldOf(() => Draw.withAlpha("#f00", 2))).toBe("alpha");
    expect(fieldOf(() => Draw.withAlpha("#f00", Number.NaN))).toBe("alpha");
    expect(Draw.withAlpha("#f00", 0.25)).toEqual({ r: 1, g: 0, b: 0, a: 0.25 });
  });

  it("measures only strings at a real size", () => {
    expect(fieldOf(() => Draw.measureText(null as never, { size: 10 }))).toBe("text");
    expect(fieldOf(() => Draw.measureText("abc", { size: Number.NaN }))).toBe("style.size");
    expect(fieldOf(() => Draw.wrapText("a b", { size: 10 }, Number.NaN))).toBe("maxWidth");
    expect(Draw.wrapText("a b", { size: 10 }, Infinity)).toEqual(["a b"]);
    expect(fieldOf(() => Draw.getRelativeLuminance({ r: 1, g: 1, b: 1 } as never))).toBe("color.a");
  });
});

describe("Draw.parseRasterFont reports a font it cannot read", () => {
  it("throws instead of returning an empty face", () => {
    // Was an empty face indistinguishable from a real one.
    expect(() => Draw.parseRasterFont(new Uint8Array(10))).toThrow(DrawError);
    expect(fieldOf(() => Draw.parseRasterFont(null as never))).toBe("data");
    expect(fieldOf(() => Draw.parseRasterFont(new Uint8Array(10), -1))).toBe("collectionIndex");
  });
});

describe("Matrix", () => {
  it("IDENTITY is frozen", () => {
    expect(Object.isFrozen(Matrix.IDENTITY)).toBe(true);
    expect(() => {
      (Matrix.IDENTITY as { e: number }).e = 50;
    }).toThrow(TypeError);
  });

  it("constructors reject non-finite input", () => {
    expect(fieldOf(() => Matrix.translate(Number.NaN, 0))).toBe("tx");
    expect(fieldOf(() => Matrix.rotate(30, Infinity))).toBe("cx");
    expect(fieldOf(() => Matrix.apply(null as never, 0, 0))).toBe("m");
  });
});

describe("Draw.Geometry and Draw.Backend check their arguments", () => {
  it("rejects NaN geometry", () => {
    // `arcToCubics` with a NaN radius returned [] — the arc silently vanished.
    expect(
      fieldOf(() =>
        Draw.Geometry.arcToCubics({ x: 0, y: 0 }, Number.NaN, 5, 0, false, true, { x: 10, y: 0 })
      )
    ).toBe("rx");
    expect(fieldOf(() => Draw.Geometry.flattenPath([{ op: "move", x: 0, y: 0 }], 0))).toBe("scale");
    expect(
      fieldOf(() => Draw.Geometry.createRectNode({ x: 0, y: 0, width: 1, height: 1 }, {}, -1))
    ).toBe("rx");
  });

  it("rejects an unusable canvas size", () => {
    expect(fieldOf(() => new Draw.Backend.BasicRasterCanvas(Number.NaN, 10))).toBe("width");
    expect(fieldOf(() => new Draw.Backend.BasicRasterCanvas(10, 2.5))).toBe("height");
    expect(fieldOf(() => Draw.Backend.downsample(new Uint8Array(64), 4, 4, 3))).toBe("source");
  });
});
