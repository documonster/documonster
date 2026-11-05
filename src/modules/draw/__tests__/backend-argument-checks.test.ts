/**
 * The published backend objects — `Draw.Backend.SvgSurface`, `Draw.Backend.BasicRasterCanvas`
 * and the surface `Draw.Backend.createRasterSurface` returns — refuse a value they cannot read
 * with a `DrawArgumentError` naming the call and the parameter, where each used to fail with a
 * `TypeError` from inside a primitive. A `NaN` still draws nothing, as it always has.
 */
import { Draw, DrawArgumentError } from "@draw/index";
import { describe, expect, it } from "vitest";

const RED = { r: 1, g: 0, b: 0, a: 1 };

function refusal(call: () => unknown): DrawArgumentError {
  try {
    call();
  } catch (err) {
    expect(err).toBeInstanceOf(DrawArgumentError);
    return err as DrawArgumentError;
  }
  throw new Error("expected a DrawArgumentError, nothing was thrown");
}

function expectRefusal(call: () => unknown, where: string, field: string): void {
  const err = refusal(call);
  expect(err.field).toBe(field);
  expect(err.message.startsWith(`${where}: ${field} must be`)).toBe(true);
}

describe("Draw.Backend.SvgSurface", () => {
  const W = "Draw.Backend.SvgSurface";
  it("names the parameter a primitive cannot read", () => {
    const s = new Draw.Backend.SvgSurface();
    expectRefusal(() => s.rect(0, 0, 1, 1, 0, null as never), `${W}#rect`, "paint");
    expectRefusal(() => s.rect(Symbol("x") as never, 0, 1, 1, 0, {}), `${W}#rect`, "x");
    expectRefusal(() => s.ellipse(0, 0, 1, 1, undefined as never), `${W}#ellipse`, "paint");
    expectRefusal(() => s.sector(0, 0, 1, 0, 0, 1, 42 as never), `${W}#sector`, "paint");
    expectRefusal(() => s.polyline(null as never, false, {}), `${W}#polyline`, "points");
    expectRefusal(() => s.polyline([null] as never, false, {}), `${W}#polyline`, "points[0]");
    expectRefusal(() => s.path({} as never, {}), `${W}#path`, "commands");
    expectRefusal(() => s.path([null] as never, {}), `${W}#path`, "commands[0]");
    expectRefusal(() => s.text(0, 0, [null] as never, { size: 10 }, 0), `${W}#text`, "lines[0]");
    expectRefusal(
      () => s.text(0, 0, [{ text: 1, dy: 0 }] as never, { size: 10 }, 0),
      `${W}#text`,
      "lines[0].text"
    );
    expectRefusal(() => s.text(0, 0, [], null as never, 0), `${W}#text`, "style");
    expectRefusal(() => s.pushClip(null as never), `${W}#pushClip`, "clip");
    expectRefusal(
      () => s.pushClip({ x: 0, y: 0, width: "1" } as never),
      `${W}#pushClip`,
      "clip.width"
    );
    expectRefusal(() => s.pushFilter({} as never), `${W}#pushFilter`, "id");
    expectRefusal(() => s.addDefs(null as never), `${W}#addDefs`, "markup");
    expectRefusal(() => s.addDefs([1] as never), `${W}#addDefs`, "markup[0]");
    // Nothing was written by the refused calls.
    expect(s.format()).toBe("");
  });

  it("still writes what the walker hands it", () => {
    const s = new Draw.Backend.SvgSurface();
    s.rect(0, 0, 2, 2, 0, { fill: RED });
    s.text(0, 0, [{ text: "a", dy: 0 }], { size: 10, fill: RED }, 0);
    expect(s.format()).toContain("<rect");
    expect(s.format()).toContain(">a</text>");
  });
});

describe("Draw.Backend.BasicRasterCanvas", () => {
  const W = "Draw.Backend.BasicRasterCanvas";
  it("names the parameter a primitive cannot read", () => {
    const c = new Draw.Backend.BasicRasterCanvas(4, 4);
    expectRefusal(() => c.fillRect(Symbol("x") as never, 0, 1, 1, "#f00"), `${W}#fillRect`, "x");
    expectRefusal(() => c.fillRect(0, 0, 1, 1, {} as never), `${W}#fillRect`, "color");
    expectRefusal(() => c.fillRoundRect(0, 0, 4, 4, 1, 42 as never), `${W}#fillRoundRect`, "color");
    expectRefusal(() => c.strokeRect(0, 0, 4, 4, 42 as never), `${W}#strokeRect`, "options");
    expectRefusal(
      () => c.strokeRect(0, 0, 4, 4, { color: 1 } as never),
      `${W}#strokeRect`,
      "options.color"
    );
    expectRefusal(() => c.drawPolyline(null as never), `${W}#drawPolyline`, "points");
    expectRefusal(
      () => c.drawPolyline([{ x: 0, y: 0 }], { dash: ["2"] } as never),
      `${W}#drawPolyline`,
      "options.dash[0]"
    );
    expectRefusal(
      () => c.drawLine(0, 0, 1, 1, "#f00", Symbol("w") as never),
      `${W}#drawLine`,
      "width"
    );
    expectRefusal(() => c.fillCircle(0, 0, 1, {} as never), `${W}#fillCircle`, "color");
    expectRefusal(() => c.fillEllipse(0, 0, 1, 1, {} as never), `${W}#fillEllipse`, "color");
    expectRefusal(() => c.strokeEllipse(0, 0, 1, 1, 42 as never), `${W}#strokeEllipse`, "options");
    expectRefusal(() => c.strokeCircle(0, 0, 1, {} as never), `${W}#strokeCircle`, "color");
    expectRefusal(() => c.fillSector(0, 0, 1, 0, 0, 1, {} as never), `${W}#fillSector`, "color");
    expectRefusal(() => c.fillPolygon(null as never, "#f00"), `${W}#fillPolygon`, "points");
    expectRefusal(
      () => c.fillRings([null] as never, "#f00", "nonzero"),
      `${W}#fillRings`,
      "rings[0]"
    );
    expectRefusal(() => c.drawText(0, 0, {} as never, { fontSize: 4 }), `${W}#drawText`, "text");
    expectRefusal(
      () => c.drawText(0, 0, "A", { fontSize: "4" } as never),
      `${W}#drawText`,
      "options.fontSize"
    );
    expectRefusal(
      () => c.drawText(0, 0, "A", { fontSize: 4, rotation: null } as never),
      `${W}#drawText`,
      "options.rotation"
    );
    expectRefusal(() => c.setFonts(null as never), `${W}#setFonts`, "fonts");
    expectRefusal(() => c.setFonts([null] as never), `${W}#setFonts`, "fonts[0]");
    expectRefusal(() => c.setFonts([], null as never), `${W}#setFonts`, "options");
    expectRefusal(() => c.pushClip(Symbol("x") as never, 0, 1, 1), `${W}#pushClip`, "x");
    expect(c.data.every(v => v === 0)).toBe(true);
  });

  it("keeps drawing nothing for a NaN, and a null colour, as before", () => {
    const c = new Draw.Backend.BasicRasterCanvas(4, 4);
    c.fillRect(Number.NaN, 0, 2, 2, "#f00");
    c.fillRect(0, 0, 2, 2, null as never);
    c.drawLine(0, 0, 3, 3, undefined);
    expect(c.data.every(v => v === 0)).toBe(true);
  });

  it("does not walk a line, circle or ellipse beyond the canvas", () => {
    // A `NaN` end never compared equal and looped forever; a far end walked one step per
    // pixel of the line, and a large radius visited its whole bounding box.
    const c = new Draw.Backend.BasicRasterCanvas(8, 8);
    const start = performance.now();
    c.drawLine(0, 0, Number.NaN, 3, "#f00");
    c.drawLine(0, 4, 1e12, 4, "#f00");
    c.fillCircle(4, 4, 1e9, "#00f");
    c.fillEllipse(4, 4, 1e9, 1e9, "#0f0");
    expect(performance.now() - start).toBeLessThan(2000);
    // The ellipse covers every pixel it was asked to.
    for (let i = 0; i < c.data.length; i += 4) {
      expect(Array.from(c.data.slice(i, i + 4))).toEqual([0, 255, 0, 255]);
    }
  });

  it("draws a clipped line through the same pixels on the canvas", () => {
    const near = new Draw.Backend.BasicRasterCanvas(16, 16);
    const far = new Draw.Backend.BasicRasterCanvas(16, 16);
    near.drawLine(0, 8, 15, 8, "#f00");
    far.drawLine(-1e9, 8, 1e9, 8, "#f00");
    expect(far.data).toEqual(near.data);
  });
});

describe("Draw.Backend.createRasterSurface(…) and normalizeSamples", () => {
  it("checks what a caller hands the returned surface", () => {
    const surface = Draw.Backend.createRasterSurface(new Draw.Backend.BasicRasterCanvas(4, 4));
    const W = "Draw.Backend.createRasterSurface(…)";
    expectRefusal(() => surface.rect(0, 0, 1, 1, 0, null as never), `${W}.rect`, "paint");
    expectRefusal(() => surface.path(null as never, {}), `${W}.path`, "commands");
    expectRefusal(() => surface.text(0, 0, null as never, { size: 1 }, 0), `${W}.text`, "lines");
    expectRefusal(() => surface.pushClip!(undefined as never), `${W}.pushClip`, "clip");
  });

  it("names the call when samples has no reading", () => {
    for (const bad of [null, "x", Number.NaN, 0, Infinity]) {
      expectRefusal(
        () => Draw.Backend.normalizeSamples(bad as never, 10, 10),
        "Draw.Backend.normalizeSamples",
        "samples"
      );
    }
    expect(Draw.Backend.normalizeSamples(undefined, 10, 10)).toBe(3);
  });

  it("renders a line reaching far beyond the list promptly through Draw.toRgba", () => {
    const image = Draw.toRgba({
      width: 4,
      height: 4,
      children: [{ kind: "line", x1: 0, y1: 1, x2: 1e9, y2: 1, paint: { stroke: RED } }]
    });
    expect(image.width).toBe(4);
  });
});
