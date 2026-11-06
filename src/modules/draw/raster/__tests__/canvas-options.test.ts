/**
 * `BasicRasterCanvas` takes its stroke and text settings as one options object. The
 * positional forms it had (`drawPolyline(points, color, width, dash, shape)`,
 * `drawText(x, y, text, fontSize, color, anchor, rotation, style)`, …) are deprecated and
 * must keep painting exactly what they painted, pixel for pixel.
 */
import { Draw } from "@draw/index";
import type { DrawRasterTextStyle } from "@draw/index";
import { describe, expect, it } from "vitest";

type Canvas = InstanceType<typeof Draw.Backend.BasicRasterCanvas>;

function paint(draw: (canvas: Canvas) => void): Uint8Array {
  const canvas = new Draw.Backend.BasicRasterCanvas(60, 40);
  // Fonts would make the text pixels depend on the host; the stroke font is the same everywhere.
  canvas.setFonts([], { useSystemFonts: false });
  draw(canvas);
  return canvas.data;
}

function inked(data: Uint8Array): boolean {
  return data.some(value => value !== 0);
}

const POINTS = [
  { x: 4, y: 4 },
  { x: 30, y: 30 },
  { x: 56, y: 6 }
];

describe("BasicRasterCanvas options forms", () => {
  it("drawPolyline: the options form paints what the positional form did", () => {
    const legacy = paint(c =>
      c.drawPolyline(POINTS, "#336699cc", 5, [6, 3], { join: "round", cap: "square" })
    );
    const current = paint(c =>
      c.drawPolyline(POINTS, {
        color: "#336699cc",
        width: 5,
        dash: [6, 3],
        join: "round",
        cap: "square"
      })
    );
    expect(inked(current)).toBe(true);
    expect(current).toEqual(legacy);
  });

  it("drawPolyline: defaults agree — a hairline, solid, miter, butt", () => {
    expect(paint(c => c.drawPolyline(POINTS, { color: "#000" }))).toEqual(
      paint(c => c.drawPolyline(POINTS, "#000"))
    );
    // No colour draws nothing, in either form.
    expect(inked(paint(c => c.drawPolyline(POINTS)))).toBe(false);
    expect(inked(paint(c => c.drawPolyline(POINTS, undefined)))).toBe(false);
  });

  it("drawLine and strokeCircle: the options form paints what the positional form did", () => {
    const line = paint(c => c.drawLine(3, 4, 50, 33, { color: "#ff0000", width: 3 }));
    expect(inked(line)).toBe(true);
    expect(line).toEqual(paint(c => c.drawLine(3, 4, 50, 33, "#ff0000", 3)));
    expect(paint(c => c.drawLine(3, 4, 50, 33, { color: "#000" }))).toEqual(
      paint(c => c.drawLine(3, 4, 50, 33, "#000"))
    );
    const circle = paint(c => c.strokeCircle(30, 20, 12, { color: "#00aa00", width: 2 }));
    expect(inked(circle)).toBe(true);
    expect(circle).toEqual(paint(c => c.strokeCircle(30, 20, 12, "#00aa00", 2)));
  });

  it("strokeRect: the options form paints what the positional form did", () => {
    const legacy = paint(c => c.strokeRect(5, 5, 40, 20, "#ff0000", 3));
    const current = paint(c => c.strokeRect(5, 5, 40, 20, { color: "#ff0000", width: 3 }));
    expect(inked(current)).toBe(true);
    expect(current).toEqual(legacy);
    expect(paint(c => c.strokeRect(5, 5, 40, 20, { color: "#ff0000" }))).toEqual(
      paint(c => c.strokeRect(5, 5, 40, 20, "#ff0000"))
    );
  });

  it("strokeEllipse: the options form paints what the positional form did", () => {
    const legacy = paint(c =>
      c.strokeEllipse(30, 20, 20, 12, "#00aa00", 4, [5, 2], { cap: "round" })
    );
    const current = paint(c =>
      c.strokeEllipse(30, 20, 20, 12, { color: "#00aa00", width: 4, dash: [5, 2], cap: "round" })
    );
    expect(inked(current)).toBe(true);
    expect(current).toEqual(legacy);
  });

  it("drawText: the options form paints what the positional form did", () => {
    const style: DrawRasterTextStyle = { family: "Arial", bold: true };
    const rotation = { angle: 20, originX: 30, originY: 20 };
    for (const anchor of ["start", "middle", "end"] as const) {
      const legacy = paint(c => c.drawText(30, 25, "Ab1", 14, "#112233", anchor, rotation, style));
      const current = paint(c =>
        c.drawText(30, 25, "Ab1", {
          fontSize: 14,
          color: "#112233",
          anchor,
          rotation,
          ...style
        })
      );
      expect(inked(current)).toBe(true);
      expect(current).toEqual(legacy);
    }
    // An omitted anchor is `start`, as an omitted (undefined) positional one was.
    expect(paint(c => c.drawText(10, 25, "Ab1", { fontSize: 14, color: "#000" }))).toEqual(
      paint(c => c.drawText(10, 25, "Ab1", 14, "#000", undefined))
    );
  });

  it("getTextWarnings answers as the deprecated textWarnings does", () => {
    const canvas = new Draw.Backend.BasicRasterCanvas(80, 20);
    canvas.setFonts([], { useSystemFonts: false });
    canvas.drawText(2, 15, "مرحبا", { fontSize: 12, color: "#000" });
    expect(canvas.getTextWarnings().length).toBeGreaterThan(0);
    expect(canvas.textWarnings()).toEqual(canvas.getTextWarnings());
  });
});

describe("SvgSurface#format", () => {
  it("serialises what the deprecated markup() did", () => {
    const surface = new Draw.Backend.SvgSurface();
    surface.pushClip({ x: 0, y: 0, width: 4, height: 4 });
    surface.rect(0, 0, 2, 2, 0, { fill: Draw.parseColor("#f00") });
    expect(surface.format()).toContain("<rect");
    expect(surface.markup()).toBe(surface.format());
  });
});
