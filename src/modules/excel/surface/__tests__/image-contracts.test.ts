/**
 * Images and shapes are checked where they are given — a value the drawing part cannot hold used to be written as
 * `NaN`, a negative extent or an unknown geometry, or to fail only when the workbook was written.
 */
import { ExcelArgumentError, Image, ExcelImageError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

function setup() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const id = Image.add(wb, { buffer: PNG, extension: "png" });
  return { wb, ws, id };
}

describe("Image.add", () => {
  it.each([
    [{}],
    [{ buffer: PNG }],
    [{ extension: "png" }],
    [{ buffer: "not bytes", extension: "png" }],
    [{ base64: 42, extension: "png" }],
    [null]
  ])("refuses %j", image => {
    const { wb } = setup();
    expect(() => Image.add(wb, image as never)).toThrow(ExcelImageError);
  });

  it("accepts bytes, base64 or a link", async () => {
    const { wb, ws } = setup();
    Image.place(
      ws,
      Image.add(wb, { base64: btoa(String.fromCharCode(...PNG)), extension: "png" }),
      "B2:C3"
    );
    Image.add(wb, { link: "https://example.com/a.png", extension: "png" });
    await expect(Workbook.toBuffer(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("Image.place", () => {
  it.each([
    [{ tl: { col: Number.NaN, row: 1 }, ext: { width: 10, height: 10 } }],
    [{ tl: { col: 1, row: 1 }, ext: { width: -10, height: 10 } }],
    [{ tl: { col: -1, row: 1 }, ext: { width: 10, height: 10 } }],
    [{ tl: "C3", br: "A1" }],
    [{ tl: { col: 3, row: 3 }, br: { col: 1, row: 1 } }],
    [{ tl: "zz", ext: { width: 1, height: 1 } }],
    [{ tl: { col: 1, row: 1 } }],
    [{ pos: { x: Number.POSITIVE_INFINITY, y: 0 }, ext: { width: 1, height: 1 } }],
    ["zz:yy"]
  ])("refuses %j and places nothing", range => {
    const { ws, id } = setup();
    expect(() => Image.place(ws, id, range as never)).toThrow();
    expect(Image.list(ws)).toEqual([]);
  });

  it("accepts the documented forms", async () => {
    const { wb, ws, id } = setup();
    Image.place(ws, id, "B2:D5");
    Image.place(ws, id, { tl: "B2", br: "C3" });
    Image.place(ws, id, { tl: { col: 1.5, row: 0.5 }, ext: { width: 20, height: 10 } });
    Image.place(ws, id, { pos: { x: 10, y: 10 }, ext: { width: 20, height: 10 } });
    expect(Image.list(ws)).toHaveLength(4);
    await expect(Workbook.toBuffer(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("Image.setModel", () => {
  it("refuses an image id the workbook does not have, and leaves the image as it was", () => {
    const { ws, id } = setup();
    Image.place(ws, id, "B2:C3");
    const image = Image.list(ws)[0];
    const before = Image.getModel(image);
    expect(() => Image.setModel(image, { ...before, imageId: "999" })).toThrow(ExcelImageError);
    expect(() => Image.setModel(image, { ...before, type: "bogus" })).toThrow(ExcelArgumentError);
    expect(() => Image.setModel(image, { ...before, range: "zz" })).toThrow();
    expect(Image.getModel(image)).toEqual(before);
  });

  it("takes the model getModel returns", () => {
    const { ws, id } = setup();
    Image.place(ws, id, "B2:C3");
    const image = Image.list(ws)[0];
    Image.setModel(image, Image.getModel(image));
    expect((Image.getModel(image) as { range?: unknown }).range).toMatchObject({
      tl: { nativeCol: 1, nativeRow: 1 }
    });
  });

  it("refuses an image that was removed", () => {
    const { ws, id } = setup();
    Image.place(ws, id, "B2:C3");
    const image = Image.list(ws)[0];
    Image.remove(ws, image);
    expect(() => Image.setModel(image, Image.getModel(image))).toThrow(ExcelArgumentError);
  });
});

describe("shapes", () => {
  it.each([
    [{ range: "B2:C3", type: "bogus" }],
    [{ range: "B2:C3", fillColor: 42 }],
    [{ range: "B2:C3", lineColor: "red" }],
    [{ range: "B2:C3", lineWidth: -1 }],
    [{ range: "zz:yy" }]
  ])("addShape refuses %j", options => {
    const { ws } = setup();
    expect(() => Image.addShape(ws, options as never)).toThrow();
    expect(Image.listShapes(ws)).toEqual([]);
  });

  it("addShape returns the shape, and removeShape takes it away", () => {
    const { ws } = setup();
    const shape = Image.addShape(ws, {
      range: "B2:C3",
      type: "flowChartProcess",
      fillColor: "#FF0000"
    });
    const other = Image.addShape(ws, { range: "E2:F3", type: "ellipse" });
    expect(Image.listShapes(ws)).toEqual([shape, other]);
    expect(Image.removeShape(ws, shape)).toBe(true);
    expect(Image.removeShape(ws, shape)).toBe(false);
    expect(Image.removeShape(ws, 0)).toBe(true);
    expect(Image.listShapes(ws)).toEqual([]);
  });
});

describe("deprecated detached records", () => {
  it("Image.create and Image.clone place nothing", () => {
    const { ws, id } = setup();
    Image.place(ws, id, "B2:C3");
    Image.create(ws, { type: "image", imageId: String(id), range: "E2:F3" });
    Image.clone(Image.list(ws)[0]);
    expect(Image.list(ws)).toHaveLength(1);
  });
});
