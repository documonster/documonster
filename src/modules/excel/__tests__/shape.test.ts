/**
 * Excel Shape Tests
 *
 * Covers Worksheet.addShape:
 * - the worksheet model surfaces shapes with resolved anchors
 * - the drawing XML carries preset geometry, fill, line and text
 * - multiple shapes and different anchor inputs
 */

import { unzip } from "@archive/index";
import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";

async function readDrawingXml(buffer: Uint8Array): Promise<string> {
  const reader = unzip(buffer);
  let drawingKey: string | undefined;
  const all: Record<string, Uint8Array> = {};
  for await (const entry of reader.entries()) {
    const bytes = await entry.bytes();
    all[entry.path] = bytes ?? new Uint8Array();
    if (/drawings\/drawing\d+\.xml$/.test(entry.path)) {
      drawingKey = entry.path;
    }
  }
  if (!drawingKey) {
    throw new Error("no drawing part found");
  }
  return new TextDecoder().decode(all[drawingKey]);
}

describe("Worksheet.addShape", () => {
  it("surfaces shapes on the worksheet model with resolved anchors", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ type: "rect", range: "B2:D5", fillColor: "FFD966", text: "Box" });

    const shapes = ws.getShapes();
    expect(shapes).toHaveLength(1);
    expect(shapes[0].shapeType).toBe("rect");
    expect(shapes[0].fillColor).toBe("FFD966");
    expect(shapes[0].text).toBe("Box");

    const model = ws.model;
    expect(model.shapes).toHaveLength(1);
    expect(model.shapes![0].anchorRange).toBeDefined();
    expect(model.shapes![0].anchorRange!.tl.nativeCol).toBe(1); // column B (0-based)
    expect(model.shapes![0].anchorRange!.tl.nativeRow).toBe(1); // row 2 (0-based)
  });

  it("defaults the preset geometry to rect", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ range: "A1:B2" });
    expect(ws.getShapes()[0].shapeType).toBe("rect");
  });

  it("writes preset geometry, fill, line and text into the drawing part", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({
      type: "rect",
      range: "B2:D5",
      fillColor: "FFD966",
      lineColor: "000000",
      lineWidth: 1,
      text: "Important"
    });

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);

    expect(xml).toContain("xdr:sp");
    expect(xml).toContain('prst="rect"');
    expect(xml).toContain('<a:srgbClr val="FFD966"');
    expect(xml).toContain('<a:srgbClr val="000000"');
    // 1pt line width → 12700 EMU
    expect(xml).toContain('w="12700"');
    expect(xml).toContain("Important");
    // A user shape must NOT be wrapped in the a14 form-control AlternateContent.
    expect(xml).not.toContain("a14:hiddenFill");
    expect(xml).not.toContain("a14:compatExt");
  });

  it("renders each preset geometry and supports multiple shapes", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ type: "rect", range: "B2:D5", fillColor: "FFD966" });
    ws.addShape({ type: "ellipse", range: "F2:H5", fillColor: "9DC3E6" });
    ws.addShape({ type: "line", range: { tl: "B7", br: "E7" }, lineColor: "FF0000", lineWidth: 2 });

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);

    expect(xml).toContain('prst="rect"');
    expect(xml).toContain('prst="ellipse"');
    expect(xml).toContain('prst="line"');
    // Three sp elements.
    expect(xml.match(/<xdr:sp\b/g) ?? []).toHaveLength(3);
    // 2pt line → 25400 EMU.
    expect(xml).toContain('w="25400"');
  });

  it("emits noFill when no fill colour is supplied", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ type: "rect", range: "A1:B2", lineColor: "000000" });

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);

    expect(xml).toContain("<a:noFill");
  });

  it("produces a structurally valid xlsx with shapes", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.getCell("A1").value = "data";
    ws.addShape({ type: "roundRect", range: "B2:D5", fillColor: "FFD966", text: "Note" });

    const buffer = await wb.xlsx.writeBuffer();
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shapes" });
  });

  it("renders a one-cell anchored shape (tl + ext)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({
      type: "rect",
      range: { tl: "B2", ext: { width: 120, height: 80 } } as never,
      fillColor: "FF0000"
    });
    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    expect(xml).toContain("xdr:oneCellAnchor");
    expect(xml).toContain('prst="rect"');
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shape-onecell" });
  });

  it("renders an absolutely-positioned shape (pos + ext)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({
      type: "ellipse",
      range: { pos: { x: 30, y: 60 }, ext: { width: 120, height: 80 } } as never,
      fillColor: "00B050"
    });
    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    expect(xml).toContain("xdr:absoluteAnchor");
    expect(xml).toContain('prst="ellipse"');
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shape-absolute" });
  });

  it("normalizes colours by stripping a leading # and upper-casing", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ type: "rect", range: "A1:B2", fillColor: "#ffd966", lineColor: "#00b050" });
    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    expect(xml).toContain('<a:srgbClr val="FFD966"');
    expect(xml).toContain('<a:srgbClr val="00B050"');
    expect(xml).not.toContain("#ffd966");
  });

  it("drops the alpha byte from 8-digit ARGB colours (srgbClr is RGB-only)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    // excelts cell fills use 8-digit ARGB; addShape must coerce to valid 6-digit RGB.
    ws.addShape({ type: "rect", range: "A1:B2", fillColor: "FFFF0000", lineColor: "FF00B050" });
    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    expect(xml).toContain('<a:srgbClr val="FF0000"');
    expect(xml).toContain('<a:srgbClr val="00B050"');
    expect(xml).not.toContain('val="FFFF0000"');
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shape-argb" });
  });

  it("rejects a range that does not cover an area with a clear error", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    // A single-cell address, an object with no br/ext/pos, and no range at all
    // must all fail fast at addShape — not crash the worksheet serializer later.
    expect(() => ws.addShape({ type: "rect", range: "B2" } as never)).toThrow(/covering an area/);
    expect(() => ws.addShape({ type: "rect", range: { tl: "B2" } } as never)).toThrow(
      /covering an area/
    );
    expect(() => ws.addShape({ type: "rect" } as never)).toThrow(/covering an area/);
    // A rejected shape leaves the worksheet serializable.
    ws.getCell("A1").value = "ok";
    expect(() => ws.model).not.toThrow();
  });

  it("keeps cNvPr ids unique when a shape shares a drawing with a chart", async () => {
    const { installChartSupport } = await import("@excel/chart/install");
    installChartSupport();
    const wb = new Workbook();
    const ws = wb.addWorksheet("mix");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.addChart(
      { type: "bar", barDir: "col", series: [{ values: "mix!$A$1:$A$2" }] } as never,
      "C1:H10"
    );
    ws.addShape({ type: "rect", range: "C12:H20", fillColor: "FFD966" });

    // Round-trip so the chart becomes a preserved graphicFrame, then re-emit.
    const buf1 = await wb.xlsx.writeBuffer();
    const wb2 = new Workbook();
    await wb2.xlsx.load(buf1 as unknown as Uint8Array);
    wb2.getWorksheet("mix")!.addShape({ type: "ellipse", range: "C22:H30", fillColor: "9DC3E6" });

    const xml = await readDrawingXml((await wb2.xlsx.writeBuffer()) as unknown as Uint8Array);
    const ids = [...xml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("escapes special characters in shape text", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.addShape({ type: "rect", range: "A1:B2", text: "a < b & c > d" });
    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    expect(xml).toContain("a &lt; b &amp; c &gt; d");
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shape-text-escape" });
  });

  it("coexists with images and charts without cNvPr id collisions", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("mixed");
    const imageId = wb.addImage({
      base64:
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      extension: "png"
    });
    ws.addImage(imageId, "A1:B2");
    ws.addShape({ type: "rect", range: "D1:E2", fillColor: "FFD966" });
    ws.addShape({ type: "ellipse", range: "G1:H2", fillColor: "9DC3E6" });

    const buffer = await wb.xlsx.writeBuffer();
    const xml = await readDrawingXml(buffer as unknown as Uint8Array);
    const ids = [...xml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m => m[1]);
    // Every cNvPr id within a drawing must be unique.
    expect(new Set(ids).size).toBe(ids.length);
    await expectValidXlsx(buffer as unknown as Uint8Array, { label: "shape-mixed" });
  });

  it("reads a file containing shapes without crashing (shapes are write-only)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("shapes");
    ws.getCell("A1").value = "keep";
    ws.addShape({ type: "rect", range: "B2:D5", fillColor: "FFD966", text: "Box" });
    const buffer = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("shapes")!;
    // Cell data survives; shapes are not parsed back (documented limitation).
    expect(ws2.getCell("A1").value).toBe("keep");
    expect(() => ws2.getShapes()).not.toThrow();
  });
});
