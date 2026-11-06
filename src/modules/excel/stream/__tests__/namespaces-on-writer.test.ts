/**
 * The document model's namespaces take a streaming writer and hand it to the writer's own implementation, so one
 * vocabulary covers both: what a namespace does to a writer reaches the file, the writer's committed-state checks
 * apply, an operation a writer cannot do is refused by name, and the deprecated class methods still write what the
 * namespace writes.
 */
import {
  Cell,
  ConditionalFormatting,
  DataValidation,
  ExcelImageError,
  ExcelStreamStateError,
  ExcelUnsupportedFeatureError,
  Image,
  Stream,
  Watermark,
  Workbook,
  Worksheet
} from "@excel/index";
import type { ConditionalFormattingOptions, DataValidationRule } from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

const CF: ConditionalFormattingOptions = {
  ref: "a1:a9",
  rules: [
    {
      type: "cellIs",
      operator: "greaterThan",
      formulae: [5],
      style: { font: { bold: true } },
      priority: 1
    }
  ]
};

/** Fill a streamed sheet with `build`, commit the workbook, and read the sheet back. */
async function roundTrip(
  build: (wb: Stream.WorkbookWriter, ws: Stream.WorksheetWriter) => void | Promise<void>
): Promise<Workbook.Handle> {
  const wb = new Stream.WorkbookWriter({ useStyles: true });
  const ws = wb.addWorksheet("S");
  await build(wb, ws);
  return Workbook.read(await wb.toBuffer());
}

function sheetOf(book: Workbook.Handle): Worksheet.Handle {
  return Workbook.getWorksheet(book, "S")!;
}

describe("namespaces on a streaming worksheet writer", () => {
  it("Worksheet.merge merges, and the deprecated mergeCells writes the same", async () => {
    for (const merge of [
      (ws: Stream.WorksheetWriter) => Worksheet.merge(ws, "a1:b2"),
      (ws: Stream.WorksheetWriter) => ws.mergeCells("A1", "B2")
    ]) {
      const book = await roundTrip((_wb, ws) => {
        ws.addRow(["x", "y"]);
        merge(ws);
      });
      expect(Worksheet.getMergedRegions(sheetOf(book))).toEqual([
        { top: 1, left: 1, bottom: 2, right: 2 }
      ]);
      expect(Cell.getValue(sheetOf(book), "A1")).toBe("x");
    }
  });

  it("Worksheet.mergeWithoutStyle leaves each cell's style", async () => {
    const book = await roundTrip((_wb, ws) => {
      const row = ws.addRow(["x", "y"]);
      Stream.setCellFont(Stream.getRowCell(row, 2), { italic: true });
      Worksheet.mergeWithoutStyle(ws, "A1:B1");
    });
    expect(Worksheet.getMergedRegions(sheetOf(book))).toEqual([
      { top: 1, left: 1, bottom: 1, right: 2 }
    ]);
    expect(Cell.getFont(sheetOf(book), "B1")?.italic).toBe(true);
  });

  it("Worksheet.merge refuses a committed sheet and a committed row", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    Stream.commitRow(ws.addRow(["a"]));
    expect(() => Worksheet.merge(ws, "A1:B1")).toThrow(/committed/);
    ws.commit();
    expect(() => Worksheet.merge(ws, "A5:B5")).toThrow(ExcelStreamStateError);
  });

  it("Worksheet.unmerge is refused by name", () => {
    const ws = new Stream.WorkbookWriter({}).addWorksheet("S");
    expect(() => Worksheet.unmerge(ws as never, "A1:B1")).toThrow(ExcelUnsupportedFeatureError);
  });

  it("Image.add and Image.place put a picture on the sheet; list and remove see it", async () => {
    const book = await roundTrip((wb, ws) => {
      const id = Image.add(wb, { buffer: PNG, extension: "png" });
      const placed = Image.place(ws, id, "b2:c3");
      Image.place(ws, id, "E5:F6");
      expect(Image.list(ws)).toHaveLength(2);
      expect(Image.list(ws)[0]).toBe(placed);
      expect(Image.remove(ws, 1)).toBe(true);
      expect(Image.list(ws)).toEqual([placed]);
    });
    expect(Image.list(sheetOf(book))).toHaveLength(1);
  });

  it("the deprecated addImage / getImages write what Image.add / Image.place write", async () => {
    const book = await roundTrip((wb, ws) => {
      const id = wb.addImage({ buffer: PNG, extension: "png" });
      ws.addImage(id, "B2:C3");
      expect(ws.getImages()).toEqual(Image.list(ws));
    });
    expect(Image.list(sheetOf(book))).toHaveLength(1);
  });

  it("Image.place names Image.place when the image is unknown, and refuses a committed sheet", () => {
    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    expect(() => Image.place(ws, 7, "A1:B2")).toThrow(/Image\.place.*no image 7/);
    const id = Image.add(wb, { buffer: PNG, extension: "png" });
    ws.commit();
    expect(() => Image.place(ws, id, "A1:B2")).toThrow(ExcelStreamStateError);
  });

  it("Image.setBackground / getBackground / removeBackground", async () => {
    const book = await roundTrip((wb, ws) => {
      const id = Image.add(wb, { buffer: PNG, extension: "png" });
      expect(Image.getBackground(ws)).toBeUndefined();
      Image.setBackground(ws, id);
      expect(Image.getBackground(ws)).toBe(String(id));
      expect(ws.getBackgroundImageId()).toBe(id);
    });
    expect(Image.getBackground(sheetOf(book))).toBeDefined();

    const removed = await roundTrip((wb, ws) => {
      Image.setBackground(ws, Image.add(wb, { buffer: PNG, extension: "png" }));
      expect(Image.removeBackground(ws)).toBe(true);
      expect(Image.removeBackground(ws)).toBe(false);
    });
    expect(Image.getBackground(sheetOf(removed))).toBeUndefined();
  });

  it("shapes are refused by name", () => {
    const ws = new Stream.WorkbookWriter({}).addWorksheet("S");
    expect(() => Image.listShapes(ws as never)).toThrow(ExcelUnsupportedFeatureError);
    expect(() => Image.addShape(ws as never, { type: "rect", range: "A1:B2" } as never)).toThrow(
      ExcelUnsupportedFeatureError
    );
  });

  it("Watermark.set / get / remove, and the deprecated methods agree", async () => {
    const book = await roundTrip((wb, ws) => {
      const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
      Watermark.set(ws, { imageId, opacity: 0.3 });
      const got = Watermark.get(ws);
      expect(got).toMatchObject({ imageId: String(imageId), mode: "overlay", opacity: 0.3 });
      expect(got).toEqual(ws.getWatermark());
      // The watermark's own drawing entry is not a placed picture.
      expect(Image.list(ws)).toEqual([]);
    });
    // The overlay is written as a transparent picture.
    expect(Image.list(sheetOf(book)).length + (Watermark.get(sheetOf(book)) ? 1 : 0)).toBe(1);

    const wb = new Stream.WorkbookWriter({});
    const ws = wb.addWorksheet("S");
    const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
    ws.addWatermark({ imageId });
    expect(Watermark.remove(ws)).toBe(true);
    expect(Watermark.get(ws)).toBeUndefined();
    expect(() => Watermark.set(ws, { imageId, mode: "header" })).toThrow(ExcelImageError);
    expect(() => Watermark.set(ws, { imageId, opacity: 2 })).toThrow(/options\.opacity/);
    ws.commit();
    expect(() => Watermark.set(ws, { imageId })).toThrow(ExcelStreamStateError);
    expect(() => Watermark.remove(ws)).toThrow(ExcelStreamStateError);
  });

  it("ConditionalFormatting.add / list / remove / clear", async () => {
    const book = await roundTrip((_wb, ws) => {
      const block = ConditionalFormatting.add(ws, CF);
      expect(block.ref).toBe("A1:A9");
      expect(ConditionalFormatting.list(ws)).toEqual([block]);
      ConditionalFormatting.add(ws, { ...CF, ref: "C1:C9" });
      expect(ConditionalFormatting.remove(ws, "c1:c9")).toBe(true);
      expect(ConditionalFormatting.remove(ws, 5)).toBe(false);
    });
    expect(ConditionalFormatting.list(sheetOf(book)).map(cf => cf.ref)).toEqual(["A1:A9"]);

    const cleared = await roundTrip((_wb, ws) => {
      ConditionalFormatting.add(ws, CF);
      ConditionalFormatting.clear(ws);
      expect(ConditionalFormatting.list(ws)).toEqual([]);
    });
    expect(ConditionalFormatting.list(sheetOf(cleared))).toEqual([]);
  });

  it("the deprecated addConditionalFormatting writes what ConditionalFormatting.add writes", async () => {
    const write = (add: (ws: Stream.WorksheetWriter) => void) =>
      roundTrip((_wb, ws) => add(ws)).then(book => ConditionalFormatting.list(sheetOf(book)));
    expect(await write(ws => ws.addConditionalFormatting(CF))).toEqual(
      await write(ws => ConditionalFormatting.add(ws, CF))
    );
  });

  it("ConditionalFormatting edits refuse a committed sheet", () => {
    const ws = new Stream.WorkbookWriter({}).addWorksheet("S");
    ws.commit();
    expect(() => ConditionalFormatting.add(ws, CF)).toThrow(ExcelStreamStateError);
    expect(() => ConditionalFormatting.remove(ws, 0)).toThrow(ExcelStreamStateError);
    expect(() => ConditionalFormatting.clear(ws)).toThrow(ExcelStreamStateError);
  });

  it("DataValidation.add reaches the file, and refuses a committed sheet", async () => {
    const rule: DataValidationRule = { type: "whole", operator: "between", formulae: [1, 9] };
    const book = await roundTrip((_wb, ws) => {
      DataValidation.add(ws, "B2:B5", rule);
      expect(DataValidation.find(ws, "B3")).toMatchObject(rule);
    });
    expect(DataValidation.find(sheetOf(book), "B4")).toMatchObject(rule);

    const ws = new Stream.WorkbookWriter({}).addWorksheet("S");
    ws.commit();
    expect(() => DataValidation.add(ws, "A1", rule)).toThrow(ExcelStreamStateError);
  });

  it("Worksheet.protect / unprotect, and the deprecated protect writes the same", async () => {
    for (const protect of [
      (ws: Stream.WorksheetWriter) => Worksheet.protect(ws, { password: "pw", formatCells: true }),
      (ws: Stream.WorksheetWriter) => ws.protect({ password: "pw", formatCells: true })
    ]) {
      const written = await roundTrip(async (_wb, ws) => {
        await protect(ws);
      });
      expect(await Worksheet.verifyPassword(sheetOf(written), "pw")).toBe(true);
      expect(sheetOf(written).sheetProtection).toMatchObject({ sheet: true });
    }

    const unprotected = await roundTrip(async (_wb, ws) => {
      await Worksheet.protect(ws, { password: "pw" });
      Worksheet.unprotect(ws);
    });
    expect(sheetOf(unprotected).sheetProtection).toBeFalsy();

    const ws = new Stream.WorkbookWriter({}).addWorksheet("S");
    ws.commit();
    await expect(Worksheet.protect(ws)).rejects.toThrow(ExcelStreamStateError);
    expect(() => Worksheet.unprotect(ws)).toThrow(ExcelStreamStateError);
  });

  it("Image.add refuses a committed workbook writer and names Image.add for a bad image", async () => {
    const wb = new Stream.WorkbookWriter({});
    expect(() => Image.add(wb, null as never)).toThrow(/Image\.add/);
    wb.addWorksheet("S");
    await wb.toBuffer();
    expect(() => Image.add(wb, { buffer: PNG, extension: "png" })).toThrow(ExcelStreamStateError);
  });
});
