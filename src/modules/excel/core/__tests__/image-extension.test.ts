import { getImage } from "@excel/core/workbook-core";
import { Image, ExcelImageError, Stream, Workbook } from "@excel/index";
import type { ImageData } from "@excel/types";
import { describe, expect, it } from "vitest";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("image extension at the API boundary", () => {
  it("Image.add rejects a MIME type and names the accepted values", () => {
    const wb = Workbook.create();
    const call = (): number =>
      Image.add(wb, { buffer: PNG, extension: "image/png" as ImageData["extension"] });
    expect(call).toThrow(ExcelImageError);
    expect(call).toThrow(/png, jpeg, gif/);
    expect(getImage(wb, 0)).toBeUndefined();
  });

  it("the streaming WorkbookWriter rejects it too", () => {
    const wb = new Stream.WorkbookWriter({});
    expect(() =>
      wb.addImage({ buffer: PNG, extension: "image/png" as ImageData["extension"] })
    ).toThrow(ExcelImageError);
  });

  it("documented extensions and jpg still work", () => {
    const wb = Workbook.create();
    for (const extension of ["png", "jpeg", "gif", "jpg"]) {
      expect(() =>
        Image.add(wb, { buffer: PNG, extension: extension as ImageData["extension"] })
      ).not.toThrow();
    }
  });
});

describe("image extension at write time", () => {
  it("a hand-built model with a MIME-type extension is refused, not written", async () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    Image.add(wb, { buffer: PNG, extension: "png" });
    const model = Workbook.getModel(wb);
    (model.media[0] as { extension: string }).extension = "image/png";
    const copy = Workbook.create();
    Workbook.setModel(copy, model);
    await expect(Workbook.toBuffer(copy)).rejects.toThrow(ExcelImageError);
    await expect(Workbook.toBuffer(copy, { format: "xlsb" })).rejects.toThrow(ExcelImageError);
  });
});
