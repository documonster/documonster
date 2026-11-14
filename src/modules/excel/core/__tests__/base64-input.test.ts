import { getImage } from "@excel/core/workbook-core";
/**
 * Base64 text given to Excel is read the same way on every platform: text that is not base64 is
 * refused where it is given, as the call's own error. Node's decoder used to skip the characters
 * it did not know and the browser's to throw a `DOMException`, so the same call wrote a damaged
 * picture under Node and failed with a foreign error in a browser — at write time, far from the
 * call that took the text.
 */
import { ExcelArgumentError, ExcelImageError, Image, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

// A 1×1 PNG
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

describe("base64 given to Excel", () => {
  it("Image.add refuses text that is not base64, with or without a data: prefix", () => {
    const wb = Workbook.create();
    for (const base64 of ["iVBOR!!w0KGgo=", "data:image/png;base64,iVBOR!!w0KGgo="]) {
      expect(() => Image.add(wb, { extension: "png", base64 })).toThrow(ExcelImageError);
    }
    expect(() => Image.add(wb, { extension: "png", base64: PNG })).not.toThrow();
    expect(() =>
      Image.add(wb, { extension: "png", base64: `data:image/png;base64,${PNG}` })
    ).not.toThrow();
  });

  it("Image.add accepts the URL-safe alphabet, as it always did on Node", async () => {
    const wb = Workbook.create();
    // Bytes whose base64 needs `+` and `/`, written in `-` and `_`
    const bytes = Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 251, 255, 191, 254);
    const id = Image.add(wb, { extension: "png", base64: "iVBOR_v_v_4=" });
    Image.place(Workbook.addWorksheet(wb, "S"), id, "A1:B2");
    const read = await Workbook.read(await Workbook.toBuffer(wb));
    expect(getImage(read, 0)?.buffer).toEqual(bytes);
  });

  it("Workbook.read refuses a { base64: true } string that is not base64", async () => {
    await expect(Workbook.read("UEsDBA!!", { base64: true })).rejects.toThrow(ExcelArgumentError);
    await expect(Workbook.read("UEsDBA!!", { base64: true })).rejects.toThrow(/not base64/);
  });

  it("a stored protection hash that is not base64 matches no password", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    await Worksheet.protect(ws, { password: "secret" });
    expect(await Worksheet.verifyPassword(ws, "secret")).toBe(true);
    // As a damaged file would carry it
    ws.sheetProtection!.hashValue = "not base64!";
    expect(await Worksheet.verifyPassword(ws, "secret")).toBe(false);
  });
});
