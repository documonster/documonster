/**
 * VML relationship parts are routed by the sheet's `legacyDrawing` / `legacyDrawingHF`
 * relationship, never by the part name: Excel names a header/footer VML `vmlDrawingN.vml`
 * exactly like a comment's.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Cell, HeaderFooterImage, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** A package with a note VML that has rels of its own, and a header VML named as Excel names it. */
async function excelShapedPackage(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setNote(ws, "A1", "a note");
  const logo = Image.add(wb, { base64: PNG, extension: "png" });
  HeaderFooterImage.set(ws, { imageId: logo, position: "CH" });
  const files = new Map<string, Uint8Array>();
  for (const [name, entry] of await extractAll(await Workbook.toBuffer(wb))) {
    files.set(name, entry.data);
  }
  const rename = (from: string, to: string) => {
    files.set(to, files.get(from)!);
    files.delete(from);
  };
  rename("xl/drawings/vmlDrawingHF1.vml", "xl/drawings/vmlDrawing2.vml");
  rename("xl/drawings/_rels/vmlDrawingHF1.vml.rels", "xl/drawings/_rels/vmlDrawing2.vml.rels");
  const sheetRels = "xl/worksheets/_rels/sheet1.xml.rels";
  files.set(
    sheetRels,
    encoder.encode(decoder.decode(files.get(sheetRels)!).replace("vmlDrawingHF1", "vmlDrawing2"))
  );
  // The note VML's own relationships, pointing at the same medium.
  files.set(
    "xl/drawings/_rels/vmlDrawing1.vml.rels",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`
    )
  );
  return createZip([...files].map(([name, data]) => ({ name, data })));
}

describe("VML relationships are routed by the sheet's relationship, not the part name", () => {
  it("keeps the header pictures of an Excel-named header VML beside a note VML with rels", async () => {
    const input = await excelShapedPackage();
    const wb = Workbook.create();
    await Workbook.read(wb, input);
    const ws = Workbook.getWorksheet(wb, "S")!;
    expect(HeaderFooterImage.list(ws)).toEqual([expect.objectContaining({ position: "CH" })]);
    expect(Cell.getNote(ws, "A1")).toBeDefined();

    const output = await Workbook.toBuffer(wb);
    await expectValidXlsx(output, { label: "note VML rels + Excel-named header VML" });
    const back = Workbook.create();
    await Workbook.read(back, output);
    const again = Workbook.getWorksheet(back, "S")!;
    expect(HeaderFooterImage.list(again)).toEqual([expect.objectContaining({ position: "CH" })]);
    expect(Cell.getNote(again, "A1")).toBeDefined();
  });
});
