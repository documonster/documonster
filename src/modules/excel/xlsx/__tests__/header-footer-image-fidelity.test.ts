/**
 * A header/footer picture keeps how Excel was told to draw it.
 *
 * Excel's Format Picture dialog writes its settings onto the VML shape: `gain` and `blacklevel` (the
 * "Washout" colour, or brightness and contrast), `grayscale` and `bilevel`, `cropleft`/`croptop`/
 * `cropright`/`cropbottom`, and the picture's own name in `o:title`. A rewrite regenerated the shape
 * from the position and size alone, so every one of them was lost and the title became "watermark".
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsb } from "@excel/__tests__/helpers/expect-valid-xlsb";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { HeaderFooterImage, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const encoder = new TextEncoder();
const decoder = new TextDecoder();

const IMAGE_DATA_ATTRIBUTES =
  'o:title="CompanyLogo" gain="19661f" blacklevel="22938f" grayscale="t" bilevel="t" cropleft="1000f" croptop="2000f" cropright="3000f" cropbottom="4000f"';

/** A sheet whose centre-header picture Excel has washed out, cropped and named. */
async function excelShapedPackage(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  HeaderFooterImage.set(ws, {
    imageId: Image.add(wb, { base64: PNG, extension: "png" }),
    position: "CH",
    width: 200,
    height: 100
  });
  const files = new Map<string, Uint8Array>();
  for (const [name, entry] of await extractAll(await Workbook.toBuffer(wb))) {
    files.set(name, entry.data);
  }
  const path = "xl/drawings/vmlDrawingHF1.vml";
  files.set(
    path,
    encoder.encode(
      decoder
        .decode(files.get(path)!)
        .replace('o:title="watermark"', IMAGE_DATA_ATTRIBUTES)
        .replace(/(<v:shape [^>]*?)>/, '$1 o:allowoverlap="f" alt="Logo">')
    )
  );
  return createZip([...files].map(([name, data]) => ({ name, data })));
}

async function headerShape(bytes: Uint8Array): Promise<string> {
  const entries = await extractAll(bytes);
  const name = [...entries.keys()].find(entry => /vmlDrawing[^/]*\.vml$/.test(entry))!;
  return /<v:shape [\s\S]*?<\/v:shape>/.exec(decoder.decode(entries.get(name)!.data))![0];
}

async function read(bytes: Uint8Array, format?: "xlsb") {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes, format ? { format } : undefined);
  return wb;
}

describe("header/footer picture fidelity", () => {
  it("keeps every attribute of the picture and its shape through an XLSX rewrite", async () => {
    const out = await Workbook.toBuffer(await read(await excelShapedPackage()));
    await expectValidXlsx(out);
    const shape = await headerShape(out);
    expect(shape).toContain(IMAGE_DATA_ATTRIBUTES);
    expect(shape).not.toContain("watermark");
    expect(shape).toContain('o:allowoverlap="f"');
    expect(shape).toContain('alt="Logo"');
    expect(shape).toContain("width:200pt;height:100pt");
    // And again: a second rewrite changes nothing.
    expect(await headerShape(await Workbook.toBuffer(await read(out)))).toBe(shape);
  });

  it("keeps them through an XLSB and back", async () => {
    const xlsb = await Workbook.toBuffer(await read(await excelShapedPackage()), {
      format: "xlsb"
    });
    await expectValidXlsb(xlsb);
    expect(await headerShape(xlsb)).toContain(IMAGE_DATA_ATTRIBUTES);
    const xlsx = await Workbook.toBuffer(await read(xlsb, "xlsb"));
    expect(await headerShape(xlsx)).toContain(IMAGE_DATA_ATTRIBUTES);
  });

  it("draws a replaced picture fresh, with none of the old one's settings", async () => {
    const wb = await read(await excelShapedPackage());
    const ws = Workbook.getWorksheet(wb, "S")!;
    HeaderFooterImage.set(ws, {
      imageId: Image.add(wb, { base64: PNG, extension: "png" }),
      position: "CH"
    });
    const shape = await headerShape(await Workbook.toBuffer(wb));
    expect(shape).not.toContain("gain=");
    expect(shape).toContain('o:title="watermark"');
  });

  it("reports the read picture through the public list unchanged", async () => {
    const ws = Workbook.getWorksheet(await read(await excelShapedPackage()), "S")!;
    expect(HeaderFooterImage.list(ws)).toEqual([
      { imageId: expect.any(String), position: "CH", width: 200, height: 100 }
    ]);
  });
});
