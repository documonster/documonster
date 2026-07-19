import { ZipArchive } from "@archive/zip";
import { Image, Workbook } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, it, expect } from "vitest";

describe("picture xfrm geometry round-trip (real non-zero geometry)", () => {
  it("preserves a non-zero <a:xfrm> through read->write (does not zero it out)", async () => {
    // 1. Build a workbook with an image and serialize.
    const src = Workbook.create();
    const ws = Workbook.addWorksheet(src, "Sheet1");
    const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
    const imageId = Image.add(src, { filename: dataPath("image.png"), extension: "png" });
    Image.place(ws, imageId, "B2:D6");
    const buf = await Workbook.toBuffer(src);

    // 2. Rewrite the drawing xml so the picture carries a REAL non-zero xfrm,
    //    simulating an externally-authored editAs="oneCell" picture.
    const { extractAll } = await import("@archive/unzip/extract");
    const entries = await extractAll(buf);
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const drawingKey = [...entries.keys()].find(k => /xl\/drawings\/drawing\d+\.xml$/.test(k))!;
    let xml = decoder.decode(entries.get(drawingKey)!.data);
    xml = xml
      .replace(/<a:off x="0" y="0"\/>/g, '<a:off x="914400" y="457200"/>')
      .replace(/<a:ext cx="0" cy="0"\/>/g, '<a:ext cx="2743200" cy="1828800"/>')
      .replace("<a:xfrm>", '<a:xfrm rot="5400000" flipH="1">')
      .replace("</xdr:spPr>", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr>');
    entries.get(drawingKey)!.data = encoder.encode(xml);

    const archive = new ZipArchive({ level: 0, reproducible: true });
    for (const [name, entry] of entries) {
      archive.add(name, entry.data);
    }
    const tampered = await archive.bytes();

    // 3. Read it back and write it out again.
    const wb = Workbook.create();
    await Workbook.read(wb, tampered);
    const out = await Workbook.toBuffer(wb);

    // 4. The re-emitted drawing must still carry the non-zero geometry.
    const outEntries = await extractAll(out);
    const outKey = [...outEntries.keys()].find(k => /xl\/drawings\/drawing\d+\.xml$/.test(k))!;
    const outXml = decoder.decode(outEntries.get(outKey)!.data);
    expect(outXml).toContain('<a:off x="914400" y="457200"/>');
    expect(outXml).toContain('<a:ext cx="2743200" cy="1828800"/>');
    expect(outXml).toContain('<a:xfrm rot="5400000" flipH="1">');
    expect(outXml).toContain('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>');
    expect(outXml).not.toContain('<a:off x="0" y="0"/>');
  });
});
