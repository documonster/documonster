import { ZipArchive } from "@archive/zip";
import { getImages } from "@excel/core/worksheet";
import { Image, Workbook } from "@excel/index";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { makeTestDataPath } from "@test/utils";
import { describe, expect, it } from "vitest";

const dec = new TextDecoder();
const enc = new TextEncoder();

/** Build an xlsx with one image, then hand back its unzipped entries. */
async function imageWorkbookEntries() {
  const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S1");
  Image.place(ws, Image.add(wb, { filename: dataPath("image.png"), extension: "png" }), "B2:D6");
  const { extractAll } = await import("@archive/unzip/extract");
  return extractAll(await Workbook.toBuffer(wb));
}

async function rezip(entries: Map<string, { data: Uint8Array }>): Promise<Uint8Array> {
  const archive = new ZipArchive({ level: 0, reproducible: true });
  for (const [name, entry] of entries) {
    archive.add(name, entry.data);
  }
  return archive.bytes();
}

function drawingKey(entries: Map<string, unknown>): string {
  return [...entries.keys()].find(k => /xl\/drawings\/drawing\d+\.xml$/.test(k))!;
}

describe("drawing round-trip: grouped shapes and picture geometry", () => {
  it("preserves an <xdr:grpSp> grouped-shape anchor through a full read->write", async () => {
    // filterDrawingAnchors used to drop any two-cell anchor without a
    // picture/shape/graphicFrame, silently deleting grouped shapes on write
    // even though the parser preserved them.
    const entries = await imageWorkbookEntries();
    const key = drawingKey(entries);
    const xml = dec
      .decode(entries.get(key)!.data)
      .replace(
        "</xdr:wsDr>",
        `<xdr:twoCellAnchor editAs="oneCell">` +
          `<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
          `<xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
          `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="99" name="G"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
          `<xdr:grpSpPr/></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
      );
    entries.get(key)!.data = enc.encode(xml);

    const wb = Workbook.create();
    await Workbook.read(wb, await rezip(entries));
    const out = await Workbook.toBuffer(wb);

    const { extractAll } = await import("@archive/unzip/extract");
    const outEntries = await extractAll(out);
    const outXml = dec.decode(outEntries.get(drawingKey(outEntries))!.data);
    expect(outXml).toContain("grpSp");
    expect(outXml).toContain('name="G"');
    // The image anchor must survive alongside the group.
    expect(outXml).toContain("<xdr:pic>");
    expect((await validateXlsxBuffer(out, { maxProblems: 50 })).problems).toEqual([]);
  });

  it("carries a picture's non-canonical spPr (rotation/flip/fill) across importSheet into another workbook", async () => {
    const entries = await imageWorkbookEntries();
    const key = drawingKey(entries);
    const xml = dec
      .decode(entries.get(key)!.data)
      .replace('<a:off x="0" y="0"/>', '<a:off x="914400" y="457200"/>')
      .replace('<a:ext cx="0" cy="0"/>', '<a:ext cx="2743200" cy="1828800"/>')
      .replace("<a:xfrm>", '<a:xfrm rot="5400000" flipH="1">')
      .replace("</xdr:spPr>", '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></xdr:spPr>');
    entries.get(key)!.data = enc.encode(xml);

    const source = Workbook.create();
    await Workbook.read(source, await rezip(entries));
    const srcSheet = Workbook.getWorksheet(source, "S1")!;

    const dest = Workbook.create();
    Workbook.importSheet(dest, srcSheet, "Copy");
    expect(getImages(Workbook.getWorksheet(dest, "Copy")!)).toHaveLength(1);

    const out = await Workbook.toBuffer(dest);
    const { extractAll } = await import("@archive/unzip/extract");
    const outEntries = await extractAll(out);
    const outXml = dec.decode(outEntries.get(drawingKey(outEntries))!.data);
    // The full spPr subtree (geometry + rotation/flip + fill) must survive the
    // cross-workbook copy, not just the four scalar geometry values.
    expect(outXml).toContain('rot="5400000"');
    expect(outXml).toContain('flipH="1"');
    expect(outXml).toContain('<a:off x="914400" y="457200"/>');
    expect(outXml).toContain('<a:srgbClr val="FF0000"/>');
    expect((await validateXlsxBuffer(out, { maxProblems: 50 })).problems).toEqual([]);
  });

  it("deduplicates one image referenced by two anchors when copied across workbooks", async () => {
    // The same imageId used by two anchors must register a single medium in the
    // destination workbook, not one per anchor.
    const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
    const src = Workbook.create();
    const ws = Workbook.addWorksheet(src, "S");
    const id = Image.add(src, { filename: dataPath("image.png"), extension: "png" });
    Image.place(ws, id, "A1:B2");
    Image.place(ws, id, "D1:E2");

    const reopened = Workbook.create();
    await Workbook.read(reopened, await Workbook.toBuffer(src));
    const dest = Workbook.create();
    Workbook.importSheet(dest, Workbook.getWorksheet(reopened, "S")!, "Copy");

    expect(getImages(Workbook.getWorksheet(dest, "Copy")!)).toHaveLength(2);
    expect(dest.media).toHaveLength(1);
    expect(
      (await validateXlsxBuffer(await Workbook.toBuffer(dest), { maxProblems: 50 })).problems
    ).toEqual([]);
  });

  it("keeps a grouped shape (with its own rel-referencing pic) alongside a top-level image", async () => {
    // Exercises the preservedRels + max-based nextRid path: the group's original
    // rIds must not collide with the freshly numbered image rel on rebuild.
    const entries = await imageWorkbookEntries();
    const key = drawingKey(entries);
    const xml = dec
      .decode(entries.get(key)!.data)
      .replace(
        "</xdr:wsDr>",
        `<xdr:twoCellAnchor editAs="oneCell">` +
          `<xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
          `<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
          `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="50" name="G"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>` +
          `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="51" name="InG"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
          `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></xdr:blipFill>` +
          `<xdr:spPr/></xdr:pic></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
      );
    entries.get(key)!.data = enc.encode(xml);

    const wb = Workbook.create();
    await Workbook.read(wb, await rezip(entries));
    const out = await Workbook.toBuffer(wb);
    const { extractAll } = await import("@archive/unzip/extract");
    const outEntries = await extractAll(out);
    const outXml = dec.decode(outEntries.get(drawingKey(outEntries))!.data);
    expect(outXml).toContain("grpSp");
    expect(outXml).toContain("<xdr:pic>");
    expect((await validateXlsxBuffer(out, { maxProblems: 50 })).problems).toEqual([]);
  });
});
