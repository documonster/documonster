/**
 * Drawing objects this library does not model — text boxes, connectors — survive a round trip verbatim.
 *
 * `content-type-element.xlsx` has three text boxes (`<xdr:sp txBox="1">`) holding a company's contact details beside a
 * picture. The anchor parsed them into nothing, so every write deleted them and kept only the picture.
 *
 * `.node` because it reads a fixture off disk.
 */
import { readFile } from "node:fs/promises";

import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip";
import { Image, Workbook } from "@excel";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

const FIXTURE = "src/modules/excel/__tests__/data/content-type-element.xlsx";
const decoder = new TextDecoder();

async function drawingXml(
  bytes: Uint8Array
): Promise<{ xml: string; rels: string; paths: string[] }> {
  const entries = await extractAll(bytes);
  const xml = decoder.decode(entries.get("xl/drawings/drawing1.xml")!.data);
  const rels = decoder.decode(entries.get("xl/drawings/_rels/drawing1.xml.rels")!.data);
  return { xml, rels, paths: [...entries.keys()] };
}

async function roundTrip(bytes: Uint8Array, format: "xlsx" | "xlsb" = "xlsx"): Promise<Uint8Array> {
  const wb = await Workbook.read(Workbook.create(), bytes);
  return Workbook.toBuffer(wb, { format, unsupported: "ignore" });
}

describe("unmodelled drawing shapes", () => {
  it("keeps the text boxes beside the picture", async () => {
    const source = new Uint8Array(await readFile(FIXTURE));
    const out = await roundTrip(source);
    const { xml } = await drawingXml(out);
    expect(xml.match(/<xdr:sp>|<xdr:sp /g)).toHaveLength(3);
    expect(xml).toContain('name="Text 2"');
    expect(xml).toContain('name="Text Box 3"');
    expect(xml).toContain('name="Text Box 4"');
    expect(xml).toContain("EVERETT (INDIA) PVT LTD");
    expect(xml).toContain("E-mail: operation@everett.co.in");
    expect(xml).toContain("<xdr:pic>");
    expect((await validateXlsxBuffer(out, { maxProblems: 50 })).problems).toEqual([]);
    // And the second write reproduces the first.
    expect((await drawingXml(await roundTrip(out))).xml).toBe(xml);
  });

  it("points a preserved shape's picture fill at the medium it named, under the name it is written as", async () => {
    const source = await extractAll(new Uint8Array(await readFile(FIXTURE)));
    const archive = new ZipArchive();
    for (const [name, entry] of source) {
      let data = entry.data;
      if (name === "xl/drawings/drawing1.xml") {
        // A text box filled with the PNG the picture also uses — the medium `image1.png` is written as `image2.png`,
        // because `image1.emf` sorts first.
        data = new TextEncoder().encode(
          decoder
            .decode(data)
            .replace(
              '<a:prstGeom prst="rect" /><a:noFill />',
              '<a:prstGeom prst="rect" /><a:blipFill><a:blip r:embed="rId1" /><a:stretch><a:fillRect /></a:stretch></a:blipFill>'
            )
        );
      }
      archive.add(name, data);
    }
    const out = await roundTrip(await archive.toBuffer());
    const { xml, rels, paths } = await drawingXml(out);
    const embed = /<a:blipFill><a:blip r:embed="(rId\d+)"/.exec(xml)?.[1];
    expect(embed).toBeDefined();
    const target = new RegExp(`Id="${embed}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
    expect(target).toMatch(/\.png$/);
    expect(paths).toContain(`xl/media/${target!.split("/").pop()}`);
  });

  it("keeps them through a conversion to XLSB", async () => {
    const source = new Uint8Array(await readFile(FIXTURE));
    const out = await roundTrip(source, "xlsb");
    const entries = await extractAll(out);
    const drawings = [...entries.keys()].filter(path =>
      /^xl\/drawings\/drawing\d+\.xml$/.test(path)
    );
    const xml = drawings.map(path => decoder.decode(entries.get(path)!.data)).join("");
    expect(xml).toContain('name="Text Box 3"');
    expect(xml).toContain("EVERETT (INDIA) PVT LTD");
    const back = await drawingXml(await roundTrip(out));
    expect(back.xml).toContain("EVERETT (INDIA) PVT LTD");
  });

  it("gives every object of the drawing its own id, and keeps a connector on the shape it joins", async () => {
    const source = await extractAll(new Uint8Array(await readFile(FIXTURE)));
    const archive = new ZipArchive();
    for (const [name, entry] of source) {
      let data = entry.data;
      if (name === "xl/drawings/drawing1.xml") {
        // The text boxes take ids 4 and 5 — what the pictures written after them are numbered — and a group holding a
        // second shape `4` and a connector from it ends the drawing.
        const group =
          "<xdr:twoCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
          "<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" +
          '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="40" name="Group"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>' +
          '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="4" name="Inner"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>' +
          '<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="41" name="Link"/><xdr:cNvCxnSpPr><a:stCxn id="4" idx="0"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr/></xdr:cxnSp>' +
          "</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>";
        data = new TextEncoder().encode(
          decoder
            .decode(data)
            .replace('id="1028"', 'id="4"')
            .replace('id="1029"', 'id="5"')
            .replace("</xdr:wsDr>", `${group}</xdr:wsDr>`)
        );
      }
      archive.add(name, data);
    }
    const wb = await Workbook.read(Workbook.create(), await archive.toBuffer());
    const ws = Workbook.getWorksheets(wb)[0]!;
    const png = source.get("xl/media/image1.png")!.data;
    Image.place(ws, Image.add(wb, { buffer: png, extension: "png" }), {
      tl: { col: 12, row: 0 },
      br: { col: 14, row: 4 }
    });
    const out = await Workbook.toBuffer(wb);
    const { xml } = await drawingXml(out);

    const ids = [...xml.matchAll(/<xdr:cNvPr [^>]*\bid="(\d+)"/g)].map(match => match[1]);
    expect(ids).toHaveLength(8);
    expect(new Set(ids).size).toBe(ids.length);
    // The text boxes keep the ids they were read with; the pictures are numbered around them.
    expect(xml).toMatch(/<xdr:cNvPr id="4" name="Text 2"/);
    expect(xml).toMatch(/<xdr:cNvPr id="5" name="Text Box 3"/);
    // The group's `Inner` collided and was renumbered, and its connector follows it.
    const inner = /<xdr:cNvPr id="(\d+)" name="Inner"/.exec(xml)![1];
    expect(inner).not.toBe("4");
    expect(xml).toContain(`<a:stCxn id="${inner}" idx="0"/>`);
    expect((await validateXlsxBuffer(out, { maxProblems: 50 })).problems).toEqual([]);
  });
});
