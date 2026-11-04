import { VmlDrawingXform } from "@excel/xlsx/xform/drawing/vml-drawing-xform";
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { XmlWriter } from "@xml/writer";
import { describe, expect, it } from "vitest";

async function parse(xml: string) {
  async function* input() {
    yield xml;
  }
  return parseXformStream(new VmlDrawingXform(), input());
}

describe("VmlDrawingXform header/footer images", () => {
  it("parses every positioned header/footer image", async () => {
    const model = await parse(`
      <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
        <v:shape id="LH" type="#_x0000_t75" style="width:20pt;height:10pt">
          <v:imagedata o:relid="rId1"/>
        </v:shape>
        <v:shape id="RF" type="#_x0000_t75" style="width:30pt;height:15pt">
          <v:imagedata o:relid="rId2"/>
        </v:shape>
      </xml>
    `);

    // Each also carries its `<v:shape>` as read, so a rewrite keeps what it says beyond these fields.
    expect(model?.headerImages).toEqual([
      {
        imageRelId: "rId1",
        width: 20,
        height: 10,
        position: "LH",
        shape: expect.objectContaining({
          tag: "v:shape",
          attrs: expect.objectContaining({ id: "LH" })
        })
      },
      {
        imageRelId: "rId2",
        width: 30,
        height: 15,
        position: "RF",
        shape: expect.objectContaining({
          tag: "v:shape",
          attrs: expect.objectContaining({ id: "RF" })
        })
      }
    ]);
  });

  it("renders every positioned header/footer image", () => {
    const writer = new XmlWriter();
    new VmlDrawingXform().render(writer, {
      headerImages: [
        { imageRelId: "rId1", width: 20, height: 10, position: "LH" },
        { imageRelId: "rId2", width: 30, height: 15, position: "RF" }
      ]
    });
    const xml = writer.toString();

    expect(xml).toContain('id="LH"');
    expect(xml).toContain('o:relid="rId1"');
    expect(xml).toContain('id="RF"');
    expect(xml).toContain('o:relid="rId2"');
    // Rendered on its own, the part has block 1 to itself.
    expect(xml).toContain('o:spid="_x0000_s1025"');
    expect(xml).toContain('o:spid="_x0000_s1026"');
  });
});

describe("VmlDrawingXform shape ids", () => {
  const shapeIds = (xml: string) =>
    [...xml.matchAll(/<v:shape\b[^>]*?\b(?:o:spid|id)="_x0000_s(\d+)"/g)].map(m => Number(m[1]));
  const idmap = (xml: string) => /<o:idmap [^>]*data="([^"]*)"/.exec(xml)?.[1];
  const note = {
    note: { texts: [{ text: "n" }], protection: {}, margins: {} },
    refAddress: { row: 1, col: 1 }
  };
  const checkbox = (shapeId: number) => ({
    shapeId,
    ctrlPropId: 1,
    tl: { col: 1, colOff: 0, row: 1, rowOff: 0 },
    br: { col: 2, colOff: 0, row: 2, rowOff: 0 },
    checked: "Unchecked",
    text: "c",
    noThreeD: true,
    print: true
  });

  function render(model: object): string {
    const writer = new XmlWriter();
    new VmlDrawingXform().render(writer, model as never);
    return writer.toString();
  }

  it("keeps checkbox ids and gives notes the free ids of the same blocks", () => {
    const xml = render({ comments: [note, note], formControls: [checkbox(1025), checkbox(2050)] });
    const ids = shapeIds(xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.sort((a, b) => a - b)).toEqual([1025, 1026, 1027, 2050]);
    expect(idmap(xml)).toBe("1,2");
  });

  it("declares the blocks notes past the first 1023 ids reach", () => {
    const xml = render({ comments: Array.from({ length: 1100 }, () => note) });
    const ids = shapeIds(xml);
    // A block holds 1023 shapes: `_x0000_s2048` would be the next block's 0th slot.
    expect(ids).not.toContain(2048);
    expect(Math.max(...ids)).toBe(2049 + 1100 - 1023 - 1);
    expect(idmap(xml)).toBe("1,2");
  });

  it("numbers a header/footer part on its own from block 1", () => {
    const xml = render({ headerImages: [{ imageRelId: "rId1", position: "CH" }] });
    expect(idmap(xml)).toBe("1");
  });

  it("uses the ids the workbook-wide allocation gave the part", () => {
    const xml = render({
      comments: [note],
      formControls: [checkbox(3073)],
      shapeIds: [3074]
    });
    expect(shapeIds(xml).sort((a, b) => a - b)).toEqual([3073, 3074]);
    expect(idmap(xml)).toBe("3");
    const header = render({
      headerImages: [{ imageRelId: "rId1", position: "CH" }],
      shapeIds: [4097]
    });
    expect(shapeIds(header)).toEqual([4097]);
    expect(idmap(header)).toBe("4");
  });

  it("keeps block 1 for an ordinary note part", () => {
    expect(idmap(render({ comments: [note] }))).toBe("1");
  });
});
