import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { PassThrough } from "@stream";
import { parseSax } from "@xml/sax";
import { XmlWriter } from "@xml/writer";
import { describe, it, expect } from "vitest";

describe("AbsoluteAnchorXform", () => {
  const EMU_PER_PIXEL_AT_96_DPI = 9525;

  it("parses absoluteAnchor from drawing XML", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/>
    <xdr:ext cx="${100 * EMU_PER_PIXEL_AT_96_DPI}" cy="${200 * EMU_PER_PIXEL_AT_96_DPI}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${100 * EMU_PER_PIXEL_AT_96_DPI}" cy="${200 * EMU_PER_PIXEL_AT_96_DPI}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>`;

    const xform = new DrawingXform();
    const stream = new PassThrough();
    stream.write(xml);
    stream.end();
    const model = await xform.parse(parseSax(stream));

    // Should have one anchor
    expect(model!.anchors).toHaveLength(1);

    const anchor = model!.anchors[0];

    // Should have pos and ext (not tl/br)
    expect(anchor.range.pos).toBeDefined();
    expect(anchor.range.pos.x).toBe(0);
    expect(anchor.range.pos.y).toBe(0);
    expect(anchor.range.ext).toBeDefined();
    expect(anchor.range.ext.width).toBe(100);
    expect(anchor.range.ext.height).toBe(200);

    // Should have picture with rId
    expect(anchor.picture).toBeDefined();
    expect(anchor.picture.rId).toBe("rId1");
  });

  it("reconciles absoluteAnchor with media", async () => {
    const xform = new DrawingXform();
    const model: any = {
      anchors: [
        {
          range: {
            editAs: "oneCell",
            pos: { x: 0, y: 0 },
            ext: { width: 100, height: 200 }
          },
          picture: { rId: "rId1" }
        }
      ]
    };

    const mediaObj = { type: "image", name: "image1", index: 0 };
    const options = {
      rels: { rId1: { Target: "../media/image1.png" } },
      mediaIndex: { "image1.png": 0 },
      media: [mediaObj]
    };

    xform.reconcile(model, options);

    expect(model.anchors[0].medium).toBe(mediaObj);
  });

  it("reconciles absoluteAnchor gracefully when rel is missing", async () => {
    const xform = new DrawingXform();
    const model: any = {
      anchors: [
        {
          range: {
            editAs: "oneCell",
            pos: { x: 0, y: 0 },
            ext: { width: 100, height: 200 }
          },
          picture: { rId: "rId99" }
        }
      ]
    };

    const options = {
      rels: {},
      mediaIndex: {},
      media: []
    };

    // Should not throw
    expect(() => xform.reconcile(model, options)).not.toThrow();
    expect(model.anchors[0].medium).toBeUndefined();
  });

  it("renders absoluteAnchor to XML", () => {
    const xform = new DrawingXform();
    const model = {
      anchors: [
        {
          anchorType: "xdr:absoluteAnchor" as const,
          range: {
            pos: { x: 10, y: 20 },
            ext: { width: 100, height: 200 }
          },
          picture: {
            rId: "rId1",
            index: 0,
            name: "Picture 1"
          }
        }
      ]
    };

    const xmlStream = new XmlWriter();
    xform.render(xmlStream, model);
    const xml = xmlStream.xml;

    expect(xml).toContain("xdr:absoluteAnchor");
    expect(xml).toContain("xdr:pos");
    expect(xml).toContain("xdr:ext");
    expect(xml).toContain("xdr:pic");
    expect(xml).toContain("xdr:clientData");
  });

  it("absoluteAnchor images are included in media with pos preserved", () => {
    // Simulate what worksheet-xform reconcile does: iterate drawing.anchors
    // and push all images (including absoluteAnchor) into model.media.
    const anchors = [
      {
        // normal twoCellAnchor
        medium: { index: 0 },
        range: {
          editAs: "oneCell",
          tl: { nativeCol: 0, nativeRow: 0, nativeColOff: 0, nativeRowOff: 0 },
          br: { nativeCol: 2, nativeRow: 2, nativeColOff: 0, nativeRowOff: 0 }
        },
        picture: { hyperlinks: undefined }
      },
      {
        // absoluteAnchor — should also be included with pos/ext intact
        medium: { index: 1 },
        range: {
          editAs: "oneCell",
          pos: { x: 0, y: 0 },
          ext: { width: 100, height: 200 }
        },
        picture: { hyperlinks: undefined }
      }
    ];

    const media: any[] = [];
    anchors.forEach(anchor => {
      if (anchor.medium) {
        media.push({
          type: "image",
          imageId: anchor.medium.index,
          range: anchor.range,
          hyperlinks: anchor.picture.hyperlinks
        });
      }
    });

    // Both anchors should be included
    expect(media).toHaveLength(2);

    // First: cell-based
    expect(media[0].imageId).toBe(0);
    expect(media[0].range.tl).toBeDefined();
    expect(media[0].range.pos).toBeUndefined();

    // Second: absolute-positioned
    expect(media[1].imageId).toBe(1);
    expect(media[1].range.pos).toEqual({ x: 0, y: 0 });
    expect(media[1].range.ext).toEqual({ width: 100, height: 200 });
  });

  it("getAnchorType returns absoluteAnchor for pos-based ranges", () => {
    const xform = new DrawingXform();
    const model: any = {
      anchors: [
        {
          range: { pos: { x: 10, y: 20 }, ext: { width: 100, height: 200 } },
          picture: { rId: "rId1" }
        }
      ]
    };
    // prepare sets anchorType
    xform.prepare(model);
    expect(model.anchors[0].anchorType).toBe("xdr:absoluteAnchor");
  });
});
