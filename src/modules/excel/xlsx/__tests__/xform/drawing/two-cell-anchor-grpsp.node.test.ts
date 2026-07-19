import { TwoCellAnchorXform } from "@excel/xlsx/xform/drawing/two-cell-anchor-xform";
import { XmlWriter } from "@xml/writer";
import { describe, expect, it } from "vitest";

/**
 * Feed a chunk of XML to a fresh xform via its streaming parser and return the
 * resulting model.
 */
async function parseXml<T>(makeXform: () => { parseStream: (s: any) => Promise<T> }, xml: string) {
  async function* one() {
    yield xml;
  }
  return makeXform().parseStream(one());
}

function render(xform: { render: (s: XmlWriter, m: any) => void }, model: any): string {
  const w = new XmlWriter();
  xform.render(w, model);
  return w.xml;
}

describe("TwoCellAnchorXform — grouped shapes and picture geometry", () => {
  const PIC_ANCHOR = `
    <xdr:twoCellAnchor editAs="oneCell"
      xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
      <xdr:pic>
        <xdr:nvPicPr>
          <xdr:cNvPr id="2" name="Picture 1"/>
          <xdr:cNvPicPr/>
        </xdr:nvPicPr>
        <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
        <xdr:spPr>
          <a:xfrm><a:off x="123456" y="654321"/><a:ext cx="1111111" cy="2222222"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    </xdr:twoCellAnchor>`;

  const GROUP_ANCHOR = `
    <xdr:twoCellAnchor
      xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
      <xdr:grpSp>
        <xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Group 9"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
        <xdr:grpSpPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/><a:chOff x="0" y="0"/><a:chExt cx="3" cy="4"/></a:xfrm></xdr:grpSpPr>
        <xdr:pic>
          <xdr:nvPicPr><xdr:cNvPr id="11" name="Inner Pic"><a:hlinkClick r:id="rId7"/></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr>
          <xdr:blipFill><a:blip r:embed="rId5"/></xdr:blipFill>
          <xdr:spPr/>
        </xdr:pic>
      </xdr:grpSp>
      <xdr:clientData/>
    </xdr:twoCellAnchor>`;

  it("preserves a picture's absolute <a:xfrm> geometry through parse", async () => {
    const model: any = await parseXml(() => new TwoCellAnchorXform(), PIC_ANCHOR);
    expect(model.picture).toBeDefined();
    expect(model.group).toBeUndefined();
    expect(model.picture.xfrmOffX).toBe(123456);
    expect(model.picture.xfrmOffY).toBe(654321);
    expect(model.picture.xfrmExtCx).toBe(1111111);
    expect(model.picture.xfrmExtCy).toBe(2222222);
  });

  it("re-emits the picture's real geometry on render (not a zero template)", async () => {
    const model: any = await parseXml(() => new TwoCellAnchorXform(), PIC_ANCHOR);
    const xml = render(new TwoCellAnchorXform(), model);
    expect(xml).toContain('<a:off x="123456" y="654321"/>');
    expect(xml).toContain('<a:ext cx="1111111" cy="2222222"/>');
    expect(xml).not.toContain('<a:off x="0" y="0"/>');
  });

  it("captures a <xdr:grpSp> group instead of misreading its nested pic", async () => {
    const model: any = await parseXml(() => new TwoCellAnchorXform(), GROUP_ANCHOR);
    // The nested <xdr:pic> must NOT be promoted to the anchor's own picture.
    expect(model.picture).toBeUndefined();
    expect(model.group).toBeDefined();
    expect(model.group.tag).toBe("xdr:grpSp");
  });

  it("round-trips a group verbatim, keeping its nested pic, rels and hyperlink", async () => {
    const model: any = await parseXml(() => new TwoCellAnchorXform(), GROUP_ANCHOR);
    const xml = render(new TwoCellAnchorXform(), model);
    expect(xml).toContain("<xdr:grpSp>");
    expect(xml).toContain('r:embed="rId5"');
    expect(xml).toContain('r:id="rId7"');
    expect(xml).toContain('name="Inner Pic"');
  });

  it("preserves text order around nested group elements", async () => {
    const xml = GROUP_ANCHOR.replace("<xdr:nvGrpSpPr>", "before<xdr:nvGrpSpPr>").replace(
      "</xdr:nvGrpSpPr>",
      "</xdr:nvGrpSpPr>after"
    );
    const model: any = await parseXml(() => new TwoCellAnchorXform(), xml);
    const rendered = render(new TwoCellAnchorXform(), model);
    expect(rendered.indexOf("before")).toBeLessThan(rendered.indexOf("<xdr:nvGrpSpPr>"));
    expect(rendered.indexOf("after")).toBeGreaterThan(rendered.indexOf("</xdr:nvGrpSpPr>"));
  });

  it("does not leak a prior anchor's picture into a following group-only anchor", async () => {
    // Reuse a SINGLE xform instance across two anchors (as the drawing parser
    // does) to prove child-model reset prevents cross-anchor bleed.
    const xform = new TwoCellAnchorXform();
    const picModel: any = await xform.parseStream(
      (async function* () {
        yield PIC_ANCHOR;
      })()
    );
    expect(picModel.picture).toBeDefined();

    const groupModel: any = await xform.parseStream(
      (async function* () {
        yield GROUP_ANCHOR;
      })()
    );
    expect(groupModel.group).toBeDefined();
    expect(groupModel.picture).toBeUndefined();
  });
});
