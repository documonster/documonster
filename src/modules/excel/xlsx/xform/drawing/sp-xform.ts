import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { EMU_PER_PX } from "@utils/units";

export interface ShapeModel {
  /** Drawing object id (must be unique within drawing part) */
  cNvPrId: number;
  /** Display name (e.g. "Check Box 1") */
  name: string;
  /** VML spid compat value (e.g. "_x0000_s1025") */
  spid: string;
  /** Text shown in the shape text box */
  text: string;
  /** Whether the shape should be hidden */
  hidden?: boolean;
}

class SpXform extends BaseXform {
  get tag(): string {
    return "xdr:sp";
  }

  render(xmlStream: any, model?: ShapeModel): void {
    if (!model) {
      return;
    }

    xmlStream.openNode(this.tag, { macro: "", textlink: "" });

    // Non-visual properties
    xmlStream.openNode("xdr:nvSpPr");
    xmlStream.openNode("xdr:cNvPr", {
      id: model.cNvPrId,
      name: model.name,
      hidden: model.hidden ? "1" : undefined
    });
    xmlStream.openNode("a:extLst");
    xmlStream.openNode("a:ext", {
      uri: "{63B3BB69-23CF-44E3-9099-C40C66FF867C}"
    });
    xmlStream.leafNode(
      "a14:compatExt",
      {
        spid: model.spid,
        "xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main"
      },
      undefined
    );
    xmlStream.closeNode(); // a:ext
    xmlStream.closeNode(); // a:extLst
    xmlStream.closeNode(); // xdr:cNvPr
    xmlStream.leafNode("xdr:cNvSpPr");
    xmlStream.closeNode(); // xdr:nvSpPr

    // Shape properties
    xmlStream.openNode("xdr:spPr", { bwMode: "auto" });
    xmlStream.openNode("a:xfrm");
    xmlStream.leafNode("a:off", { x: 0, y: 0 });
    xmlStream.leafNode("a:ext", { cx: 0, cy: 0 });
    xmlStream.closeNode(); // a:xfrm
    xmlStream.openNode("a:prstGeom", { prst: "rect" });
    xmlStream.leafNode("a:avLst");
    xmlStream.closeNode(); // a:prstGeom
    xmlStream.leafNode("a:noFill");
    xmlStream.openNode("a:ln");
    xmlStream.leafNode("a:noFill");
    xmlStream.closeNode(); // a:ln
    xmlStream.openNode("a:extLst");
    xmlStream.openNode("a:ext", { uri: "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" });
    xmlStream.openNode("a14:hiddenFill", {
      "xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main"
    });
    xmlStream.openNode("a:solidFill");
    xmlStream.leafNode("a:srgbClr", { val: "F0F0F0" });
    xmlStream.closeNode(); // a:solidFill
    xmlStream.closeNode(); // a14:hiddenFill
    xmlStream.closeNode(); // a:ext
    xmlStream.openNode("a:ext", { uri: "{91240B29-F687-4F45-9708-019B960494DF}" });
    xmlStream.openNode("a14:hiddenLine", {
      w: EMU_PER_PX,
      "xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main"
    });
    xmlStream.openNode("a:solidFill");
    xmlStream.leafNode("a:srgbClr", { val: "000000" });
    xmlStream.closeNode(); // a:solidFill
    xmlStream.leafNode("a:miter", { lim: 800000 });
    xmlStream.leafNode("a:headEnd");
    xmlStream.leafNode("a:tailEnd");
    xmlStream.closeNode(); // a14:hiddenLine
    xmlStream.closeNode(); // a:ext
    xmlStream.closeNode(); // a:extLst
    xmlStream.closeNode(); // xdr:spPr

    // Text body (label)
    xmlStream.openNode("xdr:txBody");
    xmlStream.leafNode("a:bodyPr", {
      vertOverflow: "clip",
      wrap: "square",
      lIns: 18288,
      tIns: 18288,
      rIns: 0,
      bIns: 18288,
      anchor: "ctr",
      upright: 1
    });
    xmlStream.leafNode("a:lstStyle");
    xmlStream.openNode("a:p");
    xmlStream.openNode("a:pPr", { algn: "l", rtl: 0 });
    xmlStream.leafNode("a:defRPr", { sz: 1000 });
    xmlStream.closeNode(); // a:pPr
    xmlStream.openNode("a:r");
    xmlStream.openNode("a:rPr", {
      lang: "en-US",
      sz: 800,
      b: 0,
      i: 0,
      u: "none",
      strike: "noStrike",
      baseline: 0
    });
    xmlStream.openNode("a:solidFill");
    xmlStream.leafNode("a:srgbClr", { val: "000000" });
    xmlStream.closeNode(); // a:solidFill
    xmlStream.leafNode("a:latin", { typeface: "Tahoma", pitchFamily: 2, charset: 0 });
    xmlStream.leafNode("a:ea", { typeface: "Tahoma", pitchFamily: 2, charset: 0 });
    xmlStream.leafNode("a:cs", { typeface: "Tahoma", pitchFamily: 2, charset: 0 });
    xmlStream.closeNode(); // a:rPr
    xmlStream.openNode("a:t");
    xmlStream.writeText(model.text ?? "");
    xmlStream.closeNode(); // a:t
    xmlStream.closeNode(); // a:r
    xmlStream.closeNode(); // a:p
    xmlStream.closeNode(); // xdr:txBody

    xmlStream.closeNode(); // xdr:sp
  }
}

export { SpXform };
