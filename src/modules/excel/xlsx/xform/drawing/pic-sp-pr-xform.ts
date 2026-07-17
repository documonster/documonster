import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

/**
 * `<xdr:pic><xdr:spPr>` shape properties. For a `twoCellAnchor
 * editAs="oneCell"` picture, Excel takes the picture's absolute rendered
 * position/size from this `<a:xfrm>` - the `<xdr:to>` cell in the anchor is
 * just a cache. Losing these values doesn't just fail to preserve them, it
 * makes the picture render at (0,0) with zero size (invisible) or fall back
 * to whatever a group ancestor implies.
 */
export interface PicSpPrModel {
  xfrmOffX?: number;
  xfrmOffY?: number;
  xfrmExtCx?: number;
  xfrmExtCy?: number;
}

class PicSpPrXform extends BaseXform<PicSpPrModel> {
  get tag(): string {
    return "xdr:spPr";
  }

  render(xmlStream: XmlSink, model?: PicSpPrModel): void {
    xmlStream.openNode(this.tag);
    xmlStream.openNode("a:xfrm");
    xmlStream.leafNode("a:off", { x: model?.xfrmOffX ?? 0, y: model?.xfrmOffY ?? 0 });
    xmlStream.leafNode("a:ext", { cx: model?.xfrmExtCx ?? 0, cy: model?.xfrmExtCy ?? 0 });
    xmlStream.closeNode(); // a:xfrm
    xmlStream.openNode("a:prstGeom", { prst: "rect" });
    xmlStream.leafNode("a:avLst");
    xmlStream.closeNode(); // a:prstGeom
    xmlStream.closeNode(); // xdr:spPr
  }

  parseOpen(node: ParseOpenTag): boolean {
    switch (node.name) {
      case this.tag:
        this.reset();
        return true;
      case "a:off":
        this.mergeModel({
          xfrmOffX: parseInt(node.attributes.x, 10),
          xfrmOffY: parseInt(node.attributes.y, 10)
        });
        return true;
      case "a:ext":
        this.mergeModel({
          xfrmExtCx: parseInt(node.attributes.cx, 10),
          xfrmExtCy: parseInt(node.attributes.cy, 10)
        });
        return true;
      default:
        return true;
    }
  }

  parseText(): void {}

  parseClose(name: string): boolean {
    return name !== this.tag;
  }
}

export { PicSpPrXform };
