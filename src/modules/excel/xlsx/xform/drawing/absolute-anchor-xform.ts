import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { BaseCellAnchorXform } from "@excel/xlsx/xform/drawing/base-cell-anchor-xform";
import { ExtXform } from "@excel/xlsx/xform/drawing/ext-xform";
import { GraphicFrameXform } from "@excel/xlsx/xform/drawing/graphic-frame-xform";
import { PicXform } from "@excel/xlsx/xform/drawing/pic-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";

/** https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML */
const EMU_PER_PIXEL_AT_96_DPI = 9525;

interface PosModel {
  x: number;
  y: number;
}

/**
 * Xform for the <xdr:pos> element (absolute position in EMU).
 * Converts between EMU (in XML) and pixels (in model).
 */
class PosXform extends BaseXform<PosModel> {
  declare public map: { [key: string]: any };

  constructor() {
    super();
    this.map = {};
    this.model = { x: 0, y: 0 };
  }

  get tag(): string {
    return "xdr:pos";
  }

  render(xmlStream: any, model: PosModel): void {
    xmlStream.leafNode(this.tag, {
      x: Math.floor(model.x * EMU_PER_PIXEL_AT_96_DPI),
      y: Math.floor(model.y * EMU_PER_PIXEL_AT_96_DPI)
    });
  }

  parseOpen(node: any): boolean {
    if (node.name === this.tag) {
      this.model = {
        x: parseInt(node.attributes.x ?? "0", 10) / EMU_PER_PIXEL_AT_96_DPI,
        y: parseInt(node.attributes.y ?? "0", 10) / EMU_PER_PIXEL_AT_96_DPI
      };
      return true;
    }
    return false;
  }

  parseText(): void {}

  parseClose(): boolean {
    return false;
  }
}

/**
 * Xform for <xdr:absoluteAnchor> — images positioned by absolute coordinates
 * rather than cell references.
 *
 * Structure:
 * ```xml
 * <xdr:absoluteAnchor>
 *   <xdr:pos x="0" y="0"/>
 *   <xdr:ext cx="1000000" cy="1000000"/>
 *   <xdr:pic>...</xdr:pic>
 *   <xdr:clientData/>
 * </xdr:absoluteAnchor>
 * ```
 */
class AbsoluteAnchorXform extends BaseCellAnchorXform {
  constructor() {
    super();

    this.map = {
      "xdr:pos": new PosXform(),
      "xdr:ext": new ExtXform({ tag: "xdr:ext" }),
      "xdr:pic": new PicXform(),
      // `xdr:graphicFrame` carries the chart / embedded object payload
      // for absolute-anchor sheet drawings. The oneCell / twoCell
      // counterparts already handled it; the absolute branch was
      // pic-only, so a chart authored with an absolute anchor (or
      // programmatically constructed via `{ pos, ext }`) silently
      // dropped its graphicFrame on write — the drawing XML emitted
      // `<xdr:absoluteAnchor><xdr:pos/><xdr:ext/><xdr:clientData/></xdr:absoluteAnchor>`
      // with no chart reference, so the anchor was ignored on open.
      "xdr:graphicFrame": new GraphicFrameXform(),
      "xdr:clientData": new StaticXform({ tag: "xdr:clientData" })
    };
  }

  get tag(): string {
    return "xdr:absoluteAnchor";
  }

  prepare(model: any, options: { index: number }): void {
    if (model.picture) {
      this.map["xdr:pic"].prepare(model.picture, options);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].prepare(model.graphicFrame, options);
    }
  }

  render(xmlStream: any, model: any): void {
    xmlStream.openNode(this.tag);

    this.map["xdr:pos"].render(xmlStream, model.range.pos ?? { x: 0, y: 0 });
    this.map["xdr:ext"].render(xmlStream, model.range.ext);
    if (model.picture) {
      this.map["xdr:pic"].render(xmlStream, model.picture);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].render(xmlStream, model.graphicFrame);
    }
    this.map["xdr:clientData"].render(xmlStream, {});

    xmlStream.closeNode();
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.range.pos = this.map["xdr:pos"].model;
        this.model.range.ext = this.map["xdr:ext"].model;
        this.model.picture = this.map["xdr:pic"].model;
        this.model.graphicFrame = this.map["xdr:graphicFrame"].model;
        return false;
      default:
        return true;
    }
  }

  reconcile(model: any, options: any): void {
    if (model.picture) {
      model.medium = this.reconcilePicture(model.picture, options);
    }
  }
}

export { AbsoluteAnchorXform };
