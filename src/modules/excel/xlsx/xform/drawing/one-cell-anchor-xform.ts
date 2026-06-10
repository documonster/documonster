import { BaseCellAnchorXform } from "@excel/xlsx/xform/drawing/base-cell-anchor-xform";
import { CellPositionXform } from "@excel/xlsx/xform/drawing/cell-position-xform";
import { ExtXform } from "@excel/xlsx/xform/drawing/ext-xform";
import { GraphicFrameXform } from "@excel/xlsx/xform/drawing/graphic-frame-xform";
import { PicXform } from "@excel/xlsx/xform/drawing/pic-xform";
import { ShapeXform } from "@excel/xlsx/xform/drawing/shape-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";

interface OneCellModel {
  range: {
    editAs?: string;
    tl: any;
    ext: any;
  };
  picture?: any;
  shape?: any;
  /** Graphic frame model (for charts and other embedded objects) */
  graphicFrame?: any;
}

class OneCellAnchorXform extends BaseCellAnchorXform {
  constructor() {
    super();

    this.map = {
      "xdr:from": new CellPositionXform({ tag: "xdr:from" }),
      "xdr:ext": new ExtXform({ tag: "xdr:ext" }),
      "xdr:pic": new PicXform(),
      "xdr:userShape": new ShapeXform(),
      "xdr:graphicFrame": new GraphicFrameXform(),
      "xdr:clientData": new StaticXform({ tag: "xdr:clientData" })
    };
  }

  get tag(): string {
    return "xdr:oneCellAnchor";
  }

  prepare(model: OneCellModel, options: { index: number }): void {
    if (model.picture) {
      this.map["xdr:pic"].prepare(model.picture, options);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].prepare(model.graphicFrame, options);
    }
  }

  render(xmlStream: any, model: OneCellModel): void {
    xmlStream.openNode(this.tag, { editAs: model.range.editAs ?? "oneCell" });

    this.map["xdr:from"].render(xmlStream, model.range.tl);
    this.map["xdr:ext"].render(xmlStream, model.range.ext);
    if (model.picture) {
      this.map["xdr:pic"].render(xmlStream, model.picture);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].render(xmlStream, model.graphicFrame);
    } else if (model.shape?.kind === "userShape") {
      this.map["xdr:userShape"].render(xmlStream, model.shape);
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
        this.model.range.tl = this.map["xdr:from"].model;
        this.model.range.ext = this.map["xdr:ext"].model;
        this.model.picture = this.map["xdr:pic"].model;
        this.model.graphicFrame = this.map["xdr:graphicFrame"].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model: any, options: any): void {
    if (model.picture) {
      model.medium = this.reconcilePicture(model.picture, options);
    }
  }
}

export { OneCellAnchorXform };
