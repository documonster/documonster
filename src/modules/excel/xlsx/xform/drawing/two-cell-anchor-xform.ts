import { BaseCellAnchorXform } from "@excel/xlsx/xform/drawing/base-cell-anchor-xform";
import { CellPositionXform } from "@excel/xlsx/xform/drawing/cell-position-xform";
import { GraphicFrameXform } from "@excel/xlsx/xform/drawing/graphic-frame-xform";
import { PicXform } from "@excel/xlsx/xform/drawing/pic-xform";
import { SpXform } from "@excel/xlsx/xform/drawing/sp-xform";
import { StaticXform } from "@excel/xlsx/xform/static-xform";

interface TwoCellModel {
  range: {
    editAs?: string;
    tl: any;
    br: any;
  };
  picture?: any;
  shape?: any;
  /** Graphic frame model (for charts and other embedded objects) */
  graphicFrame?: any;
  /** Wrap the anchor in mc:AlternateContent for modern drawing clients */
  alternateContent?: { requires: string };
}

class TwoCellAnchorXform extends BaseCellAnchorXform {
  constructor() {
    super();

    this.map = {
      "xdr:from": new CellPositionXform({ tag: "xdr:from" }),
      "xdr:to": new CellPositionXform({ tag: "xdr:to" }),
      "xdr:pic": new PicXform(),
      "xdr:sp": new SpXform(),
      "xdr:graphicFrame": new GraphicFrameXform(),
      "xdr:clientData": new StaticXform({ tag: "xdr:clientData" })
    };
  }

  get tag(): string {
    return "xdr:twoCellAnchor";
  }

  prepare(model: TwoCellModel, options: { index: number }): void {
    if (model.picture) {
      this.map["xdr:pic"].prepare(model.picture, options);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].prepare(model.graphicFrame, options);
    }
  }

  render(xmlStream: any, model: TwoCellModel): void {
    const wrapAlternateContent = !!model.alternateContent;
    if (wrapAlternateContent) {
      xmlStream.openNode("mc:AlternateContent", {
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
      });
      xmlStream.openNode("mc:Choice", {
        Requires: model.alternateContent?.requires,
        ...(model.alternateContent?.requires === "a14"
          ? {
              "xmlns:a14": "http://schemas.microsoft.com/office/drawing/2010/main"
            }
          : {}),
        ...(model.alternateContent?.requires === "cx"
          ? {
              "xmlns:cx": "http://schemas.microsoft.com/office/drawing/2014/chartex"
            }
          : {})
      });
    }

    const editAs = model.range.editAs ?? (model.graphicFrame ? undefined : "oneCell");
    xmlStream.openNode(this.tag, editAs ? { editAs } : {});

    this.map["xdr:from"].render(xmlStream, model.range.tl);
    this.map["xdr:to"].render(xmlStream, model.range.br);
    if (model.picture) {
      this.map["xdr:pic"].render(xmlStream, model.picture);
    } else if (model.graphicFrame) {
      this.map["xdr:graphicFrame"].render(xmlStream, model.graphicFrame);
    } else if (model.shape) {
      this.map["xdr:sp"].render(xmlStream, model.shape);
    }
    this.map["xdr:clientData"].render(xmlStream, {});

    xmlStream.closeNode();

    if (wrapAlternateContent) {
      xmlStream.closeNode(); // mc:Choice
      xmlStream.leafNode("mc:Fallback");
      xmlStream.closeNode(); // mc:AlternateContent
    }
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
        this.model.range.br = this.map["xdr:to"].model;
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
    // graphicFrame reconciliation handled at DrawingXform level
  }
}

export { TwoCellAnchorXform };
