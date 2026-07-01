import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { VmlShapeXform } from "@excel/xlsx/xform/comment/vml-shape-xform";
import type { ShapeModel } from "@excel/xlsx/xform/comment/vml-shape-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";
import { StdDocAttributes } from "@xml/writer";

interface VmlNotesModel {
  comments: ShapeModel[];
}

// This class is (currently) single purposed to insert the triangle
// drawing icons on commented cells
class VmlNotesXform extends BaseXform<VmlNotesModel> {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;

  constructor() {
    super();
    this.map = {
      "v:shape": new VmlShapeXform()
    };
    this.model = { comments: [] };
  }

  get tag(): string {
    return "xml";
  }

  render(xmlStream: XmlSink, model?: VmlNotesModel): void {
    const renderModel = model || this.model;
    xmlStream.openXml(StdDocAttributes);
    xmlStream.openNode(this.tag, VmlNotesXform.DRAWING_ATTRIBUTES);

    xmlStream.openNode("o:shapelayout", { "v:ext": "edit" });
    xmlStream.leafNode("o:idmap", { "v:ext": "edit", data: 1 });
    xmlStream.closeNode();

    xmlStream.openNode("v:shapetype", {
      id: "_x0000_t202",
      coordsize: "21600,21600",
      "o:spt": 202,
      path: "m,l,21600r21600,l21600,xe"
    });
    xmlStream.leafNode("v:stroke", { joinstyle: "miter" });
    xmlStream.leafNode("v:path", { gradientshapeok: "t", "o:connecttype": "rect" });
    xmlStream.closeNode();

    renderModel!.comments.forEach((item, index) => {
      (this.map["v:shape"] as VmlShapeXform).render(xmlStream, item, index);
    });

    xmlStream.closeNode();
  }

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          comments: []
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.comments.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(
    model: { anchors: { br?: unknown }[] },
    options: Parameters<BaseXform["reconcile"]>[1]
  ): void {
    model.anchors.forEach(anchor => {
      if (anchor.br) {
        this.map["xdr:twoCellAnchor"].reconcile(anchor, options);
      } else {
        this.map["xdr:oneCellAnchor"].reconcile(anchor, options);
      }
    });
  }

  static DRAWING_ATTRIBUTES = {
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:x": "urn:schemas-microsoft-com:office:excel"
  };
}

export { VmlNotesXform };
