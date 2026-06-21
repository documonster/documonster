import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { ExtLstXform } from "@excel/xlsx/xform/drawing/ext-lst-xform";
import { HLinkClickXform } from "@excel/xlsx/xform/drawing/hlink-click-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

interface CNvPrModel {
  index: number;
}

class CNvPrXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;
  declare public model: any;

  constructor() {
    super();

    this.map = {
      "a:hlinkClick": new HLinkClickXform(),
      "a:extLst": new ExtLstXform()
    };
  }

  get tag(): string {
    return "xdr:cNvPr";
  }

  render(xmlStream: XmlSink, model: CNvPrModel): void {
    xmlStream.openNode(this.tag, {
      id: model.index,
      name: `Picture ${model.index}`
    });
    this.map["a:hlinkClick"].render(xmlStream, model);
    this.map["a:extLst"].render(xmlStream, model);
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

  parseText(): void {}

  parseClose(name: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model = this.map["a:hlinkClick"].model;
        return false;
      default:
        return true;
    }
  }
}

export { CNvPrXform };
